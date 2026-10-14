"""Persistence for Jira connection settings and metadata mappings.

Two modes, chosen automatically by an environment variable:

- Local mode (JIRA_APP_PERSIST_LOCAL=1, set by run.bat/run.sh): persists to
  JSON files in ./data on disk, so a single local user's settings survive
  restarts. This is only appropriate for a single person running their own
  copy of the app.

- Session mode (default — used whenever that env var isn't set, e.g. on a
  shared host like Streamlit Community Cloud): persists only in Streamlit's
  per-browser-session state, never touching shared disk. Each visitor gets
  their own isolated Jira credentials and mapping that nobody else can see
  or overwrite, and nothing survives closing the tab. This is the safe
  default for any deployment more than one person can reach.
"""
import json
import os

import streamlit as st

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
CONNECTION_FILE = os.path.join(DATA_DIR, "connection.json")
MAPPING_FILE = os.path.join(DATA_DIR, "mapping.json")

# Evaluated once at import time (module import happens once per server
# process, not once per Streamlit rerun) so this is stable for the process's
# lifetime.
PERSIST_LOCAL = os.environ.get("JIRA_APP_PERSIST_LOCAL") == "1"

DEFAULT_JIRA_URL = "https://directfn-jira.atlassian.net"

DEFAULT_MAPPING = {
    "projects": {},      # excel project value -> {"key": ..., "name": ...}
    "components": {},    # jira project key -> {excel component value -> {"id": ..., "name": ...}}
    "issue_types": {},   # jira project key -> {"id": ..., "name": ...}
    "assignees": {},     # excel assignee value -> {"accountId": ..., "displayName": ...}
    "link_type": {},  # jira project key -> {"mode": "parent"} or {"mode": "link", "name": ..., "task_side": "outward"|"inward"}
    # jira project key -> {jira field id -> mapping}. A mapping's "source" is
    # either "column" (value comes from an Excel column each row) or "fixed"
    # (same value on every task — for required fields with no Excel column,
    # e.g. a mandatory custom field with no corresponding sheet data).
    "extra_fields": {},
    "title_templates": {},  # jira project key -> template string, e.g. "[NTP-SAAS][{Component}] {Description}"
    "labels": {},  # jira project key -> [label, ...] attached to every task created for that project
}


def _ensure_dir():
    os.makedirs(DATA_DIR, exist_ok=True)


def load_connection():
    if PERSIST_LOCAL:
        if not os.path.exists(CONNECTION_FILE):
            return {"base_url": DEFAULT_JIRA_URL, "email": "", "api_token": ""}
        with open(CONNECTION_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        data.setdefault("base_url", DEFAULT_JIRA_URL)
        return data
    # Session mode: this return value is only ever used the first time a
    # given browser session initializes (see app.py's ss.setdefault), so a
    # fresh default here is correct — there is nothing shared to load.
    return {"base_url": DEFAULT_JIRA_URL, "email": "", "api_token": ""}


def save_connection(base_url, email, api_token):
    if not PERSIST_LOCAL:
        # The caller already writes the same data into st.session_state
        # (ss.connection) right after calling this — that per-session state
        # IS the persistence in this mode, so there's nothing more to do.
        return
    _ensure_dir()
    with open(CONNECTION_FILE, "w", encoding="utf-8") as f:
        json.dump(
            {"base_url": base_url.strip().rstrip("/"), "email": email.strip(), "api_token": api_token.strip()},
            f,
            indent=2,
        )


def _migrate_extra_fields(extra_fields):
    """Old shape keyed each project's dict by the Excel column name, always
    implying source="column". New shape keys by Jira field id instead (so a
    mapping can also be a fixed value with no source column) — convert any
    old-shaped entries in place rather than losing already-configured mappings."""
    migrated = {}
    for pkey, entries in (extra_fields or {}).items():
        migrated[pkey] = {}
        for key, cfg in entries.items():
            if "source" in cfg:
                migrated[pkey][key] = cfg
            else:
                migrated[pkey][cfg["field_id"]] = {
                    "field_name": cfg["field_name"],
                    "source": "column",
                    "column_key": key,
                    "column_label": cfg.get("column_label", key),
                }
    return migrated


def _migrate_link_type(link_type, projects):
    """Old shape was a single global config (or None) applied to every
    project. New shape is keyed per Jira project key. A dict missing a
    "mode"/"name" key is already new-shaped (project keys never collide with
    those names); otherwise carry the old single setting forward onto every
    project already mapped, so existing behavior doesn't silently vanish."""
    if not link_type:
        return {}
    if "mode" in link_type or "name" in link_type:
        project_keys = {v["key"] for v in (projects or {}).values()}
        return {pkey: dict(link_type) for pkey in project_keys}
    return link_type


def load_mapping():
    if PERSIST_LOCAL:
        if not os.path.exists(MAPPING_FILE):
            return json.loads(json.dumps(DEFAULT_MAPPING))
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        # keep only known keys, so a schema change (e.g. renamed key) doesn't
        # leave stale data lying around in the saved file forever
        mapping = {key: data.get(key, default) for key, default in DEFAULT_MAPPING.items()}
        mapping["extra_fields"] = _migrate_extra_fields(mapping["extra_fields"])
        mapping["link_type"] = _migrate_link_type(mapping["link_type"], mapping["projects"])
        return mapping
    # Session mode: same reasoning as load_connection() above — only used on
    # a session's first init, so a fresh default is correct.
    return json.loads(json.dumps(DEFAULT_MAPPING))


def save_mapping(mapping):
    if not PERSIST_LOCAL:
        # ss.mapping's nested dicts are mutated in place by the caller before
        # this is called, so st.session_state already holds the update —
        # nothing more to persist in session mode.
        return
    _ensure_dir()
    with open(MAPPING_FILE, "w", encoding="utf-8") as f:
        json.dump(mapping, f, indent=2)
