"""Excel -> Jira Task Importer

A small Streamlit app that:
 1. Connects to a Jira Cloud site.
 2. Loads a sprint-backlog style Excel sheet.
 3. Lets the user map the sheet's free-text Project / Component / Assignee
    values onto real Jira projects / components / users (saved once, reused
    on every future import).
 4. Reviews the resolved rows and bulk-creates them as Jira issues.
"""
import hashlib
import io
import json
import os
import re
import time

import openpyxl
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

import excel_parser as xp
import storage
from jira_client import JiraClient, JiraError

BLANK = xp.BLANK
NOT_MAPPED = "— not mapped —"
NO_COMPONENT = "— no component —"
UNASSIGNED = "— unassigned —"
NO_LINK = "— none —"
NO_FIELD = "— not available in this project —"

STORY_KEY_RE = re.compile(r"^[A-Za-z][A-Za-z0-9]*-\d+$")

# Standard/system fields we never offer as an "extra" custom field mapping —
# only genuine (usually custom) fields should show up as candidates.
STANDARD_FIELD_IDS = {
    "summary", "issuetype", "project", "reporter", "assignee", "priority",
    "components", "fixVersions", "versions", "labels", "description",
    "attachment", "security", "timetracking", "duedate", "environment", "parent",
}

# Excel columns already handled by a dedicated mapping section / creation step —
# excluded from the generic "map any other column" picker.
ALREADY_CONFIGURED_COLUMNS = {
    "project_raw", "component_raw", "assignee_raw", "jira_story",
    "description", "jira_id", "estimated_hours", "row_number",
}

st.set_page_config(page_title="Jira Task Creator", page_icon="\U0001F4CB", layout="wide")

# Friendly display labels for underscored dataframe column names — falls back
# to a title-cased, space-joined version of the column name if not listed.
COLUMN_LABELS = {
    "row_number": "Row #",
    "section": "Section",
    "project_raw": "Project (Excel)",
    "component_raw": "Component (Excel)",
    "assignee_raw": "Assignee (Excel)",
    "description": "Description",
    "jira_story": "Jira Story",
    "capitalization_raw": "Capitalization (Excel)",
    "status": "Status",
    "estimated_hours": "Est. Hours",
    "jira_id": "Existing JIRA ID",
    "jira_project": "Jira Project",
    "jira_title": "Task Title",
    "jira_component": "Jira Component",
    "jira_assignee": "Jira Assignee",
    "jira_issue_type": "Issue Type",
    "jira_extra_fields": "Extra Fields",
    "story_link": "Linked Story",
    "ready": "Ready",
    "issue_key": "Issue Key",
    "url": "URL",
    "note": "Note",
}


def column_labels(columns):
    """st.dataframe column_config mapping that shows friendly labels instead
    of raw underscored column names, without renaming the underlying data."""
    return {
        col: st.column_config.Column(label=COLUMN_LABELS.get(col, col.replace("_", " ").title()))
        for col in columns
    }


# ----------------------------------------------------------------------------
# session state bootstrap
# ----------------------------------------------------------------------------
def init_state():
    ss = st.session_state
    ss.setdefault("connection", storage.load_connection())
    ss.setdefault("mapping", storage.load_mapping())
    ss.setdefault("client", None)
    ss.setdefault("me", None)
    ss.setdefault("projects", [])
    ss.setdefault("components_cache", {})
    ss.setdefault("issue_types_cache", {})
    ss.setdefault("users_cache", {})
    ss.setdefault("link_types", None)
    ss.setdefault("all_labels", None)
    ss.setdefault("create_fields_cache", {})
    ss.setdefault("df", pd.DataFrame())
    ss.setdefault("sheet_name", None)
    ss.setdefault("excel_bytes", None)
    ss.setdefault("excel_filename", None)
    ss.setdefault("col_map", {})
    ss.setdefault("header_row", None)
    ss.setdefault("create_results", None)
    ss.setdefault("_last_imported_cfg_sig", None)
    ss.setdefault("dark_mode", False)
    ss.setdefault("fixed_field_overrides", None)
    ss.setdefault("selection_reset_mode", "default")
    ss.setdefault("selection_reset_nonce", 0)
    ss.setdefault("confirm_stop", False)

    conn = ss.connection
    if ss.client is None and conn.get("base_url") and conn.get("email") and conn.get("api_token"):
        ss.client = JiraClient(conn["base_url"], conn["email"], conn["api_token"])


init_state()
ss = st.session_state


def get_projects(force=False):
    if force or not ss.projects:
        ss.projects = ss.client.get_projects()
    return ss.projects


def get_components(project_key, force=False):
    if force or project_key not in ss.components_cache:
        ss.components_cache[project_key] = ss.client.get_components(project_key)
    return ss.components_cache[project_key]


def get_link_types(force=False):
    if force or ss.link_types is None:
        ss.link_types = ss.client.get_issue_link_types()
    return ss.link_types


def get_all_labels(force=False):
    if force or ss.all_labels is None:
        ss.all_labels = ss.client.get_labels()
    return ss.all_labels


def get_issue_types(project_key, force=False):
    if force or project_key not in ss.issue_types_cache:
        ss.issue_types_cache[project_key] = ss.client.get_issue_types(project_key)
    return ss.issue_types_cache[project_key]


def get_create_fields(project_key, force=False):
    """Create-screen fields for a project, keyed by (project, issue type) since
    that's what Jira's createmeta API is scoped to. Needs an issue type to
    already be mapped for the project (done in step 3b)."""
    it = ss.mapping["issue_types"].get(project_key)
    if not it:
        return []
    cache_key = (project_key, it["id"])
    if force or cache_key not in ss.create_fields_cache:
        ss.create_fields_cache[cache_key] = ss.client.get_create_fields(project_key, it["id"])
    return ss.create_fields_cache[cache_key]


def custom_field_candidates(project_key):
    """Create-screen fields eligible for a generic 'map this Excel column to
    this Jira field' mapping — anything that isn't a standard/system field."""
    return [f for f in get_create_fields(project_key) if f.get("fieldId") not in STANDARD_FIELD_IDS]


def match_allowed_value(raw_value, allowed_values):
    """Match an Excel value like '75' against Jira dropdown options like
    '75%' by comparing the numeric part only."""
    raw_norm = re.sub(r"[^0-9.]", "", str(raw_value))
    if not raw_norm:
        return None
    for av in allowed_values or []:
        av_norm = re.sub(r"[^0-9.]", "", str(av.get("value", "")))
        if av_norm and av_norm == raw_norm:
            return av
    return None


def build_custom_field_value(field_def, raw_value):
    """Returns (value_to_set, error_note). Matches against the field's
    dropdown options if it has any, otherwise sets a plain text/number value."""
    if field_def.get("allowedValues"):
        matched = match_allowed_value(raw_value, field_def["allowedValues"])
        if matched:
            return {"id": matched["id"]}, None
        return None, f"no '{field_def['name']}' option matches value '{raw_value}'"
    schema_type = field_def.get("schema", {}).get("type")
    if schema_type == "number":
        try:
            return float(raw_value), None
        except (TypeError, ValueError):
            return None, f"'{raw_value}' is not a valid number for '{field_def['name']}'"
    return str(raw_value), None


TITLE_TEMPLATE_EXCLUDED_COLUMNS = {"row_number", "jira_id"}


def placeholder_name(col):
    """Human-friendly placeholder token for a dataframe column, e.g.
    'component_raw' -> 'Component' (used in title templates as {Component})."""
    label = COLUMN_LABELS.get(col, col.replace("_", " ").title())
    return label.replace(" (Excel)", "")


def title_template_columns(df):
    return [c for c in df.columns if c not in TITLE_TEMPLATE_EXCLUDED_COLUMNS]


def render_title_template(template, row, columns):
    """Substitutes {PlaceholderName} tokens in a title template with a row's
    values; missing/blank values become an empty string."""
    result = template
    for col in columns:
        val = clean(getattr(row, col, None))
        result = result.replace("{" + placeholder_name(col) + "}", "" if val is None else str(val))
    return result


def extra_fields_summary(row, pkey, extra_field_map, override_row=None):
    """Human-readable '<Field>=<value>' summary of every extra field mapped
    for pkey, applying override_row's values where present. Shared by the
    review table and the resolve() pass so both always agree with what the
    creation loop will actually send to Jira."""
    if not pkey:
        return None
    parts = []
    for field_id, cfg in extra_field_map.get(pkey, {}).items():
        if cfg["source"] == "fixed":
            override_val = None
            if override_row is not None:
                colname = f"fixed__{field_id}"
                if colname in override_row.index:
                    ov = clean(override_row[colname])
                    override_val = ov if ov not in (None, "") else None
            if override_val is not None:
                default_val = cfg["value"]["label"] if isinstance(cfg["value"], dict) else cfg["value"]
                default_val = default_val.strip() if isinstance(default_val, str) else default_val
                if override_val != default_val:
                    parts.append(f"{cfg['field_name']}={override_val} (overridden)")
                else:
                    parts.append(f"{cfg['field_name']}={override_val}")
            else:
                val_display = cfg["value"]["label"] if isinstance(cfg["value"], dict) else cfg["value"]
                val_display = val_display.strip() if isinstance(val_display, str) else val_display
                parts.append(f"{cfg['field_name']}={val_display}")
        else:
            val = clean(getattr(row, cfg["column_key"], None))
            if val is not None:
                parts.append(f"{cfg['field_name']}={val}")
    return "; ".join(parts) if parts else None


def get_users(project_key, force=False):
    if force or project_key not in ss.users_cache:
        ss.users_cache[project_key] = ss.client.get_assignable_users(project_key)
    return ss.users_cache[project_key]


def build_updated_workbook(original_bytes, sheet_name, header_row, col_map, results_df):
    """Writes the newly created Jira issue keys back into the JIRA ID column
    of the original sheet (adding that column if it didn't exist) and
    returns the resulting .xlsx file as bytes."""
    wb = openpyxl.load_workbook(io.BytesIO(original_bytes))
    ws = wb[sheet_name]
    jira_col = col_map.get("jira_id")
    if jira_col is None:
        jira_col = ws.max_column + 1
        ws.cell(row=header_row, column=jira_col, value="JIRA ID")
    created = results_df[results_df["issue_key"] != ""]
    for _, r in created.iterrows():
        ws.cell(row=int(r["row_number"]), column=jira_col, value=r["issue_key"])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def clean(v):
    """Normalize a pandas scalar to None if it's missing (NaN/NaT/pd.NA) —
    pandas represents missing values as NaN even in string columns, and NaN
    is truthy in plain Python, so this must be used instead of `v or ...`."""
    return None if pd.isna(v) else v


def select_value(label, options, current_value, widget_key, help_text=None):
    """options: list of (value, display_label) with value=None as the 'not mapped' placeholder."""
    labels = [o[1] for o in options]
    values = [o[0] for o in options]
    try:
        idx = values.index(current_value)
    except ValueError:
        idx = 0
    chosen_label = st.selectbox(label, labels, index=idx, key=widget_key, help=help_text)
    return values[labels.index(chosen_label)]


LIGHT_PALETTE = {
    "bg": "#FFFFFF", "bg_alt": "#F4F5F7", "text": "#172B4D", "text_muted": "#6B778C",
    "border": "rgba(9,30,66,0.13)", "primary": "#0052CC", "input_bg": "#FFFFFF",
}
DARK_PALETTE = {
    "bg": "#14181F", "bg_alt": "#1E232D", "text": "#E6E8EB", "text_muted": "#9AA4B2",
    "border": "rgba(255,255,255,0.14)", "primary": "#4C9AFF", "input_bg": "#242A35",
}


def build_custom_css(dark):
    p = DARK_PALETTE if dark else LIGHT_PALETTE
    elevation_1 = "0 1px 3px rgba(0,0,0,0.20), 0 1px 2px rgba(0,0,0,0.14)" if dark else "0 1px 3px rgba(0,0,0,0.10), 0 1px 2px rgba(0,0,0,0.08)"
    elevation_2 = "0 3px 12px rgba(0,0,0,0.32), 0 2px 4px rgba(0,0,0,0.20)" if dark else "0 3px 10px rgba(0,0,0,0.14), 0 2px 4px rgba(0,0,0,0.08)"

    # Only in dark mode do we forcibly repaint every button — in light mode
    # Streamlit's own native styling (blue primary / light secondary) already
    # works correctly and this would just flatten that distinction.
    dark_button_override = (
        f"""
[data-testid="stButton"] button, [data-testid="stButton"] button *,
[data-testid="stFormSubmitButton"] button, [data-testid="stFormSubmitButton"] button *,
[data-testid="stDownloadButton"] button, [data-testid="stDownloadButton"] button *,
[data-testid="stLinkButton"] a, [data-testid="stLinkButton"] a * {{
    background-color: {p['bg_alt']} !important;
    color: {p['text']} !important;
    border-color: {p['border']} !important;
}}
"""
        if dark
        else ""
    )

    return f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&family=Roboto+Mono&display=swap');

html, body, .stApp {{
    font-family: 'Roboto', 'Segoe UI', sans-serif;
    background-color: {p['bg']} !important;
    color: {p['text']} !important;
}}
code, pre, .stCodeBlock, div[data-testid="stMetricValue"] {{
    font-family: 'Roboto Mono', monospace;
}}

/* Text-color forcing is scoped to containers we've confirmed are plain
   content (markdown/headings/labels) rather than applied blanket to every
   p/span/li in the DOM — those tags are reused *inside* buttons, inputs
   etc. with their own background, and forcing their text color without
   also controlling that background is what caused "invisible until
   hover" controls.

   st.caption renders under data-testid="stCaptionContainer", a SEPARATE
   container from st.markdown/st.header's "stMarkdownContainer" (confirmed
   by reading Streamlit's own frontend source) — missing it here left every
   st.caption call on the page (used throughout for hint/help text) with
   Streamlit's light-theme muted-gray color on a dark background, i.e.
   exactly the "text not clear" symptom. Both are covered below. */
h1, h2, h3, h4, h5, h6 {{ color: {p['text']}; }}
div[data-testid="stMarkdownContainer"], div[data-testid="stMarkdownContainer"] *,
div[data-testid="stCaptionContainer"], div[data-testid="stCaptionContainer"] *,
div[data-testid="stWidgetLabel"], div[data-testid="stWidgetLabel"] *,
div[data-testid="stHeading"] {{
    color: {p['text']} !important;
}}
/* keep links visually distinct from body text instead of blending in
   uniformly gray/white from the blanket rule above */
div[data-testid="stMarkdownContainer"] a, div[data-testid="stCaptionContainer"] a {{
    color: {p['primary']} !important;
}}
label {{ color: {p['text']} !important; }}

footer {{visibility: hidden;}}

/* ---- Banner (fixed brand gradient, readable in both modes) -------------- */
.app-banner {{
    background: linear-gradient(135deg, #0052CC 0%, #2684FF 100%);
    padding: 20px 34px;
    border-radius: 16px;
    margin-bottom: 1.4rem;
    box-shadow: {elevation_2};
}}
.app-banner h1 {{
    color: #FFFFFF !important;
    margin: 0;
    font-size: 1.85rem;
    font-weight: 700;
    letter-spacing: -0.01em;
}}

/* ---- Page navigation (segmented control, styled as a pill tab bar) -------- */
div[data-testid="stButtonGroup"] {{
    background: {p['bg_alt']};
    padding: 5px;
    border-radius: 12px;
    gap: 4px;
    margin-bottom: 0.8rem;
    box-shadow: {elevation_1};
}}
div[data-testid="stButtonGroup"] button {{
    border: none !important;
    border-radius: 8px !important;
    font-weight: 500;
    color: {p['text_muted']} !important;
    transition: background-color 0.15s ease, color 0.15s ease;
}}
div[data-testid="stButtonGroup"] button * {{
    color: inherit !important;
}}
div[data-testid="stButtonGroup"] [data-selected] button,
div[data-testid="stButtonGroup"] [data-selected] {{
    background: {p['primary']} !important;
    color: #FFFFFF !important;
}}
div[data-testid="stButtonGroup"] [data-selected] * {{
    color: #FFFFFF !important;
    font-weight: 700;
}}

/* ---- Buttons (Material elevation + lift on hover) ------------------------
   Descendant (not direct-child) selectors: Streamlit wraps the real <button>
   one level deeper inside a tooltip/help wrapper div, confirmed by reading
   Streamlit's own frontend source — a ">" direct-child selector here never
   matched anything, so none of this elevation/hover styling was ever
   actually applied to a real button. */
.stButton button, .stDownloadButton button, .stLinkButton a, .stFormSubmitButton button {{
    border-radius: 8px;
    font-weight: 500;
    box-shadow: {elevation_1};
    transition: box-shadow 0.15s ease, transform 0.15s ease;
}}
.stButton button:hover, .stDownloadButton button:hover,
.stLinkButton a:hover, .stFormSubmitButton button:hover {{
    box-shadow: {elevation_2};
    transform: translateY(-1px);
}}
{dark_button_override}

/* ---- Inputs, selects, uploaders (native + BaseWeb) ------------------------ */
input, textarea,
div[data-baseweb="select"] > div, div[data-baseweb="base-input"],
section[data-testid="stFileUploaderDropzone"] {{
    background-color: {p['input_bg']} !important;
    color: {p['text']} !important;
    border-color: {p['border']} !important;
}}
div[data-baseweb="select"] > div *, section[data-testid="stFileUploaderDropzone"] * {{
    color: {p['text']} !important;
}}
div[data-baseweb="popover"] li, ul[role="listbox"], div[data-baseweb="menu"] {{
    background-color: {p['input_bg']} !important;
    color: {p['text']} !important;
}}

/* ---- Alert boxes (st.info/warning/success/error) -------------------------
   Left un-recolored on purpose: they use their own theme-driven pastel
   background + matching dark text (e.g. light blue bg + dark blue text for
   st.info), which stays internally legible regardless of the page's dark
   background — only rounding them for visual consistency with the rest of
   the page. */
div[data-testid="stAlert"] {{
    border-radius: 10px;
    box-shadow: {elevation_1};
}}

/* ---- Metrics as small elevated cards ------------------------------------ */
div[data-testid="stMetric"] {{
    background: {p['bg_alt']};
    border-radius: 12px;
    padding: 12px 16px;
    box-shadow: {elevation_1};
}}
div[data-testid="stMetricValue"], div[data-testid="stMetricLabel"] {{
    color: {p['text']} !important;
}}

/* ---- Expanders (used inside the component-mapping form) ----------------- */
div[data-testid="stExpander"] {{
    border-radius: 12px !important;
    box-shadow: {elevation_1};
    background: {p['bg_alt']};
    color: {p['text']};
}}
div[data-testid="stExpander"] summary, div[data-testid="stExpander"] summary * {{
    color: {p['text']} !important;
}}

/* ---- Dataframes: rounded corners (grid contents are canvas-rendered and
   may keep a light background — a Streamlit rendering limitation) ------- */
div[data-testid="stDataFrame"] {{
    border-radius: 10px;
    overflow: hidden;
    box-shadow: {elevation_1};
}}

hr {{ border-color: {p['border']} !important; }}

/* ---- Sidebar -------------------------------------------------------------- */
section[data-testid="stSidebar"] {{
    background-color: {p['bg_alt']} !important;
    border-right: 1px solid {p['border']};
}}
.sidebar-card {{
    background: {p['bg']};
    border-radius: 12px;
    padding: 14px 16px;
    margin-bottom: 12px;
    box-shadow: {elevation_1};
}}
.sidebar-card h4 {{
    margin: 0 0 8px 0;
    font-size: 0.75rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: {p['text_muted']};
    font-weight: 700;
}}
.sidebar-card .url {{
    margin-top: 4px;
    font-size: 0.78rem;
    color: {p['text_muted']};
    word-break: break-all;
}}
.status-dot {{
    display: inline-block;
    width: 8px;
    height: 8px;
    border-radius: 50%;
    margin-right: 7px;
}}
.status-dot.green {{ background: #36B37E; }}
.status-dot.amber {{ background: #FFAB00; }}
.status-dot.gray {{ background: #97A0AF; }}
</style>
"""


dark_mode = ss.get("dark_mode", False)
st.markdown(build_custom_css(dark_mode), unsafe_allow_html=True)

st.markdown(
    """
    <div class="app-banner">
        <h1>\U0001F4CB Jira Task Creator</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### \U0001F9ED Workflow Status")
    # Reserved now, filled in at the very end of the script (see
    # render_sidebar_status below) — buttons that change connection/sheet
    # state live further down in the tabs, and this placeholder lets the
    # status reflect what THIS run just did instead of lagging one rerun
    # behind (sidebar code runs before the tabs' button handlers do).
    sidebar_status_slot = st.empty()

    if ss.connection.get("base_url"):
        st.link_button("\U0001F517 Open Jira site", ss.connection["base_url"], use_container_width=True)

    st.divider()
    st.toggle("\U0001F319 Dark mode", value=dark_mode, key="dark_mode")
    if storage.PERSIST_LOCAL:
        st.caption(
            "Config is stored locally in `./data`. Keep this folder private — it contains your Jira API token in plain text."
        )
    else:
        st.caption(
            "Your Jira connection and mapping are kept only for this browser session — nobody else who opens "
            "this app can see them, and they're forgotten once you close this tab."
        )
        if ss.client or ss.connection.get("api_token"):
            if st.button("\U0001F512 Forget my Jira credentials", use_container_width=True):
                ss.connection = {"base_url": storage.DEFAULT_JIRA_URL, "email": "", "api_token": ""}
                ss.client = None
                ss.me = None
                # everything below was fetched using that token — drop it too
                # rather than leave stale Jira data attached to a cleared login
                ss.projects = []
                ss.components_cache = {}
                ss.issue_types_cache = {}
                ss.users_cache = {}
                ss.link_types = None
                ss.all_labels = None
                ss.create_fields_cache = {}
                st.success("Credentials forgotten for this session. Reconnect in the Connect tab to keep going.")
                st.rerun()
            st.caption(
                "Don't want to wait for this tab to time out on its own? Click above to wipe your token from "
                "this session's memory right now — your uploaded sheet and mapping are kept."
            )

    # The Stop button kills the whole server process — fine for a local,
    # single-user run (run.bat sets this env var), but on a shared host it
    # would take the app down for every other visitor at once. Hidden unless
    # explicitly enabled, so it's safe by default on any shared deployment.
    if os.environ.get("JIRA_APP_ALLOW_STOP") == "1":
        st.divider()
        if not ss.confirm_stop:
            if st.button("\U0001F6D1 Stop App", use_container_width=True):
                ss.confirm_stop = True
                st.rerun()
        else:
            st.warning("Stop the app? This closes the local server — you'll need to run `run.bat` again to restart it.")
            stop_col1, stop_col2 = st.columns(2)
            if stop_col1.button("Yes, stop", type="primary", use_container_width=True):
                st.info("Stopping… this tab will try to close itself, but browsers only allow that for tabs a "
                         "script opened — if it's still open in a second, just close it yourself.")
                # Best-effort only: browsers block window.close() on any tab the
                # user opened normally (typed a URL, clicked a bookmark, etc.) —
                # there is no web API that can force-close an arbitrary tab.
                components.html("<script>window.close();</script>", height=0, width=0)
                os._exit(0)
            if stop_col2.button("Cancel", use_container_width=True):
                ss.confirm_stop = False
                st.rerun()


def render_sidebar_status():
    if ss.client and ss.me:
        conn_dot, conn_text = "green", f"Connected as <b>{ss.me.get('displayName', '?')}</b>"
    elif ss.client:
        conn_dot, conn_text = "amber", "Credentials saved &mdash; not verified"
    else:
        conn_dot, conn_text = "gray", "Not connected"

    if not ss.df.empty:
        sheet_dot, sheet_text = "green", f"<b>{ss.sheet_name}</b> &middot; {len(ss.df)} rows"
    else:
        sheet_dot, sheet_text = "gray", "No sheet loaded yet"

    sidebar_status_slot.markdown(
        f"""
        <div class="sidebar-card">
            <h4>Jira Connection</h4>
            <div><span class="status-dot {conn_dot}"></span>{conn_text}</div>
            <div class="url">{ss.connection.get('base_url', '')}</div>
        </div>
        <div class="sidebar-card">
            <h4>Sprint Sheet</h4>
            <div><span class="status-dot {sheet_dot}"></span>{sheet_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


PAGES = [
    "1  Connect to Jira", "2  Upload Sheet", "3  Configure Mapping", "4  Review & Create", "5  Update Labels",
]
# A real page switch (not st.tabs) — st.tabs renders every tab's body on
# every single rerun regardless of which one is visible, which meant e.g.
# clicking "Parse sheet" also silently ran all of Configure Mapping's Jira
# API calls in that same run. Only the selected page's code below executes.
nav = st.segmented_control(
    "Navigate", PAGES, default=PAGES[0], required=True, key="active_page", label_visibility="collapsed"
)

# ============================================================================
# TAB 1 - CONNECT
# ============================================================================
if nav == PAGES[0]:
    st.subheader("\U0001F510 Connect to your Jira site")
    with st.container(border=True):
        st.caption(
            "Create an API token at "
            "[id.atlassian.com/manage-profile/security/api-tokens](https://id.atlassian.com/manage-profile/security/api-tokens)."
        )
        has_saved_token = bool(ss.connection.get("api_token"))
        with st.form("connection_form"):
            base_url = st.text_input(
                "Jira site URL",
                value=ss.connection.get("base_url", "") or storage.DEFAULT_JIRA_URL,
                placeholder="https://yourcompany.atlassian.net",
            )
            email = st.text_input("Account email", value=ss.connection.get("email", ""), placeholder="you@company.com")
            api_token_input = st.text_input(
                "API token",
                value="",
                type="password",
                placeholder="•••••••• (saved — leave blank to keep it)" if has_saved_token else "Paste your API token",
            )
            st.caption(
                "\U0001F512 A token is currently saved for this site — leave the field blank to keep it, "
                "or type a new one to replace it."
                if has_saved_token
                else "\U0001F513 No token saved yet."
            )
            col1, col2 = st.columns([1, 1])
            save_clicked = col1.form_submit_button("\U0001F4BE Save & Test Connection", type="primary", use_container_width=True)
            clear_clicked = col2.form_submit_button("Clear", use_container_width=True)

        if clear_clicked:
            storage.save_connection(storage.DEFAULT_JIRA_URL, "", "")
            ss.connection = {"base_url": storage.DEFAULT_JIRA_URL, "email": "", "api_token": ""}
            ss.client = None
            ss.me = None
            st.rerun()

        if save_clicked:
            api_token = api_token_input.strip() or ss.connection.get("api_token", "")
            if not (base_url and email and api_token):
                st.error("Site URL, email and API token are all required.")
            else:
                storage.save_connection(base_url, email, api_token)
                ss.connection = {
                    "base_url": base_url.strip().rstrip("/"),
                    "email": email.strip(),
                    "api_token": api_token,
                }
                ss.client = JiraClient(ss.connection["base_url"], ss.connection["email"], ss.connection["api_token"])
                try:
                    with st.spinner("Connecting to Jira…"):
                        ss.me = ss.client.test_connection()
                    ss.projects = []
                    st.success(f"Connected as {ss.me.get('displayName')} ({ss.me.get('emailAddress')})")
                except JiraError as e:
                    ss.me = None
                    st.error(f"Could not connect: {e}")

        if ss.client and not ss.me:
            if st.button("Test existing connection"):
                try:
                    with st.spinner("Connecting to Jira…"):
                        ss.me = ss.client.test_connection()
                    st.success(f"Connected as {ss.me.get('displayName')}")
                except JiraError as e:
                    st.error(f"Could not connect: {e}")

# ============================================================================
# TAB 2 - UPLOAD SHEET
# ============================================================================
elif nav == PAGES[1]:
    st.subheader("\U0001F4C1 Upload the sprint backlog workbook")
    with st.container(border=True):
        uploaded = st.file_uploader("Excel file (.xlsx)", type=["xlsx"])

        if uploaded is not None:
            ss.excel_bytes = uploaded.getvalue()
            ss.excel_filename = uploaded.name

        if ss.excel_bytes:
            sheet_names = xp.get_sheet_names(io.BytesIO(ss.excel_bytes))
            default_idx = 0
            for i, name in enumerate(sheet_names):
                if "sprint backlog" in name.lower():
                    default_idx = i
                    break
            col_sheet, col_btn = st.columns([3, 1])
            chosen_sheet = col_sheet.selectbox("Sheet to import", sheet_names, index=default_idx)
            col_btn.write("")
            col_btn.write("")
            if col_btn.button("\U0001F50D Parse sheet", type="primary", use_container_width=True):
                with st.spinner(f"Parsing '{chosen_sheet}'…"):
                    df, col_map, header_row = xp.parse_sheet(io.BytesIO(ss.excel_bytes), chosen_sheet)
                if df.empty:
                    st.error("No task rows found — could not detect a 'Description' column on this sheet.")
                else:
                    ss.df = df
                    ss.sheet_name = chosen_sheet
                    ss.col_map = col_map
                    ss.header_row = header_row
                    ss.create_results = None
                    st.success(f"Parsed {len(df)} task rows from '{chosen_sheet}' (header row {header_row}).")

    if not ss.df.empty:
        st.markdown(f"#### \U0001F4CA Preview — {ss.sheet_name}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Rows", len(ss.df))
        c2.metric("Distinct projects", len(xp.distinct_values(ss.df, "project_raw")))
        c3.metric("Distinct components", len(xp.distinct_values(ss.df, "component_raw")))
        c4.metric("Distinct assignees", len(xp.distinct_values(ss.df, "assignee_raw")))
        preview_cols = [
            "row_number", "section", "project_raw", "component_raw", "assignee_raw", "description",
            "jira_story", "capitalization_raw", "status", "estimated_hours",
        ]
        st.dataframe(
            ss.df[preview_cols],
            use_container_width=True,
            height=350,
            column_config=column_labels(preview_cols),
        )

# ============================================================================
# TAB 3 - METADATA MAPPING
# ============================================================================
elif nav == PAGES[2]:
    st.subheader("Configure metadata mapping")

    with st.container(border=True):
        st.markdown("#### \U0001F4E6 Backup & Restore Mapping Configuration")
        st.caption(
            "Export your Project / Component / Assignee / Capitalization / Link-type mapping to a file you "
            "can back up or hand to a teammate, or import one exported from this app. Jira credentials are "
            "never included."
        )
        ex_col, im_col = st.columns(2, gap="medium")
        with ex_col:
            st.markdown("**\U0001F4E4 Export**")
            st.download_button(
                "Export mapping config",
                json.dumps(ss.mapping, indent=2).encode("utf-8"),
                file_name="jira_mapping_config.json",
                mime="application/json",
                use_container_width=True,
            )
        with im_col:
            st.markdown("**\U0001F4E5 Import**")
            uploaded_cfg = st.file_uploader(
                "Import mapping config (.json)",
                type=["json"],
                key="mapping_config_uploader",
                label_visibility="collapsed",
            )
            if uploaded_cfg is not None:
                sig = (uploaded_cfg.name, uploaded_cfg.size)
                if ss._last_imported_cfg_sig != sig:
                    try:
                        incoming = json.load(uploaded_cfg)
                        if not isinstance(incoming, dict):
                            raise ValueError("File does not contain a mapping object.")
                        for key in storage.DEFAULT_MAPPING:
                            if key in incoming:
                                ss.mapping[key] = incoming[key]
                        storage.save_mapping(ss.mapping)
                        ss._last_imported_cfg_sig = sig
                        st.success("Mapping configuration imported and saved.")
                    except (json.JSONDecodeError, ValueError, UnicodeDecodeError) as e:
                        st.error(f"Could not read that file: {e}")

    if not ss.client:
        st.warning("Connect to Jira first (tab 1).")
    elif ss.df.empty:
        st.warning("Upload and parse a sheet first (tab 2).")
    else:
        try:
            with st.spinner("Loading Jira projects…"):
                projects = get_projects()
        except JiraError as e:
            projects = []
            st.error(f"Could not load Jira projects: {e}")

        project_options = [(None, NOT_MAPPED)] + [(p["key"], f"{p['key']} — {p['name']}") for p in projects]

        st.button("\U0001F504 Refresh Jira projects", on_click=lambda: get_projects(force=True))

        # ---- 3a. Project mapping -----------------------------------------
        with st.container(border=True):
            st.markdown("#### \U0001F5C2️ Map Excel **Project** values → Jira Project")
            project_values = xp.distinct_values(ss.df, "project_raw")
            proj_map = ss.mapping["projects"]
            with st.form("project_mapping_form"):
                new_proj_choices = {}
                for val in project_values:
                    current = proj_map.get(val, {}).get("key")
                    chosen = select_value(f"`{val}`", project_options, current, f"proj__{val}")
                    new_proj_choices[val] = chosen
                submitted = st.form_submit_button("\U0001F4BE Save project mapping", type="primary")
            if submitted:
                for val, key in new_proj_choices.items():
                    if key is None:
                        proj_map.pop(val, None)
                    else:
                        name = next(p["name"] for p in projects if p["key"] == key)
                        proj_map[val] = {"key": key, "name": name}
                storage.save_mapping(ss.mapping)
                st.success("Project mapping saved.")

        mapped_project_keys = sorted({v["key"] for v in proj_map.values() if v.get("key")})

        # projects actually present in the currently uploaded sheet (a saved
        # mapping may reference projects from an older upload that this sheet
        # doesn't use at all)
        sheet_project_raws = set(xp.distinct_values(ss.df, "project_raw"))
        sheet_project_keys = {proj_map[val]["key"] for val in sheet_project_raws if val in proj_map}
        component_relevant_keys = [k for k in mapped_project_keys if k in sheet_project_keys]

        # ---- 3b. Component & issue-type mapping ---------------------------
        with st.container(border=True):
            st.markdown("#### \U0001F9E9 Map Excel **Component** values → Jira Component (per project)")
            if not component_relevant_keys:
                st.info("Map at least one project used in this sheet above first.")
            else:
                comp_map = ss.mapping["components"]
                issue_type_map = ss.mapping["issue_types"]
                with st.spinner("Loading issue types & components from Jira…"):
                    with st.form("component_mapping_form"):
                        pending_components = {}
                        pending_issue_types = {}
                        for pkey in component_relevant_keys:
                            pname = next((p["name"] for p in projects if p["key"] == pkey), pkey)
                            excel_values_for_project = sorted(
                                {
                                    (clean(row.component_raw) or BLANK)
                                    for row in ss.df.itertuples()
                                    if (clean(row.project_raw) or BLANK)
                                    in [k for k, v in proj_map.items() if v["key"] == pkey]
                                }
                            )
                            with st.expander(f"Jira Project: {pkey} — {pname}", expanded=True):
                                try:
                                    it_options = get_issue_types(pkey)
                                except JiraError as e:
                                    it_options = []
                                    st.error(f"Could not load issue types: {e}")
                                it_choices = [(None, NOT_MAPPED)] + [(it["id"], it["name"]) for it in it_options]
                                current_it = issue_type_map.get(pkey, {}).get("id")
                                if current_it is None:
                                    task_it = next((it for it in it_options if it["name"].lower() == "task"), None)
                                    current_it = task_it["id"] if task_it else None
                                chosen_it = select_value(
                                    "Default issue type for new tasks", it_choices, current_it, f"itype__{pkey}"
                                )
                                pending_issue_types[pkey] = chosen_it

                                try:
                                    comps = get_components(pkey)
                                except JiraError as e:
                                    comps = []
                                    st.error(f"Could not load components: {e}")
                                comp_options = [(None, NO_COMPONENT)] + [(c["id"], c["name"]) for c in comps]
                                for val in excel_values_for_project:
                                    current = comp_map.get(pkey, {}).get(val, {}).get("id")
                                    chosen = select_value(
                                        f"Component `{val}`", comp_options, current, f"comp__{pkey}__{val}"
                                    )
                                    pending_components[(pkey, val)] = chosen
                        submitted_c = st.form_submit_button("\U0001F4BE Save component & issue type mapping", type="primary")
                if submitted_c:
                    for pkey, it_id in pending_issue_types.items():
                        if it_id is None:
                            issue_type_map.pop(pkey, None)
                        else:
                            it_name = next(it["name"] for it in get_issue_types(pkey) if it["id"] == it_id)
                            issue_type_map[pkey] = {"id": it_id, "name": it_name}
                    for (pkey, val), comp_id in pending_components.items():
                        comp_map.setdefault(pkey, {})
                        if comp_id is None:
                            comp_map[pkey].pop(val, None)
                        else:
                            comp_name = next(c["name"] for c in get_components(pkey) if c["id"] == comp_id)
                            comp_map[pkey][val] = {"id": comp_id, "name": comp_name}
                    storage.save_mapping(ss.mapping)
                    st.success("Component & issue type mapping saved.")

        # ---- 3c. Assignee mapping ------------------------------------------
        with st.container(border=True):
            st.markdown("#### \U0001F464 Map Excel **Assignee** names → Jira User")
            if not mapped_project_keys:
                st.info("Map at least one project above first.")
            else:
                all_users, seen = [], set()
                with st.spinner("Loading assignable users from Jira…"):
                    for pkey in mapped_project_keys:
                        try:
                            for u in get_users(pkey):
                                if u["accountId"] not in seen:
                                    seen.add(u["accountId"])
                                    all_users.append(u)
                        except JiraError as e:
                            st.error(f"Could not load users for {pkey}: {e}")
                all_users.sort(key=lambda u: u.get("displayName", ""))
                user_options = [(None, UNASSIGNED)] + [
                    (u["accountId"], f"{u.get('displayName')} ({u.get('emailAddress', 'no email')})")
                    for u in all_users
                ]
                assignee_values = xp.distinct_values(ss.df, "assignee_raw")
                assignee_map = ss.mapping["assignees"]
                with st.form("assignee_mapping_form"):
                    pending_assignees = {}
                    for val in assignee_values:
                        current = assignee_map.get(val, {}).get("accountId")
                        chosen = select_value(f"`{val}`", user_options, current, f"assignee__{val}")
                        pending_assignees[val] = chosen
                    submitted_a = st.form_submit_button("\U0001F4BE Save assignee mapping", type="primary")
                if submitted_a:
                    for val, account_id in pending_assignees.items():
                        if account_id is None:
                            assignee_map.pop(val, None)
                        else:
                            display_name = next(u["displayName"] for u in all_users if u["accountId"] == account_id)
                            assignee_map[val] = {"accountId": account_id, "displayName": display_name}
                    storage.save_mapping(ss.mapping)
                    st.success("Assignee mapping saved.")

        # ---- Task title format (per project) -----------------------------------
        with st.container(border=True):
            st.markdown("#### \U0001F4DD Task Title Format (per project)")
            st.caption(
                "Build the Jira task title (summary) from literal text plus Excel columns, per project. "
                "Pick a column below and click Insert to add it as a placeholder, or type placeholders "
                "directly. Example: `[NTP-SAAS][{Component}] {Description}` → "
                "`[NTP-SAAS][AT]Production/UAT Support - AT - Resource 1`."
            )
            if not component_relevant_keys:
                st.info("Map at least one project used in this sheet above first.")
            else:
                title_template_map = ss.mapping["title_templates"]
                placeholder_cols = title_template_columns(ss.df)
                placeholder_labels = [placeholder_name(c) for c in placeholder_cols]

                for pkey in component_relevant_keys:
                    pname = next((p["name"] for p in projects if p["key"] == pkey), pkey)
                    draft_key = f"title_template_draft__{pkey}"
                    if draft_key not in st.session_state:
                        st.session_state[draft_key] = title_template_map.get(pkey, "{Description}")

                    with st.expander(f"Jira Project: {pkey} — {pname}", expanded=False):
                        ins_col, btn_col = st.columns([3, 1])
                        insert_label = ins_col.selectbox(
                            "Insert a column as a placeholder", placeholder_labels, key=f"title_insert_choice__{pkey}"
                        )
                        btn_col.write("")
                        btn_col.write("")
                        if btn_col.button("➕ Insert", key=f"title_insert_btn__{pkey}", use_container_width=True):
                            st.session_state[draft_key] += "{" + insert_label + "}"
                            st.rerun()

                        template_text = st.text_input(
                            "Title template",
                            key=draft_key,
                            help="Type literal text freely; placeholders like {Component} or {Description} "
                            "are replaced with each row's value at creation time.",
                        )

                        project_raws_for_pkey = [k for k, v in proj_map.items() if v["key"] == pkey]
                        sample_row = next(
                            (
                                row
                                for row in ss.df.itertuples()
                                if (clean(row.project_raw) or BLANK) in project_raws_for_pkey
                            ),
                            None,
                        )
                        if sample_row is not None and template_text.strip():
                            preview_text = render_title_template(template_text, sample_row, placeholder_cols)
                            st.caption(f"Preview: `{preview_text}`")

                        if st.button("\U0001F4BE Save title template", key=f"title_save__{pkey}", type="primary"):
                            if template_text.strip():
                                title_template_map[pkey] = template_text
                            else:
                                title_template_map.pop(pkey, None)
                            storage.save_mapping(ss.mapping)
                            st.success("Title template saved.")

        # ---- Task labels (per project) -----------------------------------------
        with st.container(border=True):
            st.markdown("#### \U0001F3F7️ Task Labels (per project)")
            st.caption(
                "Labels attached to every task created for a project. Nothing is added unless configured here — "
                "leave blank for no labels. Pick from labels already used in Jira, or type a new one and press "
                "enter to add it. Jira labels can't contain spaces; spaces you type are converted to hyphens."
            )
            st.button("\U0001F504 Refresh Jira labels", on_click=lambda: get_all_labels(force=True))
            if not component_relevant_keys:
                st.info("Map at least one project used in this sheet above first.")
            else:
                try:
                    existing_labels = sorted(get_all_labels())
                except JiraError as e:
                    existing_labels = []
                    st.error(f"Could not load existing Jira labels: {e}")

                label_map = ss.mapping["labels"]
                for pkey in component_relevant_keys:
                    pname = next((p["name"] for p in projects if p["key"] == pkey), pkey)
                    current_labels = label_map.get(pkey, [])
                    # options must include any already-saved labels even if they
                    # somehow aren't in Jira's site-wide label list (yet, or anymore)
                    options = sorted(set(existing_labels) | set(current_labels))
                    with st.form(f"labels_form__{pkey}"):
                        chosen_labels = st.multiselect(
                            f"Labels for {pkey} — {pname}",
                            options=options,
                            default=current_labels,
                            accept_new_options=True,
                            key=f"labels_input__{pkey}",
                        )
                        submitted_labels = st.form_submit_button("\U0001F4BE Save labels", type="primary")
                    if submitted_labels:
                        parsed = [label.strip().replace(" ", "-") for label in chosen_labels if label.strip()]
                        if parsed:
                            label_map[pkey] = parsed
                        else:
                            label_map.pop(pkey, None)
                        storage.save_mapping(ss.mapping)
                        st.success(f"Labels saved for {pkey}." if parsed else f"Labels cleared for {pkey}.")

        # ---- 3d. Story link type --------------------------------------------
        with st.container(border=True):
            st.markdown("#### \U0001F517 Link new tasks to their Jira Story")
            st.caption(
                "When a row's 'Jira Story' value looks like a real Jira issue key (e.g. NTPSFC-16683), "
                "the created task will be linked to it using the link type chosen below."
            )
            try:
                with st.spinner("Loading issue link types from Jira…"):
                    link_types = get_link_types()
            except JiraError as e:
                link_types = []
                st.error(f"Could not load issue link types: {e}")

            link_options = [(None, NO_LINK)]
            for lt in link_types:
                link_options.append((f"{lt['name']}::outward", f"New task — {lt['outward']} — the story"))
                if lt.get("inward") != lt.get("outward"):
                    link_options.append((f"{lt['name']}::inward", f"New task — {lt['inward']} — the story"))

            current_link = ss.mapping.get("link_type")
            current_link_value = f"{current_link['name']}::{current_link['task_side']}" if current_link else None
            with st.form("link_type_form"):
                chosen_link = select_value("Link type", link_options, current_link_value, "link_type_select")
                submitted_l = st.form_submit_button("\U0001F4BE Save link type", type="primary")
            if submitted_l:
                if chosen_link is None:
                    ss.mapping["link_type"] = None
                else:
                    name, side = chosen_link.split("::")
                    ss.mapping["link_type"] = {"name": name, "task_side": side}
                storage.save_mapping(ss.mapping)
                st.success("Link type saved.")

        # ---- Additional field mappings (generic, per project) -----------------
        with st.container(border=True):
            st.markdown("#### \U0001F9E9 Additional Field Mappings")
            st.caption(
                "Map any other Jira field, per project — either from an Excel column (e.g. Capitalization, "
                "Priority) or to a fixed value used on every task (for a field Jira requires that has no "
                "matching Excel column at all, e.g. 'Payment Category'). For a dropdown-style Jira field, an "
                "Excel-column mapping selects the matching option automatically (e.g. '75' → '75%'). Columns "
                "already handled above (Project, Component, Assignee, Jira Story, Description, JIRA ID, "
                "Estimated Time) aren't offered as a column source here."
            )
            eligible_columns = [c for c in ss.df.columns if c not in ALREADY_CONFIGURED_COLUMNS]
            if not component_relevant_keys:
                st.info("Map at least one project used in this sheet above first.")
            else:
                extra_field_map = ss.mapping["extra_fields"]
                for pkey in component_relevant_keys:
                    pname = next((p["name"] for p in projects if p["key"] == pkey), pkey)
                    if not issue_type_map.get(pkey):
                        st.warning(f"Set a default issue type for {pkey} in the Component mapping section above first.")
                        continue
                    configured = extra_field_map.get(pkey, {})
                    with st.expander(f"Jira Project: {pkey} — {pname}", expanded=bool(configured)):
                        if configured:
                            st.caption("Configured mappings:")
                            for field_id, cfg in list(configured.items()):
                                if cfg["source"] == "column":
                                    desc = f"**{cfg['column_label']}** (Excel column) → {cfg['field_name']}"
                                else:
                                    val_display = cfg["value"]["label"] if isinstance(cfg["value"], dict) else cfg["value"]
                                    # strip before bolding — Markdown won't render **text **
                                    # (space touching the closing delimiter) as bold at all
                                    val_display = val_display.strip() if isinstance(val_display, str) else val_display
                                    desc = f"Fixed value **{val_display}** → {cfg['field_name']}"
                                rcol1, rcol2 = st.columns([4, 1])
                                rcol1.write(desc)
                                if rcol2.button("Remove", key=f"rmfield__{pkey}__{field_id}", use_container_width=True):
                                    configured.pop(field_id, None)
                                    storage.save_mapping(ss.mapping)
                                    st.rerun()

                        try:
                            with st.spinner(f"Loading available fields for {pkey}…"):
                                candidates = [f for f in custom_field_candidates(pkey) if f["fieldId"] not in configured]
                        except JiraError as e:
                            candidates = []
                            st.error(f"Could not load fields for {pkey}: {e}")

                        if not candidates:
                            st.caption("All available fields are mapped for this project.")
                        else:
                            field_labels = [f["name"] for f in candidates]
                            chosen_field_label = st.selectbox(
                                "Jira field", field_labels, key=f"extrafieldsel__{pkey}"
                            )
                            chosen_field = candidates[field_labels.index(chosen_field_label)]

                            source_choice = st.radio(
                                "Value comes from",
                                ["An Excel column", "A fixed value (same on every task)"],
                                key=f"extrasource__{pkey}",
                                horizontal=True,
                            )

                            if source_choice == "An Excel column":
                                used_columns = {
                                    c["column_key"] for c in configured.values() if c.get("source") == "column"
                                }
                                remaining_columns = [c for c in eligible_columns if c not in used_columns]
                                if not remaining_columns:
                                    st.caption("No remaining Excel columns to map for this project.")
                                else:
                                    col_labels_list = [
                                        COLUMN_LABELS.get(c, c.replace("_", " ").title()) for c in remaining_columns
                                    ]
                                    chosen_col_label = st.selectbox(
                                        "Excel column", col_labels_list, key=f"extracol__{pkey}"
                                    )
                                    chosen_col = remaining_columns[col_labels_list.index(chosen_col_label)]
                                    if st.button("\U0001F4BE Add mapping", key=f"extraadd__{pkey}", type="primary"):
                                        configured[chosen_field["fieldId"]] = {
                                            "field_name": chosen_field["name"],
                                            "source": "column",
                                            "column_key": chosen_col,
                                            "column_label": COLUMN_LABELS.get(chosen_col, chosen_col),
                                        }
                                        storage.save_mapping(ss.mapping)
                                        st.success(f"Mapped {COLUMN_LABELS.get(chosen_col, chosen_col)} → {chosen_field['name']}.")
                                        st.rerun()
                            elif chosen_field.get("allowedValues"):
                                opt_labels = [av["value"] for av in chosen_field["allowedValues"]]
                                chosen_opt_label = st.selectbox("Value", opt_labels, key=f"extraval__{pkey}")
                                if st.button("\U0001F4BE Add mapping", key=f"extraadd__{pkey}", type="primary"):
                                    chosen_opt = next(
                                        av for av in chosen_field["allowedValues"] if av["value"] == chosen_opt_label
                                    )
                                    configured[chosen_field["fieldId"]] = {
                                        "field_name": chosen_field["name"],
                                        "source": "fixed",
                                        "value": {"id": chosen_opt["id"], "label": chosen_opt["value"]},
                                    }
                                    storage.save_mapping(ss.mapping)
                                    st.success(f"Set fixed value {chosen_opt['value']} → {chosen_field['name']}.")
                                    st.rerun()
                            else:
                                fixed_text = st.text_input("Value", key=f"extraval__{pkey}")
                                if st.button("\U0001F4BE Add mapping", key=f"extraadd__{pkey}", type="primary"):
                                    if fixed_text.strip():
                                        configured[chosen_field["fieldId"]] = {
                                            "field_name": chosen_field["name"],
                                            "source": "fixed",
                                            "value": fixed_text.strip(),
                                        }
                                        storage.save_mapping(ss.mapping)
                                        st.success(f"Set fixed value → {chosen_field['name']}.")
                                        st.rerun()
                                    else:
                                        st.warning("Type a value first.")

# ============================================================================
# TAB 4 - REVIEW & CREATE
# ============================================================================
elif nav == PAGES[3]:
    st.subheader("\U0001F680 Review resolved tasks & create in Jira")
    if not ss.client:
        st.warning("Connect to Jira first (tab 1).")
    elif ss.df.empty:
        st.warning("Upload and parse a sheet first (tab 2).")
    else:
        proj_map = ss.mapping["projects"]
        comp_map = ss.mapping["components"]
        issue_type_map = ss.mapping["issue_types"]
        assignee_map = ss.mapping["assignees"]
        extra_field_map = ss.mapping["extra_fields"]
        title_template_map = ss.mapping["title_templates"]
        title_columns = title_template_columns(ss.df)
        link_type_cfg = ss.mapping.get("link_type")
        label_map = ss.mapping["labels"]

        # Every fixed-value extra field configured on a project used in this
        # sheet — lets the review table offer a per-task override for each.
        sheet_project_raws_t4 = set(xp.distinct_values(ss.df, "project_raw"))
        sheet_project_keys_t4 = {proj_map[v]["key"] for v in sheet_project_raws_t4 if v in proj_map}
        fixed_field_catalog = {}
        for pkey_fc in sheet_project_keys_t4:
            for field_id, cfg in extra_field_map.get(pkey_fc, {}).items():
                if cfg["source"] == "fixed" and field_id not in fixed_field_catalog:
                    field_def = next(
                        (f for f in get_create_fields(pkey_fc) if f["fieldId"] == field_id), None
                    )
                    fixed_field_catalog[field_id] = {
                        "field_name": cfg["field_name"],
                        "allowed_values": field_def.get("allowedValues") if field_def else None,
                    }

        def resolve(row):
            praw = clean(row.project_raw) or BLANK
            p = proj_map.get(praw)
            pkey = p["key"] if p else None
            pname = p["name"] if p else None
            it = issue_type_map.get(pkey) if pkey else None
            craw = clean(row.component_raw) or BLANK
            comp = comp_map.get(pkey, {}).get(craw) if pkey else None
            araw = clean(row.assignee_raw) or BLANK
            assignee = assignee_map.get(araw)
            story_raw = (clean(row.jira_story) or "").strip()
            story_key = story_raw.upper() if STORY_KEY_RE.match(story_raw) else None
            ready = bool(pkey and it)

            title_template = title_template_map.get(pkey) if pkey else None
            jira_title = (
                render_title_template(title_template, row, title_columns)
                if title_template and title_template.strip()
                else row.description
            )

            fixed_defaults = {}
            for field_id in fixed_field_catalog:
                cfg = extra_field_map.get(pkey, {}).get(field_id) if pkey else None
                if cfg and cfg["source"] == "fixed":
                    val = cfg["value"]
                    display_val = val["label"] if isinstance(val, dict) else val
                    # stripped so it exactly matches the (also-stripped) SelectboxColumn
                    # options built below — a trailing/leading space on the option's
                    # real Jira label otherwise silently breaks the pre-filled default
                    fixed_defaults[f"fixed__{field_id}"] = display_val.strip() if isinstance(display_val, str) else display_val
                else:
                    fixed_defaults[f"fixed__{field_id}"] = None

            return pd.Series(
                {
                    "jira_project": f"{pkey}" if pkey else None,
                    "jira_project_name": pname,
                    "jira_title": jira_title,
                    "jira_component": comp["name"] if comp else None,
                    "jira_issue_type": it["name"] if it else None,
                    "jira_assignee": assignee["displayName"] if assignee else None,
                    "jira_extra_fields": extra_fields_summary(row, pkey, extra_field_map),
                    "story_link": story_key,
                    "ready": ready,
                    **fixed_defaults,
                }
            )

        resolved = pd.concat([ss.df, ss.df.apply(resolve, axis=1)], axis=1)

        with st.container(border=True):
            col1, col2 = st.columns(2)
            skip_existing = col1.checkbox("Skip rows that already have a JIRA ID", value=True)
            statuses = sorted([s for s in resolved["status"].dropna().unique().tolist()])
            exclude_statuses = col2.multiselect(
                "Exclude rows with status",
                statuses,
                default=[s for s in statuses if s.lower() in ("done", "discarded")],
            )

            working = resolved.copy()
            if skip_existing:
                working = working[working["jira_id"].isna()]
            if exclude_statuses:
                working = working[~working["status"].isin(exclude_statuses)]

            not_ready = working[~working["ready"]]
            ready_rows = working[working["ready"]]

            if len(not_ready) > 0:
                st.warning(
                    f"{len(not_ready)} row(s) are missing a project or issue-type mapping and will be skipped. "
                    "Fix them in tab 3 if needed."
                )

            # Rendered before the table below so an override made here is
            # already reflected in the table's "Extra Fields" column — it's
            # built from the SAME extra_fields_summary() the creation loop
            # itself uses, so what you see here is exactly what gets sent.
            if fixed_field_catalog and len(ready_rows) > 0:
                with st.expander("\U0001F527 Override Fixed Field Values (optional)", expanded=False):
                    st.caption(
                        "These values come from Additional Field Mappings and apply to every task by default — "
                        "edit any cell to override just that one task. A cell has no effect on a row whose "
                        "project doesn't have that field mapped at all."
                    )
                    override_base_cols = ["row_number", "jira_title"] + [f"fixed__{fid}" for fid in fixed_field_catalog]
                    override_base = ready_rows[override_base_cols].reset_index(drop=True)

                    editor_column_config = {
                        "row_number": st.column_config.NumberColumn("Row #", disabled=True),
                        "jira_title": st.column_config.TextColumn("Task Title", disabled=True),
                    }
                    for field_id, meta in fixed_field_catalog.items():
                        colname = f"fixed__{field_id}"
                        if meta["allowed_values"]:
                            # de-duplicate in case two options only differ by whitespace
                            opts = list(dict.fromkeys(av["value"].strip() for av in meta["allowed_values"]))
                            editor_column_config[colname] = st.column_config.SelectboxColumn(
                                meta["field_name"], options=opts
                            )
                        else:
                            editor_column_config[colname] = st.column_config.TextColumn(meta["field_name"])

                    # Key includes a hash of the current fixed-field config so the
                    # editor re-initializes with fresh defaults whenever that
                    # config changes — st.data_editor otherwise keeps whatever it
                    # first rendered for a given key across reruns, even after the
                    # underlying default value changes (e.g. a fixed value added
                    # or edited in Configure Mapping after this tab was first seen).
                    fixed_cfg_signature = hashlib.md5(
                        json.dumps(
                            {
                                pkey_sig: {
                                    fid: cfg for fid, cfg in entries.items() if cfg.get("source") == "fixed"
                                }
                                for pkey_sig, entries in extra_field_map.items()
                            },
                            sort_keys=True,
                            default=str,
                        ).encode()
                    ).hexdigest()[:10]
                    editor_key = f"fixed_field_override_editor__{ss.sheet_name}__{fixed_cfg_signature}"

                    edited = st.data_editor(
                        override_base,
                        column_config=editor_column_config,
                        hide_index=True,
                        use_container_width=True,
                        key=editor_key,
                    )
                    ss.fixed_field_overrides = edited.set_index("row_number")
            else:
                ss.fixed_field_overrides = None

            # Recompute the extra-fields summary for ready rows now that any
            # override above is known, so the table reflects it immediately.
            if ss.fixed_field_overrides is not None and len(ready_rows) > 0:
                def _summary_with_override(r):
                    override_row = (
                        ss.fixed_field_overrides.loc[r["row_number"]]
                        if r["row_number"] in ss.fixed_field_overrides.index
                        else None
                    )
                    return extra_fields_summary(r, r["jira_project"], extra_field_map, override_row)

                working.loc[ready_rows.index, "jira_extra_fields"] = ready_rows.apply(_summary_with_override, axis=1)

            m1, m2, m3 = st.columns(3)
            m1.metric("Total in sheet", len(ss.df))
            m2.metric("Shown after filters", len(working))
            m3.metric("Ready to create", len(ready_rows))

            working["selected"] = False if ss.selection_reset_mode == "none" else working["ready"]
            st.caption(
                "Untick **Create?** on any row to skip it — even a row that's ready. "
                "Ticking it on a row that isn't ready has no effect; it still can't be created."
            )
            sel_col1, sel_col2, _sel_spacer = st.columns([1, 1, 3])
            if sel_col1.button("☑️ Select All", use_container_width=True):
                ss.selection_reset_mode = "default"
                ss.selection_reset_nonce += 1
                st.rerun()
            if sel_col2.button("⬜ Unselect All", use_container_width=True):
                ss.selection_reset_mode = "none"
                ss.selection_reset_nonce += 1
                st.rerun()

            review_cols = [
                "selected", "row_number", "section", "jira_title", "jira_project", "jira_component",
                "jira_assignee", "jira_issue_type", "jira_extra_fields", "story_link",
                "estimated_hours", "status", "ready",
            ]
            review_col_config = column_labels(review_cols)
            review_col_config["selected"] = st.column_config.CheckboxColumn("Create?", default=True)
            edited_review = st.data_editor(
                working[review_cols],
                use_container_width=True,
                height=350,
                column_config=review_col_config,
                disabled=[c for c in review_cols if c != "selected"],
                hide_index=True,
                # Key includes the active filters and the Select/Unselect All
                # nonce, so either one forces a clean re-init with fresh
                # "Create?" values instead of keeping stale per-cell edits —
                # st.data_editor otherwise ignores new default data for a key
                # it has already rendered once.
                key=(
                    f"review_selection_editor__{ss.sheet_name}__{skip_existing}"
                    f"__{sorted(exclude_statuses)}__{ss.selection_reset_nonce}"
                ),
            )
            selected_row_numbers = set(
                edited_review.loc[edited_review["selected"] & edited_review["ready"], "row_number"]
            )
            ready_rows = ready_rows[ready_rows["row_number"].isin(selected_row_numbers)]

        def create_tasks_for_rows(rows_df):
            """Runs the creation loop for the given rows (a resolved dataframe
            slice with the same columns as `resolved`) and returns a results
            dataframe. Used both for the initial bulk create and for retrying
            only the rows that failed."""
            results = []
            total = len(rows_df)
            if total == 0:
                return pd.DataFrame(results)
            progress = st.progress(0.0, text="Creating issues...")
            for i, row in enumerate(rows_df.itertuples()):
                pkey = row.jira_project
                it = issue_type_map.get(pkey)
                comp = comp_map.get(pkey, {}).get(clean(row.component_raw) or BLANK)
                assignee = assignee_map.get(clean(row.assignee_raw) or BLANK)
                story_link = clean(row.story_link)

                description_text = row.description

                original_estimate = None
                estimated_hours = clean(row.estimated_hours)
                try:
                    if estimated_hours and float(estimated_hours) > 0:
                        original_estimate = f"{int(float(estimated_hours))}h"
                except (TypeError, ValueError):
                    pass

                labels = list(label_map.get(pkey, []))

                title_template = title_template_map.get(pkey)
                summary_text = (
                    render_title_template(title_template, row, title_columns)
                    if title_template and title_template.strip()
                    else row.description
                )

                fields = ss.client.build_fields(
                    project_key=pkey,
                    issue_type_id=it["id"],
                    summary=summary_text,
                    description_text=description_text,
                    component_id=comp["id"] if comp else None,
                    labels=labels,
                    original_estimate=original_estimate,
                )

                override_row = None
                if ss.fixed_field_overrides is not None and row.row_number in ss.fixed_field_overrides.index:
                    override_row = ss.fixed_field_overrides.loc[row.row_number]

                extra_notes = []
                for field_id, cfg in extra_field_map.get(pkey, {}).items():
                    try:
                        field_def = next(
                            (f for f in get_create_fields(pkey) if f["fieldId"] == field_id), None
                        )
                        if not field_def:
                            extra_notes.append(f"'{cfg['field_name']}' is no longer available on {pkey}'s create screen")
                            continue

                        if cfg["source"] == "fixed":
                            override_val = None
                            if override_row is not None:
                                colname = f"fixed__{field_id}"
                                if colname in override_row.index:
                                    ov = clean(override_row[colname])
                                    override_val = ov if ov != "" else None

                            if field_def.get("allowedValues"):
                                target_label = override_val if override_val is not None else (
                                    cfg["value"]["label"] if isinstance(cfg["value"], dict) else cfg["value"]
                                )
                                # compare stripped — the option's real Jira label may have
                                # trailing/leading whitespace that the UI doesn't preserve
                                matched = next(
                                    (
                                        av for av in field_def["allowedValues"]
                                        if av["value"].strip() == str(target_label).strip()
                                    ),
                                    None,
                                )
                                if matched:
                                    fields[field_id] = {"id": matched["id"]}
                                else:
                                    extra_notes.append(f"'{target_label}' is not a valid option for '{cfg['field_name']}'")
                            else:
                                target_val = override_val if override_val is not None else cfg["value"]
                                value, err = build_custom_field_value(field_def, target_val)
                                if err:
                                    extra_notes.append(err)
                                else:
                                    fields[field_id] = value
                            continue

                        raw_val = clean(getattr(row, cfg["column_key"], None))
                        if raw_val is None:
                            continue
                        value, err = build_custom_field_value(field_def, raw_val)
                        if err:
                            extra_notes.append(err)
                        else:
                            fields[field_id] = value
                    except JiraError as ce:
                        extra_notes.append(f"could not load '{cfg['field_name']}' options: {ce}")

                try:
                    created = ss.client.create_issue(fields)
                    issue_key = created["key"]
                    notes = list(extra_notes)
                    if created.get("_dropped_fields"):
                        notes.append(f"dropped fields: {created['_dropped_fields']}")

                    if assignee:
                        try:
                            ss.client.assign_issue(issue_key, assignee["accountId"])
                        except JiraError as ae:
                            notes.append(f"assignment failed: {ae}")

                    if story_link and link_type_cfg:
                        try:
                            if link_type_cfg["task_side"] == "outward":
                                ss.client.create_issue_link(link_type_cfg["name"], issue_key, story_link)
                            else:
                                ss.client.create_issue_link(link_type_cfg["name"], story_link, issue_key)
                        except JiraError as le:
                            notes.append(f"story link failed: {le}")

                    results.append(
                        {
                            "row_number": row.row_number,
                            "description": row.description,
                            "status": "✅ created",
                            "issue_key": issue_key,
                            "url": f"{ss.connection['base_url']}/browse/{issue_key}",
                            "note": "; ".join(notes),
                        }
                    )
                except JiraError as e:
                    results.append(
                        {
                            "row_number": row.row_number,
                            "description": row.description,
                            "status": "❌ failed",
                            "issue_key": "",
                            "url": "",
                            "note": str(e),
                        }
                    )
                progress.progress((i + 1) / total, text=f"Creating issues... ({i + 1}/{total})")
                time.sleep(0.05)

            progress.empty()
            return pd.DataFrame(results)

        confirm = st.checkbox(f"I reviewed the table above and want to create {len(ready_rows)} Jira task(s).")
        if st.button("\U0001F680 Create Jira tasks", type="primary", disabled=not (confirm and len(ready_rows) > 0)):
            ss.create_results = create_tasks_for_rows(ready_rows)

        if ss.create_results is not None:
            with st.container(border=True):
                st.markdown("#### ✅ Results")
                n_ok = (ss.create_results["status"].str.contains("created")).sum()
                n_fail = (ss.create_results["status"].str.contains("failed")).sum()
                c1, c2 = st.columns(2)
                c1.metric("Created", int(n_ok))
                c2.metric("Failed", int(n_fail))

                if int(n_fail) > 0:
                    if st.button(f"\U0001F501 Retry {int(n_fail)} failed task(s)", type="primary"):
                        failed_row_numbers = ss.create_results.loc[
                            ss.create_results["status"].str.contains("failed"), "row_number"
                        ]
                        retry_rows = resolved[
                            resolved["row_number"].isin(failed_row_numbers) & resolved["ready"]
                        ]
                        not_retryable = set(failed_row_numbers) - set(retry_rows["row_number"])
                        if not_retryable:
                            st.warning(
                                f"Row(s) {sorted(not_retryable)} can no longer be retried — their project/issue-type "
                                "mapping is missing. Fix the mapping in Configure Mapping, then retry again."
                            )
                        if len(retry_rows) > 0:
                            retry_results = create_tasks_for_rows(retry_rows)
                            # replace each retried row's old outcome with the new one; rows
                            # not retried (e.g. already-created, or now unmappable) are kept as-is
                            kept = ss.create_results[
                                ~ss.create_results["row_number"].isin(retry_results["row_number"])
                            ]
                            ss.create_results = (
                                pd.concat([kept, retry_results], ignore_index=True)
                                .sort_values("row_number")
                                .reset_index(drop=True)
                            )
                            st.rerun()

                st.dataframe(
                    ss.create_results,
                    use_container_width=True,
                    height=300,
                    column_config=column_labels(ss.create_results.columns),
                )

                dl1, dl2 = st.columns(2)
                dl1.download_button(
                    "\U0001F4E5 Download results as CSV",
                    ss.create_results.to_csv(index=False).encode("utf-8"),
                    file_name="jira_import_results.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

                if int(n_ok) > 0:
                    updated_bytes = build_updated_workbook(
                        ss.excel_bytes, ss.sheet_name, ss.header_row, ss.col_map, ss.create_results
                    )
                    base_name = os.path.splitext(ss.excel_filename or "sprint_backlog.xlsx")[0]
                    dl2.download_button(
                        "\U0001F4E5 Download Excel with JIRA IDs filled in",
                        updated_bytes,
                        file_name=f"{base_name}_with_jira_ids.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True,
                    )
                    st.caption(
                        f"The '{ss.sheet_name}' sheet's JIRA ID column is filled in for the "
                        f"{int(n_ok)} task(s) just created; everything else in the workbook is untouched."
                    )

# ============================================================================
# TAB 5 - UPDATE LABELS (existing issues)
# ============================================================================
elif nav == PAGES[4]:
    st.subheader("\U0001F3F7️ Update labels on existing Jira tasks")
    if not ss.client:
        st.warning("Connect to Jira first (tab 1).")
    else:
        ss.setdefault("label_keys_text", "")
        ss.setdefault("label_preview", None)

        with st.container(border=True):
            st.markdown("#### 1. Which tasks?")
            st.caption("Issue keys, separated by new lines, commas or spaces (e.g. `NTP-101, NTP-102`).")
            if not ss.df.empty and "jira_id" in ss.df.columns:
                sheet_keys = [k for k in ss.df["jira_id"].dropna().astype(str).str.strip() if STORY_KEY_RE.match(k)]
                if st.button(
                    f"\U0001F4CB Load {len(sheet_keys)} key(s) from the sheet's JIRA ID column",
                    disabled=not sheet_keys,
                ):
                    ss.label_keys_text = "\n".join(dict.fromkeys(sheet_keys))
                    ss.label_preview = None
                    st.rerun()
            st.text_area("Issue keys", key="label_keys_text", height=140, label_visibility="collapsed")

        with st.container(border=True):
            st.markdown("#### 2. What to change?")
            mode = st.radio(
                "Action",
                ["Add labels", "Remove labels", "Replace all labels"],
                horizontal=True,
                key="label_update_mode",
            )
            try:
                existing_labels = sorted(get_all_labels())
            except JiraError as e:
                existing_labels = []
                st.error(f"Could not load existing Jira labels: {e}")
            chosen = st.multiselect(
                "Labels",
                options=existing_labels,
                accept_new_options=True,
                key="label_update_labels",
                help="Pick existing labels or type a new one and press enter. Spaces become hyphens.",
            )
            labels_sel = list(dict.fromkeys(l.strip().replace(" ", "-") for l in chosen if l.strip()))
            if mode == "Replace all labels":
                st.caption("Every selected task's labels are overwritten with exactly the list above (empty = clear all).")

        def new_labels_for(current):
            if mode == "Add labels":
                return current + [l for l in labels_sel if l not in current]
            if mode == "Remove labels":
                return [l for l in current if l not in labels_sel]
            return list(labels_sel)

        with st.container(border=True):
            st.markdown("#### 3. Preview & apply")
            if st.button("\U0001F50D Load tasks & preview", type="primary"):
                keys = list(dict.fromkeys(k.upper() for k in re.split(r"[\s,;]+", ss.label_keys_text) if k))
                bad = [k for k in keys if not STORY_KEY_RE.match(k)]
                if bad:
                    st.error(f"Not valid issue keys: {', '.join(bad)}")
                elif not keys:
                    st.warning("Enter at least one issue key.")
                else:
                    rows = []
                    with st.spinner(f"Loading {len(keys)} issue(s)…"):
                        for k in keys:
                            try:
                                info = ss.client.get_issue_brief(k)
                                rows.append({**info, "error": ""})
                            except JiraError as e:
                                rows.append({"key": k, "summary": "", "labels": [], "error": str(e)[:200]})
                    ss.label_preview = rows

            if ss.label_preview:
                if not labels_sel and mode != "Replace all labels":
                    st.info("Select at least one label above to see the result.")
                table = []
                for r in ss.label_preview:
                    new = new_labels_for(r["labels"])
                    table.append(
                        {
                            "Issue": r["key"],
                            "Summary": r["summary"],
                            "Current labels": ", ".join(r["labels"]),
                            "New labels": ", ".join(new) if not r["error"] else "",
                            "Changes": (new != r["labels"]) and not r["error"],
                            "Note": r["error"],
                        }
                    )
                st.dataframe(table, use_container_width=True, height=min(400, 45 + 35 * len(table)))
                todo = [
                    r for r, t in zip(ss.label_preview, table)
                    if t["Changes"]
                ]
                st.caption(f"{len(todo)} of {len(table)} task(s) will change.")
                confirmed = st.checkbox(
                    f"I confirm updating labels on {len(todo)} task(s) in Jira", key="label_update_confirm"
                )
                if st.button("✅ Apply label update", type="primary", disabled=not (todo and confirmed)):
                    ok, failed = 0, []
                    progress = st.progress(0.0)
                    for i, r in enumerate(todo, 1):
                        try:
                            if mode == "Add labels":
                                ss.client.update_labels(r["key"], add=labels_sel)
                            elif mode == "Remove labels":
                                ss.client.update_labels(r["key"], remove=labels_sel)
                            else:
                                ss.client.update_labels(r["key"], replace_with=labels_sel)
                            r["labels"] = new_labels_for(r["labels"])
                            ok += 1
                        except JiraError as e:
                            failed.append(f"{r['key']}: {str(e)[:150]}")
                        progress.progress(i / len(todo))
                    ss.all_labels = None  # new labels may now exist site-wide
                    if ok:
                        st.success(f"Updated labels on {ok} task(s).")
                    for f in failed:
                        st.error(f)

# Rendered last so it reflects anything the tabs above just changed in this
# same run (connecting, testing, parsing a sheet) rather than lagging behind.
render_sidebar_status()
