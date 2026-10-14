"""Thin wrapper around the Jira Cloud REST API (v3) used by the importer app."""
import json
import requests
from requests.auth import HTTPBasicAuth

TIMEOUT = 20


class JiraError(Exception):
    pass


def _to_adf(text):
    """Wrap plain text into Atlassian Document Format, treating blank-line
    -separated chunks as separate paragraphs — Jira does not render a literal
    '\\n' inside a single text node as a line break."""
    paragraphs = [p for p in text.split("\n\n") if p.strip()] or [text]
    return {
        "type": "doc",
        "version": 1,
        "content": [{"type": "paragraph", "content": [{"type": "text", "text": p}]} for p in paragraphs],
    }


class JiraClient:
    def __init__(self, base_url, email, api_token):
        self.base_url = base_url.strip().rstrip("/")
        self.auth = HTTPBasicAuth(email.strip(), api_token.strip())
        self.headers = {"Accept": "application/json", "Content-Type": "application/json"}

    def _url(self, path):
        return f"{self.base_url}{path}"

    def _get(self, path, params=None):
        r = requests.get(self._url(path), auth=self.auth, headers=self.headers, params=params, timeout=TIMEOUT)
        if r.status_code >= 400:
            raise JiraError(f"GET {path} failed ({r.status_code}): {r.text[:400]}")
        return r.json() if r.text else {}

    def _put(self, path, body):
        r = requests.put(self._url(path), auth=self.auth, headers=self.headers, json=body, timeout=TIMEOUT)
        if r.status_code >= 400:
            raise JiraError(f"PUT {path} failed ({r.status_code}): {r.text[:400]}")

    # ---- read-only lookups used to populate dropdowns -----------------

    def test_connection(self):
        return self._get("/rest/api/3/myself")

    def get_projects(self):
        results, start_at = [], 0
        while True:
            data = self._get("/rest/api/3/project/search", params={"maxResults": 50, "startAt": start_at})
            results.extend(data.get("values", []))
            if data.get("isLast", True) or not data.get("values"):
                break
            start_at += 50
        return sorted(results, key=lambda p: p["key"])

    def get_components(self, project_key):
        return self._get(f"/rest/api/3/project/{project_key}/components")

    def get_issue_types(self, project_key):
        try:
            project = self._get(f"/rest/api/3/project/{project_key}")
            data = self._get(
                f"/rest/api/3/issue/createmeta/{project['id']}/issuetypes", params={"maxResults": 50}
            )
            return [it for it in data.get("issueTypes", []) if not it.get("subtask")]
        except JiraError:
            return []

    def get_assignable_users(self, project_key):
        return self._get(
            "/rest/api/3/user/assignable/search", params={"project": project_key, "maxResults": 200}
        )

    def get_issue_link_types(self):
        return self._get("/rest/api/3/issueLinkType").get("issueLinkTypes", [])

    def get_labels(self):
        """All labels currently in use anywhere on the site (not project-scoped)."""
        labels, start_at = [], 0
        while True:
            data = self._get("/rest/api/3/label", params={"maxResults": 200, "startAt": start_at})
            labels.extend(data.get("values", []))
            if data.get("isLast", True) or not data.get("values"):
                break
            start_at += 200
        return labels

    def get_create_fields(self, project_key, issue_type_id):
        """All fields on the given project+issue-type's create screen, including
        `allowedValues` for select-list fields (project-scoped custom fields like
        Components or a Capitalization dropdown live here)."""
        project = self._get(f"/rest/api/3/project/{project_key}")
        fields, start_at = [], 0
        while True:
            data = self._get(
                f"/rest/api/3/issue/createmeta/{project['id']}/issuetypes/{issue_type_id}",
                params={"maxResults": 100, "startAt": start_at},
            )
            fields.extend(data.get("fields", []))
            if data.get("isLast", True) or not data.get("fields"):
                break
            start_at += 100
        return fields

    # ---- linking -----------------------------------------------------------

    def create_issue_link(self, type_name, outward_issue_key, inward_issue_key):
        """outward_issue_key <outward-phrase> inward_issue_key, per the link
        type's wording (e.g. for 'Blocks': outward 'blocks', inward 'is blocked by')."""
        r = requests.post(
            self._url("/rest/api/3/issueLink"),
            auth=self.auth,
            headers=self.headers,
            json={
                "type": {"name": type_name},
                "outwardIssue": {"key": outward_issue_key},
                "inwardIssue": {"key": inward_issue_key},
            },
            timeout=TIMEOUT,
        )
        if r.status_code >= 400:
            raise JiraError(f"issueLink failed ({r.status_code}): {r.text[:400]}")

    # ---- issue creation --------------------------------------------------

    def create_issue(self, fields):
        """POST /rest/api/3/issue, transparently dropping fields the target
        project's create screen rejects and retrying, so differently
        configured projects don't hard-fail the whole import."""
        payload = {"fields": dict(fields)}
        dropped = []
        for _ in range(6):
            r = requests.post(
                self._url("/rest/api/3/issue"),
                auth=self.auth,
                headers=self.headers,
                json=payload,
                timeout=TIMEOUT,
            )
            if r.status_code < 400:
                result = r.json()
                result["_dropped_fields"] = dropped
                return result
            try:
                body = r.json()
            except ValueError:
                raise JiraError(r.text[:500])
            errors = body.get("errors", {}) or {}
            removable = [f for f in errors if f in payload["fields"]]
            if not removable:
                raise JiraError(json.dumps(body)[:500])
            for f in removable:
                payload["fields"].pop(f, None)
                dropped.append(f)
        raise JiraError(f"Create failed after dropping fields {dropped}")

    # ---- updating existing issues ------------------------------------------

    def get_issue_brief(self, issue_key):
        """Summary + current labels of an existing issue."""
        data = self._get(f"/rest/api/3/issue/{issue_key}", params={"fields": "summary,labels"})
        fields = data.get("fields", {})
        return {"key": data.get("key", issue_key), "summary": fields.get("summary", ""), "labels": fields.get("labels", [])}

    def update_labels(self, issue_key, add=None, remove=None, replace_with=None):
        """Edit an issue's labels. `replace_with` (a list, possibly empty) sets the
        whole label list; otherwise `add` / `remove` change only those labels and
        leave the rest untouched."""
        if replace_with is not None:
            body = {"fields": {"labels": list(replace_with)}}
        else:
            ops = [{"add": lbl} for lbl in add or []] + [{"remove": lbl} for lbl in remove or []]
            body = {"update": {"labels": ops}}
        self._put(f"/rest/api/3/issue/{issue_key}", body)

    def assign_issue(self, issue_key, account_id):
        self._put(f"/rest/api/3/issue/{issue_key}/assignee", {"accountId": account_id})

    def build_fields(
        self,
        project_key,
        issue_type_id,
        summary,
        description_text=None,
        component_id=None,
        labels=None,
        original_estimate=None,
        parent_key=None,
    ):
        fields = {
            "project": {"key": project_key},
            "issuetype": {"id": issue_type_id},
            "summary": summary[:255],
        }
        if description_text:
            fields["description"] = _to_adf(description_text)
        if component_id:
            fields["components"] = [{"id": component_id}]
        if labels:
            fields["labels"] = labels
        if original_estimate:
            fields["timetracking"] = {"originalEstimate": original_estimate}
        if parent_key:
            fields["parent"] = {"key": parent_key}
        return fields
