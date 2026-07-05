import json
import os

import requests
from botbuilder.core import CardFactory, MessageFactory, TurnContext
from botbuilder.core.teams import TeamsActivityHandler
from botbuilder.schema import ChannelAccount, Mention

from src.cards import ROWS_FIELD_ID, build_request_card

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE = "https://graph.microsoft.com/.default"
GRAPH_SELECT = "$select"
OWNER_SELECT = "id,displayName,userPrincipalName,mail"


def get_graph_token():
    return requests.post(
        f"https://login.microsoftonline.com/{os.environ['MicrosoftAppTenantId']}/oauth2/v2.0/token",
        data={
            "client_id": os.environ["MicrosoftAppId"],
            "client_secret": os.environ["MicrosoftAppPassword"],
            "scope": GRAPH_SCOPE,
            "grant_type": "client_credentials",
        },
        timeout=10,
    ).json()["access_token"]


def find_group_owners(access_token: str, display_name: str) -> tuple[dict | None, list[dict]]:
    headers = {"Authorization": f"Bearer {access_token}"}
    groups = (
        requests.get(
            f"{GRAPH_BASE_URL}/groups",
            headers=headers,
            params={
                "$filter": "displayName eq '{}'".format(display_name.replace("'", "''")),
                GRAPH_SELECT: "id,displayName",
            },
            timeout=10,
        )
        .json()
        .get("value", [])
    )

    if not groups:
        return None, []

    group = groups[0]
    owners = []
    url = f"{GRAPH_BASE_URL}/groups/{group['id']}/owners/microsoft.graph.user"
    params = {GRAPH_SELECT: OWNER_SELECT}

    while url:
        data = requests.get(url, headers=headers, params=params, timeout=10).json()
        owners.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None

    return group, owners


def jira_issue_url(issue_key: str) -> str:
    return f"{os.environ['JIRA_BASE_URL'].rstrip('/')}/browse/{issue_key}"


def create_jira_issue(summary: str, description: str) -> dict:
    # Jira Cloud REST v3 wants the description as an Atlassian Document Format (ADF) doc.
    adf_description = {
        "type": "doc",
        "version": 1,
        "content": [
            {"type": "paragraph", "content": [{"type": "text", "text": line or " "}]}
            for line in description.splitlines()
        ],
    }
    return requests.post(
        f"{os.environ['JIRA_BASE_URL'].rstrip('/')}/rest/api/3/issue",
        auth=(os.environ["JIRA_EMAIL"], os.environ["JIRA_API_TOKEN"]),
        headers={"Accept": "application/json", "Content-Type": "application/json"},
        json={
            "fields": {
                "project": {"key": os.environ["JIRA_PROJECT_KEY"]},
                "summary": summary,
                "description": adf_description,
                "issuetype": {"name": os.environ["JIRA_ISSUE_TYPE"]},
            }
        },
        timeout=15,
    ).json()


def build_group_name(catalog: str, schema: str, access: str) -> str:
    # Schema-scoped when a schema is given (table is display-only), else catalog-scoped.
    if schema:
        return f"UC|Sch|{catalog}.{schema}|{access}"
    return f"UC|Cat|{catalog}|{access}"


def row_from_submitted_value(value: dict) -> dict | None:
    row = {
        k: str(value.get(k, "")).strip()
        for k in ("workspace", "catalog", "schema", "table", "access", "message")
    }
    if not row["catalog"]:
        return None
    row["access"] = row["access"] or "Write"
    return row


def _detail_lines(
    requester_name: str, group_name: str, row: dict, jira_issue: dict | None = None
) -> list[str]:
    # Shared by the Jira description and the Teams owner message, so both show the same facts.
    lines = [
        f"Requested by: {requester_name or 'A user'}",
        f"Group: {group_name}",
        f"Workspace: {row['workspace'] or 'Not provided'}",
        f"Catalog: {row['catalog']}",
        f"Schema: {row['schema'] or 'Not provided'}",
        f"Table: {row['table'] or 'Not provided'}",
        f"Access: {row['access']}",
        f"Message: {row['message'] or 'No message provided.'}",
    ]
    if jira_issue:
        lines.append(f"Jira: {jira_issue['key']} {jira_issue_url(jira_issue['key'])}")
    lines.append("As you are the owner, kindly look into it.")
    return lines


def build_access_request_text(
    owners: list[dict], requester_name: str, group_name: str, row: dict
) -> str:
    owner_references = ", ".join(
        f"{owner['displayName']} <{owner.get('userPrincipalName') or owner['mail']}>"
        for owner in owners
    )
    return "\n".join(
        [f"Owners: {owner_references}"] + _detail_lines(requester_name, group_name, row)
    )


def build_owner_request_activity(
    owners: list[dict], requester_name: str, group_name: str, row: dict, jira_issue: dict
):
    mentions, owner_texts = [], []
    for owner in owners:
        name, owner_id = owner["displayName"], owner.get("id")
        if owner_id:
            mentions.append(
                Mention(
                    mentioned=ChannelAccount(id=owner_id, name=name, aad_object_id=owner_id),
                    text=f"<at>{name}</at>",
                    type="mention",
                )
            )
        owner_texts.append(f"<at>{name}</at>" if owner_id else name)

    header = f"{', '.join(owner_texts)}, please review this access request."
    activity = MessageFactory.text(
        "\n".join([header] + _detail_lines(requester_name, group_name, row, jira_issue))
    )
    activity.entities = mentions
    return activity


class DPSBot(TeamsActivityHandler):
    async def on_message_activity(self, turn_context: TurnContext) -> None:
        value = turn_context.activity.value
        action = value.get("action") if isinstance(value, dict) else None

        if action in ("add_row", "submit_requests"):
            await self._handle_form_submission(turn_context, value, action)
            return

        await turn_context.send_activity(
            MessageFactory.attachment(CardFactory.adaptive_card(build_request_card()))
        )

    async def _handle_form_submission(
        self, turn_context: TurnContext, value: dict, action: str
    ) -> None:
        try:
            rows = json.loads(value.get(ROWS_FIELD_ID) or "[]")
        except ValueError:
            rows = []

        current_row = row_from_submitted_value(value)
        if current_row:
            rows.append(current_row)

        if action == "add_row":
            if not current_row:
                await turn_context.send_activity(
                    "Please enter at least a catalog before adding another request."
                )
                return
            await turn_context.send_activity(
                MessageFactory.attachment(CardFactory.adaptive_card(build_request_card(rows)))
            )
            return

        if not rows:
            await turn_context.send_activity(
                "Please enter at least one request (catalog) before submitting."
            )
            return

        requester = getattr(turn_context.activity.from_property, "name", None)
        access_token = get_graph_token()
        for row in rows:
            await self._submit_row(turn_context, access_token, requester, row)

    async def _submit_row(
        self, turn_context: TurnContext, access_token: str, requester: str, row: dict
    ) -> None:
        group_name = build_group_name(row["catalog"], row["schema"], row["access"])
        group, owners = find_group_owners(access_token, group_name)
        if not group:
            await turn_context.send_activity(f"No matching group found for {group_name}.")
            return
        if not owners:
            await turn_context.send_activity(f"No owners found for {group_name}.")
            return

        description = build_access_request_text(owners, requester, group_name, row)
        jira_issue = create_jira_issue(f"DPSBot access request - {group_name}", description)
        await turn_context.send_activity(
            build_owner_request_activity(owners, requester, group_name, row, jira_issue)
        )
