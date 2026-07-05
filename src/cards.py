import json

INPUT_TEXT = "Input.Text"
ROWS_FIELD_ID = "rows_json"
ACCESS_CHOICES = [{"title": "Read", "value": "Read"}, {"title": "Write", "value": "Write"}]


def _row_summary(row: dict) -> str:
    target = (
        row["catalog"]
        + (f".{row['schema']}" if row.get("schema") else "")
        + (f".{row['table']}" if row.get("table") else "")
    )
    return f"{row['workspace'] or 'Not provided'} / {target} ({row['access']})"


def build_request_card(rows: list[dict] | None = None) -> dict:
    rows = rows or []
    summary = (
        [
            {
                "type": "TextBlock",
                "text": "Requests added so far:",
                "weight": "Bolder",
                "wrap": True,
            },
            {
                "type": "TextBlock",
                "text": "\n".join(f"- {_row_summary(row)}" for row in rows),
                "wrap": True,
            },
        ]
        if rows
        else []
    )

    return {
        "type": "AdaptiveCard",
        "version": "1.2",  # conservative for Azure Portal Web Chat compatibility
        "fallbackText": "DPSBot request form.",
        "body": summary
        + [
            {
                "type": "TextBlock",
                "text": "Add a data access request",
                "weight": "Bolder",
                "spacing": "Medium",
                "wrap": True,
            },
            {"type": INPUT_TEXT, "id": "workspace", "placeholder": "Workspace"},
            {"type": INPUT_TEXT, "id": "catalog", "placeholder": "Catalog"},
            {"type": INPUT_TEXT, "id": "schema", "placeholder": "Schema"},
            {"type": INPUT_TEXT, "id": "table", "placeholder": "Table"},
            {
                "type": "Input.ChoiceSet",
                "id": "access",
                "style": "compact",
                "value": "Write",
                "choices": ACCESS_CHOICES,
            },
            {"type": INPUT_TEXT, "id": "message", "placeholder": "Message", "isMultiline": True},
            # Hidden inputs still round-trip on Action.Submit, carrying
            # accumulated rows forward statelessly.
            {
                "type": INPUT_TEXT,
                "id": ROWS_FIELD_ID,
                "value": json.dumps(rows),
                "isVisible": False,
            },
        ],
        "actions": [
            {"type": "Action.Submit", "title": "Add another", "data": {"action": "add_row"}},
            {"type": "Action.Submit", "title": "Submit all", "data": {"action": "submit_requests"}},
        ],
    }
