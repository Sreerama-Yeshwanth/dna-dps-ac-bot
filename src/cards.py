INPUT_TEXT = "Input.Text"


def build_request_card() -> dict:
    # This dictionary is the Adaptive Card form shown when the user mentions DPSBot.
    return {
        "type": "AdaptiveCard",
        # Version 1.2 is conservative for Azure Portal Web Chat compatibility.
        "version": "1.2",
        "fallbackText": "DPSBot request form.",
        "body": [
            {
                "type": INPUT_TEXT,
                "id": "catalog",
                "placeholder": "Catalog",
            },
            {"type": INPUT_TEXT, "id": "schema", "placeholder": "Schema"},
            {"type": INPUT_TEXT, "id": "table", "placeholder": "Table"},
            {
                "type": INPUT_TEXT,
                "id": "message",
                "placeholder": "Message",
                "isMultiline": True,
            },
        ],
        "actions": [
            # Input values are posted back in activity.value when the user clicks Submit.
            {
                "type": "Action.Submit",
                "title": "Submit",
                "data": {"action": "submit_table_details"},
            }
        ],
    }
