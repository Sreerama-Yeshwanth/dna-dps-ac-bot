import logging
import os

import azure.functions as func
from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings, TurnContext
from botbuilder.schema import Activity

from src.dps_bot import DPSBot

logging.basicConfig(level=logging.INFO)

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)

SETTINGS = BotFrameworkAdapterSettings(
    app_id=os.environ.get("MicrosoftAppId", ""),
    app_password=os.environ.get("MicrosoftAppPassword", ""),
    channel_auth_tenant=os.environ.get("MicrosoftAppTenantId", ""),
)

ADAPTER = BotFrameworkAdapter(SETTINGS)
BOT = DPSBot()

@app.route(route="messages", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
async def messages(req: func.HttpRequest) -> func.HttpResponse:
    # Azure Bot Service posts each incoming activity to this /api/messages route.
    body = req.get_json()

    # Convert the request JSON into a Bot Framework Activity object.
    activity = Activity().deserialize(body)
    auth_header = req.headers.get("Authorization", "")

    # The adapter validates the request, creates a turn context, and calls BOT.on_turn.
    await ADAPTER.process_activity(activity, auth_header, BOT.on_turn)

    # Normal message activities are acknowledged with HTTP 200 after the bot sends its reply.
    return func.HttpResponse(status_code=200)