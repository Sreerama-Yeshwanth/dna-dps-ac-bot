import logging
import os

import azure.functions as func
from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings
from botbuilder.schema import Activity

from src.dps_bot import DPSBot

logging.basicConfig(level=logging.INFO)

app = func.FunctionApp(http_auth_level=func.AuthLevel.ANONYMOUS)
ADAPTER = BotFrameworkAdapter(
    BotFrameworkAdapterSettings(
        app_id=os.environ.get("MicrosoftAppId", ""),
        app_password=os.environ.get("MicrosoftAppPassword", ""),
        channel_auth_tenant=os.environ.get("MicrosoftAppTenantId", ""),
    )
)
BOT = DPSBot()


@app.route(route="messages", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
async def messages(req: func.HttpRequest) -> func.HttpResponse:
    activity = Activity().deserialize(req.get_json())
    await ADAPTER.process_activity(activity, req.headers.get("Authorization", ""), BOT.on_turn)
    return func.HttpResponse(status_code=200)
