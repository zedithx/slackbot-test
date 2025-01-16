import os
from slack_sdk import WebClient
from slack_sdk.socket_mode import SocketModeClient
from slack_sdk.socket_mode.request import SocketModeRequest
from slack_sdk.socket_mode.response import SocketModeResponse
from slack_sdk.errors import SlackApiError
from dotenv import load_dotenv
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)  # Set to DEBUG for more details
logger = logging.getLogger(__name__)

# Load environment variables from .env file
load_dotenv()

# Initialize Slack clients
slack_bot_token = os.getenv("SLACK_BOT_TOKEN")
slack_app_token = os.getenv("SLACK_APP_TOKEN")

web_client = WebClient(token=slack_bot_token)
socket_client = SocketModeClient(app_token=slack_app_token, web_client=web_client)

# A simple list to store logs of user messages
message_log = []

# Function to handle messages containing "log this" or "note this"
def handle_message(event_data):
    message = event_data.get("text", "")
    user = event_data.get("user", "")
    channel = event_data.get("channel", "")

    if "log this" in message.lower() or "note this" in message.lower():
        # Save the message to the log
        message_log.append({
            "user": user,
            "text": message,
            "timestamp": datetime.utcnow().isoformat(),
        })

        # Send confirmation to the channel
        try:
            web_client.chat_postMessage(
                channel=channel,
                text=f"Got it, <@{user}>! I've noted your message: \"{message}\"."
            )
        except SlackApiError as e:
            logger.info(f"Error posting message: {e.response['error']}")

# Function to handle app mentions for "show notes"
def handle_app_mention(event_data):
    text = event_data.get("text", "")
    channel = event_data.get("channel", "")

    if "show notes" in text.lower():
        notes_summary = "\n".join(
            [f"{i+1}. <@{entry['user']}>: \"{entry['text']}\"" for i, entry in enumerate(message_log)]
        ) if message_log else "No notes recorded yet."

        try:
            web_client.chat_postMessage(
                channel=channel,
                text=f"Here are the recorded notes:\n{notes_summary}"
            )
        except SlackApiError as e:
            logger.info(f"Error posting message: {e.response['error']}")

# Handle events from the Socket Mode client
def handle_events(client, payload):
    # Log the payload for debugging
    logger.debug(f"Received payload: {payload}")

    if payload.type == "events_api":
        event = payload.payload.get("event", {})
        logger.info(f"Processing event: {event}")

        if event.get("type") == "message" and not event.get("bot_id"):  # Ignore bot messages
            handle_message(event)
        elif event.get("type") == "app_mention":
            handle_app_mention(event)

        # Acknowledge the event
        client.send_socket_mode_response(SocketModeResponse(envelope_id=payload.envelope_id))

# Start the Socket Mode client
if __name__ == "__main__":
    socket_client.socket_mode_request_listeners.append(handle_events)
    logger.info("⚡️ Slack Bot is running!")

    # Connect to Slack
    socket_client.connect()

    # Prevent the script from exiting
    try:
        import time
        while True:
            time.sleep(1)  # Keep the script alive
    except KeyboardInterrupt:
        logger.info("Bot stopped manually.")