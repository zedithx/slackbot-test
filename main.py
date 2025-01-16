import os
from slack_sdk import WebClient
from slack_sdk.socket_mode import SocketModeClient
from slack_sdk.socket_mode.response import SocketModeResponse
from slack_sdk.errors import SlackApiError
from dotenv import load_dotenv
from datetime import datetime
import logging

# For excel writing
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


# Function to initialize the Excel file if it doesn't exist
def initialize_excel_file(path):
    # Format the filename as "DD-MM_CIT.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Checkin Log"
    # Write headers
    sheet.append(["User", "Message", "Timestamp"])
    workbook.save(path)


def check_duplicate_user(excel_file, user):
    # Load the workbook and access the active sheet
    workbook = load_workbook(excel_file)
    sheet = workbook.active

    # Check if the user already exists in the "User" column
    user_column_index = 1  # Assuming "User" is in the first column
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=user_column_index, max_col=user_column_index):
        for cell in row:
            if cell.value == user:
                logger.info(f"Duplicate user found: {user}")
                return True

    logger.info(f"No duplicate found for user: {user}")
    return False


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

def handle_message(event_data):
    message = event_data.get("text", "")
    user = event_data.get("user", "")
    channel = event_data.get("channel", "")

    if "checkin" in message.lower() or "check in" in message.lower():
        # Save the message to the log
        message_log.append({
            "user": user,
            "text": message,
            "timestamp": datetime.utcnow().strftime("%H:%M"),
        })
        # Write to an excel instead
        row_data = [user, datetime.utcnow().strftime("%H:%M")]
        # Send confirmation to the channel
        # This will determine file path
        current_date = datetime.now().strftime("%d-%m_CIT")
        excel_file = f"checkin_records/{current_date}.xlsx"
        if not os.path.exists(excel_file):
            initialize_excel_file(excel_file)
        elif check_duplicate_user(excel_file, user):
            # return error message
            web_client.chat_postMessage(
                channel=channel,
                text=f"Hi @{user}. You have already checked in for today! Please do not send duplicate messages 😡."
            )
        else:
            workbook = load_workbook(excel_file)
            sheet = workbook.active
            sheet.append(row_data)  # Append the row
            workbook.save(excel_file)  # Save changes
            logger.info(f"Check in for @{user} has been recorded in the excel")
        try:
            web_client.chat_postMessage(
                channel=channel,
                text=f"@{user} has checked in! Have a wonderful day at the office today!."
            )
            logger.info(f"Check in for @{user} is successful")
        except SlackApiError as e:
            logger.info(f"Error posting message: {e.response['error']}")

# Function to handle app mentions for "show notes"
# def handle_app_mention(event_data):
#     text = event_data.get("text", "")
#     channel = event_data.get("channel", "")
#
#     if "" in text.lower():
#         notes_summary = "\n".join(
#             [f"{i+1}. <@{entry['user']}>: \"{entry['text']}\"" for i, entry in enumerate(message_log)]
#         ) if message_log else "No notes recorded yet."
#
#         try:
#             web_client.chat_postMessage(
#                 channel=channel,
#                 text=f"Here are the recorded notes:\n{notes_summary}"
#             )
#         except SlackApiError as e:
#             logger.info(f"Error posting message: {e.response['error']}")

# Handle events from the Socket Mode client
def handle_events(client, payload):
    # Log the payload for debugging
    # logger.debug(f"Received payload: {payload}")

    if payload.type == "events_api":
        event = payload.payload.get("event", {})
        logger.info(f"Processing event: {event}")

        if event.get("type") == "message" and not event.get("bot_id"):  # Ignore bot messages
            handle_message(event)
        # elif event.get("type") == "app_mention":
        #     handle_app_mention(event)

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