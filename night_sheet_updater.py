import os
import base64
from collections import defaultdict
from datetime import datetime

import environ
import requests
import pandas as pd

from office365_api import Sharepoint

# === ENVIRONMENT SETUP ===
env = environ.Env()
env.read_env()

API_KEY = env("MAZEVO_API_KEY")
GET_EVENTS_URL = env("GET_EVENTS_URL")
GET_BOOKING_DETAILS_URL = env("GET_BOOKING_DETAILS_URL")
GET_DIAGRAM_URL = env("GET_DIAGRAM_URL")

BUILDING_IDS = [334, 335, 340, 341]


# === UTILITIES ===

def format_date(dt: datetime) -> str:
    return dt.strftime("%Y-%m-%dT00:00:00-05:00")

def get_date_str(date_str: str) -> str:
    try:
        # return date string in mm_dd_yyyy format
        dt = datetime.fromisoformat(date_str)
        return dt.strftime("%m_%d_%Y")
    except ValueError:
        raise ValueError(f"Invalid date format: {date_str}. Expected ISO format.")


def fetch_api_data(API_URL: str, body, method='POST') -> dict:
    headers = {
        "X-API-Key": API_KEY,
        "Content-Type": "application/json"
    }

    try:
        if method.upper() == 'GET':
            response = requests.get(API_URL, headers=headers)
        else:
            response = requests.post(API_URL, json=body, headers=headers)
        response.raise_for_status()
        print(f"✅ API call to {API_URL} successful")
        return response.json()
    except Exception as e:
        raise Exception(f"API call failed: {e}")


def save_diagram_and_upload(base64_string: str, file_name: str, upload_folder: str) -> str:
    """
    Decodes a base64-encoded image and uploads it directly to SharePoint.
    Returns the uploaded SharePoint file path.
    """
    try:
        # Ensure the file has a .png extension
        if not file_name.lower().endswith(".png"):
            file_name = file_name.rsplit('.', 1)[0] + ".png"

        # Decode base64 to binary
        binary_data = base64.b64decode(base64_string)

        # Optionally save locally (temporary)
        local_path = f"local_directory/temp_diagrams/{file_name}"
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(binary_data)

        # Upload to SharePoint
        sharepoint = Sharepoint()
        result = sharepoint.upload_file(file_name, upload_folder, binary_data)
        if result["error"]:
            raise Exception(f"Upload failed: {result['error']}")

        # Return the SharePoint path
        return f"https://mavsuta.sharepoint.com/sites/EOTEventOperationsTeam181/Shared%20Documents/{upload_folder}/{file_name}"
    except Exception as e:
        raise Exception(f"Error handling diagram upload: {e}")


# === DATA PROCESSING ===

def compose_bookings(raw_data: list) -> list:
    composed = []
    for item in raw_data:
        # Download and upload diagram if it exists
        if item.get("hasDiagram"):
            try:
                diagram_url = GET_DIAGRAM_URL + '' + str(item.get("bookingId"))
                response = fetch_api_data(diagram_url, {}, method='GET')
                base64_str = response.get("file")
                if base64_str:
                    date_of_event = get_date_str(item.get("dateTimeStart"))
                    diagram_file_name = date_of_event + '_' + str(item.get("bookingId"))
                    diagram_file_name += '_' + response["fileName"]
                    uploaded_path = save_diagram_and_upload(base64_str, diagram_file_name, "General/EventSetupDiagrams/Mazevo/RoomDiagrams")
                    item["diagramPath"] = uploaded_path
            except Exception as e:
                print(f"⚠️ Failed to handle diagram for booking {item.get('bookingId')}: {e}")
                item["diagramPath"] = None

        # Compose Booking
        booking = {
            "bookingId": item.get("bookingId"),
            "diagramPath": item.get("diagramPath"),
            "buildingDescription": item.get("buildingDescription"),
            "roomDescription": item.get("roomDescription"),
            "dateTimeStart": item.get("dateTimeStart"),
            "dateTimeEnd": item.get("dateTimeEnd"),
            "hasDiagram": item.get("hasDiagram"),
            "setupNotes": item.get("setupNotes"),
            "setupStyle": item.get("setupStyle"),
            "setupCount": item.get("setupCount"),
            "eventType": item.get("eventType"),
            "bookingDetails": []
        }

        for detail in item.get("bookingDetails", []):
            booking["bookingDetails"].append({
                "bookingId": item.get("bookingId"),
                "bookingDetailId": detail.get("bookingDetailId"),
                "resourceId": detail.get("resourceId"),
                "resource": detail.get("resource"),
                "serviceProvider": detail.get("serviceProvider"),
                "serviceProviderId": detail.get("serviceProviderId"),
                "quantity": detail.get("quantity"),
                "notes": detail.get("notes"),
                "specialInstructions": detail.get("specialInstructions"),
                "serviceStartTime": detail.get("serviceStartTime"),
                "serviceEndTime": detail.get("serviceEndTime"),
            })

        composed.append(booking)
    return composed


def group_bookings_by_date(bookings: list) -> dict:
    grouped = defaultdict(list)
    for booking in bookings:
        try:
            dt = datetime.fromisoformat(booking["dateTimeStart"])
            key = dt.strftime("%m_%d_%Y")
            grouped[key].append(booking)
        except Exception as e:
            print(f"⚠️ Skipping booking with bad date: {e}")
    return grouped


def process_excel(bookings: list, file_path: str):
    grouped = group_bookings_by_date(bookings)
    print("📅 Grouped bookings by date:")
    for date, items in grouped.items():
        print(f" - {date}: {len(items)} bookings")

    # You can modify this to update Excel using pandas or openpyxl
    # Placeholder:
    df = pd.read_excel(file_path)
    df.loc[len(df)] = ["Modified by script"] + [None] * (len(df.columns) - 1)
    df.to_excel(file_path, index=False)


def filter_events(api_data):
    booking_ids = []
    for item in api_data:
        if item.get("statusDescription") == "Confirmed" and item.get("eventType") != "Maintenance":
            bid = item.get("bookingId")
            if bid not in booking_ids:
                booking_ids.append(bid)
    return booking_ids


# === MAIN FUNCTION ===

def run_on_sharepoint_file(folder_path: str, file_name: str, start_date: datetime, end_date: datetime) -> str:
    sharepoint = Sharepoint()

    get_events_body = {
        "start": format_date(start_date),
        "end": format_date(end_date),
        "buildingIds": BUILDING_IDS
    }

    events_data = fetch_api_data(GET_EVENTS_URL, get_events_body)
    booking_ids = filter_events(events_data)

    get_booking_details_body = {
        "bookingIds": booking_ids
    }

    booking_details_data = fetch_api_data(GET_BOOKING_DETAILS_URL, get_booking_details_body)
    structured_bookings = compose_bookings(booking_details_data)

    # Download from SharePoint
    result = sharepoint.download_file(file_name, folder_path)
    if result["error"]:
        raise Exception(f"Download failed: {result['error']}")
    local_path = result["downloaded_file_path"]

    # Process file
    process_excel(structured_bookings, local_path)

    # Upload back
    with open(local_path, "rb") as f:
        content = f.read()
    upload_result = sharepoint.upload_file(file_name, folder_path, content)
    if upload_result["error"]:
        raise Exception(f"Upload failed: {upload_result['error']}")

    return f"✅ Processed and uploaded '{file_name}' to '{folder_path}'"

if __name__ == "__main__":
    # Example usage
    folder_path = "Apps/Mazevo"
    file_name = "Night Sheet - Multi Day Test.xlsx"
    start_date = datetime(2025, 6, 16)
    end_date = datetime(2025, 6, 19)

    try:
        result = run_on_sharepoint_file(folder_path, file_name, start_date, end_date)
        print(result)
    except Exception as e:
        print(f"❌ Error: {e}")
