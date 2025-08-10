import time
import requests
import tempfile
import win32clipboard
import keyboard
import pyautogui
import webbrowser

from google.oauth2 import service_account
from googleapiclient.discovery import build

# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = "service_account.json"
SHEET_ID = "sheet id"
SHEET_RANGE = "ranges"  # Include column M now
SCOPES = ["spreadsheet link"]

PLATE_RECOGNIZER_API_KEY = "your_api"  # Replace with your API key
WHATSAPP_GROUP_NAME = "contact name"


# === GOOGLE SHEET FUNCTIONS ===

def get_first_pending_row(service):
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE
    ).execute()
    rows = result.get("values", [])
    for idx, row in enumerate(rows, start=2):
        while len(row) < 13:  # Ensure up to column M
            row.append("")
        if row[11].strip().upper() == "N":  # Column L
            return idx, row
    return None, None


def update_flag_and_plate(service, row_index, plate_text):
    body = {"values": [["Y", plate_text]]}
    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=f"Sheet2!L{row_index}:M{row_index}",
        valueInputOption="RAW",
        body=body
    ).execute()
    print(f"✅ Updated flag to Y and plate to '{plate_text}' at row {row_index}")


# === IMAGE DOWNLOAD + OCR ===

def download_image_from_drive_link(link):
    if "uc?id=" in link:
        file_id = link.split("uc?id=")[-1].split("&")[0]
        direct_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    elif "open?id=" in link:
        file_id = link.split("open?id=")[-1].split("&")[0]
        direct_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    elif "/file/d/" in link:
        file_id = link.split("/file/d/")[1].split("/")[0]
        direct_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    else:
        direct_url = link

    response = requests.get(direct_url, stream=True)
    if response.status_code == 200:
        temp = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg")
        with open(temp.name, 'wb') as f:
            for chunk in response.iter_content(1024):
                f.write(chunk)
        return temp.name
    else:
        print("❌ Failed to download image:", response.status_code)
        return None


def recognize_plate_from_image(image_path, threshold=80.0):
    url = 'url of model'

    with open(image_path, 'rb') as image_file:
        response = requests.post(
            url,
            files={'upload': image_file},
            data={'regions': 'in'},
            headers={'Authorization': f'Token {PLATE_RECOGNIZER_API_KEY}'}
        )

    if response.status_code in [200, 201]:
        result = response.json()
        results = result.get("results", [])
        if not results:
            print("⚠️ No plate detected.")
            return "NOT_DETECTED"
        
        plate_data = results[0]
        plate_number = plate_data['plate'].upper()
        confidence = plate_data['score'] * 100

        print(f"🔍 Plate: {plate_number} | Confidence: {confidence:.1f}%")

        if confidence >= threshold:
            return plate_number
        else:
            print("⚠️ Confidence below threshold. Marking as NOT_DETECTED.")
            return "NOT_DETECTED"
    else:
        print("❌ PlateRecognizer API Error:", response.status_code)
        return "API_ERROR"



# === WHATSAPP SENDER ===

def send_whatsapp_with_image_and_message(image_drive_link, group_name, message_text, number_plate):
    webbrowser.open(image_drive_link)
    time.sleep(8)
    keyboard.press_and_release('ctrl+c')  # Copy image
    time.sleep(1)

    # Open WhatsApp
    keyboard.press_and_release('win')
    time.sleep(1)
    keyboard.write('WhatsApp')
    time.sleep(1)
    keyboard.press_and_release('enter')
    time.sleep(5)

    # Search group
    keyboard.press_and_release('ctrl+f')
    time.sleep(1)
    keyboard.write(group_name)
    time.sleep(2)
    keyboard.press_and_release('enter')
    time.sleep(2)
    keyboard.press_and_release('down')
    time.sleep(2)
    keyboard.press_and_release('enter')
    time.sleep(2)

    # Paste image
    keyboard.press_and_release('ctrl+v')
    time.sleep(2)

    # Append number plate to message
    final_message = f"{message_text.strip()}\n*Number Plate*: {number_plate}"

    # Copy message
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(final_message)
    win32clipboard.CloseClipboard()
    print("📋 Message copied to clipboard.")

    # Paste & send
    keyboard.press_and_release('ctrl+v')
    time.sleep(1)
    keyboard.press_and_release('enter')
    print("✅ WhatsApp message sent with image and number plate.")


# === MAIN LOOP ===

def main():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build("sheets", "v4", credentials=creds)

    while True:
        row_index, row = get_first_pending_row(service)
        if not row:
            print("⏳ No rows with Flag 'N'. Waiting 15 seconds...")
            time.sleep(15)
            continue

        drive_link = row[9]       # Column J
        message_text = row[10]    # Column K

        print(f"\n➡️ Processing row {row_index}...")

        local_image = download_image_from_drive_link(drive_link)
        if not local_image:
            print("⚠️ Skipping row due to image download failure.")
            time.sleep(5)
            continue

        number_plate = recognize_plate_from_image(local_image)
        update_flag_and_plate(service, row_index, number_plate)

        send_whatsapp_with_image_and_message(drive_link, WHATSAPP_GROUP_NAME, message_text, number_plate)

        print("⏳ Waiting 3 seconds before checking next...")
        time.sleep(3)


if __name__ == "__main__":
    main()
