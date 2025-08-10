import time
import win32clipboard
import keyboard
import pyautogui

from google.oauth2 import service_account
from googleapiclient.discovery import build
import webbrowser

# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = "service_account.json" 
""
SHEET_ID = ""
SHEET_RANGE = ""
SCOPES = ["
"]

def get_first_pending_row(service):
    sheet = service.spreadsheets()
    result = sheet.values().get(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE
    ).execute()
    rows = result.get("values", [])
    for idx, row in enumerate(rows, start=2):
        while len(row) < 12:  # Ensure up to column L
            row.append("")
        if row[11].strip().upper() == "N":  # Check column L
            return idx, row
    return None, None

def update_flag(service, row_index):
    cell_range = f"Sheet2!L{row_index}"  # Correct: Column L
    body = {"values": [["Y"]]}
    service.spreadsheets().values().update(
        spreadsheetId=SHEET_ID,
        range=cell_range,
        valueInputOption="RAW",
        body=body
    ).execute()
    print(f"✅ Updated flag to Y at row {row_index} (column L)")

def send_whatsapp_with_image_and_message(drive_link, whatsapp_name, message_text):
    webbrowser.open(drive_link)
    time.sleep(8)
    keyboard.press_and_release('ctrl+c')  # Copy image from browser
    time.sleep(1)
    # Step 2: Open WhatsApp
    keyboard.press_and_release('win')
    time.sleep(1)
    keyboard.write('WhatsApp')
    time.sleep(1)
    keyboard.press_and_release('enter')
    time.sleep(5)
    # Step 3: Search and open chat
    keyboard.press_and_release('ctrl+f')
    time.sleep(1)
    keyboard.write(whatsapp_name)
    time.sleep(2)
    keyboard.press_and_release('enter')
    time.sleep(2)
    keyboard.press_and_release('down')
    time.sleep(2)
    keyboard.press_and_release('enter')
    time.sleep(2)
    # Step 4: Paste image
    keyboard.press_and_release('ctrl+v')
    time.sleep(2)
    # Step 5: Paste only the message from the cell (column G)
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardText(str(message_text))
    win32clipboard.CloseClipboard()
    print("✅ Copied to clipboard:", message_text)
    time.sleep(1)
    keyboard.press_and_release('ctrl+v')   # Paste the cell message
    time.sleep(1)
    keyboard.press_and_release('enter')
    print("✅ WhatsApp image and cell message sent!")

def main():
    
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    service = build("sheets", "v4", credentials=creds)
    
    while True:
        row_index, row = get_first_pending_row(service)
        if not row:
            print("No pending messages with Flag N found. Waiting 15 seconds before checking again...")
            time.sleep(15)
            continue
        
        drive_link = row[9]         # Column J
        message_text = row[10] 
        whatsapp_name = "Security Alerts"    # <-- Take WhatsApp group/contact from the sheet

        print(f"➡️ Sending for: row {row_index} to WhatsApp: {whatsapp_name}")
        print("Message text:", message_text)

        send_whatsapp_with_image_and_message(drive_link, whatsapp_name, message_text)
        update_flag(service, row_index)
        print("⏳ Waiting 3 seconds before next check...\n")
        time.sleep(3)

if __name__ == "__main__":
    main()
    
