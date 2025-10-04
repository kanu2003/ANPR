***Project Overview

This Python script orchestrates a powerful automation pipeline designed to process real-time events, recognize vehicle number plates from attached images, update a central database (Google Sheet), and immediately dispatch alerts via WhatsApp.

It is ideal for implementing a low-cost, automated tracking or notification system for maintenance workflows, security checks, or logistics monitoring.

*** Architecture and Workflow

The system operates on a loop, linking three distinct services via a centralized control script:

Google Sheets (Input & Output):

Input: Monitored continuously for new records where the processing flag (Column L) is set to 'N' (No). It extracts the Drive Image Link (Column J) and Alert Message (Column K).

Output: Updates the processing flag to 'Y' (Yes) and writes the detected Number Plate to Column M.

External ANPR API:

Downloads the image from the Google Drive link.

Sends the image to the Plate Recognizer API for Automatic Number Plate Recognition (ANPR).

Returns the detected plate text with a confidence score.

Desktop Automation (WhatsApp):

Uses keyboard, pyautogui, and webbrowser to automate the desktop workflow:

Opens the image link to copy the image to the clipboard.

Opens the local WhatsApp application.

Searches for the configured group name.

Pastes the image and sends a final alert message including the detected number plate.
