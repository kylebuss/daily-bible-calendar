import time
import pandas as pd
from datetime import datetime, timedelta
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pickle
import os
import re
from constants import SCOPES, BOOK_MAPPINGS

def authenticate_google_calendar():
    """Authenticate and return Google Calendar service object"""
    creds = None

    # Load existing credentials
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    # If no valid credentials, get new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # update file path to your secret.json file
            flow = InstalledAppFlow.from_client_secrets_file("daily-bible-calendar\secret.json", SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    return build("calendar", "v3", credentials=creds)

# Formats a Book/Chapter(s) link for Berean Standard Bible on stepbible.org
def format_bible_link(book, sc, ec):
    book_clean = book.strip().replace(" ", "")
    book_formatted = BOOK_MAPPINGS.get(book_clean.lower())

    # if a differnt Bible version is preferred, change the URL accordingly
    # Example: https://www.stepbible.org/?q=version=ESV@reference{book_formatted[1]}.{sc}-{ec}&options=VHNUG

    # if a different Bible app is preferred, change the URL accordingly
    # Example: https://www.bible.com/bible/59/{book_formatted[2]}.{sc}
    # NOTE: bible.com only supports single chapters, so if a range is provided, it will only link to the starting chapter unless you have updated the create_calendar_events method to handle ranges
    if(ec != sc):
        return (
            f"Step Bible link for {book} {sc}-{ec}: https://www.stepbible.org/?q=version=BSB@reference={book_formatted[1]}.{sc}-{ec}&options=VHNUG\n\n"
        )
    else:
        return (
            f"Step Bible link for {book} {sc}: https://www.stepbible.org/?q=version=BSB@reference={book_formatted[1]}.{sc}\n\n"
        )

# Formats a Book/Chapter link for Berean Standard Bible openbible.com/audio/souer/
def format_audio_link(book, chapter):
    book_clean = book.strip().replace(" ", "").lower()
    book_formatted = BOOK_MAPPINGS.get(book_clean.lower())

    if book_formatted is None:
        print(f"Book '{book}' not found in mappings.")
        return ""
    audio_chapter = chapter
    if audio_chapter < 10:
        audio_chapter = f"00{audio_chapter}"
    elif audio_chapter < 100:
        audio_chapter = f"0{audio_chapter}"
    else:
        audio_chapter = str(audio_chapter)

    book_number = book_formatted[0]
    if book_number < 10:
        book_number = f"0{book_number}"

    # if a different Bible version is preferred, change the URL accordingly
    # Example: https://www.bible.com/audio-bible/59/{book_formatted[2]}.{audio_chapter}.ESV
    # NOTE: book_formatted[2] is the openbible.com book abbreviation, you may have to add to the BOOK_MAPPINGS dictionary if you decide to use a different Bible version book_formatted[3] 
    return (
        f"Audio link for {book} {chapter}: https://www.openbible.com/audio/souer/BSB_{book_number}_{book_formatted[2]}_{audio_chapter}.mp3\n"
    )

def parse_excel_file(file_path):
    """Parse Excel file and extract Bible reading data"""
    try:
        # Read Excel file
        df = pd.read_excel(file_path)

        # Expected columns
        required_columns = ["Date","Book","SC","EC", "PS", "PR"]

        # Check if required columns exist
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(f"Missing required columns: {missing_columns}")
            print(f"Available columns: {list(df.columns)}")
            return None

        # Clean and validate data
        df = df.dropna(subset=required_columns)

        # Convert Date column to datetime
        df["Date"] = pd.to_datetime(df["Date"])

        return df

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


def create_calendar_event(service, event_data, day, calendar_id="Colum-Buss Family"):

    date = event_data["Date"]
    book = event_data["Book"]
    sc = int(event_data["SC"])
    ec = int(event_data["EC"])
    psalm = int(event_data["PS"])
    proverb = int(event_data["PR"])

    # Create StepBible link for main reading
    step_bible_link = ""
    print(f"Processing: {book} {sc}-{ec}")
    if sc != 0:
        step_bible_link = format_bible_link(book, sc, ec)

    # Create audio links for main reading
    audio_links = ""
    if ec > sc:
        for chapter in range(sc, ec + 1):
            audio_links += format_audio_link(book, chapter)
    elif sc != 0:
        audio_links = format_audio_link(book, sc)

    # Psalm/Proverb links
    psalm_or_proverb_link = ""
    psalm_or_proverb_audio_link = ""
    if psalm != 0:
        psalm_or_proverb_link = format_bible_link("Psalms", psalm, psalm)
        print(f"Processing Psalm: {psalm}")
        psalm_or_proverb_audio_link = format_audio_link("Psalms", psalm)
    elif proverb != 0:
        psalm_or_proverb_link = format_bible_link("Proverbs", proverb, proverb)
        print(f"Processing Proverb: {proverb}")
        psalm_or_proverb_audio_link = format_audio_link("Proverbs", proverb)

    # Avoid hitting API rate limits     
    time.sleep(1)  

    # Format calendar event
    event = {
        "summary": f"Bible Reading Day {day}",
        "description": f"{step_bible_link}{audio_links}\n{psalm_or_proverb_link}{psalm_or_proverb_audio_link}",
        "start": {
            "date": date.strftime("%Y-%m-%d"),
            "time": "06:00:00",  # Set a default time, adjust as needed
            "timeZone": "America/Chicago",  # Adjust timezone as needed
        },
        "end": {
            "date": date.strftime("%Y-%m-%d"),
            "time": "07:00:00",  # Set a default end time, adjust as needed
            "timeZone": "America/Chicago",  # Adjust timezone as needed
        },
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "popup", "minutes": 10},
            ],
        },
    }

    try:
        created_event = (
            service.events().insert(calendarId=calendar_id, body=event).execute()
        )
        print(f"Event created for {date.strftime('%Y-%m-%d')}")
        return created_event
    except Exception as e:
        print(f"Error creating event for {date.strftime('%Y-%m-%d')}: {e}")
        return None


def main():
    # update the file path to your Excel file
    excel_file_path = "daily-bible-calendar/2026_Bible_Reading_Plan.xlsx"

    # Use 'primary' for main calendar or specific calendar ID
    calendar_id = "Colum-Buss Family"  

    print("Starting Bible Reading Calendar Creator...")

    # Parse Excel file
    print(f"Reading Excel file: {excel_file_path}")
    df = parse_excel_file(excel_file_path)

    if df is None:
        print("Failed to parse Excel file. Please check the file format.")
        return
    
    if len(df) == 0:
        print("No reading entries found in the Excel file.")
        return
    
    print(f"Found {len(df)} reading entries")

    # Authenticate Google Calendar
    print("Authenticating with Google Calendar...")
    try:
        service = authenticate_google_calendar()
        print("Successfully authenticated!")
    except Exception as e:
        print(f"Authentication failed: {e}")
        print("Make sure you have:")
        print("1. Created a Google Cloud project")
        print("2. Enabled the Calendar API")
        print("3. Downloaded secret.json file")
        return

    # Create calendar events
    print("Creating calendar events...")
    successful_events = 0

    for index, row in df.iterrows():
        event_created = create_calendar_event(service, row, index + 1, calendar_id)
        if event_created:
            successful_events += 1

    print(
        f"\nCompleted! Successfully created {successful_events} out of {len(df)} events."
    )

if __name__ == "__main__":
    main()
