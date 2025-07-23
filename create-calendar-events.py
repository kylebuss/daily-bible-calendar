import pandas as pd
from datetime import datetime, timedelta
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import pickle
import os
import re

# Google Calendar API setup
SCOPES = ["https://www.googleapis.com/auth/calendar"]


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
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    return build("calendar", "v3", credentials=creds)


def format_bible_link(book, chapter):
    """Create StepBible.org link for given book and chapter"""
    # Clean book name and format for URL
    book_clean = book.strip().replace(" ", "")

    # Common book abbreviations mapping
    book_mappings = {
        "1samuel": "1Sam",
        "2samuel": "2Sam",
        "1kings": "1Kgs",
        "2kings": "2Kgs",
        "1chronicles": "1Chr",
        "2chronicles": "2Chr",
        "1corinthians": "1Cor",
        "2corinthians": "2Cor",
        "1thessalonians": "1Thess",
        "2thessalonians": "2Thess",
        "1timothy": "1Tim",
        "2timothy": "2Tim",
        "1peter": "1Pet",
        "2peter": "2Pet",
        "1john": "1John",
        "2john": "2John",
        "3john": "3John",
    }

    book_formatted = book_mappings.get(book_clean.lower(), book_clean)
    return (
        f"https://www.stepbible.org/?q=version=ESV|reference={book_formatted}.{chapter}"
    )


def parse_excel_file(file_path):
    """Parse Excel file and extract Bible reading data"""
    try:
        # Read Excel file
        df = pd.read_excel(file_path)

        # Expected columns: Date, Book, Chapter, Title (optional)
        required_columns = ["Date", "Book", "Chapter"]

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

        # Ensure Chapter is integer
        df["Chapter"] = df["Chapter"].astype(int)

        return df

    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


def create_calendar_event(service, event_data, calendar_id="primary"):
    """Create a single calendar event"""

    date = event_data["Date"]
    book = event_data["Book"]
    chapter = event_data["Chapter"]
    title = event_data.get("Title", f"{book} {chapter}")

    # Create StepBible link
    bible_link = format_bible_link(book, chapter)

    # Format event
    event = {
        "summary": f"Bible Reading: {title}",
        "description": f"Read {book} Chapter {chapter}\n\nStepBible Link: {bible_link}",
        "start": {
            "date": date.strftime("%Y-%m-%d"),
            "timeZone": "America/New_York",  # Adjust timezone as needed
        },
        "end": {
            "date": date.strftime("%Y-%m-%d"),
            "timeZone": "America/New_York",
        },
        "reminders": {
            "useDefault": False,
            "overrides": [
                {"method": "popup", "minutes": 30},
            ],
        },
    }

    try:
        created_event = (
            service.events().insert(calendarId=calendar_id, body=event).execute()
        )
        print(f"Event created: {title} on {date.strftime('%Y-%m-%d')}")
        return created_event
    except Exception as e:
        print(f"Error creating event for {title}: {e}")
        return None


def main():
    """Main function to process Excel file and create calendar events"""

    # Configuration
    excel_file_path = "bible_reading_plan.xlsx"  # Update with your file path
    calendar_id = "primary"  # Use 'primary' for main calendar or specific calendar ID

    print("Starting Bible Reading Calendar Creator...")

    # Parse Excel file
    print(f"Reading Excel file: {excel_file_path}")
    df = parse_excel_file(excel_file_path)

    if df is None:
        print("Failed to parse Excel file. Please check the file format.")
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
        print("3. Downloaded credentials.json file")
        return

    # Create calendar events
    print("Creating calendar events...")
    successful_events = 0

    for index, row in df.iterrows():
        event_created = create_calendar_event(service, row, calendar_id)
        if event_created:
            successful_events += 1

    print(
        f"\nCompleted! Successfully created {successful_events} out of {len(df)} events."
    )


# Example Excel file structure helper
def create_sample_excel():
    """Create a sample Excel file for testing"""
    sample_data = {
        "Date": ["2024-01-01", "2024-01-02", "2024-01-03", "2024-01-04", "2024-01-05"],
        "Book": ["Genesis", "Genesis", "Genesis", "Exodus", "Exodus"],
        "Chapter": [1, 2, 3, 1, 2],
        "Title": [
            "Creation",
            "Garden of Eden",
            "The Fall",
            "Moses Born",
            "Burning Bush",
        ],
    }

    df = pd.DataFrame(sample_data)
    df.to_excel("sample_bible_reading_plan.xlsx", index=False)
    print("Created sample_bible_reading_plan.xlsx")


if __name__ == "__main__":
    # Uncomment the line below to create a sample Excel file first
    # create_sample_excel()

    main()
