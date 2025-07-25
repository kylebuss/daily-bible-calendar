import time
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

# openbible.com book number, stepbible.org book abbreviation, openbible.com book abbreviation
BOOK_MAPPINGS = {
    "genesis": [1, "Gen", "Gen"],
    "exodus": [2, "Exod", "Exo"],
    "leviticus": [3, "Lev", "Lev"],
    "numbers": [4, "Num", "Num"],
    "deuteronomy": [5, "Deut", "Deu"],
    "joshua": [6, "Josh", "Jos"],
    "judges": [7, "Judg", "Jdg"],
    "ruth": [8, "Ruth", "Rut"],
    "1samuel": [9,"1Sam", "1Sa"],
    "2samuel": [10,"2Sam", "2Sa"],
    "1kings": [11,"1Kngs", "1Ki"],
    "2kings": [12,"2Kngs", "2Ki"],
    "1chronicles": [13,"1Chr", "1Ch"],
    "2chronicles": [14,"2Chr", "2Ch"],
    "ezra": [15,"Ezra", "Ezr"],
    "nehemiah": [16,"Neh", "Neh"],
    "esther": [17,"Esth", "Est"],
    "job": [18,"Job", "Job"],
    "psalms": [19,"Psalm", "Psa"],
    "proverbs": [20,"Prov", "Pro"],
    "ecclesiastes": [21,"Eccl", "Ecc"],
    "songofsongs": [22,"Song", "Sng"],
    "isaiah": [23,"Isa", "Isa"],
    "jeremiah": [24,"Jer", "Jer"],
    "lamentations": [25,"Lam", "Lam"],
    "ezekiel": [26,"Ezek", "Ezk"],
    "daniel": [27,"Dan", "Dan"],
    "hosea": [28,"Hosea", "Hos"],
    "joel": [29,"Joel", "Jol"],
    "amos": [30,"Amos", "Amo"],
    "obadiah": [31,"Obad", "Oba"],
    "jonah": [32,"Jonah", "Jon"],
    "micah": [33,"Micah", "Mic"],
    "nahum": [34,"Nahum", "Nam"],
    "habakkuk": [35,"Hab", "Hab"],
    "zephaniah": [36,"Zeph", "Zep"],
    "haggai": [37,"Hag", "Hag"],
    "zechariah": [38,"Zech", "Zec"],
    "malachi": [39,"Mal", "Mal"],
    "matthew": [40,"Matt", "Mat"],
    "mark": [41,"Mark", "Mrk"],
    "luke": [42,"Luke", "Luk"],
    "john": [43,"John", "Jhn"],
    "acts": [44,"Acts", "Act"],
    "romans": [45,"Rom", "Rom"],
    "1corinthians": [46,"1Cor", "1Co"],
    "2corinthians": [47,"2Cor", "2Co"],
    "galatians": [48,"Gal", "Gal"],
    "ephesians": [49,"Eph", "Eph"],
    "philippians": [50,"Phil", "Php"],
    "colossians": [51,"Col", "Col"],
    "1thessalonians": [52,"1Thess", "1Th"],
    "2thessalonians": [53,"2Thess", "2Th"],
    "1timothy": [54,"1Tim", "1Ti"],
    "2timothy": [55,"2Tim", "2Ti"],
    "titus": [56,"Titus", "Tts"],
    "philemon": [57,"Phlm", "Phm"],
    "hebrews": [58,"Heb", "Heb"],
    "james": [59,"James", "Jas"],
    "1peter": [60,"1Pet", "1Pe"],
    "2peter": [61,"2Pet", "2Pe"],
    "1john": [62,"1John", "1Jn"],
    "2john": [63,"2John", "2Jn"],
    "3john": [64,"3John", "3Jn"],
    "jude": [65,"Jude", "Jud"],
    "revelation": [66,"Rev", "Rev"],
}

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
            flow = InstalledAppFlow.from_client_secrets_file("daily-bible-calendar\secret.json", SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for next run
        with open("token.pickle", "wb") as token:
            pickle.dump(creds, token)

    return build("calendar", "v3", credentials=creds)

# Formats a Book/Chapter(s) link for stepbible.org
def format_bible_link(book, sc, ec):
    book_clean = book.strip().replace(" ", "")
    book_formatted = BOOK_MAPPINGS.get(book_clean.lower())

    if(ec != sc):
        return (
            f"Step Bible link for {book} {sc}-{ec}: https://www.stepbible.org/?q=version=BSB@reference={book_formatted[1]}.{sc}-{ec}&options=VHNUG\n\n"
        )
    else:
        return (
            f"Step Bible link for {book} {sc}: https://www.stepbible.org/?q=version=BSB@reference={book_formatted[1]}.{sc}\n\n"
        )

# Formats a Book/Chapter link for openbible.com/audio/souer/
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
    return (
        f"Audio link for {book} {chapter}: https://www.openbible.com/audio/souer/BSB_{book_number}_{book_formatted[2]}_{audio_chapter}.mp3\n"
    )

def parse_excel_file(file_path):
    """Parse Excel file and extract Bible reading data"""
    try:
        # Read Excel file
        df = pd.read_excel(file_path)

        # Expected columns: Date, Book, Chapter, Title (optional)
        required_columns = ["Date", "OT", "OTSC", "OTEC", "NT", "NTSC", "NTEC", "PS", "PR"]

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


def create_calendar_event(service, event_data, day, calendar_id="primary"):
    """Create a single calendar event"""

    date = event_data["Date"]
    ot_book = event_data["OT"]
    ot_starting_chapter = event_data["OTSC"]
    ot_ending_chapter = event_data["OTEC"]
    nt_book = event_data["NT"]
    nt_starting_chapter = event_data["NTSC"]
    nt_ending_chapter = event_data["NTEC"]
    psalm = event_data["PS"]
    proverb = event_data["PR"]

    # Create StepBible link
    step_bible_ot_link = ""
    print(f"Processing OT: {ot_book} {ot_starting_chapter}-{ot_ending_chapter}")
    if ot_book != "None" and ot_starting_chapter != 0:
        step_bible_ot_link = format_bible_link(ot_book, ot_starting_chapter, ot_ending_chapter)
    
    step_bible_nt_link = ""
    print(f"Processing NT: {nt_book} {nt_starting_chapter}-{nt_ending_chapter}")
    if nt_book != "None" and nt_starting_chapter != 0:
        step_bible_nt_link = format_bible_link(nt_book, nt_starting_chapter, nt_ending_chapter)

    psalm_or_proverb_link = ""
    if psalm != 0:
        psalm_or_proverb_link = format_bible_link("Psalms", psalm, psalm)
    elif proverb != 0:
        psalm_or_proverb_link = format_bible_link("Proverbs", proverb, proverb)

    # Create audio links
    ot_audio_links = ""
    nt_audio_links = ""
    if ot_book != "None" and ot_ending_chapter > ot_starting_chapter:
        for chapter in range(ot_starting_chapter, ot_ending_chapter + 1):
            ot_audio_links += format_audio_link(ot_book, chapter)
    elif ot_book != "None" and ot_starting_chapter != 0:
        ot_audio_links = format_audio_link(ot_book, ot_starting_chapter)
    if nt_book != "None" and nt_ending_chapter > nt_starting_chapter:
        for chapter in range(nt_starting_chapter, nt_ending_chapter + 1):
            nt_audio_links += format_audio_link(nt_book, chapter)
    elif nt_book != "None" and nt_starting_chapter != 0:
        nt_audio_links = format_audio_link(nt_book, nt_starting_chapter)
    
    psalm_or_proverb_audio_link = ""
    if psalm != 0:
        print(f"Processing Psalm: {psalm}")
        psalm_or_proverb_audio_link = format_audio_link("Psalms", psalm)
    elif proverb != 0:
        print(f"Processing Proverb: {proverb}")
        psalm_or_proverb_audio_link = format_audio_link("Proverbs", proverb)

    # Avoid hitting API rate limits     
    time.sleep(0.5)  
    
    # Format event
    event = {
        "summary": f"Bible Reading Day {day}",
        "description": f"{step_bible_ot_link}{ot_audio_links}\n{step_bible_nt_link}{nt_audio_links}\n{psalm_or_proverb_link}{psalm_or_proverb_audio_link}",
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
    # Configuration
    excel_file_path = "daily-bible-calendar\\2025 Bible Reading Plan.xlsx"  # Update with your file path
    calendar_id = "primary"  # Use 'primary' for main calendar or specific calendar ID

    print("Starting Bible Reading Calendar Creator...")

    # Parse Excel file
    print(f"Reading Excel file: {excel_file_path}")
    df = parse_excel_file(excel_file_path)

    if df is None:
        print("Failed to parse Excel file. Please check the file format.")
        return
    print(df.head())
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
