from datetime import datetime, timedelta
import math
from openpyxl import Workbook

# /c:/Source/prodigals-playground/daily-bible-calendar/create-excel-plan.py
# GitHub Copilot

import csv

BOOK_TOTAL_CHAPTERS = {
    'Genesis': 50, 'Exodus': 40, 'Leviticus': 27, 'Numbers': 36, 'Deuteronomy': 34,
    'Joshua': 24, 'Judges': 21, 'Ruth': 4, '1 Samuel': 31, '2 Samuel': 24,
    '1 Kings': 22, '2 Kings': 25, '1 Chronicles': 29, '2 Chronicles': 36,
    'Ezra': 10, 'Nehemiah': 13, 'Esther': 10, 'Job': 42, 'Psalms': 150,
    'Proverbs': 31, 'Ecclesiastes': 12, 'Song of Solomon': 8, 'Isaiah': 66,
    'Jeremiah': 52, 'Lamentations': 5, 'Ezekiel': 48, 'Daniel': 12,
    'Hosea': 14, 'Joel': 3, 'Amos': 9, 'Obadiah': 1, 'Jonah': 4,
    'Micah': 7, 'Nahum': 3, 'Habakkuk': 3, 'Zephaniah': 3, 'Haggai': 2,
    'Zechariah': 14, 'Malachi': 4, 'Matthew': 28, 'Mark': 16, 'Luke': 24,
    'John': 21, 'Acts': 28, 'Romans': 16, '1 Corinthians': 16, '2 Corinthians': 13,
    'Galatians': 6, 'Ephesians': 6, 'Philippians': 4, 'Colossians': 4,
    '1 Thessalonians': 5, '2 Thessalonians': 3, '1 Timothy': 6, '2 Timothy': 4,
    'Titus': 3, 'Philemon': 1, 'Hebrews': 13, 'James': 5, '1 Peter': 5,
    '2 Peter': 3, '1 John': 5, '2 John': 1, '3 John': 1, 'Jude': 1, 'Revelation': 22
}

def parse_chrono_csv(csv_path):
    # Returns list of (Book, [chapter numbers]) in chrono order
    books = []
    with open(csv_path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            book = row['Book'].strip()
            chapters = row['Chapters'].strip()
            if not chapters:
                # If no chapters, use all chapters for this book
                total = BOOK_TOTAL_CHAPTERS.get(book)
                if total:
                    chs = list(range(1, total+1))
                    books.append((book, chs))
                continue
            if '-' in chapters:
                # Range, e.g. 1-11
                parts = chapters.split('-')
                start = int(parts[0])
                end = int(parts[1])
                chs = list(range(start, end+1))
                books.append((book, chs))
            else:
                # Single or comma-separated, or a single number (means all chapters)
                try:
                    chs = [int(c) for c in chapters.split(',') if c.strip()]
                except Exception:
                    chs = []
                if len(chs) == 1 and chs[0] == BOOK_TOTAL_CHAPTERS.get(book, -1):
                    # If the only entry is the total, use all chapters
                    chs = list(range(1, chs[0]+1))
                elif not chs:
                    # fallback: all chapters
                    total = BOOK_TOTAL_CHAPTERS.get(book)
                    if total:
                        chs = list(range(1, total+1))
                books.append((book, chs))
    return books

def generate_daily_plan(books_chapters, days=365):
    # Flatten all chapters in chrono order: [(Book, chapter)]
    all_chapters = []
    for book, chs in books_chapters:
        for ch in chs:
            all_chapters.append((book, ch))
    daily = []
    idx = 0
    total = len(all_chapters)
    for d in range(days):
        # 3 chapters for 3 days, then 2 chapters every 4th day (2.75 avg)
        if (d % 4) == 3:
            num = 2
        else:
            num = 3
        day_chaps = all_chapters[idx:idx+num]
        daily.append(day_chaps)
        idx += num
        if idx >= total:
            break
    # If fewer days than requested, pad with empty
    while len(daily) < days:
        daily.append([])
    # If there are leftover chapters, add them to the last non-empty day
    if idx < total:
        for i in range(idx, total):
            # Add to last non-empty day
            for j in range(len(daily)-1, -1, -1):
                if daily[j]:
                    daily[j].append(all_chapters[i])
                    break
    return daily[:days]

def get_psalm_for_day(day):
    # Only for days 0-361, every day except every 6th day
    if day > 361:
        return 0
    if (day+1) % 6 == 0:
        return 0
    # Ensure Psalm 150 is included
    psalm = (day - (day // 6)) + 1
    if psalm > 150:
        return 0
    return psalm

def get_proverb_for_day(day):
    # Only for days 0-361, only every 6th day
    if day > 361:
        return 0
    if (day+1) % 6 == 0:
        # Count how many 6th days have occurred so far
        pr_day = (day + 1) // 6
        if pr_day > 31:
            return 0
        return pr_day
    return 0

def write_custom_excel_plan(csv_path, filename, start_date, days=365):
    books_chapters = parse_chrono_csv(csv_path)
    daily_plan = generate_daily_plan(books_chapters, days)
    wb = Workbook()
    ws = wb.active
    ws.append(["Date", "Book", "SC", "EC", "PS", "PR"])
    for i in range(days):
        date = start_date + timedelta(days=i)
        chaps = daily_plan[i]
        if chaps:
            # Group by book
            book = chaps[0][0]
            sc = chaps[0][1]
            ec = chaps[-1][1]
        else:
            book = ""
            sc = 0
            ec = 0
        ps = get_psalm_for_day(i)
        pr = get_proverb_for_day(i)
        ws.append([date.strftime("%Y-%m-%d"), book, sc, ec, ps, pr])
    wb.save(filename)

if __name__ == "__main__":
    # Example usage
    csv_path = "C:/Source/prodigals-playground/daily-bible-calendar/chrono.csv"
    filename = "C:/Source/prodigals-playground/daily-bible-calendar/2026_Bible_Reading_Plan.xlsx"
    start_date = datetime(2026, 1, 1)
    write_custom_excel_plan(csv_path, filename, start_date, days=365)