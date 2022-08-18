import json
from locale import DAY_1
import re
import requests
import unicodedata
import datetime

from urllib.error import HTTPError
from bs4 import BeautifulSoup

BASE_URL = "https://wol.jw.org"

OCLM_PARTICIPANT_REQUIREMENTS = {
    'Bible Reading': 1,
    'Initial Call Video': 0,
    'Initial Call': 2,
    'Return Visit Video': 0,
    'Return Visit': 2,
    'Bible Study': 2
}

MONTHS = {
    'January': '01',
    'February': '02',
    'March': '03',
    'April': '04',
    'May': '05',
    'June': '06',
    'July': '07',
    'August': '08',
    'September': '09',
    'October': '10',
    'November': '11',
    'December': '12'
}

DAY_OF_WEEK = {
    'monday': 0,
    'tuesday': 1,
    'wednesday': 2,
    'thursday': 3,
    'friday': 4,
    'saturday': 5,
    'sunday': 6
}

SONGS_ID_RANGE = [
    1102016801,
    1102016952
]

COUNSEL_POINT_ID_RANGE = [
    1102018441,
    1102018460
]

def _clean_unicode_to_ascii(description):
    return description.replace(
        '\u200b','').replace(
        '\u201c',"'").replace(
        '\u201d',"'").replace(
        '\u2013',"'").replace(
        '\u2019',"'").replace(
        '\u2014',"-").replace(
        '\u00b6',"p")

def scrape_songs():
    songs = {}
    for i in range(SONGS_ID_RANGE[0], SONGS_ID_RANGE[1]):
        song_number = i - (SONGS_ID_RANGE[0] - 1)
        url = f"https://wol.jw.org/en/wol/d/r1/lp-e/{i}"
        print(f"getting song number: {song_number} from {url}")
        song = requests.get(url)
        soup = BeautifulSoup(song.content, 'html.parser')
        song_index = unicodedata.normalize('NFKD', soup.find(id="p1").find("strong").text).lower()
        song_title = _clean_unicode_to_ascii(unicodedata.normalize('NFKD', soup.find(id="p2").find("strong").text))
        ref_scripture = unicodedata.normalize('NFKD', soup.find(id="p3").find("a").text)
        image_rel_url = soup.find_all("div", {"class": "thumbnail"})[0].find("img")['src']
        image_url = f"https://wol.jw.org{image_rel_url}"
        songs[song_index] = {
            'index': song_index,
            'number': song_number,
            'title': song_title,
            'ref_scripture': ref_scripture,
            'url': url,
            'image_url': image_url
        }
        print(f"\t adding: {songs[song_index]}")
    with open('songs.json', 'w') as song_json_file:
        song_json_file.write(json.dumps(songs, indent=4, sort_keys=False))
    return songs


def scrape_counsel_points():
    studies = {}
    for i in range(COUNSEL_POINT_ID_RANGE[0], COUNSEL_POINT_ID_RANGE[1]):
        study_number = i - (COUNSEL_POINT_ID_RANGE[0] - 1)
        url = f"https://wol.jw.org/en/wol/d/r1/lp-e/{i}"
        print(f"getting counsel study: {study_number} from {url}")
        study = requests.get(url)
        soup = BeautifulSoup(study.content, 'html.parser')
        study_index = unicodedata.normalize('NFKD', soup.find(id="p1").find("strong").text).lower()
        study_title = _clean_unicode_to_ascii(unicodedata.normalize('NFKD', soup.find(id="p2").find("strong").text))
        theme_scripture = unicodedata.normalize('NFKD', soup.find(id="p3").find("strong").text)
        studies[study_index] = {
            'index': study_index,
            'number': study_number,
            'title': study_title,
            'theme_scripture': theme_scripture,
            'url': url
        }
        print(f"\t adding: {studies[study_index]}")
    with open('studies.json', 'w') as study_json_file:
        study_json_file.write(json.dumps(studies, indent=4, sort_keys=False))
    return studies


def _get_weekly_schedule_links_en(month, year):
    wburl = f"{BASE_URL}/en/wol/library/r1/lp-e/all-publications/meeting-workbooks/life-and-ministry-meeting-workbook-{year}/{month.lower()}"
    page = requests.get(wburl)
    soup = BeautifulSoup(page.content, 'html.parser')
    months_regex = re.compile("(January [123456789][123456789]?-|February [123456789][123456789]?-|March [123456789][123456789]?-|April [123456789][123456789]?-|May [123456789][123456789]?-|June [123456789][123456789]?-|July [123456789][123456789]?-|August [123456789][123456789]?-|September [123456789][123456789]?-|October [123456789][123456789]?-|November [123456789][123456789]?-|December [123456789][123456789]?-)")
    weekly_schedule_links: list = []
    for a in soup.find_all("a"):
        if hasattr(a, 'href'):
            if months_regex.search(str(a)):
                weekly_schedule_links.append(f"{BASE_URL}{a.get('href')}")
    return weekly_schedule_links


def _get_meeting_datetime(scraped_week_string, year, meeting_day_of_week='Wednesday', meeting_24hr_start_time='19:30'):
    date_split = scraped_week_string.split(' ')
    month = MONTHS[date_split[0]]
    monday_day = date_split[1].split('-')[0]
    # WOL formats split months with a long dash, not a short dash they
    # use when the week is all in the same month.
    if date_split[1].find('–') > 0:
        monday_day = date_split[1].split('–')[0]
    if int(monday_day) < 10:
        monday_day = f"0{monday_day}"
    meeting_datetime = datetime.datetime.fromisoformat(f"{year}-{month}-{monday_day} {meeting_24hr_start_time}:00")
    meeting_datetime = meeting_datetime + datetime.timedelta(days=DAY_OF_WEEK[meeting_day_of_week.lower()])
    return meeting_datetime


def _get_meeting_soup(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return BeautifulSoup(response.content, 'html.parser')
    except HTTPError as he:
        print(f"could not fetch schedule for {url}: {he}")
        return None


def _get_song_details(index):
    songs_dict = {}
    with open('songs.json', 'r') as songs_json:
       songs_dict = json.load(songs_json)
    if index.lower() in songs_dict:
        return songs_dict[index.lower()]
    return {}


def _get_study_detail(index):
    studies_dict = {}
    with open('studies.json', 'r') as studies_json:
       studies_dict = json.load(studies_json)
    if index.lower() in studies_dict:
        return studies_dict[index.lower()]
    return {}


def build_parts_dictionary(month, year, meeting_day_of_week='Wednesday', meeting_24hr_start_time='19:30'):
    monthly_schedule = []
    print(f"retrieving weekly schedules from wol.jw.org for workbooks for {month} - {year}")
    for url in _get_weekly_schedule_links_en(month, year):
        try:
            soup = _get_meeting_soup(url)
            weekof = unicodedata.normalize('NFKD', soup.find(id="p1").find("strong").text)
            meeting_datetime =  _get_meeting_datetime(weekof, year, meeting_day_of_week, meeting_24hr_start_time)
            schedule = {
                'weekof': weekof,
                'url': url,
                'meeting_datetime': meeting_datetime.isoformat(),
                'meeting_date': meeting_datetime.strftime('%B %-d'),
                'bible_reading': unicodedata.normalize('NFKD', soup.find(id="p2").find("strong").text),
                'chairman': None,
                'second_school_chairman': None,
                'opening_song': {
                    'details': _get_song_details(unicodedata.normalize('NFKD', soup.find(id="p3").find("strong").text)),
                    'start': meeting_datetime.strftime('%-I:%M'),
                    'end': (meeting_datetime + datetime.timedelta(minutes=5)).strftime('%-I:%M')
                },
                'opening_prayer': None,
                'opening_comments': {
                    'duration_min': 1,
                    'start':  (meeting_datetime + datetime.timedelta(minutes=5)).strftime('%-I:%M'),
                    'end': (meeting_datetime + datetime.timedelta(minutes=6)).strftime('%-I:%M')
                },
                'treasures_from_gods_word': {
                    'parts': []
                },
                'apply_yourself_to_the_field_ministry': {
                    'parts': []
                },
                'middle_song': {},
                'living_as_christians': {
                    'parts': []
                },
                'concluding_comments': {
                    'duration_min': 3,
                    'start': (meeting_datetime + datetime.timedelta(minutes=95)).strftime('%-I:%M'),
                    'end': (meeting_datetime + datetime.timedelta(minutes=98)).strftime('%-I:%M')
                },
                'closing_song': {},
                'closing_prayer': None
            }
            schedule['treasures_from_gods_word']['parts'].append(
                {
                    'type': 'treasures',
                    'assigned': None,
                    'duration_min': 10,
                    'start': (meeting_datetime + datetime.timedelta(minutes=6)).strftime('%-I:%M'),
                    'stop': (meeting_datetime + datetime.timedelta(minutes=16)).strftime('%-I:%M'),
                    'theme': _clean_unicode_to_ascii(unicodedata.normalize('NFKD', soup.find(id="p6").find("a").find("strong").text))
                }
            )
            schedule['treasures_from_gods_word']['parts'].append(
                {
                    'type': 'digging_for_spiritual_gems',
                    'assigned': None,
                    'duration_min': 10,
                    'start': (meeting_datetime + datetime.timedelta(minutes=16)).strftime('%-I:%M'),
                    'stop': (meeting_datetime + datetime.timedelta(minutes=26)).strftime('%-I:%M'),
                }
            )
            schedule['treasures_from_gods_word']['parts'].append(
                {
                    'type': 'bible_reading',
                    'main_hall_assigned': None,
                    'second_school_assigned': None,
                    'duration_min': 4,
                    'start': (meeting_datetime + datetime.timedelta(minutes=26)).strftime('%-I:%M'),
                    'stop': (meeting_datetime + datetime.timedelta(minutes=30)).strftime('%-I:%M'),
                    'reading': _clean_unicode_to_ascii(unicodedata.normalize('NFKD', soup.find(id="p10").find("a").text)),
                    'counsel_point': _get_study_detail(unicodedata.normalize('NFKD', soup.find(id="p10").find_all("a")[-1].text.replace('th ','')))
                }
            )
            first_apply_part = 12
            last_apply_part = 14
            middle_song_part = 16
            first_living_part = 17
            last_living_part = 18
            cbs_part = 19
            closing_song_part = 21
            for p in range(11,23):
                part_found = soup.find(id=f"p{p}")
                if part_found:
                    part_text = unicodedata.normalize('NFKD', part_found.text)
                    if part_text and part_text == "APPLY YOURSELF TO THE FIELD MINISTRY":
                        first_apply_part = p + 1
                    if part_text and part_text == "LIVING AS CHRISTIANS":
                        last_apply_part = p - 1
                        middle_song_part = p + 1
                        first_living_part = p + 2
                    if part_text and part_text.find("Concluding Comments") > -1:
                        last_living_part = p - 2
                        cbs_part = p - 1
                        closing_song_part = p + 1
            schedule['middle_song'] = {
                'details': _get_song_details(unicodedata.normalize('NFKD', soup.find(id=f"p{middle_song_part}").find("strong").text)),
                'start': (meeting_datetime + datetime.timedelta(minutes=45)).strftime('%-I:%M'),
                'end': (meeting_datetime + datetime.timedelta(minutes=50)).strftime('%-I:%M')
            }
            schedule['closing_song'] = {
                'details': _get_song_details(unicodedata.normalize('NFKD', soup.find(id=f"p{closing_song_part}").find("strong").text)),
                'start': (meeting_datetime + datetime.timedelta(minutes=98)).strftime('%-I:%M'),
                'end': (meeting_datetime + datetime.timedelta(minutes=115)).strftime('%-I:%M')
            }
            part_start_datetime = None
            part_stop_datetime = None
            previous_part_has_counsel = False
            for p in range(first_apply_part, last_apply_part + 1):
                part_type = unicodedata.normalize('NFKD', soup.find(id=f"p{p}").find("strong").text)[0:-1]
                part_description = _clean_unicode_to_ascii(unicodedata.normalize('NFKD', soup.find(id=f"p{p}").text))
                min_match = re.search('[123456789][0123456789]? min.', part_description)
                part_mins = int(part_description[min_match.span()[0]:min_match.span()[1]-5])
                counsel_point_details = None
                counsel_point_found = re.search('th study', part_description)
                if counsel_point_found:
                    counsel_point_details = _get_study_detail(unicodedata.normalize('NFKD', soup.find(id=f"p{p}").find_all("a")[-1].text.replace('th ','')))
                if p == first_apply_part:
                    part_start_datetime = meeting_datetime + datetime.timedelta(minutes=30)
                else:
                    if previous_part_has_counsel:
                        part_start_datetime = part_stop_datetime + datetime.timedelta(minutes=1)
                    else:
                        part_start_datetime = part_stop_datetime
                part_stop_datetime = part_start_datetime + datetime.timedelta(minutes=part_mins)
                if counsel_point_details:
                    previous_part_has_counsel = True
                else:
                    previous_part_has_counsel = False
                    counsel_point_details = {}
                schedule['apply_yourself_to_the_field_ministry']['parts'].append(
                    {
                        'main_hall_assigned': None,
                        'second_school_assigned': None,
                        'type': part_type,
                        'duration_min': part_mins,
                        'start': part_start_datetime.strftime('%-I:%M'),
                        'stop': part_stop_datetime.strftime('%-I:%M'),
                        'description': part_description,
                        'counsel_point': counsel_point_details
                    }
                )
            part_start_datetime = None
            part_stop_datetime = None
            for p in range(first_living_part, last_living_part + 1):
                part_description = _clean_unicode_to_ascii(unicodedata.normalize('NFKD', soup.find(id=f"p{p}").text))
                min_match = re.search('[123456789][0123456789]? min.', part_description)
                part_mins = int(part_description[min_match.span()[0]:min_match.span()[1]-5])
                if p == first_living_part:
                    part_start_datetime = meeting_datetime + datetime.timedelta(minutes=50)
                part_stop_datetime = part_start_datetime + datetime.timedelta(minutes=part_mins)
                schedule['living_as_christians']['parts'].append(
                    {
                        'type': 'part',
                        'assigned': None,
                        'duration_min': part_mins,
                        'start': part_start_datetime.strftime('%-I:%M'),
                        'stop': part_stop_datetime.strftime('%-I:%M'),
                        'description': part_description
                    }
                )
            schedule['living_as_christians']['parts'].append(
                {
                    'type': 'congregation_bible_study',
                    'conductor': None,
                    'reader': None,
                    'duration_min': 30,
                    'start': (meeting_datetime + datetime.timedelta(minutes=65)).strftime('%-I:%M'),
                    'stop': (meeting_datetime + datetime.timedelta(minutes=95)).strftime('%-I:%M'),
                    'description': _clean_unicode_to_ascii(unicodedata.normalize('NFKD', soup.find(id=f"p{cbs_part}").text))
                }
            )
            monthly_schedule.append(schedule)
        except HTTPError as he:
            print(f"could not fetch schedule for {url}: {he}")
    with open(f"{month}.json", "w") as schedule_json_file:
        schedule_json_file.write(json.dumps(monthly_schedule, indent=4, sort_keys=False))
    return monthly_schedule
