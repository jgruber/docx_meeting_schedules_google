# docx_meeting_schedules_google
Scrape Meeting Schedules and Render docx Schedules That Look Descent on Google Docs

This is a simple python script to scrape songs, counsel point studies, OCLM schedules, and Watchtower schedules from wol.jw.org and render `docx` schedule files which can be upload and edited on Google Docs.

## Install requirements

```
python3 -m venv .venv
. .venv/bin/activate
(.venv) pip3 install -r requirements.txt
```

## Usage

## Scraped Songs, Counsel Points, and Public Talk Titles

This repo comes with WOL scraped dictionaries for meeting songs, counsel points,
and public talk titles. You can update them from the verions in this repository.

### Sing Out Joyfully Songs Dictionary

```
python3
>>> import scrape
>>> scrape.scrape_songs()
```

This will overwrite the included `songs.json` file.

### Apply Yourself to Reading and Teaching Counsel Points Dictionary

```
python3
>>> import scrape
>>> scrape.scrape_scrape_counsel_points()
```

This will overwrite the included `studies.json` file.

### Public Talk Titles Dictionary

You will need to download the `S-99_E.docx` file which contains
all the public talk titles in numeric order from `jw.org`. Place
this file in the top repository directory and then scrape that file.

```
python3
>>> import scrape
>>> scrape.scrape_public_talk_titles()
```

This will overwrite the included `public_talks.json` file.

### Build Midweek Schedules for a Given Month

You will be building midweek meeting schedules by inputing the month name, year, the day of the week name for your meeting, and the 24hr format of your meeting
time.

The following example is for meetings in the month of `September` for year `2022` for a congregation which meets on `Wednesday` night at `19:10` (7:30 PM).

```
python3
>>> import scrape
>>> scrape.build_midweek_schedule_doc('September', 2022, 'Wednesday', '19:30')
```

### Build Weekend / Weeklong Schedules for a Given Month

You will be building midweek meeting schedules by inputing the month name, year, the day of the week name for your meeting, and the 24hr format of your meeting
time.

The following example is for meetings in the month of `September` for year `2022` for a congregation which meets on `Saturday` night at `10:30` (10:30 AM).


```
python3
>>> import scrape
>>> scrape.build_weeklong_schedule_doc('September', 2022, 'Saturday', '10:30')
```

## Preserving JSON of the WOL Scraped Data

Simply add a `True` argument (named argument is `use_cache`) to your scrape method. In example:

```
python3
>>> import scrape
>>> scrape.build_midweek_schedule_doc('September', 2022, 'Wednesday', '19:30', True)
```

## Editing Template docx Files

The `docx` template files and their associated rendering python modules are in the `templates` directory. 

All replaement variables from the scraped WOL data are in `{{ variable_name }}` mustashe format. 

Upload the template files to Google Docs, make alterations in the formatting, retaining the variable substitions you which, and then edit the redering python modules to insert the variables in the document.
