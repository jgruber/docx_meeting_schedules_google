# oclmscraper
Scrape OCLM schedules

This is a simple python script to scrape songs, counsel point studies, and OCLM schedules from wol.jw.org.

## Install requirements

```
python3 -m venv .venv
. .venv/bin/activate
(.venv) pip3 install -r requirements.txt
```

## Usage

### Priminary Task: Create a JSON dictionary of Song details

```
python3
>>> import scrape
>>> scrape.scrape_songs()
```

### Priminary Task: Create a JSON dictionary of Counsel Point details

```
python3
>>> import scrape
>>> scrape.scrape_scrape_counsel_points()
```

### Retrieve OCLM Schedule for January, March, May, July, September, Novemeber

The method `build_parts_dictionary(month, year, meeting_day_of_week='Wednesday', meeting_24hr_start_time='19:30')`

Takes arguments:

`month`: The month name for the 2 month schedule. Should be `January`, `March`, `May`, `July`, `Septemeber` or `November`.

`year`: Integer year to scrape. Example: `2022`.

`meeting_day_of_week`: The day of the week of your meeting. Should be `Monday`, `Tuesday`, `Wednesday`, `Thursday`, `Friday`, `Saturday`, or `Sunday`.

`meeting_24hr_start_time`: This is the `HH:mm` your meeting begins in 24hr format. Example for 7:30 PM meeting it would be `19:30`.


```
python3
>>> import scrape
>>> scrape.build_meeting_parts('January', 2022, 'Wednesday', '19:30')
```
