import logging

logging.basicConfig(
    level=logging.NOTSET,
    filemode="w",
    filename="calendar2ics.log",
    format="%(asctime)s - %(levelname)s - %(name)s - %(filename)s - %(funcName)s - %(lineno)d - %(message)s",
)

import re
import datetime
import pdf2docx
import argparse
from pathlib import Path
from docx.api import Document
from ics import Calendar, Event


def process_split_months(day_num, month):
    if "/" in month:
        logging.debug(f"Processing ambiguitiy in month name: {month}")
        month = month.split("/")
    day_num = int(day_num)
    if isinstance(month, list):
        if day_num <= 5:
            month = month[1]
        else:
            month = month[0]
        logging.debug(f"Processed month name: {month}")
    else:
        month = month
    return month


def convert_pdf_to_docx(calendar_pdf_filepath, calendar_docx_filepath):
    try:
        logging.info(
            f"Attempting to convert {calendar_pdf_filepath} to {calendar_docx_filepath}"
        )
        pdf_file = pdf2docx.Converter(calendar_pdf_filepath)
        pdf_file.convert(calendar_docx_filepath)
        pdf_file.close()
        logging.info(
            f"Successfully converted {calendar_pdf_filepath} to {calendar_docx_filepath}"
        )
    except Exception as e:
        logging.error(
            f"Failed to convert {calendar_pdf_filepath} to {calendar_docx_filepath}:\n{e}"
        )
        exit(0)


def load_calendar(calendar_docx_filepath):
    logging.info("Loading Calendar Events for .ics export")
    calendar_data = dict()
    month_names = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Sept",
        "Oct",
        "Nov",
        "Dec",
    ]
    document = Document(calendar_docx_filepath)
    table = document.tables[0]

    calendar_data = dict()
    weeks = list()
    new_year = datetime.datetime.now().year

    logging.debug("Processing all rows to identify possible weeks")
    for row in table.rows:
        text = list(cell.text for cell in row.cells)
        text = list(map(str.strip, text))
        text[1] = text[1].replace("Sept", "Sep")
        if text[1].split("/")[0] in month_names:
            weeks.append(text)
        elif " " in text[1] and text[1].split(" ")[0] in month_names:
            weeks.append(text)
            if text[1].split(" ")[0] == "Jan":
                new_year = int(text[1].split(" ")[1])

    first_date = weeks[0][2]
    first_date_month = process_split_months(first_date, weeks[0][1])
    last_date = weeks[-1][-4]
    last_date_month = process_split_months(last_date, weeks[-1][1])

    logging.debug(f"First Date in Calendar: {first_date_month} {first_date}")
    logging.debug(f"Last Date in Calendar: {last_date_month} {last_date}")

    current_date_obj = datetime.datetime.strptime(
        f"{first_date.zfill(2)} {first_date_month} 2022", r"%d %b %Y"
    ).date()
    last_date_obj = datetime.datetime.strptime(
        f"{last_date.zfill(2)} {last_date_month} {new_year}", r"%d %b %Y"
    ).date()

    while current_date_obj <= last_date_obj:
        calendar_data[current_date_obj] = list()
        current_date_obj = current_date_obj + datetime.timedelta(days=1)

    calendar_events = list()
    current_year = datetime.datetime.now().year

    logging.info("Processing weeks to identify calendar events")
    for week in weeks:
        month = week[1]
        events = week[-1].split("\n")
        days = list()
        for d in week[2:9]:
            if d not in days:
                days.append(d)
        day_nums = list(map(lambda x: int(x[:2]), days))

        logging.debug(f"Days in week: {days}")
        logging.debug(f"Month: {month}")
        logging.debug(f"Events: {events}")

        covered_dates = list()
        logging.info("Adding events to calendar")
        for event in events:
            temp_event = re.split(r"[\-\–]", event)
            temp_event = list(map(str.strip, temp_event))

            # Event name does not contain a date -- event spans entire week
            if not any([not x or x[0].isdigit() for x in temp_event]):
                event_name = " - ".join(temp_event)
                week_start_day = day_nums[0]
                week_end_day = day_nums[-1]
                if " " in month and month.split()[0] == "Jan":
                    week_start_month = month.split()[0]
                    week_start_year = datetime.datetime.now().year + 1
                    current_year = week_start_year
                else:
                    week_start_month = process_split_months(week_start_day, month)
                    week_start_year = current_year

                if " " in month and month.split()[0] == "Jan":
                    week_end_month = month.split()[0]
                    week_end_year = datetime.datetime.now().year + 1
                    current_year = week_end_year
                else:
                    week_end_month = process_split_months(week_end_day, month)
                    week_end_year = current_year

                logging.debug(f"Event Name: {event_name}")
                logging.debug(
                    f"Week start date: {week_start_day} {week_start_month} {week_start_year}"
                )
                logging.debug(
                    f"Week end date: {week_end_day} {week_end_month} {week_end_year}"
                )

                calendar_event = Event(
                    name=event_name,
                    begin=datetime.datetime.strptime(
                        f"{day_nums[0]} {week_start_month} {week_start_year}",
                        r"%d %b %Y",
                    ).date(),
                    end=datetime.datetime.strptime(
                        f"{day_nums[-1]} {week_end_month} {week_end_year}",
                        r"%d %b %Y",
                    ).date(),
                )
                calendar_event.make_all_day()
                calendar_events.append(calendar_event)

                start_date = datetime.datetime.strptime(
                    f"{day_nums[0]} {week_start_month} {week_start_year}",
                    r"%d %b %Y",
                ).date()
                end_date = datetime.datetime.strptime(
                    f"{day_nums[-1]} {week_end_month} {week_end_year}",
                    r"%d %b %Y",
                ).date()
                covered_dates.extend(
                    [
                        start_date + datetime.timedelta(days=x)
                        for x in range(0, (end_date - start_date).days + 1)
                    ]
                )

            # check first two elements start with a number
            elif all([x and x[0].isdigit() for x in temp_event[:2]]):
                event_name = temp_event[2:]
                event_name = list(map(str.strip, event_name))
                event_name = " - ".join(event_name)
                event_start_day = int(temp_event[0][:-2])
                event_end_day = int(temp_event[1][:-2])

                if " " in month and month.split()[0] == "Jan":
                    event_start_month = month.split()[0]
                    event_start_year = datetime.datetime.now().year + 1
                    current_year = event_start_year
                else:
                    event_start_month = process_split_months(event_start_day, month)
                    event_start_year = current_year

                if " " in month and month.split()[0] == "Jan":
                    event_end_month = month.split()[0]
                    event_end_year = datetime.datetime.now().year + 1
                    current_year = event_end_year
                else:
                    event_end_month = process_split_months(event_end_day, month)
                    event_end_year = current_year

                logging.debug(f"Event Name: {event_name}")
                logging.debug(
                    f"Event start date: {event_start_day} {event_start_month} {event_start_year}"
                )
                logging.debug(
                    f"Event end date: {event_end_day} {event_end_month} {event_end_year}"
                )

                calendar_event = Event(
                    name=event_name,
                    begin=datetime.datetime.strptime(
                        f"{event_start_day} {event_start_month} {event_start_year}",
                        r"%d %b %Y",
                    ).date(),
                    end=datetime.datetime.strptime(
                        f"{event_end_day} {event_end_month} {event_end_year}",
                        r"%d %b %Y",
                    ).date(),
                )
                calendar_event.make_all_day()
                calendar_events.append(calendar_event)

                start_date = datetime.datetime.strptime(
                    f"{event_start_day} {event_start_month} {event_start_year}",
                    r"%d %b %Y",
                ).date()
                end_date = datetime.datetime.strptime(
                    f"{event_end_day} {event_end_month} {event_end_year}",
                    r"%d %b %Y",
                ).date()
                covered_dates.extend(
                    [
                        start_date + datetime.timedelta(days=x)
                        for x in range(0, (end_date - start_date).days + 1)
                    ]
                )

            # check first element starts with a number
            elif temp_event[0] and temp_event[0][0].isdigit():
                event_name = temp_event[1:]
                event_name = list(map(str.strip, event_name))
                event_name = " - ".join(event_name)
                event_day = int(temp_event[0][:-2])

                if " " in month and month.split()[0] == "Jan":
                    event_month = month.split()[0]
                    event_year = datetime.datetime.now().year + 1
                    current_year = event_year
                else:
                    event_month = process_split_months(event_day, month)
                    event_year = current_year

                logging.debug(f"Event Name: {event_name}")
                logging.debug(f"Event date: {event_day} {event_month} {event_year}")

                calendar_event = Event(
                    name=event_name,
                    begin=datetime.datetime.strptime(
                        f"{event_day} {event_month} {event_year}", r"%d %b %Y"
                    ).date(),
                    end=datetime.datetime.strptime(
                        f"{event_day} {event_month} {event_year}", r"%d %b %Y"
                    ).date(),
                )
                calendar_event.make_all_day()
                calendar_events.append(calendar_event)
                covered_dates.append(
                    datetime.datetime.strptime(
                        f"{event_day} {event_month} {event_year}", r"%d %b %Y"
                    ).date()
                )

        logging.info("Adding unlisted events to calendar")
        for day in days:
            day = day.split("\n")
            day_num = int(day[0])
            events_on_day = day[1:]
            if " " in month and month.split()[0] == "Jan":
                event_month = month.split()[0]
                event_year = datetime.datetime.now().year + 1
                current_year = event_year
            else:
                event_month = process_split_months(day_num, month)
                event_year = current_year

            if events_on_day:
                logging.debug(f"Events: {events_on_day}")
                logging.debug(f"Day: {day_num} {event_month} {event_year}")

                for event_name in events_on_day:
                    if event_name == "H":
                        event_name = "Holiday"
                    event_datetime = datetime.datetime.strptime(
                        f"{day_num} {event_month} {event_year}", r"%d %b %Y"
                    ).date()

                    if event_datetime not in covered_dates:
                        logging.debug(
                            f"Adding {event_name} on {event_datetime} since it is not covered"
                        )
                        calendar_event = Event(
                            name=event_name,
                            begin=event_datetime,
                            end=event_datetime,
                        )
                        calendar_event.make_all_day()
                        calendar_events.append(calendar_event)
                        covered_dates.append(event_datetime)
                    else:
                        found = False
                        for event in calendar_events:
                            event_begin = datetime.datetime.strptime(
                                f"{event.begin.day} {event.begin.month} {event.begin.year}",
                                r"%d %m %Y",
                            ).date()
                            event_end = datetime.datetime.strptime(
                                f"{event.end.day} {event.end.month} {event.end.year}",
                                r"%d %m %Y",
                            ).date()
                            if (
                                event_begin == event_end - datetime.timedelta(days=1)
                                and event_datetime == event_begin
                            ):
                                event_old_name = event.name
                                event.name = event.name + " - " + event_name
                                logging.debug(
                                    f"Updated {event_old_name} on {event_datetime} to {event.name}"
                                )
                                found = True
                        if not found:
                            logging.debug(
                                f"Adding {event_name} on {event_datetime} since it is not covered"
                            )
                            calendar_event = Event(
                                name=event_name,
                                begin=event_datetime,
                                end=event_datetime,
                            )
                            calendar_event.make_all_day()
                            calendar_events.append(calendar_event)
                            covered_dates.append(event_datetime)

    return calendar_events


if __name__ == "__main__":
    parser = argparse.ArgumentParser("PESU ICS Calendar Export")
    parser.add_argument(
        "-i", "--input", help="Input calendar PDF file", required=True, type=str
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Output calendar ics file",
        default="calendar.ics",
        type=str,
    )
    args = parser.parse_args()

    calendar_pdf_filepath = Path(args.input)
    calendar_ics_filepath = Path(args.output)
    if not calendar_pdf_filepath.exists():
        logging.error("Input file does not exist")
        exit(1)
    else:
        logging.info("Beginning calendar export")
        convert_pdf_to_docx(calendar_pdf_filepath, "calendar.docx")
        calendar_events = load_calendar("calendar.docx")
        calendar = Calendar(events=set(calendar_events))
        with open(calendar_ics_filepath, "w") as f:
            f.writelines(calendar.serialize_iter())
        Path("calendar.docx").unlink()
        logging.info("Calendar export complete")
