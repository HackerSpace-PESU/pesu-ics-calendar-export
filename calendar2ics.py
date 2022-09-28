import re
import datetime
import pdf2docx
from docx.api import Document

def process_split_months(day_num, month):
    if '/' in month:
        month = month.split('/')
    day_num = int(day_num)
    if isinstance(month, list):
        if day_num <= 5:
            month = month[1]
        else:
            month = month[0]
    else:
        month = month
    return month

def convert_pdf_to_docx(calendar_pdf_filepath, calendar_docx_filepath):
    try:
        pdf_file = pdf2docx.Converter(calendar_pdf_filepath)
        pdf_file.convert(calendar_docx_filepath)
        pdf_file.close()
    except Exception as e:
        print(f"Error converting {calendar_pdf_filepath} to {calendar_docx_filepath}: {e}")
        exit(0)

def load_calendar(calendar_docx_filepath):
    calendar_data = dict()
    month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Sept", "Oct", "Nov", "Dec"]
    document = Document(calendar_docx_filepath)
    table = document.tables[0]

    calendar_data = dict()
    weeks = list()
    new_year = datetime.datetime.now().year
    for row in table.rows:
        text = list(cell.text for cell in row.cells)
        text = list(map(str.strip, text))
        # print(text)
        text[1] = text[1].replace("Sept", "Sep")
        if text[1].split('/')[0] in month_names:
            # print(text)
            weeks.append(text)
        elif ' ' in text[1] and text[1].split(' ')[0] in month_names:
            # print(text)
            weeks.append(text)
            if text[1].split(' ')[0] == 'Jan':
                new_year = int(text[1].split(' ')[1])

    first_date = weeks[0][2]
    first_date_month = process_split_months(first_date, weeks[0][1])
    last_date = weeks[-1][-4]
    last_date_month = process_split_months(last_date, weeks[-1][1])

    # print(f"first_date_month: {first_date_month}")
    # print(f"first_date: {first_date}")
    # print(f"last_date_month: {last_date_month}")
    # print(f"last_date: {last_date}")

    temp_date = datetime.datetime.strptime(
        f"{first_date.zfill(2)} {first_date_month} 2022", r"%d %b %Y").date()
    last_date_obj = datetime.datetime.strptime(
        f"{last_date.zfill(2)} {last_date_month} {new_year}", r"%d %b %Y").date()
    # print(f"temp_date: {temp_date}")
    # print(f"last_date_obj: {last_date_obj}")

    while temp_date <= last_date_obj:
        calendar_data[temp_date] = list()
        temp_date = temp_date + datetime.timedelta(days=1)
    # print(calendar_data)

    current_year = datetime.datetime.now().year
    for week in weeks:
        month = week[1]
        events = week[-1]
        days = list(set(week[2:9]))
        # print(days)

        if '/' in month:
            # print(month)
            month = month.split('/')

        for day in days:
            if '\n' in day:
                day_num = int(day.split('\n')[0])
                day_events = day.split('\n')[1:]
                if ' ' in month and month.split()[0] == 'Jan':
                    event_month = month.split()[0]
                    current_year = datetime.datetime.now().year + 1
                else:
                    event_month = process_split_months(day_num, month)
                # print(day_num, day_events, event_month, current_year)
                day_num = str(day_num)
                key = datetime.datetime.strptime(
                    f"{day_num.zfill(2)} {event_month} {current_year}", r"%d %b %Y").date()
                # print(key)
                for de in day_events:
                    calendar_data[key].append(de)
        # print(calendar_data)
        # exit()
        
        # print(events.split('\n'))
        for event in events.split('\n'):
            event = event.strip()
            event = re.split(r'[\-\–]', event)
            event = list(map(str.strip, event))
            # print(event)
            # check if no element starts with a number
            if not any([not x or x[0].isdigit() for x in event]):
                # print(event)
                #TODO
                pass
            
            # check first two elements start with a number
            elif all([x and x[0].isdigit() for x in event[:2]]):
                # print(event)
                #TODO
                pass

            # check first element starts with a number
            elif event[0] and event[0][0].isdigit():
                print(event)
                #TODO
                pass

        
            continue













            event = re.sub(r' +', ' ', event)
            if event:
                # print(event)
                m = re.findall(r'^\d+', event)
                if m:
                    # print("Matches:", m)
                    # print(m.groups())
                    # day_num = [int(m.groups())]
                    # print(day_num)
                    if chr(8211) in event:
                        event = event.split(chr(8211))[-1].strip()
                    else:
                        event = event.split(chr(45))[-1].strip()
                    event = ('H', event)
                    # print(event)
                else:
                    day_num = list(
                        map(lambda x: int(x.split('\n')[0].strip()), days))

                for day in day_num:
                    if isinstance(month, list):
                        if day <= 5:
                            event_month = month[1]
                        else:
                            event_month = month[0]
                    else:
                        event_month = month
                    day = str(day)
                    key = datetime.datetime.strptime(
                        f"{day.zfill(2)} {event_month} 2022", r"%d %b %Y").date()
                    if isinstance(event, tuple) and event[0] == 'H' and 'H' in calendar_data[key]:
                        calendar_data[key].remove('H')
                    calendar_data[key].append(event)
    return calendar_data

if __name__ == '__main__':
    calendar_pdf_filepath = 'data/calendar.pdf'
    calendar_docx_filepath = 'data/calendar.docx'
    # convert_pdf_to_docx(calendar_pdf_filepath, calendar_docx_filepath)
    calendar_data = load_calendar(calendar_docx_filepath)
    # print(calendar_data)
