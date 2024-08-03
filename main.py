import icalendar
import xlsxwriter
import datetime
from bs4 import BeautifulSoup

def parse_ics(file_path):
    with open(file_path, 'rb') as f:
        calendar = icalendar.Calendar.from_ical(f.read())
    
    events = []
    for component in calendar.walk():
        if component.name == "VEVENT":
            dtstart = component.get('DTSTART').dt
            dtend = component.get('DTEND').dt

            if isinstance(dtstart, datetime.date) and not isinstance(dtstart, datetime.datetime):
                dtstart = datetime.datetime(dtstart.year, dtstart.month, dtstart.day)
            if isinstance(dtend, datetime.date) and not isinstance(dtend, datetime.datetime):
                dtend = datetime.datetime(dtend.year, dtend.month, dtend.day)
            
            if dtstart.tzinfo is not None:
                dtstart = dtstart.replace(tzinfo=None)
            if dtend.tzinfo is not None:
                dtend = dtend.replace(tzinfo=None)
            
            event = {
                'SUMMARY': component.get('SUMMARY'),
                'DESCRIPTION': component.get('DESCRIPTION'),
                'DTSTART': dtstart,
                'DTEND': dtend,
                'DATE': dtstart.date()
            }
            events.append(event)
    
    return events

def html_to_excel_format(html_text, workbook):
    if not html_text:
        return [{'text': ''}]
    
    soup = BeautifulSoup(html_text, 'html.parser')
    fragments = []
    
    def handle_element(element):
        if isinstance(element, str):
            fragments.append({'text': element})
        elif element.name == 'br':
            fragments.append({'text': '\n'})
        elif element.name == 'b':
            fragments.append({'text': element.get_text(), 'format': workbook.add_format({'bold': True})})
        elif element.name == 'i':
            fragments.append({'text': element.get_text(), 'format': workbook.add_format({'italic': True})})
        elif element.name == 'u':
            fragments.append({'text': element.get_text(), 'format': workbook.add_format({'underline': True})})
        elif element.name == 'a':
            href = element.get('href', '')
            text = element.get_text()
            fragments.append({'text': f'{text} ({href})', 'format': workbook.add_format({'font_color': 'blue', 'underline': True})})
        elif element.name == 'font':
            size = element.get('size')
            if size:
                try:
                    size = int(size)
                except ValueError:
                    size = None
            fragments.append({'text': element.get_text(), 'format': workbook.add_format({'font_size': size}) if size else {}})
        else:
            for sub_element in element.descendants:
                handle_element(sub_element)
    
    for element in soup:
        if element.name in ['ul', 'ol']:
            list_items = element.find_all('li')
            for idx, li in enumerate(list_items):
                if element.name == 'ol':
                    fragments.append({'text': f'{idx + 1}. {li.get_text()}\n'})
                else:
                    fragments.append({'text': f'• {li.get_text()}\n'})
        else:
            handle_element(element)
    
    return fragments

def write_to_excel(events, output_path):
    events.sort(key=lambda event: event['DTSTART'])
    
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet("Events")
    
    headers = ['DATE', 'SUMMARY', 'DESCRIPTION', 'DTSTART', 'DTEND']
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    # Cell formatting: top-left alignment
    cell_format = workbook.add_format()
    cell_format.set_align('left')
    cell_format.set_align('top')

    # Collect data for width calculation
    max_widths = [0] * len(headers)  # Initialize list to track maximum width for each column

    for row_num, event in enumerate(events, 1):
        worksheet.write(row_num, 0, event['DATE'].strftime("%Y-%m-%d"), cell_format)
        worksheet.write(row_num, 1, event['SUMMARY'], cell_format)
        
        description_fragments = html_to_excel_format(event['DESCRIPTION'], workbook)
        if len(description_fragments) == 1 and 'format' not in description_fragments[0]:
            description_text = description_fragments[0]['text']
            worksheet.write(row_num, 2, description_text, cell_format)
            max_widths[2] = max(max_widths[2], len(description_text))
        else:
            fragments = []
            for fragment in description_fragments:
                if 'format' in fragment:
                    fragments.append(fragment['format'])
                fragments.append(fragment['text'])
            worksheet.write_rich_string(row_num, 2, *fragments)
        
        dtstart_str = event['DTSTART'].strftime("%Y-%m-%d %H:%M:%S")
        dtend_str = event['DTEND'].strftime("%Y-%m-%d %H:%M:%S")
        worksheet.write(row_num, 3, dtstart_str, cell_format)
        worksheet.write(row_num, 4, dtend_str, cell_format)
        
        # Update maximum widths
        max_widths[0] = max(max_widths[0], len(event['DATE'].strftime("%Y-%m-%d")))
        max_widths[1] = max(max_widths[1], len(event['SUMMARY']))
        max_widths[3] = max(max_widths[3], len(dtstart_str))
        max_widths[4] = max(max_widths[4], len(dtend_str))
    
    # Adjust column widths
    for col_num, max_width in enumerate(max_widths):
        worksheet.set_column(col_num, col_num, max_width + 2)  # Add some padding
    
    workbook.close()

def main():
    ics_file = 'path_to_your_file.ics'
    xlsx_file = 'output.xlsx'
    
    events = parse_ics(ics_file)
    write_to_excel(events, xlsx_file)

if __name__ == "__main__":
    main()
