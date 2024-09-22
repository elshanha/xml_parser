import xml.etree.ElementTree as ET
import pandas as pd
import csv
from openpyxl import Workbook

# Function to parse a xml file which structured in sample.xml
def xml_to_csv(xml_file, csv_path):
    
    tree = ET.parse(xml_file)
    root = tree.getroot()

    for item in root.iter('measInfo'):

        attrib_keys = list(item.attrib.keys())
        gp_keys = list(item.find('granPeriod').keys())
        gp_keys_first = gp_keys[0]
        gp_keys_second = gp_keys[1]

        key_name = attrib_keys[0]
        job = item.find('job').tag
        duration = gp_keys_first
        beginTime = gp_keys_second
        repPeriod = item.find('repPeriod').tag
        measTypes = item.find('measTypes').tag
        measValue = item.find('measValue').tag
        measResults = item.find('measValue/measResults').tag
        suspect = item.find('measValue/suspect').tag
        
        hs = [key_name, job, duration, beginTime, repPeriod, measTypes, measValue, measResults, suspect]

        fss = [format_string(s) for s in hs]

        headers = fss

    with open(csv_path, 'w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(headers) 
        
        for item in root.iter('measInfo'):
            try:
                measInfoId = item.get('measInfoId') if item is not None else 'N/A'
                job = item.find('job').get('jobId') if item.find('job') is not None else 'N/A'
                duration = item.find('granPeriod').get('duration') if item.find('granPeriod') is not None else 'N/A'
                beginTime = item.find('granPeriod').get('beginTime') if item.find('granPeriod') is not None else 'N/A'
                repPeriod = item.find('repPeriod').get('duration') if item.find('repPeriod') is not None else 'N/A'
                measTypes = item.find('measTypes').text if item.find('measTypes') is not None else 'N/A'
                measValue = item.find('measValue').get('measObjLdn') if item.find('measValue') is not None else 'N/A'
                measResults = item.find('measValue/measResults').text if item.find('measValue/measResults') is not None else 'N/A'
                suspect = item.find('measValue/suspect').text if item.find('measValue/suspect') is not None else 'N/A'
                
                writer.writerow([measInfoId, job, duration, beginTime , repPeriod, measTypes, measValue, measResults, suspect])
                
            except AttributeError as e:
                print(f"Error: {e}")
                continue

# Function to give better look to header row
def format_string(s):
    formatted = ''.join(' ' + char if char.isupper() and index > 0 and s[index - 1].islower() else char
                        for index, char in enumerate(s))
    
    return formatted.strip().capitalize()


# CSV to Excel without pandas
def csv_to_excel(csv_file, excel_file):
    if not csv_file.lower().endswith('.csv'):
        print('Expected a CSV File')
        return 0
    if not excel_file.lower().endswith('.xlsx'):
        print('Expected a Excel file')
        return 0
    
    wb = Workbook()
    ws = wb.active
    try:
        with open(csv_file, 'r') as f:
            for row in csv.reader(f):
                ws.append(row)
        wb.save(excel_file)
    except FileNotFoundError:
        print('Error: File not found')


# CSV to Excel with pandas
def pd_csv_to_excel(input, output):

    if not input.lower().endswith('.csv'):
        print('Expected a CSV File')
        return 0
    if not output.lower().endswith('.xlsx'):
        print('Expected a Excel file')
        return 0
    
    rf = pd.read_csv(input)
  
    rf.to_excel(output, index=False, header=True, sheet_name='Data')

def main(xml_file, csv_file, excel_file):
    xml_to_csv(xml_file, csv_file)
    # csv_to_excel(csv_file, excel_file)
    pd_csv_to_excel(csv_file, excel_file)

# Add file names (actual xml file that must be converted, others can be anything with only proper extensions)
if __name__ == "__main__":
    main('sample_data.xml', 'data.csv', 'Newdata.xlsx')