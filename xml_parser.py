import xml.etree.ElementTree as ET
import pandas as pd
import csv
from openpyxl import Workbook

# Function to parse a xml file which structured in sample_data.xml
def xml_to_csv(xml_file, csv_path):
    
    tree = ET.parse(xml_file)
    root = tree.getroot()

    ns = {'ns': root.tag.split('}')[0].strip('{')}

    headers = None

    for item in root.findall('.//ns:measInfo', namespaces=ns):
        attrib_keys = list(item.attrib.keys())
        gp_keys = list(item.find('ns:granPeriod', namespaces=ns).attrib.keys())

        key_name = attrib_keys[0] if attrib_keys else 'N/A'
        job = 'jobId'
        duration = gp_keys[0] if gp_keys else 'N/A'
        beginTime = gp_keys[1] if len(gp_keys) > 1 else 'N/A'
        repPeriod = 'repPeriod'
        measTypes = 'measTypes'
        measValue = 'measObjLdn'
        measResults = 'measResults'
        suspect = 'suspect'

        hs = [key_name, job, duration, beginTime, repPeriod, measTypes, measValue, measResults, suspect]

        fss = [format_string(s) for s in hs]

        headers = fss

        break

    if headers:
        with open(csv_path, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(headers) 
            
            for item in root.findall('.//ns:measInfo', namespaces=ns):
                try:
                    measInfoId = item.get('measInfoId', 'N/A')

                    job_elem = item.find('ns:job', namespaces=ns)
                    job = job_elem.get('jobId', 'N/A') if job_elem is not None else 'N/A'

                    granPeriod = item.find('ns:granPeriod', namespaces=ns)
                    duration = granPeriod.get('duration', 'N/A') if granPeriod is not None else 'N/A'
                    beginTime = granPeriod.get('beginTime', 'N/A') if granPeriod is not None else 'N/A'

                    repPeriod_elem = item.find('ns:repPeriod', namespaces=ns)
                    repPeriod = repPeriod_elem.get('duration', 'N/A') if repPeriod_elem is not None else 'N/A'

                    measTypes_elem = item.find('ns:measTypes', namespaces=ns)
                    measTypes = measTypes_elem.text if measTypes_elem is not None else 'N/A'

                    measValue_elem = item.find('ns:measValue', namespaces=ns)
                    measObjLdn = measValue_elem.get('measObjLdn', 'N/A') if measValue_elem is not None else 'N/A'
                    measResults_elem = measValue_elem.find('ns:measResults', namespaces=ns) if measValue_elem is not None else None
                    measResults = measResults_elem.text if measResults_elem is not None else 'N/A'

                    suspect_elem = measValue_elem.find('ns:suspect', namespaces=ns) if measValue_elem is not None else None
                    suspect = suspect_elem.text if suspect_elem is not None else 'N/A'

                    writer.writerow([measInfoId, job, duration, beginTime, repPeriod, measTypes, measObjLdn, measResults, suspect])

                except AttributeError as e:
                    print(f"Error processing measInfo: {e}")
                    continue
    else:
        print("No valid headers found, skipping CSV writing.")

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
    main('some_data.xml', 'newvalue.csv', 'Newvalue.xlsx')