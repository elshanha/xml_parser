import xml.etree.ElementTree as ET
import pandas as pd
import csv
from openpyxl import Workbook

# Function to extract metadata from a xml file
def extract_metadata(xml_file):
    tree = ET.parse(xml_file)
    root = tree.getroot()

    ns = {'ns': root.tag.split('}')[0].strip('{')}

    fh = root.find('ns:fileHeader', namespaces=ns)
    dnPrefix = fh.get('dnPrefix', 'N/A') if fh is not None else 'N/A'
    vendorName = fh.get('vendorName', 'N/A') if fh is not None else 'N/A'
    fileFormatVersion = fh.get('fileFormatVersion', 'N/A') if fh is not None else 'N/A'

    fileSender = fh.find('ns:fileSender', namespaces=ns) if fh is not None else 'N/A'
    flDn = fileSender.get('localDn', 'N/A') if fileSender is not None else 'N/A'
    elementType = fileSender.get('elementType', 'N/A') if fileSender is not None else 'N/A'

    measCollec = fh.find('ns:measCollec', namespaces=ns) if fh is not None else 'N/A'
    beginT = measCollec.get('beginTime', 'N/A') if measCollec is not None else 'N/A'

    fileFooter = root.find('ns:fileFooter', namespaces=ns)
    mC = fileFooter.find('ns:measCollec', namespaces=ns) if fileFooter is not None else 'N/A'
    endTime = mC.get('endTime', 'N/A') if mC is not None else 'N/A'

    new_data = {
        'Dnprefix': dnPrefix,
        'Vendor name' : vendorName,
        'File format version' : fileFormatVersion,
        'LocalDn' : flDn,
        'Element type' : elementType,
        'Begin time' : beginT,
        'End time' : endTime
    }

    return new_data

# Function to parse a xml file which structured in sample_data.xml
def xml_to_csv(xml_file, csv_path, metadata=None):

    tree = ET.parse(xml_file)
    root = tree.getroot()

    ns = {'ns': root.tag.split('}')[0].strip('{')}

    headers = None

    for meas_data in root.findall('.//ns:measData', namespaces=ns):
        managed_element = list(meas_data.find('ns:managedElement', namespaces=ns).keys())

        localDn = managed_element[0] if managed_element else 'N/A'
        userLabel = managed_element[1] if managed_element else 'N/A'
        
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

            hs = [localDn, userLabel, key_name, job, duration, beginTime, repPeriod, measTypes, measValue, measResults, suspect]

            fss = [format_string(s) for s in hs]

            headers = fss

            break

    if headers:

        with open(csv_path, 'w', newline='', encoding='utf-8') as csv_file:
            writer = csv.writer(csv_file)
            writer.writerow(headers) 

            for meas_data in root.findall('.//ns:measData', namespaces=ns):

                managed_element = meas_data.find('ns:managedElement', namespaces=ns)
                localDn = managed_element.get('localDn', 'N/A') if managed_element is not None else 'N/A'
                userLabel = managed_element.get('userLabel', 'N/A') if managed_element is not None else 'N/A'

                for item in meas_data.findall('ns:measInfo', namespaces=ns):
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

                        writer.writerow([localDn, userLabel, measInfoId, job, duration, beginTime, repPeriod, measTypes, measObjLdn, measResults, suspect])


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

# CSV to Excel with pandas
def pd_csv_to_excel(input, output, metadata):

    if not input.lower().endswith('.csv'):
        print('Expected a CSV File')
        return 0
    if not output.lower().endswith('.xlsx'):
        print('Expected a Excel file')
        return 0
    
    rf = pd.read_csv(input)
  
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        rf.to_excel(writer, index=False, header=True, sheet_name='Data')

        if metadata:
            extra_df = pd.DataFrame([metadata])
            extra_df.to_excel(writer, index=False, header=True, sheet_name='Metadata')


def main(xml_file, csv_file, excel_file):
    xml_to_csv(xml_file, csv_file)
    m_data = extract_metadata(xml_file)
    pd_csv_to_excel(csv_file, excel_file, metadata=m_data)

# Add file names (actual xml file that must be converted, others can be anything with only proper extensions)
if __name__ == "__main__":
    main('sample_data.xml', 'newvalue.csv', 'Newvalue.xlsx')