from __future__ import division
import csv
import os
import operator
import datetime
import xlsxwriter
import argparse
from pprint import pprint
import pandas as pd

parser = argparse.ArgumentParser()

parser.add_argument('sitename', help = 'The name of the site the report will be generated for with "quotes" around it.')
parser.add_argument('filename', help = 'The csv filename: it must be in the same directory as this script')

args = parser.parse_args()

sitename = '{}'.format(args.sitename)
filename = '{}'.format(args.filename)

PATH = os.path.dirname(os.path.abspath(__file__))

monthlist = []
yearlist = []
timelist = []

def csv_import(filename):
    with open(filename, 'r') as file:
        reader = csv.reader(file, delimiter=',')
        # remove header line from CSV if manually ran
        next(reader)
        next(reader)
        loc_params = next(reader)
        # Build list of location dictionaries
        client_list = []
        for row in reader:
            # location dictionary
            data = {}
            for x in range(len(loc_params)):
                # for each location parameters add key and value to dictionary
                data[loc_params[x]] = str(row[x])
            client_list.append(data)
        return client_list

def calculate_connected_time(start_time, end_time):
    global monthlist
    global yearlist
    global timelist
    #start_time = datetime.datetime.strptime(start_time, '%m/%d/%y %H:%M')
    start_time = datetime.datetime.strptime(start_time, '%Y-%m-%d %H:%M:%S')
    monthlist.append(start_time.strftime('%B'))
    yearlist.append(start_time.strftime('%Y'))
    timelist.append(end_time)
    #end_time = datetime.datetime.strptime(end_time, '%m/%d/%y %H:%M')
    end_time = datetime.datetime.strptime(end_time, '%Y-%m-%d %H:%M:%S')
    connected_time = (end_time - start_time).total_seconds()
    return connected_time


print('gathering data from csv')
client_list = csv_import(filename)
print('processing data')

df = pd.DataFrame(client_list)
df['connected_time'] = df.apply(lambda x: calculate_connected_time(x.start_time, x.end_time), axis=1)

  
monthstr = "{} - {}".format(max(set(monthlist), key= monthlist.count), max(set(yearlist), key= yearlist.count))
# Used for start and end times off to the side of report
timeset = set(timelist)
#timeset = sorted(timeset, key=lambda timeset: datetime.datetime.strptime(timeset, '%m/%d/%y %H:%M'))
timeset = sorted(timeset, key=lambda timeset: datetime.datetime.strptime(timeset, '%Y-%m-%d %H:%M:%S'))


print("creating excel report")
excelname = os.path.splitext(filename)[0]
excelname += '.xlsx'
workbook = xlsxwriter.Workbook('{}'.format(excelname))
workbook.set_size(1600, 2000)
worksheet = workbook.add_worksheet('Report')
# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20.5)
worksheet.set_column('B:F', 14.8)
worksheet.set_column('K:N', 22)
worksheet.set_row(1, 13)
worksheet.set_row(2, 13)
worksheet.set_row(3, 13)
worksheet.set_row(4, 13)
worksheet.set_row(5, 13)
worksheet.set_row(6, 13)
worksheet.set_column('K:L', None, None, {'hidden': True})

# Create a format to use in the merged range.
merge_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '5C5B5A',
    'font_color': 'white',
    'font_size': 14})

Label_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '5C5B5A',
    'font_color': 'white',
    'font_size': 12,
    'text_wrap': 1
})
header_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '5C5B5A',
    'font_color': 'white',
    'font_size': 10
})
centered_hyp_format = workbook.add_format({
    'align': 'center',
    'color': '#0000EE',
    'underline': 1,
    'font_size': 10
})
bold_format = workbook.add_format({
    'align': 'center',
    'font_size': 10,
    'bold': 1
})
#center_format = workbook.add_format({
#    'align': 'center'
#})
bottom_border_title = workbook.add_format({
    'align': 'center',
    'font_size': 10,
    'underline': 1
})
main_site_format = workbook.add_format({
    'bold':1,
    'bottom':1,
    'font_size': 10,
    'bottom_color': '#0000EE',
    'align': 'right',
    'num_format': '#,###,##0'
})
main_site_location_format = workbook.add_format({
    'bold':1,
    'bottom':1,
    'font_size': 10,
    'bottom_color': '#0000EE',
    'align': 'left'
})
sub_site_format = workbook.add_format({
    'bottom':1,
    'align':'right',
    'font_size': 10,
    'bottom_color': '#800080',
    'num_format': '#,###,##0'
})
sub_site_location_format = workbook.add_format({
    'bottom':1,
    'align':'left',
    'font_size': 10,
    'bottom_color': '#800080'
})
ssid_format = workbook.add_format({
    'align':'right',
    'font_size': 10,
    'bg_color': '#C0C0C0',
    'num_format': '#,###,##0'
})
ssid_name_format = workbook.add_format({
    'align':'left',
    'bg_color': '#C0C0C0',
    'font_size': 10
})
bold_only_format = workbook.add_format({
    'bold': 1
})

# Merge cells on row 1.
worksheet.merge_range('A1:E1', '{} - WiFi Statistics Summary Report'.format(monthstr), merge_format) #Change to F1 if using time
worksheet.merge_range('A2:A7','{}'.format(sitename),Label_format)
worksheet.merge_range('B2:E2','')
#worksheet.write_url('B2', 'http://www.library.ca.gov/Content/pdf/services/toLibraries/SurveyInstructions2017-18.pdf', centered_hyp_format)

# Add Total
worksheet.write('C4', 'Client User Summary',bold_format)
worksheet.write('C5', 'Number of Sessions', bottom_border_title)
worksheet.write('D5', 'Number of Users', bottom_border_title)
#worksheet.write('E5', 'Sum of Time (hours)', bottom_border_title)
#worksheet.write('F5', 'Sum of Time (minutes)', bottom_border_title)
worksheet.write('C6', len(df.client_mac))
worksheet.write('D6', len(df['client_mac'].unique()))
#worksheet.write('E6', round(df['connected_time'].sum()/3600))
#worksheet.write('F6', round(df['connected_time'].sum()/60))


# Build Table header
worksheet.write('A8', 'Locations', header_format)
worksheet.write('B8', 'SSID', header_format ) 
worksheet.write('C8', 'Number of Sessions', header_format) 
worksheet.write('D8', 'Number or Users', header_format) 
worksheet.write('E8', '', header_format) # remove if adding client sum
#worksheet.write('E8', 'Sum of Time (hours)', header_format)
#worksheet.write('F8', 'Sum of Time (minutes)', header_format)

# Print Start and End times off to the side of the Report
worksheet.write('G3', 'Time Stamps from Client Summary', bold_only_format) # Change to H if adding Times
worksheet.write('G4', 'Start time:') # Change to H if adding Times
worksheet.write('H4', ' {}'.format(timeset[0])) # Change to I if adding Times
worksheet.write('G5', 'End time:') # Change to H if adding Times
worksheet.write('H5',' {}'.format(timeset[-1])) # Change to I if adding Times

cursor_line = 8

location_list = df.location.unique().tolist()

for location in location_list:
    filt = df['location'] == location
    location_df = df.loc[filt]
    cursor_line += 1
    main_location_name = (location_df['location'].unique()[0])
    main_location_sessions = len(location_df.client_mac)
    main_location_unique_count = len(location_df['client_mac'].unique())
    worksheet.write('A{}'.format(cursor_line), "    {}".format(main_location_name), main_site_location_format)
    worksheet.write('B{}'.format(cursor_line), "", main_site_location_format)
    worksheet.write('C{}'.format(cursor_line), main_location_sessions, main_site_format) 
    worksheet.write('D{}'.format(cursor_line), main_location_unique_count, main_site_format) 
    worksheet.write('E{}'.format(cursor_line), "", main_site_location_format)# remove if adding client sum
    #worksheet.write('E{}'.format(cursor_line), round(location_df['connected_time'].sum()/3600), main_site_format)
    #worksheet.write('F{}'.format(cursor_line), round(location_df['connected_time'].sum()/60), main_site_format)
    ssid_loc_list = location_df.ssid.unique().tolist()
    for ssid_loc in ssid_loc_list:
        filt = location_df['ssid'] == ssid_loc
        ssid_loc_df = location_df[filt]
        cursor_line += 1
        ssid_loc_name = (ssid_loc_df['ssid'].unique()[0])
        ssid_loc_sessions = len(ssid_loc_df.client_mac)
        ssid_loc_unique_count = len(ssid_loc_df['client_mac'].unique())
        worksheet.write('A{}'.format(cursor_line), "", ssid_name_format)
        worksheet.write('B{}'.format(cursor_line), "    {}".format(ssid_loc_name), ssid_name_format)
        worksheet.write('C{}'.format(cursor_line), ssid_loc_sessions, ssid_format) # change to B{} if adding client sum 
        worksheet.write('D{}'.format(cursor_line), ssid_loc_unique_count, ssid_format) # change to C{} if adding client sum
        worksheet.write('E{}'.format(cursor_line), "", ssid_name_format)
        #worksheet.write('E{}'.format(cursor_line), round(ssid_loc_df['connected_time'].sum()/3600), ssid_name_format)
        #worksheet.write('F{}'.format(cursor_line), round(ssid_loc_df['connected_time'].sum()/60), ssid_name_format)
    sub_location_list = location_df.sublocation.unique().tolist()
    for sub_location in sub_location_list:
        filt = location_df['sublocation'] == sub_location
        sub_loc_df = location_df[filt]
        cursor_line += 1
        sub_loc_name = (sub_loc_df['sublocation'].unique()[0])
        sub_loc_sessions = len(sub_loc_df.client_mac)
        sub_loc_unique_count = len(sub_loc_df['client_mac'].unique())
        worksheet.write('A{}'.format(cursor_line), "        {}".format(sub_loc_name), sub_site_location_format)
        worksheet.write('B{}'.format(cursor_line), "", sub_site_location_format)
        worksheet.write('C{}'.format(cursor_line), sub_loc_sessions, sub_site_format) 
        worksheet.write('D{}'.format(cursor_line), sub_loc_unique_count, sub_site_format)
        worksheet.write('E{}'.format(cursor_line), "", sub_site_location_format)
        #worksheet.write('E{}'.format(cursor_line), round(sub_loc_df['connected_time'].sum()/3600), sub_site_format)
        #worksheet.write('F{}'.format(cursor_line), round(sub_loc_df['connected_time'].sum()/60), sub_site_format)
        
ssids = {}
ssid_list = df.ssid.unique().tolist()
for ssid in ssid_list:
    filt = df['ssid'] == ssid
    ssid_df = df[filt]
    ssids[ssid] = len(ssid_df['client_mac'].unique())

#print("There are {} unique clients".format(sum(ssids.values())))
sorted_ssids = sorted(ssids.items(), key=operator.itemgetter(1), reverse=True)
cursor_line += 5
worksheet.merge_range('A{}:E{}'.format(cursor_line,cursor_line), 'Unique Clients by SSID', merge_format)
ssidline = cursor_line
otherline = ssidline -1
other_total = 0
if len(sorted_ssids) > 9:
    worksheet.write('A{}'.format(cursor_line), 'Unique Clients by SSID (Top 10)', merge_format)
    worksheet.write('M{}'.format(otherline), 'OTHER SSIDs', bold_format)
    for x in range(len(sorted_ssids)-1, 9, -1):
        otherline += 1
        other_total += sorted_ssids[x][1]
        worksheet.write('M{}'.format(otherline),'{} - {:,}'.format(sorted_ssids[x][0], sorted_ssids[x][1]))
        worksheet.write('N{}'.format(otherline),sorted_ssids[x][1])
        sorted_ssids.remove(sorted_ssids[x])
    sorted_ssids.append(tuple(('OTHER SSIDs', other_total)))
for ssid in sorted_ssids:
    worksheet.write('K{}'.format(ssidline),'{} - {:,}'.format(ssid[0], ssid[1]))
    worksheet.write('L{}'.format(ssidline),ssid[1])
    ssidline+=1

# Create a chart object.
chart = workbook.add_chart({'type': 'pie'})
chart.show_hidden_data()
chart.add_series({
    'categories': '=Report!$K${}:$K${}'.format(cursor_line,cursor_line+len(sorted_ssids)-1),
    'values':     '=Report!$L${}:$L${}'.format(cursor_line,cursor_line+len(sorted_ssids)-1),
})
cursor_line += 1
chart.set_style(10)
chart.set_size({'width': 540, 'height': 432})
worksheet.insert_chart('A{}'.format(cursor_line), chart, {'x_offset': 25, 'y_offset': 15})


workbook.close()
print("completed - saved as {}".format(excelname))
