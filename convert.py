from openpyxl import Workbook
import re
import argparse
import string 
import csv

#parse the command line arguments 
parser = argparse.ArgumentParser()
parser.add_argument("input_filename",
					help="the filename of csv file to be converted",
                    type=str)
parser.add_argument("-o",
					"--output",
					help="the filename of converted Excel file",
					type=str)
args = parser.parse_args()


filename = args.input_filename
filename2 = "output.xlsx"
if args.output:
	filename2 = args.output



# start row no for the excel 
start_row = 2
end_row = 1000
get_severity_level_column = "J"
root_cause_column = "K"
when_discovered_column = "L"
priority_column = "E"
key_column = "B"
summary_column = "C"

default_pre_release = "Pre-Release"
default_post_release = "Post-Release"
default_unknown_root_cause = "UnKnown"

dict_priority_sev = {"Unprioritized":"L4",
				"Low":"L4",
				"None":"L4",
				"Trivial":"L4",
				"Minor":"L3",
				"Medium":"L3",
				"Normal":"L3",
				"High":"L2",
				"Major":"L2",
				"Urgent":"L1",
				"Critical":"L1",
				"Blocker":"L1"}

read_data=[]
with open(filename, 'rb') as f:
    reader = csv.reader(f)
    for row in reader:
        read_data.append(row)


end_row = len(read_data)
wb2 = Workbook()
ws2 = wb2.active

# fill the first row
ws2["A1"] = "Issue Type"
ws2["B1"] = "Issue key"
ws2["C1"] = "Summary"
ws2["D1"] = "Assignee"
ws2["E1"] = "Priority"
ws2["F1"] = "Status"
ws2["G1"] = "Created"
ws2["H1"] = "Resolution"
ws2["I1"] = "Resolved"
ws2["J1"] = "GET Severity Level"
ws2["K1"] = "Root Cause"
ws2["L1"] = "When Discovered"
ws2["M1"] = "Reporter"



#Automatically fill columns
for row in range(start_row, end_row):
	data = read_data[row-1]

	# add data
	for k in range(0,2):
		the_cell = "{}{}".format(string.ascii_uppercase[k], row)
		ws2[the_cell] = data[k]
	for k in range(3,10):
		the_cell = "{}{}".format(string.ascii_uppercase[k-1], row)
		ws2[the_cell] = data[k]
	for k in range(10,12):
		the_cell = "{}{}".format(string.ascii_uppercase[k], row)
		ws2[the_cell] = ""
	the_cell = "{}{}".format(string.ascii_uppercase[12], row)
	ws2[the_cell] = data[12]


	severity_cell = "{}{}".format(get_severity_level_column, row)
	
	priority_cell = "{}{}".format(priority_column, row)
	
	key_cell = "{}{}".format(key_column, row)
	if ws2[key_cell].value is None:
		break

	# fill severity column based on priority column
	the_priority = ws2[priority_cell].value
	if the_priority is None:
		the_priority = "None"
	the_priority = the_priority.strip()
	#print the_priority
	the_severity = dict_priority_sev[the_priority]
	ws2[severity_cell] = the_severity

	# fill when discovered column based on summary column 
	summary_cell = "{}{}".format(summary_column, row)
	when_discovered_cell = "{}{}".format(when_discovered_column, row)
	the_summary = ws2[summary_cell].value
	if len(re.findall("Prod ", the_summary, re.I)) > 0:
		the_when_discovered = default_post_release
	else:
		the_when_discovered = default_pre_release
	ws2[when_discovered_cell] = the_when_discovered

	root_cause_cell = "{}{}".format(root_cause_column, row)
	ws2[root_cause_cell] = default_unknown_root_cause


wb2.save(filename2)
