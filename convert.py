from openpyxl import load_workbook
import re
import argparse

#parse the command line arguments 
parser = argparse.ArgumentParser()
parser.add_argument("input_filename",
					help="the filename of Excel file to be converted",
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
start_row = 5
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


wb2 = load_workbook(filename)
ws2 = wb2["general_report"]

#Automatically fill columns
for row in range(start_row, end_row):
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