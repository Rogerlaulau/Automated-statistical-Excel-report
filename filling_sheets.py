from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import json


# If 5 arguments are not provided alltogether, stop execution

# Include standard modules
import argparse

reminder = 'This program aims at generating an Excel file on Agent performance and FAQ statistic given essential arguments.\nFor example, command line:\npython filename.py -s session.json -f FAQ.json -b "2021-01-09 10:00:00" -e "2021-01-11 22:10:00" -o myreport'

# Initiate the parser
parser = argparse.ArgumentParser(description=reminder)

# Add long and short argument
parser.add_argument("--session", "-s", required=True, help="set session filename, e.g. session.json")
parser.add_argument("--faq", "-f", required=True, help="set FAQ filename, e.g. FAQ.json")
parser.add_argument("--begin", "-b", required=True, help="set report beginning date-time, e.g. 2021-01-09 10:00:00")
parser.add_argument("--end", "-e", required=True, help="set report ending date-time, e.g. 2021-01-11 22:10:00")
parser.add_argument("--out", "-o", required=True, help="set output filename, e.g. myreport")

# Read arguments from the command line
args = parser.parse_args()

# check the arguments 
#if args.session and args.faq and args.begin and args.end and args.out:
print("Set session filename   :",args.session)
print("Set faq filename       :",args.faq)
print("Set beginning date-time:",args.begin)
print("Set ending date-time   :",args.end)
print("Set output filename    :",args.out)

if args.session.split(".")[-1] != "json":
    raise Exception(f"wrong file extension {args.session}, should be json only")
if args.faq.split(".")[-1] != "json":
    raise Exception(f"wrong file extension: {args.faq}, should be json only")
if len(args.begin) != 19:
    raise Exception(f"Incorrect beginning datetime {args.begin}, should be length of 19 only")
if len(args.end) != 19:
    raise Exception(f"Incorrect ending datetime {args.end}, should be length of 19 only")
if len(args.out) < 1:
    raise Exception(f"Incorrect output filename {args.out}") 


# session_filename = 'session.json'
# FAQ_filename = "FAQ.json"
# start_time = "2021-01-09 10:00:00"
# end_time = "2021-01-11 02:00:00"
# out_filename = "statistic5.xlsx"

out_filename = args.out.strip()
if out_filename.endswith(".xlsx"):
    out_filename = out_filename.split(".")[0]

session_filename = args.session
FAQ_filename     = args.faq
start_time       = args.begin
end_time         = args.end


from initiate_sheets import Initiate_Sheets

print('\nGenerating Excel ...')
mysheets = Initiate_Sheets(session_filename, start_time, end_time)

mysheets.create_statistic_ws(FAQ_filename)
mysheets.create_agent_ws()
mysheets.create_utilization_ws(session_filename)
mysheets.create_overall_ws()


# IF YOU WANNA MAKE MODIFICATION TO ANY CELL IN A SHEET:
#mysheets.overall_ws['A1'].value = "THIS IS TESTING CASE"


# REMEMBER: MODIFICATION TAKES EFFECT ONLY PRIOR TO SAVE FILE

mysheets.save_file(out_filename+'.xlsx')
print(f'Done saving [{out_filename}.xlsx]')



'''
COMMAND LINE:
python filling_sheets.py -s session.json -f FAQ.json -b "2021-01-09 10:00:00" -e "2021-01-11 22:10:00" -o myreport
'''









