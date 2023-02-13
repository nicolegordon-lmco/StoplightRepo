#!/usr/bin/env python
# coding: utf-

# imports
import os
import xlsxwriter
import argparse
import datetime as dt
import regex as re
from Excel_Functions import *
from Aggregate import get_aggregated_data, get_stoplight_data

import sys


def main(curSprint, lastCompleteSprint, PI,
        newJiraFile, prevJiraFile, baseJiraFile, 
        stoplight_wdir, PILookupFile):
    # Directory where two excel files will be output
    stoplight_dir = os.path.join(stoplight_wdir, f'Stoplight_{dt.datetime.now().strftime("%y%m%d_%H%M%S")}')
    if not os.path.exists(stoplight_dir):
        os.makedirs(stoplight_dir)
        
    # Get all data
    # Epics
    epics = ['ACE-1 CLIN 2013: Maps', 
           'ACE-1 CLIN 2013: Rapid Adaptive Planning (RAP)',
            'ACE-1 CLIN 2013: RAPSAW SEIT',
            'ACE-1 CLIN 2013: Resource Deconf. (RD) / Resource Viewer (RV)',
            'ACE-1 CLIN 2016: CSS HW Engineering',
            'ACE-1 CLIN 2016: ESS Solution',
            'ACE-1 CLIN 2016: SEIT',
            'ACE-1 CLIN 2016: SIEM and IDS',
            'ACE-1 CLIN 2018: CAMD',
            'ACE-1 CLIN 2018: CSS Extension',
            'ACE-1 CLIN 2018: DevEnv Products',
            'ACE-1 CLIN 2018: DevSecOps',
            'ACE-1 CLIN 2018: EA SEIT',
            'ACE-1 CLIN 2018: Enterprise Architecture HW Engineering'
            ]
    # Clins
    pattern = re.compile("CLIN \d{4}")
    clins = sorted(list(set([re.search(pattern, epic).group() for epic in epics])))
    data = get_aggregated_data(lastCompleteSprint, 
                               newJiraFile, prevJiraFile, 
                               baseJiraFile, PILookupFile,
                               epics, clins, PI)

    # Uncomment this return when testing data aggregation if you don't want to print excel sheets
    # return
    
    # Write all data to excel
    create_excel(data, lastCompleteSprint, stoplight_dir,
                 newJiraFile, prevJiraFile, baseJiraFile)

    # Separate by CLIN
    
    stoplight_dict = {}
    for clin in clins:
        stoplight, change_BL = get_stoplight_data(data, clin)
        stoplight_dict[clin] = {'Data': stoplight, 
                               'Change_BL': change_BL}
        
    # Write to stoplight excel
    stoplight_file = os.path.join(stoplight_dir,
                                  f'Stoplight_Graphics_{dt.datetime.now().strftime("%y%m%d_%H%M%S")}.xlsx')
    wb = xlsxwriter.Workbook(stoplight_file)
    for clin in clins:
        create_stoplight_sheet(wb, stoplight_dict, curSprint, clin)
    wb.close()
    
    return 


if __name__ == "__main__":
    # File that contains dates of sprints
    default_PILookupFile = r"C:\Users\e439931\PMO\Stoplight\PI_Lookup.xlsx"

    # Directory where stoplight files will be output
    default_stoplight_wdir = r'C:\Users\e439931\PMO\Stoplight\Stoplights'

    # Get last completed sprint
    sprint_file = r"C:\Users\e439931\PMO\Stoplight\Sprints.xlsx"
    sprints = pd.read_excel(sprint_file, header=0)
    today = dt.datetime.now()
    PI_sprint = sprints[(today > sprints.End)].iloc[-1].Sprint
    last_complete_sprint = PI_sprint.split('.')[-1]
    if last_complete_sprint == "IP":
        last_complete_sprint = 6
    PI = '.'.join(PI_sprint.split('.')[:2])

    description = ("A script that takes in the sprint number and the current, previous, and baseline Jira exports "
                    + "and computes calculations on the data. Outputs one spreadsheet with the Jira data"
                    + "and pivot tables, and one spreadsheet with the Stoplight graphic for each CLIN.")
    parser = argparse.ArgumentParser(description=description)

    parser.add_argument('newJiraFile', 
                        help='Path to new Jira export')
    parser.add_argument('prevJiraFile', 
                        help='Path to previous Jira export')
    parser.add_argument('baseJiraFile', 
                        help='Path to baseline Jira export')
    parser.add_argument('-sprint', 
                        help='INT: Current sprint number', 
                        type=int)
    parser.add_argument('-lastCompleteSprint', 
                        help='INT: Last complete sprint number', 
                        type=int,
                        default=last_complete_sprint)
    parser.add_argument('-PILookupFile', 
                        help='Path to PI Lookup file with dates for each PI', 
                        default=default_PILookupFile)
    parser.add_argument('-stoplightWdir', 
                        help='Working directory where Stoplight files will be output', 
                        default=default_stoplight_wdir)
    args= parser.parse_args()

    # Command to run
    # python Stoplight.py 
    # C:\Users\e439931\PMO\Stoplight\Roadmaps\COOLR_ACE_1_Roadmap_230209_Cur.xlsx 
    # C:\Users\e439931\PMO\Stoplight\Roadmaps\COOLR_ACE_1_Roadmap_230209_Prev.xlsx 
    # C:\Users\e439931\PMO\Stoplight\Roadmaps\COOLR_ACE_1_Roadmap_230209_Baseline.xlsx 
    # -sprint=3

    data = main(args.sprint, args.lastCompleteSprint, PI, args.newJiraFile.strip('"'), args.prevJiraFile.strip('"'), args.baseJiraFile.strip('"'),
     args.stoplightWdir.strip('"'), args.PILookupFile.strip('"'))





