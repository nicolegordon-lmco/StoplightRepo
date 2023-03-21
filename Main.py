import os
import regex as re
import pandas as pd
import datetime as dt
import xlsxwriter
import argparse
import sys
from PivotClass import Pivot

def main(curSprint, lastCompleteSprint, PI,
        newJiraFile, prevJiraFile, baseJiraFile, 
        stoplightWdir, PILookupFile,
        printContributors):

    # Directory where two excel files will be output
    stoplightDir = os.path.join(stoplightWdir, 
                                f'Stoplight_{dt.datetime.now().strftime("%y%m%d_%H%M%S")}')
    if not os.path.exists(stoplightDir):
        os.makedirs(stoplightDir)
        
    # Get all data
    # Epics
    # epics = ['ACE-1 CLIN 2013: Maps', 
    #          'ACE-1 CLIN 2013: Rapid Adaptive Planning (RAP)',
    #         'ACE-1 CLIN 2013: RAPSAW SEIT',
    #         'ACE-1 CLIN 2013: Resource Deconf. (RD) / Resource Viewer (RV)',
    #         'ACE-1 CLIN 2016: CSS HW Engineering',
    #         'ACE-1 CLIN 2016: ESS Solution',
    #         'ACE-1 CLIN 2016: Multi-Factor Authentication (MFA) Solution',
    #         'ACE-1 CLIN 2016: SEIT',
    #         'ACE-1 CLIN 2016: SIEM and IDS',
    #         'ACE-1 CLIN 2018: CAMD',
    #         'ACE-1 CLIN 2018: CSS Extension',
    #         'ACE-1 CLIN 2018: DevEnv Products',
    #         'ACE-1 CLIN 2018: DevSecOps',
    #         'ACE-1 CLIN 2018: EA SEIT',
    #         'ACE-1 CLIN 2018: Enterprise Architecture HW Engineering'
    #         ]
    epics = [
            'ACE-1 CLIN 2013: LAE BCB-1505',
            'ACE-1 CLIN 2013: Rapid Adaptive Planning (RAP)',
            'ACE-1 CLIN 2013: RAPSAW SEIT',
            'ACE-1 CLIN 2013: Resource Deconf. (RD) / Resource Viewer (RV)',
            'ACE-1 CLIN 2016: CSS HW Engineering',
            'ACE-1 CLIN 2016: ESS Solution',
            'ACE-1 CLIN 2016: Multi-Factor Authentication (MFA) Solution',
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
    clinPattern = re.compile("CLIN \d{4}")
    clins = sorted(list(set([re.search(clinPattern, epic).group() for epic in epics])))

    # Instantiate pivots from current, previous, baseline weeks
    cur = Pivot(newJiraFile, PILookupFile, epics, clins, PI, jira="Current", stoplightDir=stoplightDir)
    prev = Pivot(prevJiraFile, PILookupFile, epics, clins, PI, jira="Previous", stoplightDir=stoplightDir)
    baseline = Pivot(baseJiraFile, PILookupFile, epics, clins, PI, jira="Baseline", stoplightDir=stoplightDir)

    if printContributors is not None:
        for pivot in [cur, prev, baseline]:
            pivot.pivotDf.to_excel(os.path.join(stoplightDir, f"{pivot.jira}Contributors.xlsx"), index=False)

    # Set slips
    cur.set_slip(prev.JiraDf, baseline.JiraDf)
    prev.set_slip(prev.JiraDf, baseline.JiraDf)

    # Set new stories
    cur.set_new(prev.JiraDf, baseline.JiraDf)

    # Set changes since last week
    cur.set_weekly_change(prev.pivotTable)

    # Set cumulative metrics
    cur.set_cum_metrics()
    prev.set_cum_metrics()
    baseline.set_cum_metrics()

    # Set sprint metrics
    cur.set_sprint_metrics(curSprint, lastCompleteSprint,
                            baseline.cumSum, baseline.cumPer)
    prev.set_sprint_metrics(curSprint, lastCompleteSprint,
                            baseline.cumSum, baseline.cumPer)
    
    # Write to excel
    excelFile = os.path.join(stoplightDir,
                             f'Ground_Dev_ART_STOPLIGHT_{dt.datetime.now().strftime("%y%m%d_%H%M%S")}.xlsx')
    writer = pd.ExcelWriter(excelFile,
                            engine='xlsxwriter')  
    cur.excel_pivot(writer)
    prev.excel_pivot(writer)
    baseline.excel_pivot(writer)

    cur.excel_Jira(writer)
    prev.excel_Jira(writer, cur=cur)
    baseline.excel_Jira(writer, cur=cur)
    writer.book.close()

    # Separate by CLIN
    cur.stoplightDict = {}
    for clin in clins:
        cur.set_stoplight_data(baseline, clin)
        cur.stoplightDict[clin] = {'Data': cur.stoplightData, 
                                    'Change_BL': cur.changeBL}
        
    # Write to stoplight excel
    stoplightFile = os.path.join(stoplightDir,
                                  f'Stoplight_Graphics_{dt.datetime.now().strftime("%y%m%d_%H%M%S")}.xlsx')
    wb = xlsxwriter.Workbook(stoplightFile)
    for clin in clins:
        cur.create_stoplight_sheet(wb, clin)
    wb.close()

if __name__ == "__main__":
    # File that contains dates of sprints
    defaultPILookupFile = r"C:\Users\e439931\PMO\Stoplight\PI_Lookup.xlsx"

    # Directory where stoplight files will be output
    defaultStoplightWdir = r'C:\Users\e439931\PMO\Stoplight\Stoplights'

    # sprint info
    sprintFile = r"C:\Users\e439931\PMO\Stoplight\Sprints.xlsx"
    sprints = pd.read_excel(sprintFile, header=0)
    today = dt.datetime.today()
    
    # Get current sprint
    # get most recent tuesday
    today = dt.datetime.strptime(dt.datetime.today().date().strftime('%Y-%m-%d %H:%M:%S'), 
                                 '%Y-%m-%d %H:%M:%S')
    todayDay = today.weekday()
    if todayDay < 1:
        # it is monday
        lastTuesday = today - dt.timedelta(days=6)
    else:
        # it is not monday
        lastTuesday = today - dt.timedelta(days=todayDay+1)
    PISprint = sprints[((lastTuesday > sprints.Start) 
                        & (lastTuesday <= sprints.End))].iloc[0].Sprint
    curSprint = PISprint.split('.')[-1]
    if curSprint == "IP":
        curSprint = 6

    # Get last completed sprint
    PISprint = sprints[(today > sprints.End)].iloc[-1].Sprint
    if curSprint == 1:
        lastCompleteSprint = 0
    else:
        lastCompleteSprint = PISprint.split('.')[-1]
        if lastCompleteSprint == "IP":
            lastCompleteSprint = 6
    
    curSprint = int(curSprint)
    lastCompleteSprint = int(lastCompleteSprint)
    PI = '.'.join(PISprint.split('.')[:2])

    description = ("A script that takes in the sprint number and the "
                   + "current, previous, and baseline Jira exports "
                   + "and computes calculations on the data. Outputs "
                   + "one spreadsheet with the Jira data and pivot tables,"
                   + " and one spreadsheet with the Stoplight graphic for each CLIN.")
    parser = argparse.ArgumentParser(description=description)

    parser.add_argument('newJiraFile', 
                        help='Path to new Jira export')
    parser.add_argument('prevJiraFile', 
                        help='Path to previous Jira export')
    parser.add_argument('baseJiraFile', 
                        help='Path to baseline Jira export')
    parser.add_argument('--sprint', 
                        help='INT: Current sprint number', 
                        type=int,
                        default=curSprint)
    parser.add_argument('--lastCompleteSprint', 
                        help='INT: Last complete sprint number', 
                        type=int,
                        default=lastCompleteSprint)
    parser.add_argument('--PI', 
                        help='str: current PI', 
                        type=str,
                        default=PI)
    parser.add_argument('--PILookupFile', 
                        help='Path to PI Lookup file with dates for each PI', 
                        default=defaultPILookupFile)
    parser.add_argument('--stoplightWdir', 
                        help='Working directory where Stoplight files will be output', 
                        default=defaultStoplightWdir)
    parser.add_argument('--printContributors', 
                        help='Flag to print items that contributed to the pivots', 
                        action=argparse.BooleanOptionalAction)
    args = parser.parse_args()
    
    # Check if --sprint was input but --lastCompleteSprint was not
    if ('--sprint' in ''.join(sys.argv)) & ('--lastCompleteSprint' not in ''.join(sys.argv)):
        print("--sprint was input but --lastCompleteSprint was not. \
              \nIf --sprint is input, --lastCompleteSprint must also be input. Now exiting...")
        sys.exit()

    data = main(args.sprint, 
                args.lastCompleteSprint, 
                args.PI, 
                args.newJiraFile.strip('"'), 
                args.prevJiraFile.strip('"'),
                args.baseJiraFile.strip('"'),
                args.stoplightWdir.strip('"'), 
                args.PILookupFile.strip('"'),
                args.printContributors)
