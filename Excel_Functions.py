#!/usr/bin/env python
# coding: utf-

# imports
import pandas as pd
import os
import string
import datetime as dt
from format import formats

def format_keys(df, keys, format, ws):
    for row in range(df.shape[0]):
        key = df.loc[row, 'Key']
        if key in keys:
            ws.conditional_format(f'B{row+2}', 
                                    {'type': 'no_errors',
                                    'format': format})

def create_stoplight_sheet(wb, stoplightDict, curSprint, lastCompleteSprint, clin):
    letters = string.ascii_uppercase
    stoplightData = stoplightDict[clin]['Data'].copy()
    stoplightChange = stoplightDict[clin]['Change_BL'].copy()
    PISprintCompleted = f"{stoplightData.columns[(lastCompleteSprint-1) * 2][:4]}.{lastCompleteSprint}"
    stoplightNumCols = stoplightData.shape[1]

    for idx in stoplightData.index:
        change = stoplightChange.loc[idx, 'Change Since BL']
        if change > 0:
            stoplightData.loc[idx, 'Current Total Pts'] = f"{stoplightData.loc[idx, 'Current Total Pts']} (+{int(change)})"
        elif change < 0:
            stoplightData.loc[idx, 'Current Total Pts'] = f"{stoplightData.loc[idx, 'Current Total Pts']} ({int(change)})"

    ws = wb.add_worksheet(f'{clin} Stoplight')

    # Merge cells for headers
    mergedHeaders = ['A1:A1', 'B1:C1', 'D1:E1', 'F1:G1',
                      'H1:I1', 'J1:K1', 'L1:M1', 'N1:N1', 'O1:R1',
                        f'S1:{letters[stoplightNumCols]}1']
    for cellRange in mergedHeaders:
        if cellRange[0] != cellRange[3]:
            ws.merge_range(cellRange, '')

    # Write and format headers
    headerFormat = wb.add_format(formats['header'])
    
    headers = ['CLIN 2013',
               'Sprint 1', 'Sprint 2', 'Sprint 3', 
               'Sprint 4', 'Sprint 5', 'Sprint 6 (Planning Sprint)',
               'Current Total Pts (Change Since PI Planning)',
              f'Points Analysis (End Sprint {PISprintCompleted})',
              f'Assumed Velocity Forcast (End Sprint {PISprintCompleted}']

    for idx, header in zip(mergedHeaders, headers):
        ws.write(idx, header, headerFormat)

    # Write and format subheaders
    subheaderFormat = wb.add_format(formats['subheader'])
    for col in range(1,13):
        if col % 2 == 1:
            subheader = 'Baseline'
        else:
            if (col/2) <= curSprint:
                subheader = 'Actual'
            else:
                subheader = 'Projected'
        ws.write(1, col, subheader, subheaderFormat)

    extraSubheaders = stoplightData.columns[12:] # Anything after sprint 1-6 data
    for col in range(13,stoplightNumCols+1):
        ws.write(1, col, extraSubheaders[col-13], subheaderFormat)

    # Column widths
    # ws.set_column(start_col, end_col, width) (columns are 0 indexed)
    ws.set_column(1, stoplightNumCols, 10) # default
    ws.set_column(0, 0, 30) # clin categories
    ws.set_column(13, 13, 18) # current total
    ws.set_column(14, stoplightNumCols, 12) # point metrics

    # Add epics
    epicFormat = wb.add_format(formats['epic'])
    numRows = stoplightData.index.shape[0]
    for i in range(numRows):
        ws.write(f'A{i+3}', stoplightData.index[i], epicFormat)

    # Add stoplight data
    dataFormatNumDelta = wb.add_format(formats['SLNumDelta'])
    dataFormatNum = wb.add_format(formats['SLNum'])
    dataFormatDec = wb.add_format(formats['SLDec'])
    dataFormatPer = wb.add_format(formats['SLPer'])
    dataFormatGreen = wb.add_format(formats['SLGreen'])
    dataFormatYellow = wb.add_format(formats['SLYellow'])
    dataFormatRed = wb.add_format(formats['SLRed'])
    dataFormatBlue = wb.add_format(formats['SLBlue'])

    for i in range(numRows):
        for col in range(1,stoplightNumCols+1):
            cellData = stoplightData.iloc[i, col-1]
            # Sprints 1-6
            if col in range(1,13):
                # Col is <= current sprint and is actual/projected
                if (col <= curSprint*2) & (col % 2 == 0):
                    corrBLData = stoplightData.iloc[i, col-2]
                    diff = cellData - corrBLData
                    # On track/ahead
                    if diff > -0.05:
                        dataFormat = dataFormatGreen
                    # Slightly off track
                    if (diff <= -0.05) & (diff > -0.1):
                        dataFormat = dataFormatYellow
                    # Off track
                    if (diff <= -0.1):
                        dataFormat = dataFormatRed
                else:  
                    dataFormat = dataFormatPer
            # Points metrics
            elif col == 13:
                dataFormat = dataFormatNumDelta
            elif col == stoplightNumCols:
                dataFormat = dataFormatDec
            else:
                dataFormat = dataFormatNum
            # If a nan, write empty cell
            try:
                ws.write(i+2, col, cellData, dataFormat)
            except:
                ws.write(i+2, col, "", dataFormat)
            
    return


def create_excel(data, sprint, stoplightDir, 
                 newDataFile, prevDataFile, baseDataFile):
    letters = string.ascii_uppercase
    excelFile = os.path.join(stoplightDir,
                             f'Ground_Dev_ART_STOPLIGHT_{dt.datetime.now().strftime("%y%m%d_%H%M%S")}.xlsx')
    writer = pd.ExcelWriter(excelFile,
                            engine='xlsxwriter')   
    wb = writer.book
    for sheet in data.keys():
        ws = wb.add_worksheet(sheet)
        writer.sheets[sheet] = ws
        data[sheet]['Pivot'].to_excel(writer, sheet_name=sheet, startrow=1 , startcol=0, freeze_panes=(0,1)) 
        numEpics = data[sheet]['Pivot'].shape[0]
        numColsSum = data[sheet]["Pivot"].shape[1]
        numColsCum = data[sheet]["Cum Sum"].shape[1]
        
        lastColSum = letters[numColsSum]
        lastColCum = letters[numColsCum]

        cumSumStartRow = numEpics + 4
        data[sheet]['Cum Sum'].to_excel(writer, sheet_name=sheet, startrow=cumSumStartRow, startcol=0) 
        cumPerStartRow = numEpics*2+6
        data[sheet]['Cum Per'].to_excel(writer, sheet_name=sheet, startrow=cumPerStartRow, startcol=0) 
        clinStartRow = numEpics*3+7
        data[sheet]['CLIN Per'].to_excel(writer, sheet_name=sheet, startrow=clinStartRow, startcol=0) 

        # header format
        titleFormat = wb.add_format(formats['title'])

        # Pergentage format
        percentFormat = wb.add_format({'num_format': '0%'})
        
        ws.conditional_format(f'B{cumPerStartRow+2}:{lastColCum}{cumPerStartRow+2+numEpics}', 
                                {'type': 'no_errors',
                                'format': percentFormat})
        ws.conditional_format(f'B{clinStartRow+2}:{lastColCum}{clinStartRow+4}', 
                                {'type': 'no_errors',
                                'format': percentFormat})

        # Extra tables
        if (sheet == 'Current Pivot') | (sheet == 'Previous Pivot'):
            numColsSprint = data[sheet]["Sprint Metrics"].shape[1]
            numColsRem = data[sheet]["Remaining Metrics"].shape[1]
            firstColChange = letters[numColsSum+2]
            lastColSprint = letters[numColsSum+numColsSprint+2]
            firstColRem = letters[numColsSum+numColsSprint+4]
            lastColRem = letters[numColsSum+numColsSprint+numColsRem+3]

            # Sprint table and remaining table
            ws.merge_range(f'{firstColChange}{cumPerStartRow}:{lastColSprint}{cumPerStartRow}',
                       f'Sprint {sprint}', titleFormat)
            ws.merge_range(f'{firstColRem}{cumPerStartRow}:{lastColRem}{cumPerStartRow}',
                       f'Sprint {sprint}', titleFormat)
            
            data[sheet]['Sprint Metrics'].to_excel(writer, sheet_name=sheet,
                                                   startrow=cumPerStartRow, startcol=numColsSum+2) 
            data[sheet]['Remaining Metrics'].to_excel(writer, sheet_name=sheet,
                                                        startrow=cumPerStartRow, startcol=numColsSum+numColsSprint+4,
                                                        index=False) 
            ws.set_column(numColsSum+2, numColsSum+2, 60)

            # Round format
            roundFormat = wb.add_format({'num_format': '#,##0'})
            ws.conditional_format(f'{letters[numColsSum+5]}{cumPerStartRow+2}:{letters[numColsSum+numColsSprint+numColsRem+2]}{cumPerStartRow+2+numEpics}',
                                    {'type': 'no_errors',
                                    'format': roundFormat})
            round2Format = wb.add_format({'num_format': '#,##0.00'})
            ws.conditional_format(f'{lastColRem}{cumPerStartRow+2}:{lastColRem}{cumPerStartRow+2+numEpics}',
                                    {'type': 'no_errors',
                                    'format': round2Format})

            # Columnd widths
            ws.set_column(numColsSum+3, numColsSum+numColsSprint+numColsRem+3, 18)
 
            if sheet == 'Current Pivot':
                numColsChange = data[sheet]['Changes Since Last Week'].shape[1]
                lastColChange = letters[numColsSum+2+numColsChange]
                # Changes since last week
                data[sheet]['Changes Since Last Week'].to_excel(writer, sheet_name=sheet, 
                                                                startrow=1 , startcol=numColsSum+2)  
                
                # Add header
                ws.merge_range(f'{firstColChange}1:{lastColChange}1',
                               'Changes Since Last Week', titleFormat)
                
                # Add cell formatting to changes since last week
                redFormat = wb.add_format(formats['redDelta'])
                lastRow = data[sheet]['Changes Since Last Week'].shape[0] + 2
                ws.conditional_format(f'{letters[numColsSum+3]}3:{lastColChange}{lastRow}', 
                                        {'type': 'cell',
                                        'criteria': '!=',
                                        'value': 0,
                                        'format': redFormat})

        # Column widths and formats
        catFormat = wb.add_format({'align': 'left',
                                   'bold': False})
        ws.set_column(0, 0, 60, catFormat)
        ws.set_column(1, numColsSum, 10, catFormat)

        # Merge and add headers
        letters = string.ascii_uppercase
        
        ws.merge_range(f'A1:{lastColSum}1',
                       'Sum of Story Points', titleFormat)
        ws.merge_range(f'A{cumSumStartRow}:{lastColCum}{cumSumStartRow}',
                       'Cumulative', titleFormat)
        ws.merge_range(f'A{cumPerStartRow}:{lastColCum}{cumPerStartRow}',
                       'Percentage', titleFormat)
        

    sheets = ['Current Jira Export', 'Previous Jira Export', 'Baseline Jira Export']
    files = [newDataFile, prevDataFile, baseDataFile]

    slipFormat = wb.add_format(formats['slipStories'])
    slipStories = data['Current Pivot']['Slip'].Key.values

    newFormat = wb.add_format(formats['newStories'])
    newStories = data['Current Pivot']['New'].Key.values

    for (file, sheet) in zip(files, sheets):
        df = pd.read_excel(file, usecols='A:P')
        df.to_excel(writer, sheet_name=sheet, index=False)   

        ws = writer.sheets[sheet]
        # Add conditional formatting for new and slips
        if sheet == 'Current Jira Export':
            format_keys(df, newStories, newFormat, ws)
        else:
            format_keys(df, slipStories, slipFormat, ws)
    wb.close()
    
    return