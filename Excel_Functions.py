#!/usr/bin/env python
# coding: utf-

# imports
import pandas as pd
import os
import xlsxwriter
import string
import datetime as dt

def create_stoplight_sheet(wb, stoplight_dict, sprint, clin):
    letters = string.ascii_uppercase
    stoplight_data = stoplight_dict[clin]['Data'].copy()
    stoplight_change = stoplight_dict[clin]['Change_BL'].copy()
    PI_sprint = f"{stoplight_data.columns[(sprint-1) * 2][:4]}.{sprint}"

    for idx in stoplight_data.index:
        change = stoplight_change.loc[idx, 'Change Since BL']
        if change > 0:
            stoplight_data.loc[idx, 'Current Total Pts'] = f"{stoplight_data.loc[idx, 'Current Total Pts']} (+{int(change)})"
        elif change < 0:
            stoplight_data.loc[idx, 'Current Total Pts'] = f"{stoplight_data.loc[idx, 'Current Total Pts']} ({int(change)})"

    ws = wb.add_worksheet(f'{clin} Stoplight')

    # Merge cells for headers
    merged_headers = ['A1:A1', 'B1:C1', 'D1:E1', 'F1:G1',
                      'H1:I1', 'J1:K1', 'L1:M1', 'N1:N1', 'O1:Q1']
    for cell_rng in merged_headers:
        if cell_rng[0] != cell_rng[3]:
            ws.merge_range(cell_rng, '')

    # Write and format headers
    header_format = wb.add_format({'bold': True, 'font_color': 'white',
                                 'font': 'Calibri', 'font_size': 12,
                                   'align': 'center', 'valign': 'top',
                                   'text_wrap': True, 
                                 'bg_color': '#7A869A',
                                  'border': True, 'border_color': '#cccccc'})
    headers = ['CLIN 2013',
               'Sprint 1', 'Sprint 2', 'Sprint 3', 
               'Sprint 4', 'Sprint 5', 'Sprint 6 (Planning Sprint)',
               'Current Total Pts (Change Since PI Planning)',
              f'Points Analysis (End Sprint {PI_sprint})']

    for idx, header in zip(merged_headers, headers):
        ws.write(idx, header, header_format)

    # Write and format subheaders
    subheader_format = wb.add_format({'bold': True, 'font_color': 'black',
                                     'font': 'Calibri', 'font_size': 12,
                                       'align': 'center', 'valign': 'top',
                                       'text_wrap': True, 
                                     'bg_color': '#ffffff',
                                     'border': True, 'border_color': '#cccccc'})
    for col in range(1,13):
        if col % 2 == 1:
            subheader = 'Baseline'
        else:
            if (col/2) <= sprint:
                subheader = 'Actual'
            else:
                subheader = 'Projected'
        ws.write(1, col, subheader, subheader_format)
    extra_subheaders = ['Current Total Pts', 'Points Expected', 
                        'Points Completed', 'Delta Points']
    for col in range(13,17):
        ws.write(1, col, extra_subheaders[col-13], subheader_format)

    # Column widths
    ws.set_column(1, 16, 10)
    ws.set_column(0, 0, 30)
    ws.set_column(13, 13, 18)

    # Add category data
    cat_format = wb.add_format({'bold': True, 'font_color': 'black',
                                 'font': 'Calibri', 'font_size': 12,
                               'align': 'center', 'valign': 'top',
                               'text_wrap': True,
                               'bg_color': '#ffffff',
                               'border': True, 'border_color': '#cccccc'})
    num_rows = stoplight_data.index.shape[0]
    for i in range(num_rows):
        ws.write(f'A{i+3}', stoplight_data.index[i], cat_format)

    # Add stoplight data
    data_format_num_delta = wb.add_format({'align': 'center',
                                            'valign': 'top',
                                            'font_color': 'black',
                                            'font': 'Calibri', 
                                            'font_size': 11,
                                            'bg_color': '#ffffff',
                                            'border': True, 'border_color': '#cccccc'})
    data_format_num = wb.add_format({'align': 'center',
                                    'valign': 'top',
                                     'font_color': 'black',
                                     'font': 'Calibri', 
                                     'font_size': 11,
                                    'bg_color': '#ffffff',
                                    'border': True, 'border_color': '#cccccc',
                                    'num_format': '#,##0'})
    data_format_per = wb.add_format({'align': 'center',
                                    'valign': 'top',
                                     'font_color': 'black',
                                     'font': 'Calibri', 
                                     'font_size': 11,
                                     'border': True, 'border_color': '#cccccc',
                                     'bg_color': '#ffffff',
                                    'num_format': '0%'})
    data_format_green = wb.add_format({'align': 'center',
                                    'valign': 'top',
                                     'font_color': 'black',
                                     'font': 'Calibri', 
                                     'font_size': 11,
                                       'border': True, 'border_color': '#cccccc',
                                     'bg_color': '#57D9A3',
                                    'num_format': '0%'})
    data_format_yellow = wb.add_format({'align': 'center',
                                    'valign': 'top',
                                     'font_color': 'black',
                                     'font': 'Calibri', 
                                     'font_size': 11,
                                    'border': True, 'border_color': '#cccccc',
                                     'bg_color': '#FFE380',
                                    'num_format': '0%'})
    data_format_red = wb.add_format({'align': 'center',
                                    'valign': 'top',
                                     'font_color': 'black',
                                     'font': 'Calibri', 
                                     'font_size': 11,
                                     'border': True, 'border_color': '#cccccc',
                                     'bg_color': '#DE350B',
                                    'num_format': '0%'})
    data_format_blue = wb.add_format({'align': 'center',
                                    'valign': 'top',
                                     'font_color': 'black',
                                     'font': 'Calibri', 
                                     'font_size': 11,
                                      'border': True, 'border_color': '#cccccc',
                                     'bg_color': '#2684FF',
                                    'num_format': '0%'})

    data_list = stoplight_data.values.ravel()
    for i in range(num_rows):
        for col in range(1,17):
            cell_data = data_list[16*i + col - 1]
            # Sprints 1-6
            if col in range(1,13):
                # Col is <= current sprint and is actual/projected
                if (col <= sprint*2) & (col % 2 == 0):
                    corr_BL_data = data_list[16*i + col - 2]
                    diff = cell_data - corr_BL_data
                    # On track/ahead
                    if diff > -0.05:
                        data_format = data_format_green
                    # Slightly off track
                    if (diff <= -0.05) & (diff > -0.1):
                        data_format = data_format_yellow
                    # Off track
                    if (diff <= -0.1):
                        data_format = data_format_red
                else:  
                    data_format = data_format_per
            # Points metrics
            elif col == 13:
                data_format_num_delta
            else:
                data_format = data_format_num
            ws.write(i+2, col, cell_data, data_format)
            
    return


def create_excel(data, sprint, stoplight_dir, 
                 newDataFile, prevDataFile, baseDataFile):
    letters = string.ascii_uppercase
    excelFile = os.path.join(stoplight_dir,
                             f'Ground_Dev_ART_STOPLIGHT_{dt.datetime.now().strftime("%y%m%d_%H%M%S")}.xlsx')
    writer = pd.ExcelWriter(excelFile,
                            engine='xlsxwriter')   
    wb = writer.book
    for sheet in data.keys():
        ws = wb.add_worksheet(sheet)
        writer.sheets[sheet] = ws
        data[sheet]['Pivot'].to_excel(writer, sheet_name=sheet, startrow=1 , startcol=0) 
        num_epics = data[sheet]['Pivot'].shape[0]
        num_cols_sum = data[sheet]["Pivot"].shape[1]
        num_cols_cum = data[sheet]["Cum Sum"].shape[1]
        
        last_col_sum = letters[num_cols_sum]
        last_col_cum = letters[num_cols_cum]

        cumsum_startrow = num_epics + 4
        data[sheet]['Cum Sum'].to_excel(writer, sheet_name=sheet, startrow=cumsum_startrow, startcol=0) 
        cumper_startrow = num_epics*2+6
        data[sheet]['Cum Per'].to_excel(writer, sheet_name=sheet, startrow=cumper_startrow, startcol=0) 
        clin_startrow = num_epics*3+7
        data[sheet]['CLIN Per'].to_excel(writer, sheet_name=sheet, startrow=clin_startrow, startcol=0) 

        # header format
        title_format = wb.add_format({'font': 'Calibri', 
                                      'font_color': 'white',
                                     'font_size': 12, 
                                      'bold': True,
                                     'bg_color': '#003cb3',
                                     'align': 'center',
                                     'border': True,
                                     'border_color': 'black'})

        # Pergentage format
        percent_format = wb.add_format({'num_format': '0%'})
        
        ws.conditional_format(f'B{cumper_startrow+2}:{last_col_cum}{cumper_startrow+2+num_epics}', 
                                {'type': 'no_errors',
                                'format': percent_format})
        ws.conditional_format(f'B{clin_startrow+2}:{last_col_cum}{clin_startrow+4}', 
                                {'type': 'no_errors',
                                'format': percent_format})

        # Extra tables
        if (sheet == 'Current Pivot') | (sheet == 'Previous Pivot'):
            num_cols_sprint = data[sheet]["Sprint Metrics"].shape[1]
            first_col_change = letters[num_cols_sum+2]
            last_col_sprint = letters[num_cols_sum+num_cols_sprint+2]

            # Sprint table
            ws.merge_range(f'{first_col_change}{cumper_startrow}:{last_col_sprint}{cumper_startrow}',
                       f'Sprint {sprint}', title_format)
            
            data[sheet]['Sprint Metrics'].to_excel(writer, sheet_name=sheet,
                                                   startrow=cumper_startrow, startcol=10) 
            ws.set_column(10, 10, 60)

            # Round format
            round_format = wb.add_format({'num_format': '#,##0'})
            ws.conditional_format(f'{letters[num_cols_sum+num_cols_sprint]}{cumper_startrow+2}:{last_col_sprint}{cumper_startrow+2+num_epics}',
                                    {'type': 'no_errors',
                                    'format': round_format})
 
            if sheet == 'Current Pivot':
                num_cols_change = data[sheet]['Changes Since Last Week'].shape[1]
                last_col_change = letters[num_cols_sum+2+num_cols_change]
                # Changes since last week
                data[sheet]['Changes Since Last Week'].to_excel(writer, sheet_name=sheet, 
                                                                startrow=1 , startcol=10)   
                ws.set_column(11, 20, 15)
                
                # Add header
                ws.merge_range(f'{first_col_change}1:{last_col_change}1',
                               'Changes Since Last Week', title_format)
                
                # Add cell formatting to changes since last week
                red_format = wb.add_format({'bg_color': '#FFC7CE',
                                           'font_color': '#9C0006'})
                last_row = data[sheet]['Changes Since Last Week'].shape[0] + 2
                ws.conditional_format(f'{letters[num_cols_sum+3]}3:{last_col_change}{last_row}', 
                                        {'type': 'cell',
                                        'criteria': '!=',
                                        'value': 0,
                                        'format': red_format})

        # Column widths and formats
        cat_format = wb.add_format({'align': 'left',
                                   'bold': False})
        ws.set_column(0, 0, 60, cat_format)
        ws.set_column(1, 9, 10, cat_format)

        # Merge and add headers
        letters = string.ascii_uppercase
        
        ws.merge_range(f'A1:{last_col_sum}1',
                       'Sum of Story Points', title_format)
        ws.merge_range(f'A{cumsum_startrow}:{last_col_cum}{cumsum_startrow}',
                       'Cumulative', title_format)
        ws.merge_range(f'A{cumper_startrow}:{last_col_cum}{cumper_startrow}',
                       'Percentage', title_format)
        

    sheets = ['Current Jira Export', 'Previous Jira Export', 'Baseline Jira Export']
    files = [newDataFile, prevDataFile, baseDataFile]
    for (file, sheet) in zip(files, sheets):
        df = pd.read_excel(file, usecols='A:P')
        df.to_excel(writer, sheet_name=sheet, index=False)   
    wb.close()
    
    return