import pandas as pd
import numpy as np
import sys
import re
import shutil
import os
import string
from format import formats

class Pivot:
    def exit(self):
        if os.path.exists(self.stoplightDir):
            shutil.rmtree(self.stoplightDir)
        sys.exit()
    def __init__(self, jiraFile, PILookupFile, epics, clins, PI, jira, stoplightDir):
        self.JiraDf = pd.read_excel(jiraFile, usecols='A:P')
        self.stoplightDir = stoplightDir

        # Check if Jira file is empty. If yes, quit
        if self.JiraDf.size == 0:
            print(f"{jira} Jira file is empty. Please check the export file. " \
                    "Now exiting...")
            if os.path.exists(self.stoplightDir):
                shutil.rmtree(self.stoplightDir)
            sys.exit()
        self.PILookupDf = pd.read_excel(PILookupFile, 
                                        sheet_name = 'PI Lookup', 
                                        parse_dates=['Start', 'End'])
        self.epics = epics
        self.clins = clins
        self.PI = PI
        self.jira = jira
        self.sheetPivot = f"{self.jira} Pivot"
        self.sheetJira = f"{self.jira} Jira Export"

        self.clean_data()
        self.set_attributes()
        self.pivotTable = self.get_pivot()
        return
    
    def clean_data(self):
        # Fill dates for start and end
        self.JiraDf['Planned Start Date'].fillna(method='pad', inplace=True)
        epic = (self.JiraDf['Issue Type'] == 'Portfolio Epic')
        self.JiraDf.loc[epic, 'Planned Start Date'] = pd.NaT
        self.JiraDf['Planned End Date'].fillna(method='pad', inplace=True)
        self.JiraDf.loc[epic, 'Planned End Date'] = pd.NaT
        return

    def set_attributes(self, featureLevel=3):
        # Make sure all indexes are valid
        self.JiraDf.Index.apply(lambda x: self.testIndexes(x))

        # Index Level
        self.JiraDf['Index Level'] = self.JiraDf.Index.str.count("\.") + 1

        # Feature Level
        self.JiraDf['Feature Level'] = featureLevel

        # Epic
        self.JiraDf['Epic'] = self.split_level(self.JiraDf, 
                                               self.JiraDf['Issue Type'] == 'Portfolio Epic', 
                                               lambda x: x[0], 
                                               'No Epic')

        # Capability
        capabilityTemp = pd.DataFrame({'Index': self.JiraDf.Index,
                                       'Summary':
                                            self.split_level(self.JiraDf, 
                                                            self.JiraDf.Key.apply(lambda x: x[:9]) == "pcmCoolr", 
                                                            lambda x: '.'.join(x[:2]), 
                                                            np.nan)})
        capabilityFill = self.split_level(capabilityTemp, 
                                          self.JiraDf.Key.apply(lambda x: x[:9]) == "pcmCoolr", 
                                          lambda x: x[0], 
                                          self.JiraDf.Summary)
        capabilityTemp.Summary.fillna(capabilityFill, inplace=True)
        self.JiraDf['Capability'] = capabilityTemp.Summary

        # ID: Capability
        idCapabilityTemp = pd.DataFrame({'Index': self.JiraDf.Index,
                                         'Key':
                                            self.split_level(self.JiraDf, 
                                                            self.JiraDf.Key.apply(lambda x: x[:9]) == "pcmCoolr", 
                                                            lambda x: '.'.join(x[:2]), 
                                                            np.nan,
                                                            split_by='Key')})
        idCapabilityFill = self.split_level(idCapabilityTemp, 
                                            self.JiraDf.Key.apply(lambda x: x[:9]) == "pcmCoolr", 
                                            lambda x: x[0], 
                                            self.JiraDf.Key,
                                            split_by='Key')
        idCapabilityTemp.Key.fillna(idCapabilityFill, inplace=True)
        self.JiraDf['ID: Capability'] = idCapabilityTemp.Key + ": " + self.JiraDf['Capability']

        # Feature
        featuresDf = self.JiraDf[self.JiraDf['Index Level'] == 3] [['Index', 'Key', 'Summary']]
        featuresDf['feature'] = featuresDf.Key + ": " + featuresDf.Summary
        featuresDict = (featuresDf
                        .set_index('Index')
                        .to_dict()['feature'])
        self.JiraDf['Feature'] = (self.JiraDf.Index.str
                                    .split('.')
                                    .apply(lambda x: '.'.join(x[:3]))
                                    .map(featuresDict)
                                    .fillna(np.nan))

        self.JiraDf['Features'] = self.JiraDf['Key'] + ": " + self.JiraDf['Summary']
        self.JiraDf['Features'].where(self.JiraDf['Index Level'] == self.JiraDf['Feature Level'], 
                                        np.nan, inplace=True)

        # PI Lookup
        self.JiraDf['PI Lookup'] = self.get_PILookup()

        # N-Sprint
        self.JiraDf['N-Sprint'] = self.JiraDf.Sprint.str.count(',') + 1

        # PI
        self.JiraDf['PI'] = np.nan
        definedSprint = (self.JiraDf['N-Sprint'] >= 1)
        pattern = re.compile(r"PI \d{2}\.\d")
        self.JiraDf.loc[definedSprint, 'PI'] = self.JiraDf.Sprint.apply(self.get_PI_sprint, 
                                                                        args=(pattern,))
        backlog = self.JiraDf.Sprint.str.startswith('Backlog').fillna(False)
        self.JiraDf.loc[backlog, 'PI'] = ['Backlog'] * backlog.sum()
        self.JiraDf['PI'].fillna(self.JiraDf['PI Lookup'], inplace=True)

        # Sprint Num
        pattern = re.compile(r"PI \d{2}\.\d - S\d")
        self.JiraDf['Sprint Num'] = self.JiraDf.Sprint.apply(self.get_PI_sprint, 
                                                            args=(pattern, False,))

        # PI-Sprint
        pattern = re.compile(r"\d{2}\.\d")
        pi = self.JiraDf.PI.apply(self.get_PI_sprint, args=(pattern, ))
        self.JiraDf['PI-Sprint'] = pi + "-" + self.JiraDf['Sprint Num'].fillna('')
        self.JiraDf.loc[backlog, 'PI-Sprint'] = ['Backlog'] * backlog.sum()

        # Team
        self.JiraDf['Team'] = self.JiraDf.Sprint.apply(self.get_team)

        # Level
        self.JiraDf["Level"] = np.nan
        pcmCoolr = (self.JiraDf.Key.str.startswith('pcmCoolr'))
        space = (self.JiraDf.Key.str.startswith('SPACE'))
        team = (self.JiraDf.Key.str.count('_') > 1)
        self.JiraDf.loc[pcmCoolr,"Level"] = 'Solution'
        self.JiraDf.loc[space,"Level"] = 'Portfolio'
        self.JiraDf.loc[team,"Level"] = 'Team'
        self.JiraDf.Level.fillna('ART', inplace=True)
        return

    def get_pivot(self, df=None):
        if df is None:
            df = self.JiraDf
        # Filters 
        PIFilter = (df.PI == f"PI {self.PI}")
        levelFilter = (df.Level == "Team")
        issueTypeFilter = ((df['Issue Type'] == "Enabler") 
                        | (df['Issue Type'] == "Story"))
        epicFilter = df.Epic.apply(lambda x: x in self.epics)
        filters = (PIFilter & levelFilter & issueTypeFilter & epicFilter)

        # Pivot table
        dfFiltered = df.copy()
        dfFiltered = dfFiltered[filters]
        marginsName = 'Grand Total'

        # If no stories have slipped, add row of 0s to prevent error
        if dfFiltered.shape[0] == 0:
            dfFiltered.loc[0] = 0
        dfFiltered['Σ Story Points'].fillna(0, inplace=True)

        summaryPivot = dfFiltered.pivot_table(values='Σ Story Points', 
                                                index='Epic', 
                                                columns='PI-Sprint', 
                                                aggfunc='sum',
                                                margins=True,
                                                margins_name=marginsName)

        # Need to add epics that don't have slip points
        for epic in self.epics:
            if epic not in summaryPivot.index.values:
                summaryPivot.loc[epic] = 0

        pivotOrder = self.epics + [marginsName]
        summaryPivot = summaryPivot.loc[pivotOrder, :].fillna(0)
        return summaryPivot

    def testIndexes(self, x):
        """Function to test if Index is valid"""
        if isinstance(x, float): 
            print(f"Invalid Index: {x}. Please check the {self.jira} Jira " \
                "export file for missing or extra rows and rerun. " \
                    "Now exiting...")
            if os.path.exists(self.stoplightDir):
                shutil.rmtree(self.stoplightDir)
            sys.exit()

    def split_level(self, df, filterCondition, applyFunc, fillNA, split_by='Summary'):
        """Function to split into levels by Summary""" 
        featureDict = (df[filterCondition][['Index', split_by]]
                        .set_index('Index')
                        .to_dict()[split_by])
        newSeries = (df.Index.str.split('.').apply(applyFunc)
                    .map(featureDict)
                    .fillna(fillNA))
        return newSeries
    
    def get_PILookup(self):
        """Get PI dates from excel file"""
        PILookup = {}
        for idx, date in zip(self.JiraDf['Index'], self.JiraDf['Planned Start Date']):
            if pd.isna(date):
                PILookup[idx] = np.nan
            elif isinstance(date, (float, int)):
                print(f"Invalid PLanned Start Date: {date}. " \
                    f"Please check the {self.jira} Jira export file" \
                    " for invalid dates. Now exiting...")   
                if os.path.exists(self.stoplightDir):
                    shutil.rmtree(self.stoplightDir) 
                sys.exit()
            else: 
                for i in self.PILookupDf.index:
                    if ((date >= self.PILookupDf.loc[i, 'Start']) 
                        & (date <= self.PILookupDf.loc[i, 'End'])):
                        PILookup[idx] = self.PILookupDf.loc[i, 'PI']
                        break
        return list(PILookup.values())
    
    def get_PI_sprint(self, string, pattern, PI=True):
        matches = pattern.findall(str(string))
        if not matches:
            return np.nan
        if PI:
            return matches[-1]
        else:
            return matches[-1][-2:]
        
    def get_team(self, string):
        match_idx = str(string).rfind("PCM_GD_")
        if match_idx == -1:
            return np.nan
        return string[match_idx+7:]
    
    def set_slip(self, prevJiraDf, baselineDf):
        # Only consider current PI stories
        curDf = self.JiraDf[(self.JiraDf.PI == f"PI {self.PI}")]
        prevJiraDf = prevJiraDf[(prevJiraDf.PI == f"PI {self.PI}")]
        baselineDf = baselineDf[(baselineDf.PI == f"PI {self.PI}")]

        # Keys in each df
        curKey = curDf.Key
        prevKey = prevJiraDf.Key
        baseKey = baselineDf.Key

        # Keys that have slipped from previous or baseline
        prevSlip = prevJiraDf[(~prevKey.isin(curKey))]
        baseSlip = baselineDf[(~baseKey.isin(curKey))]

        # Start slip df with slips from previous Jira df
        slipDf = prevSlip.copy()
        # Add in slips from the baseline
        for idx in baseSlip.index:
            if baseSlip.loc[idx, 'Key'] in prevSlip.Key.values:
                continue
            else:
                slipDf = pd.concat((slipDf, 
                                    pd.DataFrame(baseSlip.loc[idx]).transpose()), 
                                    ignore_index=True)
        self.slipDf = slipDf
        self.slipDfPrev = prevSlip
        self.slipDfBaseline = baseSlip
        self.slipPivotTable = self.get_pivot(df=self.slipDf)
        self.pivotTable['Slip'] = self.slipPivotTable['Grand Total']
        return

    def set_new(self, prevJiraDf, baselineDf):
        # Only consider current PI stories
        curDf = self.JiraDf[(self.JiraDf.PI == f"PI {self.PI}")]
        prevJiraDf = prevJiraDf[(prevJiraDf.PI == f"PI {self.PI}")]
        baselineDf = baselineDf[(baselineDf.PI == f"PI {self.PI}")]

        # Keys in each df
        curKey = curDf.Key
        prevKey = prevJiraDf.Key
        baseKey = baselineDf.Key

        # New keys not in previous or baseline
        new = curDf[(~curKey.isin(prevKey) & ~curKey.isin(baseKey))]
        self.newDf = new
        return

    def set_weekly_change(self, prevPivot):
        # Changes since last week
        pattern = re.compile(r'\d{2}\.\d-S\d')
        changesSinceLastWeek = self.pivotTable - prevPivot
        cols = (changesSinceLastWeek.columns.to_series()
                .apply(self.get_PI_sprint, args=(pattern,))
                .dropna().values)
        cols = np.append(cols, ['Grand Total'])
        changesSinceLastWeek = changesSinceLastWeek.loc[self.epics, 
                                                        cols]
        self.changesWeek = changesSinceLastWeek
        return

    def get_cumsum(self, pivot):
        return pivot.cumsum(axis=1)

    def get_cumper(self):
        # Total points in assigned sprint columns will be last column of 
        # the cumulative sum df plus the slipped points
        if "Slip" in self.pivotTable.columns:
            totals = (self.cumSum.iloc[:, -1] 
                      + self.pivotTable.Slip.drop('Grand Total'))
        else:
            totals = self.cumSum.iloc[:, -1]

        # Reshape
        totals = (totals
                .values
                .repeat(self.cumSum.shape[1])
                .reshape(self.cumSum.shape))
        cumPer = (self.cumSum / totals)
        return cumPer
    
    @staticmethod
    def get_clin(df, clin):
        clinIdx = df.index.to_series().apply(lambda x: clin in x)
        return df[clinIdx]
    
    def get_clinPer(self, clin):
        # Only use assigned sprint columns, plus slip if applicable
        curClin = self.get_clin(self.cumSum, clin)
        if 'Slip' in self.pivotTable.columns:
            curClinSlip = self.get_clin(self.pivotTable, clin).Slip
            total = (curClin.iloc[:, -1] + curClinSlip).values.sum()
        else:
            total = curClin.iloc[:, -1].values.sum() 
        clinPer = self.get_clin(self.cumSum, clin).sum() / total
        return clinPer

    def set_cum_metrics(self):
        # Only use actual sprint data
        pattern = re.compile(r'\d{2}\.\d-S\d')
        cols = (self.pivotTable.columns
                .to_series().apply(self.get_PI_sprint, 
                                    args=(pattern,))
                                    .dropna().values)
        pivotSprints = self.pivotTable.copy()
        pivotSprints = pivotSprints.loc[self.epics, cols]
        
        # Points cumulative sum 
        self.cumSum = self.get_cumsum(pivotSprints)
        
        # Points cumulative percentage
        self.cumPer = self.get_cumper()
        
        # CLIN breakout
        CLINDf = pd.DataFrame()
        for clin in self.clins:
            clinPer = self.get_clinPer(clin)
            CLINDf[clin] = clinPer
        CLINDf = CLINDf.transpose()
        self.clinDf = CLINDf
        return

    def set_sprint_metrics(self, curSprint, lastCompleteSprint,
                            baselineCumSum, baselineCumPer):
        # Set last complete sprint
        self.curSprint = curSprint
        self.lastCompleteSprint = lastCompleteSprint

        # Current total (Points in sprint plus slip)
        cumSumTot = self.cumSum.iloc[:, -1].values 
        slip = self.pivotTable.Slip.drop(['Grand Total'])
        curTotal = cumSumTot + slip
        
        # Change since baseline
        baselineTotal = baselineCumSum.iloc[:, -1].values
        baselineChange = curTotal - baselineTotal
        
        # Points expected
        lastPattern = re.compile(fr'\d\d\.\d-S{lastCompleteSprint}')
        col = (baselineCumPer.columns.to_series()
            .apply(self.get_PI_sprint, args=(lastPattern,))
            .dropna().values[0])
        pointsExpected = (baselineCumPer.loc[self.epics, col] * curTotal)
        
        # Points completed
        col = (self.cumSum.columns.to_series()
            .apply(self.get_PI_sprint, args=(lastPattern,))
            .dropna().values[0])
        pointsCompleted = self.cumSum.loc[self.epics, col]

        # Current completed
        curPattern = re.compile(fr'\d\d\.\d-S{curSprint}')
        col = (self.cumSum.columns.to_series()
            .apply(self.get_PI_sprint, args=(curPattern,))
            .dropna().values[0])
        curPointsCompleted = self.cumSum.loc[self.epics, col]
        
        # Delta
        delta = pointsCompleted - pointsExpected

        # Points remaining
        pointsRem = curTotal - pointsCompleted

        # Velocity
        vel = pointsCompleted / lastCompleteSprint

        # Sprints left
        sprintsRem = pointsRem / vel
        
        # Data frame
        sprintMetricsDf = pd.DataFrame({'Current Total Pts': curTotal,
                                        'Change Since BL': baselineChange,
                                        'Points Expected': pointsExpected,
                                        'Points Completed': pointsCompleted,
                                        'Delta Points': delta,
                                        'Current Completed': curPointsCompleted})
        remainingSprintMetrics = pd.DataFrame({'Points Remaining': pointsRem,
                                                'Velocity': vel, 
                                                'Sprints Remaining': sprintsRem})
        self.sprintMetrics = sprintMetricsDf
        self.remainingSprintMetrics = remainingSprintMetrics
        return

    def excel_pivot(self, writer):
        letters = string.ascii_uppercase
        wb = writer.book

        ws = wb.add_worksheet(self.sheetPivot)
        writer.sheets[self.sheetPivot] = ws
        self.pivotTable.to_excel(writer, sheet_name=self.sheetPivot, startrow=1 , 
                                    startcol=0, freeze_panes=(0,1)) 
        numEpics = self.pivotTable.shape[0]
        numColsSum = self.pivotTable.shape[1]
        numColsCum = self.cumSum.shape[1]
        
        lastColSum = letters[numColsSum]
        lastColCum = letters[numColsCum]

        cumSumStartRow = numEpics + 4
        self.cumSum.to_excel(writer, sheet_name=self.sheetPivot, 
                                        startrow=cumSumStartRow, startcol=0) 
        cumPerStartRow = numEpics*2+6
        self.cumPer.to_excel(writer, sheet_name=self.sheetPivot, 
                                        startrow=cumPerStartRow, startcol=0) 
        clinStartRow = numEpics*3+7
        self.clinDf.to_excel(writer, sheet_name=self.sheetPivot, 
                                        startrow=clinStartRow, startcol=0) 

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
        if (self.sheetPivot == 'Current Pivot') | (self.sheetPivot == 'Previous Pivot'):
            numColsSprint = self.sprintMetrics.shape[1]
            numColsRem = self.remainingSprintMetrics.shape[1]
            firstColChange = letters[numColsSum+2]
            lastColSprint = letters[numColsSum+numColsSprint+2]
            firstColRem = letters[numColsSum+numColsSprint+4]
            lastColRem = letters[numColsSum+numColsSprint+numColsRem+3]

            # Sprint table and remaining table
            ws.merge_range(f'{firstColChange}{cumPerStartRow}:{lastColSprint}{cumPerStartRow}',
                    f'Sprint {self.lastCompleteSprint}', titleFormat)
            ws.merge_range(f'{firstColRem}{cumPerStartRow}:{lastColRem}{cumPerStartRow}',
                    f'Sprint {self.lastCompleteSprint}', titleFormat)
            
            self.sprintMetrics.to_excel(writer, sheet_name=self.sheetPivot,
                                                startrow=cumPerStartRow, startcol=numColsSum+2) 
            self.remainingSprintMetrics.to_excel(writer, sheet_name=self.sheetPivot,
                                                        startrow=cumPerStartRow, 
                                                        startcol=numColsSum+numColsSprint+4,
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

            if self.sheetPivot == 'Current Pivot':
                numColsChange = self.changesWeek.shape[1]
                lastColChange = letters[numColsSum+2+numColsChange]
                # Changes since last week
                self.changesWeek.to_excel(writer, 
                                                                sheet_name=self.sheetPivot, 
                                                                startrow=1 , 
                                                                startcol=numColsSum+2)  
                
                # Add header
                ws.merge_range(f'{firstColChange}1:{lastColChange}1',
                            'Changes Since Last Week', titleFormat)
                
                # Add cell formatting to changes since last week
                redFormat = wb.add_format(formats['redDelta'])
                lastRow = self.changesWeek.shape[0] + 2
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
        return
    
    def format_keys(self, keys, format, ws):
        for row in range(self.JiraDf.shape[0]):
            key = self.JiraDf.loc[row, 'Key']
            if key in keys:
                ws.conditional_format(f'B{row+2}', 
                                        {'type': 'no_errors',
                                        'format': format})
        return
                
    def excel_Jira(self, writer, cur=None):
        wb = writer.book

        self.JiraDf.to_excel(writer, sheet_name=self.sheetJira, index=False)   

        ws = writer.sheets[self.sheetJira]
        # Add conditional formatting for new and slips
        if self.sheetJira == 'Current Jira Export':
            newFormat = wb.add_format(formats['newStories'])
            newStories = self.newDf.Key.values
            self.format_keys(newStories, newFormat, ws)
        else:
            slipFormat = wb.add_format(formats['slipStories'])
            if self.sheetJira == 'Previous Jira Export':
                slipStories = cur.slipDfPrev.Key.values
            else:
                slipStories = cur.slipDfBaseline.Key.values
            self.format_keys(slipStories, slipFormat, ws)
        
        return

    @staticmethod
    def merge_baseline_cur(baseline, cur, overall=False):
        # Merge baseline and current
        df = pd.merge(baseline, cur, left_index=True, right_index=True, 
                    suffixes=('_BL', '_Cur'))
        # Sort columns
        df = df.reindex(sorted(df.columns), axis=1)
        # Rename overall index
        if overall:
            df.rename(index={df.index[0]:'Overall'}, inplace=True)
        return df

    def set_stoplight_data(self, baseline, clin):
        baselinePer = self.get_clin(baseline.cumPer, clin)
        curPer = self.get_clin(self.cumPer, clin)
        baselineTot = self.get_clin(baseline.clinDf, clin)
        curTot = self.get_clin(self.clinDf, clin)

        sprintsEpics = self.merge_baseline_cur(baselinePer, curPer)
        sprintsTot = self.merge_baseline_cur(baselineTot, curTot, overall=True)

        sprintsData = pd.concat((sprintsEpics, sprintsTot))
        
        pointsEpics = self.get_clin(self.sprintMetrics, clin)
        pointsTot = pd.DataFrame({'Overall': pointsEpics.sum()}).transpose()
        pointsData = pd.concat((pointsEpics, pointsTot))

        remEpics = self.get_clin(self.remainingSprintMetrics, clin)
        remTot = (pd.DataFrame({'Overall': [np.nan] * remEpics.shape[1]})
                .transpose().rename(columns={0: remEpics.columns[0],
                                                1: remEpics.columns[1],
                                                2: remEpics.columns[2]}))
        remData = pd.concat((remEpics, remTot))

        pointsData = pd.concat((pointsData, remData), axis=1)
        
        stoplightData = pd.concat((sprintsData, pointsData), axis=1)
        
        for idx in stoplightData.index:
            stoplightData.rename(index={idx:idx.split(': ')[-1]}, inplace=True)
            
        changeSinceBL = stoplightData[['Change Since BL']].copy()
        stoplightData.drop(columns='Change Since BL', inplace=True)

        self.changeBL = changeSinceBL
        self.stoplightData = stoplightData

        return
        

    def create_stoplight_sheet(self, wb, clin):
        letters = string.ascii_uppercase
        stoplightData = self.stoplightDict[clin]['Data'].copy()
        stoplightChange = self.stoplightDict[clin]['Change_BL'].copy()
        PISprintCompleted = f"{stoplightData.columns[(self.lastCompleteSprint-1) * 2][:4]}.{self.lastCompleteSprint}"
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
                if (col/2) <= self.curSprint:
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
                    if (col <= self.curSprint*2) & (col % 2 == 0):
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