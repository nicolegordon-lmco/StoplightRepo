import pandas as pd
import numpy as np
import sys
import re
import string
from format import formats

class Pivot:
    def __init__(self, jiraFile, PILookupFile, epics, clins, PI, jira):
        self.JiraDf = pd.read_excel(jiraFile, usecols='A:P')
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
    
    def clean_data(self):
        # Fill dates for start and end
        self.JiraDf['Planned Start Date'].fillna(method='pad', inplace=True)
        epic = (self.JiraDf['Issue Type'] == 'Portfolio Epic')
        self.JiraDf.loc[epic, 'Planned Start Date'] = pd.NaT
        self.JiraDf['Planned End Date'].fillna(method='pad', inplace=True)
        self.JiraDf.loc[epic, 'Planned End Date'] = pd.NaT

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

    def get_pivot(self, df=None, slip=False):
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

        # If slip pivot, need to add epics that don't have slip points
        if slip:
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
        # Keys in each df
        curKey = self.JiraDf.Key
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
        self.slipPivotTable = self.get_pivot(df=self.slipDf, slip=True)
        self.pivotTable['Slip'] = self.slipPivotTable['Grand Total']

    def set_new(self, prevJiraDf, baselineDf):
        # Keys in each df
        curKey = self.JiraDf.Key
        prevKey = prevJiraDf.Key
        baseKey = baselineDf.Key

        # New keys not in previous or baseline
        new = self.JiraDf[(~curKey.isin(prevKey) & ~curKey.isin(baseKey))]
        self.newDf = new

    def set_weekly_change(self, prevPivot):
        # Changes since last week
        pattern = re.compile(r'\d{2}\.\d-S\d')
        changesSinceLastWeek = self.pivotTable - prevPivot
        cols = (changesSinceLastWeek.columns.to_series()
                .apply(self.get_PI_sprint, args=(pattern,))
                .dropna().values)
        cols = np.append(cols, ['Slip', 'Grand Total'])
        changesSinceLastWeek = changesSinceLastWeek.loc[self.epics, 
                                                        cols]
        self.changesWeek = changesSinceLastWeek

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

    def set_sprint_metrics(self, curSprint, lastCompleteSprint,
                            baselineCumSum, baselineCumPer):
        # Set last complete sprint
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
        wb.close()
    
    def format_keys(self, keys, format, ws):
        for row in range(self.JiraDf.shape[0]):
            key = self.JiraDf.loc[row, 'Key']
            if key in keys:
                ws.conditional_format(f'B{row+2}', 
                                        {'type': 'no_errors',
                                        'format': format})
                
    def excel_Jira(self, writer):
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
            slipStories = self.slipDfKey.values
            self.format_keysformat_keys(slipStories, slipFormat, ws)
        wb.close()

    