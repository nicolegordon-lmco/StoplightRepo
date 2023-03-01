import pandas as pd
import numpy as np
import sys
import re

class Pivot:
    def __init__(self, jiraFile, PILookupFile, epics, PI, jira=None, slip=False):
        self.JiraDf = pd.read_excel(jiraFile, usecols='A:P')
        self.PILookupDf = pd.read_excel(PILookupFile, 
                                        sheet_name = 'PI Lookup', 
                                        parse_dates=['Start', 'End'])
        self.epics = epics
        self.PI = PI
        self.jira = jira
        self.slip = slip

        # If slipDf, data is already cleaned and attributes added
        # Only need to create the pivot
        if self.slip: 
            self.set_pivot()

        self.clean_data()
        self.set_attributes()
        self.set_pivot()
    
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

    def set_pivot(self):
        # Filters 
        PIFilter = (self.JiraDf.PI == self.PI)
        levelFilter = (self.JiraDf.Level == "Team")
        issueTypeFilter = ((self.JiraDf['Issue Type'] == "Enabler") 
                        | (self.JiraDf['Issue Type'] == "Story"))
        epicFilter = self.JiraDf.Epic.apply(lambda x: x in self.epics)
        filters = (PIFilter & levelFilter & issueTypeFilter & epicFilter)

        # Pivot table
        dfFiltered = self.JiraDf.copy()
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
        if self.slip:
            for epic in self.epics:
                if epic not in summaryPivot.index.values:
                    summaryPivot.loc[epic] = 0

        pivotOrder = self.epics + [marginsName]
        summaryPivot = summaryPivot.loc[pivotOrder, :].fillna(0)
        self.pivotTable = summaryPivot

    def testIndexes(self, x):
        """Function to test if Index is valid"""
        if isinstance(x, float): 
            print(f"Invalid Index: {x}. Please check the {self.jira} Jira " \
                "export file for missing or extra rows and rerun. " \
                    "Now exiting...")
            sys.exit()

    def split_level(df, filterCondition, applyFunc, fillNA, split_by='Summary'):
        """Function to split into levels by Summary""" 
        featureDict = (df[filterCondition] [['Index', split_by]]
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
    
    def get_PI_sprint(string, pattern, PI=True):
        matches = pattern.findall(str(string))
        if not matches:
            return np.nan
        if PI:
            return matches[-1]
        else:
            return matches[-1][-2:]
        
    def get_team(string):
        match_idx = str(string).rfind("PCM_GD_")
        if match_idx == -1:
            return np.nan
        return string[match_idx+7:]