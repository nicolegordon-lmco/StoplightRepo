#!/usr/bin/env python
# coding: utf-

# imports
import pandas as pd
import numpy as np
import regex as re
import sys

def testIndexes(x, jira):
    """Function to test if Index is valid"""
    if isinstance(x, float): 
        print(f"Invalid Index: {x}. Please check the {jira} Jira export file" \
               " for missing or extra rows and rerun. Now exiting...")
        sys.exit()

def split_level(df, filterCondition, applyFunc, fillNA):
    """Function to split into levels by Summary""" 
    featureDict = (df[filterCondition] [['Index', 'Summary']]
                     .set_index('Index')
                     .to_dict()['Summary'])
    newSeries = (df.Index.str.split('.').apply(applyFunc)
                   .map(featureDict)
                   .fillna(fillNA))
    return newSeries

def split_level_key(df, filterCondition, applyFunc, fillNA):
    """Function to split into levels by Key""" 
    featureDict = (df[filterCondition] [['Index', 'Key']]
                     .set_index('Index')
                     .to_dict()['Key'])
    newSeries = (df.Index.str.split('.').apply(applyFunc)
                   .map(featureDict)
                   .fillna(fillNA))
    return newSeries

def get_PILookup(df, PILookupDf, jira):
    """Get PI dates from excel file"""
    PILookup = {}
    for idx, date in zip(df['Index'], df['Planned Start Date']):
        if pd.isna(date):
            PILookup[idx] = np.nan
        elif isinstance(date, (float, int)):
            print(f"Invalid PLanned Start Date: {date}. Please check the {jira} Jira export file" \
                    " for invalid dates. Now exiting...")    
            sys.exit()
        else: 
            for i in PILookupDf.index:
                if ((date >= PILookupDf.loc[i, 'Start']) & (date <= PILookupDf.loc[i, 'End'])):
                    PILookup[idx] = PILookupDf.loc[i, 'PI']
                    break
    return list(PILookup.values())

def find_PI_sprint(string, pattern, PI=True):
    matches = pattern.findall(str(string))
    if not matches:
        return np.nan
    if PI:
        return matches[-1]
    else:
        return matches[-1][-2:]
    
def find_team(string):
    match_idx = str(string).rfind("PCM_GD_")
    if match_idx == -1:
        return np.nan
    return string[match_idx+7:]

def get_attributes(df, PILookupDf, jira, featureLevel=3):
    # Make sure all indexes are valid
    df.Index.apply(lambda x: testIndexes(x, jira))

    # Index Level
    df['Index Level'] = df.Index.str.count("\.") + 1

    # Feature Level
    df['Feature Level'] = featureLevel

    # Epic
    df['Epic'] = split_level(df, 
                            df['Issue Type'] == 'Portfolio Epic', 
                            lambda x: x[0], 
                            'No Epic')

    # Capability
    capabilityTemp = pd.DataFrame({'Index': df.Index,
                                    'Summary':
                                    split_level(df, 
                                    df.Key.apply(lambda x: x[:9]) == "pcmCoolr", 
                                    lambda x: '.'.join(x[:2]), 
                                    np.nan)})
    capabilityFill = split_level(capabilityTemp, 
                                  df.Key.apply(lambda x: x[:9]) == "pcmCoolr", 
                                  lambda x: x[0], 
                                  df.Summary)
    capabilityTemp.Summary.fillna(capabilityFill, inplace=True)
    df['Capability'] = capabilityTemp.Summary

    # ID: Capability
    idCapabilityTemp = pd.DataFrame({'Index': df.Index,
                                        'Key':
                                        split_level_key(df, 
                                        df.Key.apply(lambda x: x[:9]) == "pcmCoolr", 
                                        lambda x: '.'.join(x[:2]), 
                                        np.nan)})
    idCapabilityFill = split_level_key(idCapabilityTemp, 
                                          df.Key.apply(lambda x: x[:9]) == "pcmCoolr", 
                                          lambda x: x[0], 
                                          df.Key)
    idCapabilityTemp.Key.fillna(idCapabilityFill, inplace=True)
    df['ID: Capability'] = idCapabilityTemp.Key + ": " + df['Capability']

    # Feature
    featuresDf = df[df['Index Level'] == 3] [['Index', 'Key', 'Summary']]
    featuresDf['feature'] = featuresDf.Key + ": " + featuresDf.Summary
    featuresDict = (featuresDf
                     .set_index('Index')
                     .to_dict()['feature'])
    df['Feature'] = (df.Index.str.split('.').apply(lambda x: '.'.join(x[:3]))
                               .map(featuresDict)
                               .fillna(np.nan))

    df['Features'] = df['Key'] + ": " + df['Summary']
    df['Features'].where(df['Index Level'] == df['Feature Level'], np.nan, inplace=True)

    # PI Lookup
    df['PI Lookup'] = get_PILookup(df, PILookupDf, jira)

    # N-Sprint
    df['N-Sprint'] = df.Sprint.str.count(',') + 1

    # PI
    df['PI'] = np.nan
    definedSprint = (df['N-Sprint'] >= 1)
    pattern = re.compile(r"PI \d{2}\.\d")
    df.loc[definedSprint, 'PI'] = df.Sprint.apply(find_PI_sprint, args=(pattern,))
    backlog = df.Sprint.str.startswith('Backlog').fillna(False)
    df.loc[backlog, 'PI'] = ['Backlog'] * backlog.sum()
    df['PI'].fillna(df['PI Lookup'], inplace=True)

    # Sprint Num
    pattern = re.compile(r"PI \d{2}\.\d - S\d")
    df['Sprint Num'] = df.Sprint.apply(find_PI_sprint, args=(pattern, False,))

    # PI-Sprint
    pattern = re.compile(r"\d{2}\.\d")
    pi = df.PI.apply(find_PI_sprint, args=(pattern, ))
    df['PI-Sprint'] = pi + "-" + df['Sprint Num'].fillna('')
    df.loc[backlog, 'PI-Sprint'] = ['Backlog'] * backlog.sum()

    # Team
    df['Team'] = df.Sprint.apply(find_team)

    # Level
    df["Level"] = np.nan
    pcmCoolr = (df.Key.str.startswith('pcmCoolr'))
    space = (df.Key.str.startswith('SPACE'))
    team = (df.Key.str.count('_') > 1)
    df.loc[pcmCoolr,"Level"] = 'Solution'
    df.loc[space,"Level"] = 'Portfolio'
    df.loc[team,"Level"] = 'Team'
    df.Level.fillna('ART', inplace=True)
    
    return df
