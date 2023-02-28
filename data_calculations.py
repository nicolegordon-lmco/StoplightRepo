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

def split_level(df, filter_condition, apply_func, fill_na):
    """Function to split into levels by Summary""" 
    feature_dict = (df[filter_condition] [['Index', 'Summary']]
                     .set_index('Index')
                     .to_dict()['Summary'])
    new_series = (df.Index.str.split('.').apply(apply_func)
                   .map(feature_dict)
                   .fillna(fill_na))
    return new_series

def split_level_key(df, filter_condition, apply_func, fill_na):
    """Function to split into levels by Key""" 
    feature_dict = (df[filter_condition] [['Index', 'Key']]
                     .set_index('Index')
                     .to_dict()['Key'])
    new_series = (df.Index.str.split('.').apply(apply_func)
                   .map(feature_dict)
                   .fillna(fill_na))
    return new_series

def get_PI_Lookup(df, PILookup_df, jira):
    """Get PI dates from excel file"""
    PI_lookup = {}
    for idx, date in zip(df['Index'], df['Planned Start Date']):
        if pd.isna(date):
            PI_lookup[idx] = np.nan
        elif isinstance(date, (float, int)):
            print(f"Invalid PLanned Start Date: {date}. Please check the {jira} Jira export file" \
                    " for invalid dates. Now exiting...")    
            sys.exit()
        else: 
            for i in PILookup_df.index:
                if ((date >= PILookup_df.loc[i, 'Start']) & (date <= PILookup_df.loc[i, 'End'])):
                    PI_lookup[idx] = PILookup_df.loc[i, 'PI']
                    break
    return list(PI_lookup.values())

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

def get_attributes(df, PILookup_df, jira, feature_level=3):
    # Make sure all indexes are valid
    df.Index.apply(lambda x: testIndexes(x, jira))

    # Index Level
    df['Index Level'] = df.Index.str.count("\.") + 1

    # Feature Level
    df['Feature Level'] = feature_level

    # Epic
    df['Epic'] = split_level(df, 
                            df['Issue Type'] == 'Portfolio Epic', 
                            lambda x: x[0], 
                            'No Epic')

    # Capability
    capability_temp = pd.DataFrame({'Index': df.Index,
                                    'Summary':
                                    split_level(df, 
                                    df.Key.apply(lambda x: x[:9]) == "PCM_COOLR", 
                                    lambda x: '.'.join(x[:2]), 
                                    np.nan)})
    capability_fill = split_level(capability_temp, 
                                  df.Key.apply(lambda x: x[:9]) == "PCM_COOLR", 
                                  lambda x: x[0], 
                                  df.Summary)
    capability_temp.Summary.fillna(capability_fill, inplace=True)
    df['Capability'] = capability_temp.Summary

    # ID: Capability
    id_capability_temp = pd.DataFrame({'Index': df.Index,
                                        'Key':
                                        split_level_key(df, 
                                        df.Key.apply(lambda x: x[:9]) == "PCM_COOLR", 
                                        lambda x: '.'.join(x[:2]), 
                                        np.nan)})
    id_capability_fill = split_level_key(id_capability_temp, 
                                          df.Key.apply(lambda x: x[:9]) == "PCM_COOLR", 
                                          lambda x: x[0], 
                                          df.Key)
    id_capability_temp.Key.fillna(id_capability_fill, inplace=True)
    df['ID: Capability'] = id_capability_temp.Key + ": " + df['Capability']

    # Feature
    features_df = df[df['Index Level'] == 3] [['Index', 'Key', 'Summary']]
    features_df['feature'] = features_df.Key + ": " + features_df.Summary
    features_dict = (features_df
                     .set_index('Index')
                     .to_dict()['feature'])
    df['Feature'] = (df.Index.str.split('.').apply(lambda x: '.'.join(x[:3]))
                               .map(features_dict)
                               .fillna(np.nan))

    df['Features'] = df['Key'] + ": " + df['Summary']
    df['Features'].where(df['Index Level'] == df['Feature Level'], np.nan, inplace=True)

    # PI Lookup
    df['PI Lookup'] = get_PI_Lookup(df, PILookup_df, jira)

    # N-Sprint
    df['N-Sprint'] = df.Sprint.str.count(',') + 1

    # PI
    df['PI'] = np.nan
    defined_sprint = (df['N-Sprint'] >= 1)
    pattern = re.compile(r"PI \d{2}\.\d")
    df.loc[defined_sprint, 'PI'] = df.Sprint.apply(find_PI_sprint, args=(pattern,))
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
    pcm_coolr = (df.Key.str.startswith('PCM_COOLR'))
    space = (df.Key.str.startswith('SPACE'))
    team = (df.Key.str.count('_') > 1)
    df.loc[pcm_coolr,"Level"] = 'Solution'
    df.loc[space,"Level"] = 'Portfolio'
    df.loc[team,"Level"] = 'Team'
    df.Level.fillna('ART', inplace=True)
    
    return df
