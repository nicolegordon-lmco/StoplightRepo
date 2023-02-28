#!/usr/bin/env python
# coding: utf-

# imports
import pandas as pd
from data_calculations import *

def clean_data(df):
    # Fill dates for start and end
    df['Planned Start Date'].fillna(method='pad', inplace=True)
    epic = (df['Issue Type'] == 'Portfolio Epic')
    df.loc[epic, 'Planned Start Date'] = pd.NaT
    df['Planned End Date'].fillna(method='pad', inplace=True)
    df.loc[epic, 'Planned End Date'] = pd.NaT
    return df

def create_pivot(df, epics, PI, slip=False):
    # Filters 
    PIFilter = (df.PI == "PI 23.1")
    levelFilter = (df.Level == "Team")
    issueTypeFilter = ((df['Issue Type'] == "Enabler") | (df['Issue Type'] == "Story"))
    epicFilter = df.Epic.apply(lambda x: x in epics)
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
        for epic in epics:
            if epic not in summaryPivot.index.values:
                summaryPivot.loc[epic] = 0

    pivotOrder = epics + [marginsName]
    summaryPivot = summaryPivot.loc[pivotOrder, :].fillna(0)
    return summaryPivot


def pivot_from_df(df, PILookupDf, epics, PI, slip=False, jira=None):
    # If slipDf, data is already cleaned and attributes added
    # Only need to create the pivot
    if slip: 
        pivotTable = create_pivot(df, epics, PI, slip)
        return pivotTable, df
    df = clean_data(df)
    df = get_attributes(df, PILookupDf, jira)
    pivotTable = create_pivot(df, epics, PI)
    return pivotTable, df

def get_slip(curJiraDf, prevJiraDf, baselineDf):
    # Keys in each df
    curKey = curJiraDf.Key
    prevKey = prevJiraDf.Key
    baseKey = baselineDf.Key

    # Keys that have slipped from previous or baseline
    prevSlip = prevJiraDf[(~prevKey.isin(curKey))]
    baseSlip = baselineDf[(~baseKey.isin(curKey))]

    # Start slip df with slips from previous Jira df
    slip = prevSlip.copy()
    # Add in slips from the baseline
    for idx in baseSlip.index:
        if baseSlip.loc[idx, 'Key'] in prevSlip.Key.values:
            continue
        else:
            slip = pd.concat((slip, pd.DataFrame(baseSlip.loc[idx]).transpose()), ignore_index=True)
    return slip

def get_new(curJiraDf, prevJiraDf, baselineDf):
    # Keys in each df
    curKey = curJiraDf.Key
    prevKey = prevJiraDf.Key
    baseKey = baselineDf.Key

    # New keys not in previous or baseline
    new = curJiraDf[(~curKey.isin(prevKey) & ~curKey.isin(baseKey))]
    return new
    
def all_pivots(newDataFile, prevDataFile, baseDataFile, PILookupFile, epics, PI):
    # Import new and previous data
    curJiraDf = pd.read_excel(newDataFile, usecols='A:P')
    prevJiraDf = pd.read_excel(prevDataFile, usecols='A:P')
    baselineDf = pd.read_excel(baseDataFile, usecols='A:P')
    PILookupDf = pd.read_excel(PILookupFile, sheet_name = 'PI Lookup', parse_dates=['Start', 'End'])
    
    # Create pivots
    curPivot, curJiraDf = pivot_from_df(curJiraDf, PILookupDf, epics, PI, jira="Current")
    prevPivot, prevJiraDf = pivot_from_df(prevJiraDf, PILookupDf, epics, PI, jira="Previous")
    baselinePivot, baselineDf = pivot_from_df(baselineDf, PILookupDf, epics, PI, jira="Baseline")

    # Add slip
    curSlipDf = get_slip(curJiraDf, prevJiraDf, baselineDf)
    curSlipPivot, curSlipDf = pivot_from_df(curSlipDf, PILookupDf, epics, PI, slip=True)
    curPivot['Slip'] = curSlipPivot['Grand Total']

    prevSlipDf = get_slip(prevJiraDf, prevJiraDf, baselineDf)
    prevSlipPivot, prevSlipDf = pivot_from_df(prevSlipDf, PILookupDf, epics, PI, slip=True)
    prevPivot['Slip'] = prevSlipPivot['Grand Total']

    # Get new stories
    curNewDf = get_new(curJiraDf, prevJiraDf, baselineDf)

    return curPivot, prevPivot, baselinePivot, curSlipDf, curNewDf