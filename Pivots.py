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
    PI_filter = (df.PI == "PI 23.1")
    Level_filter = (df.Level == "Team")
    IssueType_filter = ((df['Issue Type'] == "Enabler") | (df['Issue Type'] == "Story"))
    Epic_filter = df.Epic.apply(lambda x: x in epics)
    filters = (PI_filter & Level_filter & IssueType_filter & Epic_filter)

    # Pivot table
    df_filtered = df.copy()
    df_filtered = df_filtered[filters]
    margins_name = 'Grand Total'

    # If no stories have slipped, add row of 0s to prevent error
    if df_filtered.shape[0] == 0:
        df_filtered.loc[0] = 0
    df_filtered['Σ Story Points'].fillna(0, inplace=True)

    summary_pivot = df_filtered.pivot_table(values='Σ Story Points', 
                                              index='Epic', 
                                              columns='PI-Sprint', 
                                              aggfunc='sum',
                                              margins=True,
                                              margins_name=margins_name)

    # If slip pivot, need to add epics that don't have slip points
    if slip:
        for epic in epics:
            if epic not in summary_pivot.index.values:
                summary_pivot.loc[epic] = 0

    pivot_order = epics + [margins_name]
    summary_pivot = summary_pivot.loc[pivot_order, :].fillna(0)
    return summary_pivot


def pivot_from_df(df, PILookup_df, epics, PI, slip=False, jira=None):
    # If slip_df, data is already cleaned and attributes added
    # Only need to create the pivot
    if slip: 
        pivot_table = create_pivot(df, epics, PI, slip)
        return pivot_table, df
    df = clean_data(df)
    df = get_attributes(df, PILookup_df, jira)
    pivot_table = create_pivot(df, epics, PI)
    return pivot_table, df

def get_slip(curJira_df, prevJira_df, baseline_df):
    # Keys in each df
    curKey = curJira_df.Key
    prevKey = prevJira_df.Key
    baseKey = baseline_df.Key

    # Keys that have slipped from previous or baseline
    prevSlip = prevJira_df[(~prevKey.isin(curKey))]
    baseSlip = baseline_df[(~baseKey.isin(curKey))]

    # Start slip df with slips from previous Jira df
    slip = prevSlip.copy()
    # Add in slips from the baseline
    for idx in baseSlip.index:
        if baseSlip.loc[idx, 'Key'] in prevSlip.Key.values:
            continue
        else:
            slip = pd.concat((slip, pd.DataFrame(baseSlip.loc[idx]).transpose()), ignore_index=True)
    return slip
    
def all_pivots(newDataFile, prevDataFile, baseDataFile, PILookupFile, epics, PI):
    # Import new and previous data
    curJira_df = pd.read_excel(newDataFile, usecols='A:P')
    prevJira_df = pd.read_excel(prevDataFile, usecols='A:P')
    baseline_df = pd.read_excel(baseDataFile, usecols='A:P')
    PILookup_df = pd.read_excel(PILookupFile, sheet_name = 'PI Lookup', parse_dates=['Start', 'End'])
    
    # Create pivots
    cur_pivot, curJira_df = pivot_from_df(curJira_df, PILookup_df, epics, PI, jira="Current")
    prev_pivot, prevJira_df = pivot_from_df(prevJira_df, PILookup_df, epics, PI, jira="Previous")
    baseline_pivot, baseline_df = pivot_from_df(baseline_df, PILookup_df, epics, PI, jira="Baseline")

    # Add slip
    curSlip_df = get_slip(curJira_df, prevJira_df, baseline_df)
    curSlip_pivot, curSlip_df = pivot_from_df(curSlip_df, PILookup_df, epics, PI, slip=True)
    cur_pivot['Slip'] = curSlip_pivot['Grand Total']

    prevSlip_df = get_slip(prevJira_df, prevJira_df, baseline_df)
    prevSlip_pivot, prevSlip_df = pivot_from_df(prevSlip_df, PILookup_df, epics, PI, slip=True)
    prev_pivot['Slip'] = prevSlip_pivot['Grand Total']

    return cur_pivot, prev_pivot, baseline_pivot