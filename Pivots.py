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

def create_pivot(df, epics, PI):
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
    summary_pivot = df_filtered.pivot_table(values='Σ Story Points', 
                                              index='Epic', 
                                              columns='PI-Sprint', 
                                              aggfunc='sum',
                                              margins=True,
                                              margins_name=margins_name)
    pivot_order = epics + [margins_name]
    summary_pivot = summary_pivot.loc[pivot_order, :].fillna(0)
    return summary_pivot


def pivot_from_df(df, PILookup_df, epics, PI):
    df = clean_data(df)
    df = get_attributes(df, PILookup_df)
    pivot_table = create_pivot(df, epics, PI)
    return pivot_table

def all_pivots(newDataFile, prevDataFile, baseDataFile, PILookupFile, epics, PI):
    # Import new and previous data
    curJira_df = pd.read_excel(newDataFile, usecols='A:P')
    prevJira_df = pd.read_excel(prevDataFile, usecols='A:P')
    baseline_df = pd.read_excel(baseDataFile, usecols='A:P')
    PILookup_df = pd.read_excel(PILookupFile, sheet_name = 'PI Lookup', parse_dates=['Start', 'End'])
    
    # Create pivots
    cur_pivot = pivot_from_df(curJira_df, PILookup_df, epics, PI)
    prev_pivot = pivot_from_df(prevJira_df, PILookup_df, epics, PI)
    baseline_pivot = pivot_from_df(baseline_df, PILookup_df, epics, PI)
    
    return cur_pivot, prev_pivot, baseline_pivot