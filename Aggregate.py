#!/usr/bin/env python
# coding: utf-

# imports
import pandas as pd
import numpy as np
import regex as re
from Pivots import *
from data_calculations import find_PI_sprint

def get_cumsum(pivot):
    return pivot.cumsum(axis=1)

def get_cumper(pivot, cum_sum, epics):
    # Total points in assigned sprint columns will be last column of the cumulative sum df
    # plus the slipped points
    if "Slip" in pivot.columns:
        totals = cum_sum.iloc[:, -1] + pivot.Slip.drop('Grand Total')
    else:
        totals = cum_sum.iloc[:, -1]

    # Reshape
    totals = (totals
             .values
             .repeat(cum_sum.shape[1])
             .reshape(cum_sum.shape))
    cum_per = (cum_sum / totals)
    return cum_per

def get_clin(df, clin):
    clin_idx = df.index.to_series().apply(lambda x: clin in x)
    return df[clin_idx]

def get_clin_per(cum_sum, pivot, clin):
    # Only use assigned sprint columns, plus slip if applicable
    cur_clin = get_clin(cum_sum, clin)
    if 'Slip' in pivot.columns:
        cur_clin_slip = get_clin(pivot, clin).Slip
        total = (cur_clin.iloc[:, -1] + cur_clin_slip).values.sum()
    else:
        total = cur_clin.iloc[:, -1].values.sum() 
    clin_per = get_clin(cum_sum, clin).sum() / total
    return clin_per

def get_cum_metrics(pivot, epics, clins):
    # Only use actual sprint data
    pattern = re.compile(r'\d{2}\.\d-S\d')
    cols = pivot.columns.to_series().apply(find_PI_sprint, args=(pattern,)).dropna().values
    pivot_sprints = pivot.copy()
    pivot_sprints = pivot_sprints.loc[epics, cols]
    
    # Points cumulative sum 
    cum_sum = get_cumsum(pivot_sprints)
    
    # Points cumulative percentage
    cum_per = get_cumper(pivot, cum_sum, epics)
    
    # CLIN breakout
    CLIN_df = pd.DataFrame()
    for clin in clins:
        clin_per = get_clin_per(cum_sum, pivot, clin)
        CLIN_df[clin] = clin_per
    CLIN_df = CLIN_df.transpose()
    
    return cum_sum, cum_per, CLIN_df

def get_sprint_metrics(curSprint, lastCompleteSprint, pivot, pivot_cumsum,
                         baseline_cumsum, baseline_cumper, epics):
    # Current total (Points in sprint plus slip)
    cumsum_tot = pivot_cumsum.iloc[:, -1].values 
    slip = pivot.Slip.drop(['Grand Total'])
    cur_total = cumsum_tot + slip
    
    # Change since baseline
    baseline_total = baseline_cumsum.iloc[:, -1].values
    baseline_change = cur_total - baseline_total
    
    # Points expected
    last_pattern = re.compile(fr'\d\d\.\d-S{lastCompleteSprint}')
    col = (baseline_cumper.columns.to_series()
           .apply(find_PI_sprint, args=(last_pattern,))
           .dropna().values[0])
    points_expected = (baseline_cumper.loc[epics, col] * cur_total)
    
    # Points completed
    col = (pivot_cumsum.columns.to_series()
           .apply(find_PI_sprint, args=(last_pattern,))
           .dropna().values[0])
    points_completed = pivot_cumsum.loc[epics, col]

    # Current completed
    cur_pattern = re.compile(fr'\d\d\.\d-S{curSprint}')
    col = (pivot_cumsum.columns.to_series()
           .apply(find_PI_sprint, args=(cur_pattern,))
           .dropna().values[0])
    cur_points_completed = pivot_cumsum.loc[epics, col]
    
    # Delta
    delta = points_completed - points_expected

    # Points remaining
    pts_rem = cur_total - points_completed

    # Velocity
    vel = points_completed / lastCompleteSprint

    # Sprints left
    sprints_rem = pts_rem / vel
    
    # Data frame
    sprint_metrics_df = pd.DataFrame({'Current Total Pts': cur_total,
                                     'Change Since BL': baseline_change,
                                     'Points Expected': points_expected,
                                     'Points Completed': points_completed,
                                     'Delta Points': delta,
                                     'Current Completed': cur_points_completed})
    remaining_sprint_metrics = pd.DataFrame({'Points Remaining': pts_rem,
                                            'Velocity': vel, 
                                            'Sprints Remaining': sprints_rem})

    return sprint_metrics_df, remaining_sprint_metrics

def get_aggregated_data(curSprint, lastCompleteSprint, 
                        newDataFile, prevDataFile, 
                        baseDataFile, PILookupFile,
                        epics, clins, PI):
    # Get pivot tables
    (cur_pivot, prev_pivot, baseline_pivot) = all_pivots(newDataFile, prevDataFile, 
                                                        baseDataFile, PILookupFile, 
                                                        epics, PI)
    
    # Changes since last week
    pattern = re.compile(r'\d{2}\.\d-S\d')
    changes_since_last_week = cur_pivot - prev_pivot
    cols = (changes_since_last_week.columns.to_series()
            .apply(find_PI_sprint, args=(pattern,))
            .dropna().values)
    cols = np.append(cols, ['Slip', 'Grand Total'])
    changes_since_last_week = changes_since_last_week.loc[epics, cols]
    
    # Get cumulative metrics
    (cur_cum_sum, cur_cum_per, cur_CLIN_df) = get_cum_metrics(cur_pivot, epics, clins)
    (prev_cum_sum, prev_cum_per, prev_CLIN_df) = get_cum_metrics(prev_pivot, epics, clins)
    (baseline_cum_sum, baseline_cum_per, baseline_CLIN_df) = get_cum_metrics(baseline_pivot, epics, clins)
    
    # Get sprint metrics
    cur_sprint_metrics, cur_rem_metrics = get_sprint_metrics(curSprint, lastCompleteSprint, cur_pivot,  
                                                            cur_cum_sum, baseline_cum_sum, 
                                                            baseline_cum_per,
                                                            epics)
    prev_sprint_metrics, prev_rem_metrics = get_sprint_metrics(curSprint, lastCompleteSprint, prev_pivot,
                                                                prev_cum_sum, baseline_cum_sum, 
                                                                baseline_cum_per,
                                                                epics)
    # Summary dict
    tables = {'Current Pivot': {'Pivot': cur_pivot,
                                'Changes Since Last Week': changes_since_last_week,
                                'Cum Sum': cur_cum_sum,
                               'Cum Per': cur_cum_per,
                               'CLIN Per': cur_CLIN_df,
                               'Sprint Metrics': cur_sprint_metrics,
                               'Remaining Metrics': cur_rem_metrics},
             'Previous Pivot': {'Pivot': prev_pivot,
                                'Cum Sum': prev_cum_sum,
                               'Cum Per': prev_cum_per,
                               'CLIN Per': prev_CLIN_df,
                               'Sprint Metrics': prev_sprint_metrics,
                               'Remaining Metrics': prev_rem_metrics},
             'Baseline Pivot': {'Pivot': baseline_pivot,
                                'Cum Sum': baseline_cum_sum,
                               'Cum Per': baseline_cum_per,
                               'CLIN Per': baseline_CLIN_df}}
    
    return tables


def merge_baseline_cur(baseline, cur, overall=False):
    # Merge baseline and current
    df = pd.merge(baseline, cur, left_index=True, right_index=True, suffixes=('_BL', '_Cur'))
    # Sort columns
    df = df.reindex(sorted(df.columns), axis=1)
    # Rename overall index
    if overall:
        df.rename(index={df.index[0]:'Overall'}, inplace=True)
    return df

def get_stoplight_data(data, clin):
    baseline_per = get_clin(data['Baseline Pivot']['Cum Per'], clin)
    cur_per = get_clin(data['Current Pivot']['Cum Per'], clin)
    baseline_tot = get_clin(data['Baseline Pivot']['CLIN Per'], clin)
    cur_tot = get_clin(data['Current Pivot']['CLIN Per'], clin)

    sprints_cats = merge_baseline_cur(baseline_per, cur_per)
    sprints_tot = merge_baseline_cur(baseline_tot, cur_tot, overall=True)

    sprints_data = pd.concat((sprints_cats, sprints_tot))
    
    points_cats = get_clin(data['Current Pivot']['Sprint Metrics'], clin)
    points_tot = pd.DataFrame({'Overall': points_cats.sum()}).transpose()
    points_data = pd.concat((points_cats, points_tot))

    rem_cats = get_clin(data['Current Pivot']['Remaining Metrics'], clin)
    rem_tot = (pd.DataFrame({'Overall': [np.nan] * rem_cats.shape[1]})
               .transpose().rename(columns={0: rem_cats.columns[0],
                                            1: rem_cats.columns[1],
                                            2: rem_cats.columns[2]}))
    rem_data = pd.concat((rem_cats, rem_tot))

    points_data = pd.concat((points_data, rem_data), axis=1)
    
    stoplight_data = pd.concat((sprints_data, points_data), axis=1)
    
    for idx in stoplight_data.index:
        stoplight_data.rename(index={idx:idx.split(': ')[-1]}, inplace=True)
        
    change_since_BL = stoplight_data[['Change Since BL']].copy()
    stoplight_data.drop(columns='Change Since BL', inplace=True)
    
    return stoplight_data, change_since_BL
