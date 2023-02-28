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

def get_cumper(pivot, cumSum, epics):
    # Total points in assigned sprint columns will be last column of 
    # the cumulative sum df plus the slipped points
    if "Slip" in pivot.columns:
        totals = cumSum.iloc[:, -1] + pivot.Slip.drop('Grand Total')
    else:
        totals = cumSum.iloc[:, -1]

    # Reshape
    totals = (totals
             .values
             .repeat(cumSum.shape[1])
             .reshape(cumSum.shape))
    cumPer = (cumSum / totals)
    return cumPer

def get_clin(df, clin):
    clinIdx = df.index.to_series().apply(lambda x: clin in x)
    return df[clinIdx]

def get_clinPer(cumSum, pivot, clin):
    # Only use assigned sprint columns, plus slip if applicable
    curClin = get_clin(cumSum, clin)
    if 'Slip' in pivot.columns:
        curClinSlip = get_clin(pivot, clin).Slip
        total = (curClin.iloc[:, -1] + curClinSlip).values.sum()
    else:
        total = curClin.iloc[:, -1].values.sum() 
    clinPer = get_clin(cumSum, clin).sum() / total
    return clinPer

def get_cum_metrics(pivot, epics, clins):
    # Only use actual sprint data
    pattern = re.compile(r'\d{2}\.\d-S\d')
    cols = pivot.columns.to_series().apply(find_PI_sprint, 
                                           args=(pattern,)).dropna().values
    pivotSprints = pivot.copy()
    pivotSprints = pivotSprints.loc[epics, cols]
    
    # Points cumulative sum 
    cumSum = get_cumsum(pivotSprints)
    
    # Points cumulative percentage
    cumPer = get_cumper(pivot, cumSum, epics)
    
    # CLIN breakout
    CLINDf = pd.DataFrame()
    for clin in clins:
        clinPer = get_clinPer(cumSum, pivot, clin)
        CLINDf[clin] = clinPer
    CLINDf = CLINDf.transpose()
    
    return cumSum, cumPer, CLINDf

def get_sprint_metrics(curSprint, lastCompleteSprint, pivot, pivotCumSum,
                         baselineCumSum, baselineCumPer, epics):
    # Current total (Points in sprint plus slip)
    cumSumTot = pivotCumSum.iloc[:, -1].values 
    slip = pivot.Slip.drop(['Grand Total'])
    curTotal = cumSumTot + slip
    
    # Change since baseline
    baselineTotal = baselineCumSum.iloc[:, -1].values
    baselineChange = curTotal - baselineTotal
    
    # Points expected
    lastPattern = re.compile(fr'\d\d\.\d-S{lastCompleteSprint}')
    col = (baselineCumPer.columns.to_series()
           .apply(find_PI_sprint, args=(lastPattern,))
           .dropna().values[0])
    pointsExpected = (baselineCumPer.loc[epics, col] * curTotal)
    
    # Points completed
    col = (pivotCumSum.columns.to_series()
           .apply(find_PI_sprint, args=(lastPattern,))
           .dropna().values[0])
    pointsCompleted = pivotCumSum.loc[epics, col]

    # Current completed
    curPattern = re.compile(fr'\d\d\.\d-S{curSprint}')
    col = (pivotCumSum.columns.to_series()
           .apply(find_PI_sprint, args=(curPattern,))
           .dropna().values[0])
    curPointsCompleted = pivotCumSum.loc[epics, col]
    
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

    return sprintMetricsDf, remainingSprintMetrics

def get_aggregated_data(curSprint, lastCompleteSprint, 
                        newDataFile, prevDataFile, 
                        baseDataFile, PILookupFile,
                        epics, clins, PI):
    # Get pivot tables
    (curPivot, prevPivot, baselinePivot,
      curSlip, curNew) = all_pivots(newDataFile, prevDataFile, 
                                    baseDataFile, PILookupFile, 
                                    epics, PI)
    
    # Changes since last week
    pattern = re.compile(r'\d{2}\.\d-S\d')
    changesSinceLastWeek = curPivot - prevPivot
    cols = (changesSinceLastWeek.columns.to_series()
            .apply(find_PI_sprint, args=(pattern,))
            .dropna().values)
    cols = np.append(cols, ['Slip', 'Grand Total'])
    changesSinceLastWeek = changesSinceLastWeek.loc[epics, cols]
    
    # Get cumulative metrics
    (curCumSum, curCumPer, curClinDf) = get_cum_metrics(curPivot, 
                                                        epics, 
                                                        clins)
    (prevCumSum, prevCumPer, prevCLINDf) = get_cum_metrics(prevPivot, 
                                                           epics, 
                                                           clins)
    (baselineCumSum, baselineCumPer, baselineCLINDf) = get_cum_metrics(baselinePivot,
                                                                       epics, 
                                                                       clins)
    
    # Get sprint metrics
    curSprintMetrics, curRemMetrics = get_sprint_metrics(curSprint, 
                                                         lastCompleteSprint, 
                                                         curPivot, 
                                                         curCumSum, 
                                                         baselineCumSum, 
                                                         baselineCumPer,
                                                         epics)
    prevSprintMetrics, prevRemMetrics = get_sprint_metrics(curSprint, 
                                                           lastCompleteSprint, 
                                                           prevPivot, 
                                                           prevCumSum,
                                                           baselineCumSum, 
                                                           baselineCumPer,
                                                           epics)
    # Summary dict
    tables = {
        'Current Pivot': {'Pivot': curPivot,
                            'Changes Since Last Week': changesSinceLastWeek,
                            'Cum Sum': curCumSum,
                            'Cum Per': curCumPer,
                            'CLIN Per': curClinDf,
                            'Sprint Metrics': curSprintMetrics,
                            'Remaining Metrics': curRemMetrics,
                            'Slip': curSlip,
                            'New': curNew},
        'Previous Pivot': {'Pivot': prevPivot,
                            'Cum Sum': prevCumSum,
                            'Cum Per': prevCumPer,
                            'CLIN Per': prevCLINDf,
                            'Sprint Metrics': prevSprintMetrics,
                            'Remaining Metrics': prevRemMetrics},
        'Baseline Pivot': {'Pivot': baselinePivot,
                            'Cum Sum': baselineCumSum,
                            'Cum Per': baselineCumPer,
                            'CLIN Per': baselineCLINDf}
    }
    
    return tables


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

def get_stoplight_data(data, clin):
    baselinePer = get_clin(data['Baseline Pivot']['Cum Per'], clin)
    curPer = get_clin(data['Current Pivot']['Cum Per'], clin)
    baselineTot = get_clin(data['Baseline Pivot']['CLIN Per'], clin)
    curTot = get_clin(data['Current Pivot']['CLIN Per'], clin)

    sprintsEpics = merge_baseline_cur(baselinePer, curPer)
    sprintsTot = merge_baseline_cur(baselineTot, curTot, overall=True)

    sprintsData = pd.concat((sprintsEpics, sprintsTot))
    
    pointsEpics = get_clin(data['Current Pivot']['Sprint Metrics'], clin)
    pointsTot = pd.DataFrame({'Overall': pointsEpics.sum()}).transpose()
    pointsData = pd.concat((pointsEpics, pointsTot))

    remEpics = get_clin(data['Current Pivot']['Remaining Metrics'], clin)
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
    
    return stoplightData, changeSinceBL
