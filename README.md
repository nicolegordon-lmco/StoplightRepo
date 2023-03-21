# StoplightRepo

Scripts to automate the creation of the Stoplight pivot tables and Stoplight graphic. 

The main script is Stoplight.py. This is what is run to create the output. The format to run the script is
"python Stoplight.py pathToCurrentJiraExport pathToPreviousJiraExport pathToBaslineJiraExport" 

The optional flags are
    --sprint: int
        Current sprint number. The default is the sprint number found based on the most recent Tuesday's date.
    --lastCompleteSprint: int 
        Last complete sprint number. The default is the last completed sprint found based on today's date.
    --PI: str
        Current PI. The default is the current PI found based on the most recent Tuesday's date.
    --PILookupFile: str
        Path to PI Lookup file with dates for each PI. This is defaulted in the script.
    --stoplightWdir: str
        Working directory where Stoplight files will be output. This is defaulted in the script.
    --printContributors: flag 
        Flag to print items that contributed to the pivots. Default is off. To turn on, just specify "--printContributors" in the command line.

The PivotClass.py script contains the Pivot Class which is where most of the code is stored. 
The format.py script contains a dictionary with all of the formats used in the excel output files.