# com.castsoftware.uc.NewViolationsReport
Generate an Excel file with the list of new and fixed violations between the last 2 snapshots.
This Excel file contains 2 Excel sheets : a summary sheet on the new and fixed violations and a sheet with the list of all the new and 	fixed violations between the last 2 snapshots for the all CAST quality rules that have been checked.

The report can be generated in 2 ways :
    - Automatically generated in CAST Management Studio after a snapshot generation (the extension needs be installed with Server Manager)
    - In a Windows command line executing a batch program (can be automated in schedulers like Jenkins)
