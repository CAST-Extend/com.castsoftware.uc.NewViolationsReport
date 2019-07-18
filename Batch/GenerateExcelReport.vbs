'=======================================================================================================
'=======================================================================================================
' TEST EXECUTION PARAMETERS
appName ="Report_Dev"
csvDate="201907171311" 'CStr(Year(Now)) & CStr(Month(Now)) & CStr(Day(Now)) & Cstr(Hour(Now))& Cstr(Minute(Now))
path = "C:\Users\eca\Desktop\Support\Education"
logpath="C:\Users\eca\Desktop\Support\Education"
reportpath="C:\Users\eca\Desktop\Support\Education"
'=============================================================================================================
'================================================EXECUTION====================================================
'=============================================================================================================
'Init System Variables passed throughh previous Batch file
Dim appname, csvDate, path, logpath, reportpath
	set args = WScript.Arguments
	' Parse args
	If args.Count > 0 Then
			appName = args(0)
			csvDate = args(1)
			path = args(2)
			logpath = args(3)
			reportpath = args(4)
	End If
	Init appName,csvDate,path, logpath, reportpath
 
'=======================================================================================================
'=======================================================================================================
'=======================================================================================================
'=======================================================================================================
'==============General constants that can be modified with caution

'Constants relative to the Excel formatting
Public Const QR_SHEET_PREFIX = "CI REVIEW "
Public Const VIOLATION_SHEET_PREFIX = "CI RAW DATA "
Public Const TEMPLATE_FILE_NAME = "APP_NAME - CAST - CI Review - DATE.xlsx"
Public Const FINAL_FILE_TEMPLATENAME = " - CAST - CI Review - "
Public Const MAX_NB_QR_LINES = 1000
Public Const MAX_NB_VI_LINES = 6000
Public Const LAST_SNAPSHOT_COL = "Last Snapshot"
Public Const PREV_SNAPSHOT_COL = "Previous Snapshot"
Public Const URL_COL_NUM = 12
Public Const CSV2_LAST_COL_ID = "L"
Public Const CSV1_CRIT_COL_ID = "B"
Public Const CSV1_PRIO_COL_ID = ""
Public Const CSV1_NV_COL_ID = "E"
Public Const CSV1_LAST_COL_ID = "F"
Public Const CSV1_WEIGHT_COL_ID = "C"
Public Const CSV2_CRIT_COL_ID = "C"
Public Const CSV1_RNAME_COL_ID = "D"
Public Const S1_RNAME_COL_ID = "E"
Public Const S0_SNAPORDER_COL_NUM = 2
Public Const S0_LASTVNAME_COL_NUM = 4
Public Const S0_SNAPDATE_COL_NUM = 5
Public Const S1_FIRST_LINE_NUM = 10
Public Const S2_LAST_COL_ID = "K"
Public Const S2_TEMPL_LAST_ROW = 4
Public Const S2_OBJFNAME_COL_NUM = 6
Public Const S2_OBJFNAME_COL_ID = "F"
Public Const S1_TEMPL_LAST_ROW = 13
Public Const S1_LAST_COL_ID = "G"
Public Const S1_EVOL_COL_ID = "D"
Public Const S1_RNAME_COL_NUM = 5
Public Const S1_NV_COL_NUM = 6
Public Const S2_RNAME_COL_ID = "E"
Public Const S1_FV_COL_NUM = 7
Public Const S1_NV_COL_ID = "F"
Public Const S1_FV_COL_ID = "G"

'=======================================================================================================
'Init System Variables 
Dim XLAPP
Dim logFile
Dim FSOAPP
Dim logs
	logs = ""
'=======================================================================================================
'EXCEL Constants - cannot be modified
Public Const xlDelimited = 1
Public Const xlTextQualifierDoubleQuote = 1
Public Const xlPasteValues = -4163 '(&HFFFFEFBD)
Public Const xlPasteFormulas = -4123 '(&HFFFFEFE5)
Public Const xlPasteFormats = -4122 '(&HFFFFEFE6)
Public Const xlDown=-4121
Public Const xlToLeft = -4159
Public Const xlToRight = -4161
Public Const xlUp = -4162
Public Const xlEdgeBottom = 9
Public Const xlSortOnValues = 0
Public Const xlAscending = 1
Public Const xlDescending = 2
Public Const xlSortNormal = 0
Public Const xlGuess = 0
Public Const xlTopToBottom = 1
Public Const xlPinYin = 1
Public Const xlAutomatic = -4105
Public Const xlUnderlineStyleNone = -4142
'=======================================================================================================
'Init Variables provided as executable parameter
Public mainFolder 'As String
Public reportFolder 'As String
Public logFolder 'As String
Public applicationName 'As String
Public latestVersionName 'As String
Public previousVersionName 'As String
Public latestVersionDate 'As String
Public previousVersionDate 'As String

'=======================================================================================================
'Defining the Program Variables deducted from Init Variables above
Public cQRCSVFileName 'As String
Public cNVCSVFileName 'As String
Public cFVCSVFileName 'As String
Public cSRCSVFileName
Public reportGenerationDate 'As String
Public finalFileName 'As String
Public tab_qr_evolution 'As String
Public tab_violation_list 'As String
Public lastNVRow 'As Integer
Public lastVLRow 'As Integer
Public IsNVSheetEmpty
Public IsFVSheetEmpty
Public numberOfQR
Public numberOfNV
Public numberOfFV


'=======================================================================================================
'=======================================================================================================
'===========================================UTILITIES FUNCTIONS ========================================
'=======================================================================================================
'=======================================================================================================
Sub logText(strText)
     'logFile.WriteLine(Now & " - " & strText)
	 logs = logs & vbCrlf & Now & " - " & strText
	 'Msgbox strText
End Sub

'=======================================================================================================
'=======================================================================================================

Function openFile(FilePath)
On Error Resume Next
	logText "		OPENFILE: " & FilePath
	If FSOAPP.FileExists(FilePath) then
		'Function Open(Filename As String, [UpdateLinks], [ReadOnly], [Format], [Password], [WriteResPassword], [IgnoreReadOnlyRecommended], [Origin], [Delimiter], [Editable], [Notify], [Converter], [AddToMru], [Local], [CorruptLoad])
         XLAPP.Workbooks.Open FilePath,0,False
		openFile = True
	Else
		WScript.echo "The file " & FilePath & " was not found." & vbCrLf & "Please make sure that it has not been moved or deleted."
		logText "ERROR::OPENFILE:The file " & FilePath & " was not found.Please make sure that it has not been moved or deleted."
		openFile = False
    End If
	If Err Then
		logText "ERROR::OPENFILE " & Err.Description
		openFile = False
	End If
On Error Goto 0
End Function

'=======================================================================================================
'=======================================================================================================

Function CloseOpenedFiles()
	i = 1
		Do While i <= XLAPP.Workbooks.Count
				If 	XLAPP.Workbooks.Item(i).Name = cQRCSVFileName Or _
						XLAPP.Workbooks.Item(i).Name = cNVCSVFileName Or _
						XLAPP.Workbooks.Item(i).Name = cSRCSVFileName Or _
						XLAPP.Workbooks.Item(i).Name = cFVCSVFileName Or _
						XLAPP.Workbooks.Item(i).Name = TEMPLATE_FILE_NAME Then 
						XLAPP.Workbooks.Item(i).Close False 'Close([SaveChanges], [Filename], [RouteWorkbook])
				Else
					If 		XLAPP.Workbooks.Item(i).Name = finalFileName Then 		XLAPP.Workbooks.Item(i).Close True
				End If
				i = i + 1
	   Loop
End Function

'=======================================================================================================
'=======================================================================================================

'This function opens a CSV file, splits it by semi-column character into a target sheet of the main file
Function openAndSplitCSVFile(CSVFilePath, CSVFileName)
On Error Resume Next
	logText "		OPENANDSPLITCSVFILE: " & CSVFilePath
    If openFile(CSVFilePath) Then
		With XLAPP.Workbooks(CSVFileName).ActiveSheet
			'Delete the first line containing the headlines
			If 	.UsedRange.Rows.Count <= 1 Then 
				logText "INFO::OPENANDSPLITCSVFILE: CSV File '" & CSVFilePath & "' is empty."
				openAndSplitCSVFile = "Empty"			
			Else
				.Rows.Item(1).Delete
				If 	XLAPP.WorksheetFunction.IsError( XLAPP.WorksheetFunction.FindB(";", .Cells(1,1).Value, 1)) Then '.Columns.Count > 1 Then
					logText "INFO::OPENANDSPLITCSVFILE: CSV File '" & CSVFilePath & "' is already splitted."
				Else					
					.Columns.Item(1).TextToColumns .Cells(1,1),xlDelimited, xlTextQualifierDoubleQuote, False, False, True, False,False,False,False
				End If
				openAndSplitCSVFile = True
				logText "DONE::OPENANDSPLITCSVFILE - CSV File '" & CSVFilePath & "' has been splitted."
			End If
		End With
	Else
		logText "ERROR::OPENANDSPLITCSVFILE - CSV File '" & CSVFilePath & "' could not be opened"
		openAndSplitCSVFile = False
	End If

If Err Then
	logText "ERROR::OPENANDSPLITCSVFILE '" & CSVFilePath & " : " & Err.Description
	openAndSplitCSVFile = False
End If
On Error Goto 0
End Function

'=======================================================================================================
'=======================================================================================================

'copy the content of CSV files into the Final file
'all cell references are hardwritten in this function
Function copyCSVToFinal(FileName, DestinationSheet, DestinationCellNumber, IsFirstSheet)
On Error Resume Next
	Dim lastRow
	logText "		COPYCSVTOFINAL:: " & FileName & "," & DestinationSheet & "," & DestinationCellNumber & "," & IsFirstSheet
	With XLAPP.Workbooks(FileName).ActiveSheet	
		If IsFirstSheet Then
			lastRow = .UsedRange.Rows.Count
			'Msgbox "lastRow in first sheet " & lastRow 
			If lastRow > MAX_NB_QR_LINES Then lastRow = MAX_NB_QR_LINES	
			.Sort.SortFields.Clear
			'order by rule criticality then by weight
			.Sort.SortFields.Add .range(CSV1_CRIT_COL_ID & "1",CSV1_CRIT_COL_ID & lastRow) , xlSortOnValues, xlDescending ', , xlSortNormal
			.Sort.SortFields.Add .range(CSV1_WEIGHT_COL_ID & "1",CSV1_WEIGHT_COL_ID & lastRow) , xlSortOnValues, xlDescending ', , xlSortNormal
			.Sort.SetRange .range("A1", CSV1_LAST_COL_ID & lastRow)
			.Sort.Header = xlGuess
			.Sort.MatchCase = False
			.Sort.Orientation = xlTopToBottom
			.Sort.SortMethod = xlPinYin
			.Sort.Apply
			
			'replace criticality value
			.range(CSV1_CRIT_COL_ID & "1", CSV1_CRIT_COL_ID & lastRow).Replace CStr(1), "X"
			.range(CSV1_CRIT_COL_ID & "1", CSV1_CRIT_COL_ID & lastRow).Replace CStr(0), ""	'CRITICALITY COLUMN REPLACEMENT

			'
			'First 3 columns are pasted separately, first column of the CSV must always contain NOT NULL values (object Id or Name)
			.range("A1", CSV1_WEIGHT_COL_ID & lastRow).Copy
			XLAPP.Workbooks(finalFileName).Worksheets(DestinationSheet).range("A" & DestinationCellNumber).PasteSpecial xlPasteValues
        'Other columns are pasted
			.range(CSV1_RNAME_COL_ID & "1", CSV1_LAST_COL_ID & lastRow).Copy
			XLAPP.Workbooks(finalFileName).Worksheets(DestinationSheet).range(S1_RNAME_COL_ID & DestinationCellNumber).PasteSpecial xlPasteValues
		Else
		'whole content of other sheets are pasted
			lastRow = 	.UsedRange.Rows.Count '.range("A1").End(xlDown).Row
			If lastRow > MAX_NB_VI_LINES Then lastRow = MAX_NB_VI_LINES
			.range(CSV2_CRIT_COL_ID & "1", CSV2_CRIT_COL_ID & lastRow).Replace CStr(1), "X"
			.range(CSV2_CRIT_COL_ID & "1", CSV2_CRIT_COL_ID & lastRow).Replace CStr(0), ""	'CRITICALITY COLUMN REPLACEMENT
			.range("A1",CSV2_LAST_COL_ID & lastRow).Copy
			XLAPP.Workbooks(finalFileName).Worksheets(DestinationSheet).range("A" & DestinationCellNumber).PasteSpecial xlPasteValues
		End If
	End With
    XLAPP.Workbooks(finalFileName).Save
    XLAPP.CutCopyMode = False
	XLAPP.Workbooks(FileName).Close False
	
	If Err Then 
		logText "ERROR::COPYCSVTOFINAL on file "& FileName & " : " & Err.Description
		copyCSVToFinal = -1 'False
	Else	
		copyCSVToFinal = lastRow 'True
	End If
On Error Goto 0
End Function




'=============================================================================================================
'=============================================================================================================
'=============================================================================================================
'================================================MAIN SUB ====================================================
'=============================================================================================================
'=============================================================================================================
'populate variables, launch the generation
Sub Init(paramAppName, paramCSVDate, paramPath, paramLogPath, paramReportPath)
	
	Dim snapshotSheet
    'Variables provided by the execution
    applicationName = paramAppName
    mainFolder = paramPath 
	logFolder = paramLogPath
	reportFolder = paramReportPath
	dateSuffix = paramCSVDate 'Left(paramCSVDate,8)
		
	Set FSOAPP = CreateObject("Scripting.FileSystemObject")
	Set XLAPP = CreateObject("Excel.Application")
    XLAPP.AddCustomList Array("extreme", "high", "moderate", "low")
	
	logText ">>>>>>InitVariables: " & applicationName & "," & dateSuffix & "," & mainFolder

    'Change following Values depending on testing or production phases
	XLAPP.Visible = False
    XLAPP.DisplayAlerts = False
	
	
	'Just in case the values are not properly affected, we will not block the generation
	latestVersionName = "Latest Snapshot"
	latestVersionDate = "YYYY-MM-DD"
	previousVersionName = "Previous Snapshot"
	previousVersionDate = "YYYY-MM-DD"
        
    'built variables
	tab_qr_evolution = QR_SHEET_PREFIX & dateSuffix
	tab_violation_list = VIOLATION_SHEET_PREFIX & dateSuffix
    reportGenerationDate = CStr(Year(Now)) & "/" & CStr(Month(Now)) & "/" & CStr(Day(Now))
    finalFileName = applicationName & FINAL_FILE_TEMPLATENAME & dateSuffix & ".xlsx"
    
    'get all CSV files
    cSRCSVFileName = applicationName & "_SnapshotReport_" & paramCSVDate & ".csv"
    cQRCSVFileName = applicationName & "_SummaryReport_" & paramCSVDate & ".csv"
    cNVCSVFileName = applicationName & "_NewViolations_" & paramCSVDate & ".csv"
    cFVCSVFileName = applicationName & "_FixedViolations_" & paramCSVDate & ".csv"
	
	If FSOAPP.FileExists(reportFolder & "\" & cSRCSVFileName) And _
		FSOAPP.FileExists(reportFolder & "\" & cQRCSVFileName) And _
		FSOAPP.FileExists(reportFolder & "\" & cNVCSVFileName) And _
		FSOAPP.FileExists(reportFolder & "\" & cFVCSVFileName) Then
	
		dim tmpLdate, tmpPDate
		If openAndSplitCSVFile(reportFolder & "\" & cSRCSVFileName,cSRCSVFileName) Then
			Set snapshotSheet = XLAPP.Workbooks(cSRCSVFileName).ActiveSheet
			Select Case snapshotSheet.Cells(1,S0_SNAPORDER_COL_NUM).Value
				Case LAST_SNAPSHOT_COL
					latestVersionName = snapshotSheet.Cells(1,S0_LASTVNAME_COL_NUM).Value
					tmpLdate = CDate(snapshotSheet.Cells(1,S0_SNAPDATE_COL_NUM).Value)
					previousVersionName = snapshotSheet.Cells(2,S0_LASTVNAME_COL_NUM).Value
					tmpPDate = CDate(snapshotSheet.Cells(2,S0_SNAPDATE_COL_NUM).Value)
				Case PREV_SNAPSHOT_COL
					latestVersionName = snapshotSheet.Cells(2,S0_LASTVNAME_COL_NUM).Value
					tmpLdate = CDate(snapshotSheet.Cells(2,S0_SNAPDATE_COL_NUM).Value)
					previousVersionName = snapshotSheet.Cells(1,S0_LASTVNAME_COL_NUM).Value
					tmpPDate = CDate(snapshotSheet.Cells(1,S0_SNAPDATE_COL_NUM).Value)
				Case Else					
			End Select
		End If
		XLAPP.Workbooks(cSRCSVFileName).Close False
		'Date in format: 10/12/2018  10:00:00
		latestVersionDate =  Year(tmpLdate) & "-" &  Right("0" & Month(tmpLdate), 2)& "-" &  Right("0" & Day(tmpLdate),2)
		previousVersionDate = Year(tmpPDate)& "-" & Right("0" & Month(tmpPDate),2)& "-" & Right("0" & Day(tmpPDate),2)
		GenerateEducationReport
		CloseOpenedFiles		
	Else
			logText "ERROR::One of the Input Files does not exists. Report Generation has been aborted"
			WScript.echo "One of the Input Files does not exists. Report Generation has been aborted: " & reportFolder & "\" & cSRCSVFileName
	End If
	XLAPP.DisplayAlerts = True
	XLAPP.Quit
	
	Set logFile = FSOAPP.CreateTextFile(logFolder & "\ExcelEducationReport_log_" & applicationName & " - " & paramCSVDate & ".txt", True)
	logFile.WriteLine(logs)
	Set logFile = Nothing
	Set FSOAPP = Nothing
	Set XLAPP = Nothing
End Sub


'=============================================================================================================
'=============================================================================================================
'=============================================================================================================
'=====================================================MAJOR FUNCTION==========================================
'=============================================================================================================
'=============================================================================================================

Function GenerateEducationReport()
On Error Resume Next
		
	logText ">>>>>>Process started with application : " & applicationName & _
            vbCrLf & ">>>>>>CSV Files : " & cQRCSVFileName & ", " & cNVCSVFileName & ", " & cFVCSVFileName & _
            vbCrLf & ">>>>>>FinalFileName : " & finalFileName
    
    
	If openAndSplitCSVFile(reportFolder & "\" & cQRCSVFileName,cQRCSVFileName) = True Then
		
		If openFile(mainFolder & "\" & TEMPLATE_FILE_NAME) = True Then
			
			XLAPP.Workbooks(TEMPLATE_FILE_NAME).SaveCopyAs reportFolder & "\" & finalFileName
			XLAPP.Workbooks(TEMPLATE_FILE_NAME).Close False
			If openFile(reportFolder & "\" & finalFileName) = True  Then
			
				XLAPP.Workbooks(finalFileName).Worksheets.Item(1).Name = tab_qr_evolution 'Renaming sheets with final names
				XLAPP.Workbooks(finalFileName).Worksheets.Item(1).DisplayPageBreaks = False
				'import the content of Main CSV file.
				numberOfQR = copyCSVToFinal (cQRCSVFileName, tab_qr_evolution, S1_FIRST_LINE_NUM, True)
				'XLAPP.Workbooks(cQRCSVFileName).Close False
				If  numberOfQR > 0 Then	
					'proceed to other CSV sheets
					If openAndSplitCSVFile(reportFolder & "\" & cNVCSVFileName,cNVCSVFileName) = "Empty" Then IsNVSheetEmpty = True

					If openAndSplitCSVFile(reportFolder & "\" & cFVCSVFileName,cFVCSVFileName) = "Empty" Then IsFVSheetEmpty = True

					If IsNVSheetEmpty = True And IsFVSheetEmpty = True Then 
						XLAPP.Workbooks(cNVCSVFileName).Close False
						XLAPP.Workbooks(cFVCSVFileName).Close False
						XLAPP.Workbooks(finalFileName).Worksheets.Item(2).Delete				
						FormatFirstSheet
					Else
						XLAPP.Workbooks(finalFileName).Worksheets.Item(2).Name = tab_violation_list 'Renaming sheets with final names
						XLAPP.Workbooks(finalFileName).Worksheets.Item(2).DisplayPageBreaks = False
						'CSV File N°1 - New violations
						lastNVRow = 1
						If IsNVSheetEmpty = False Then 
							numberOfNV = copyCSVToFinal(cNVCSVFileName, tab_violation_list, 2, False)
							If numberOfNV > -1 Then
							lastNVRow = numberOfNV +1
'XLAPP.Workbooks(finalFileName).Worksheets(tab_violation_list).UsedRange.Rows.Count '.range("A1").End(xlDown).Row
							Else 
								logText "ERROR::GenerateEducationReport: Failed to copy CSV file N°1.Generation aborted"
								WScript.echo "GenerateEducationReport: Failed to copy CSV file Item 1.Generation aborted"
								GenerateEducationReport = False
								Exit Function				
							End If
						End If
						'Msgbox "lastNVRow= "& lastNVRow
						'XLAPP.Workbooks(cNVCSVFileName).Close False
						'CSV File N°2 - Fixed Violations
						lastVLRow = lastNVRow
						If IsFVSheetEmpty = False Then 
							numberOfFV = copyCSVToFinal (cFVCSVFileName, tab_violation_list, lastNVRow + 1, False) 
							If numberOfFV > -1 Then
								lastVLRow = lastNVRow + numberOfFV
								'lastVLRow = XLAPP.Workbooks(finalFileName).Worksheets(tab_violation_list).UsedRange.Rows.Count '.range("A1").End(xlDown).Row
							Else 
								logText "ERROR::GenerateEducationReport: Failed to copy CSV file N°2.Generation aborted"
								WScript.echo "GenerateEducationReport: Failed to copy CSV file Item 2.Generation aborted"
								GenerateEducationReport = False
								Exit Function				
							End If
						End If
						'XLAPP.Workbooks(cFVCSVFileName).Close False
						
						FormatSecondSheet 	
						XLAPP.Workbooks(finalFileName).Save			
						FormatFirstSheet
						XLAPP.Workbooks(finalFileName).Save
						XLAPP.Workbooks(finalFileName).Worksheets.Item(2).DisplayPageBreaks = True
					End If
					XLAPP.Workbooks(finalFileName).Worksheets.Item(1).DisplayPageBreaks = True
					XLAPP.Workbooks(finalFileName).Close True
					logText ">>>>>>Process ended with application : " & applicationName & vbCrLf & ">>>>>>FinalFileName : " & finalFileName
				Else
					Err.Raise 32811, "GenerateEducationReport", "Impossible to copy data from Main CSV File. Generation aborted"
				End If
			Else
				Err.Raise 32811, "GenerateEducationReport", "Input File could not be opened. Generation aborted"
			End If
		Else
			Err.Raise 32811, "GenerateEducationReport", "Template File could not be opened. Generation aborted"
		End If	
	Else
		Err.Raise 32811, "GenerateEducationReport", "Main sheet is empty or invalid. Generation aborted"
	End If
	
	If Err Then
		logText ("ERROR::N°" & Err.Number & " - " & Err.Description & " ; Error source : " & Err.Source)
		WScript.echo "An error occured during final Excel generation, please contact CAST Support." & vbCrLf & vbCrLf & _
                "Error code : " & Err.Number & vbCrLf & _
               "Error description : " & Err.Description & vbCrLf & vbCrLf & _
               "Error source : " & Err.Source & vbCrLf & _
               "Error context : " & Err.HelpContext _
               , vbOKOnly, "CAST Education Rules"
    
		logText "ERROR::Report Generation has been aborted"
		GenerateEducationReport = False
	Else 
		logText "SUCCESS::Report has been generated: " & finalFileName
		WScript.echo "Report has been generated: " & finalFileName
		GenerateEducationReport = True
	End If   
On Error Goto 0
End Function


'=======================================================================================================
'=======================================================================================================
'================================================FUNCTIONS TO GENERATE FINAL REPORT =====================
'=======================================================================================================
'=======================================================================================================
'=======================================================================================================

Function FormatSecondSheet ()
On Error Resume Next
	'>>>>>>>>>>>>>Formatting Sheet N2°
			'First column should never contain empty values
	logText "		FORMAT SECOND SHEET"
	Dim link, keepLineStyle, keepLineWeight, keepLineColor
		
		With XLAPP.Workbooks(finalFileName).Worksheets(tab_violation_list) 
			.Activate	
			'Copy last row format
			keepLineStyle =  .range("A" & S2_TEMPL_LAST_ROW, S2_LAST_COL_ID & S2_TEMPL_LAST_ROW).Borders(xlEdgeBottom).LineStyle
			keepLineWeight = .range("A" & S2_TEMPL_LAST_ROW, S2_LAST_COL_ID & S2_TEMPL_LAST_ROW).Borders(xlEdgeBottom).Weight
			keepLineColor = .range("A" & S2_TEMPL_LAST_ROW, S2_LAST_COL_ID & S2_TEMPL_LAST_ROW).Borders(xlEdgeBottom).Color
			
			
			'Format all lines
			.range("A2", S2_LAST_COL_ID & "3").Copy
			.range("A2", S2_LAST_COL_ID & lastVLRow).PasteSpecial xlPasteFormats
			
			'Clear all leftover lines formatting
			If S2_TEMPL_LAST_ROW > lastVLRow Then  
				.range("A" & lastVLRow+1,S2_LAST_COL_ID & S2_TEMPL_LAST_ROW).ClearFormats
				.range("A" & lastVLRow+1,S2_LAST_COL_ID & S2_TEMPL_LAST_ROW).ClearContents
			End If
			'adding URLs to the last column only for new violations
			i = 2				
			Do While i <= lastNVRow
					link = .range(CSV2_LAST_COL_ID & i).Value
						If link <> "" Then .Hyperlinks.Add .Cells(i, S2_OBJFNAME_COL_NUM),link
						i = i + 1
			Loop
			
			If IsFVSheetEmpty = False Then
				'Format fixed violations without no hyperlink	
				With .range(S2_OBJFNAME_COL_ID & Cstr(lastNVRow+1), S2_OBJFNAME_COL_ID & lastVLRow).Font
					.ColorIndex = xlAutomatic
					.TintAndShade = 0
					.Underline = xlUnderlineStyleNone
				End With
			End If

			.range("A" & lastVLRow, S2_LAST_COL_ID & lastVLRow).Borders(xlEdgeBottom).LineStyle = keepLineStyle 
			.range("A" & lastVLRow, S2_LAST_COL_ID & lastVLRow).Borders(xlEdgeBottom).Weight = keepLineWeight 
			.range("A" & lastVLRow, S2_LAST_COL_ID & lastVLRow).Borders(xlEdgeBottom).Color = keepLineColor 
			'logText "Loop:URL column"
			
			.Columns.Item(URL_COL_NUM).Clear
			.Rows(1).Select
			.Rows(1).AutoFilter	
			.UsedRange.Columns.AutoFit	
			'.Cells(1,1).Select
			 XLAPP.Goto .Cells(1,1), True
		End With

		If Err Then logText "ERROR::FORMAT SECOND SHEET: "& Err.Description
On Error Goto 0
End Function


'=======================================================================================================
'=======================================================================================================
        
Function FormatFirstSheet ()
On Error Resume Next
	logText "		FORMAT FIRST SHEET"

	'>>>>>>>>>>>>>Formatting Sheet N1°
	Dim keepLineStyle
	Dim keepLineWeight
	Dim keepLineColor
	Dim oldDate
	Dim lastQRRow 'As Integer
	Dim lastRow 'As Integer


	If IsNVSheetEmpty = False or IsFVSheetEmpty = False Then 
		Set secondSheet = XLAPP.Workbooks(finalFileName).Worksheets(tab_violation_list)
	End If
	
	With XLAPP.Workbooks(finalFileName).Worksheets(tab_qr_evolution)
		.Activate
		
		'logText "Filling the En-tete"
		'logtext S1_RNAME_COL_ID
		.range(S1_RNAME_COL_ID & "2").Value = applicationName
		.range(S1_RNAME_COL_ID & "3").Value = latestVersionName & " / " & latestVersionDate
		.range(S1_RNAME_COL_ID & "4").Value = previousVersionName & " / " & previousVersionDate
		.range("F1").Value = reportGenerationDate
		lastQRRow = numberOfQR + S1_FIRST_LINE_NUM - 1 '.Range("A" & S1_FIRST_LINE_NUM).Rows.Count '.End(xlDown).Row
		'Last Row Style is copied
		'keepLineStyle = 	.range("A" & S1_TEMPL_LAST_ROW, S1_LAST_COL_ID & S1_TEMPL_LAST_ROW).Borders(xlEdgeBottom).LineStyle
		'keepLineWeight = 	.range("A" & S1_TEMPL_LAST_ROW, S1_LAST_COL_ID & S1_TEMPL_LAST_ROW).Borders(xlEdgeBottom).Weight
		'keepLineColor = 	.range("A" & S1_TEMPL_LAST_ROW, S1_LAST_COL_ID & S1_TEMPL_LAST_ROW).Borders(xlEdgeBottom).Color
		'First and last Column should never contain empty values

				logText "format last row"
		.range("A" & CStr(S1_TEMPL_LAST_ROW-1), S1_LAST_COL_ID & S1_TEMPL_LAST_ROW).Copy
		.range("A" & lastQRRow, S1_LAST_COL_ID & CStr(lastQRRow+1)).PasteSpecial(xlPasteFormats)
		
		'Clear all formatting for QR rule lines
		if S1_TEMPL_LAST_ROW > lastQRRow Then  
			.range("A" & lastQRRow+2,S1_LAST_COL_ID & S1_TEMPL_LAST_ROW).ClearFormats
			.range("A" & lastQRRow+2,S1_LAST_COL_ID & S1_TEMPL_LAST_ROW).ClearContents
		End If
		
		logText "format other rows"
		.range("A" & S1_FIRST_LINE_NUM, S1_LAST_COL_ID & Cstr(S1_FIRST_LINE_NUM+1)).Copy
		.range("A" & S1_FIRST_LINE_NUM, S1_LAST_COL_ID & lastQRRow).PasteSpecial(xlPasteFormats)
		
		'.range("A" & lastQRRow+1, S1_LAST_COL_ID & lastQRRow+1).Borders(xlEdgeBottom).LineStyle = keepLineStyle 
		'.range("A" & lastQRRow+1, S1_LAST_COL_ID & lastQRRow+1).Borders(xlEdgeBottom).Weight = keepLineWeight 
		'.range("A" & lastQRRow+1, S1_LAST_COL_ID & lastQRRow+1).Borders(xlEdgeBottom).Color = keepLineColor 
		
		
				logText "copy the formulAs in 4th column"
		.range(S1_EVOL_COL_ID & S1_FIRST_LINE_NUM).Copy
		.range(S1_EVOL_COL_ID & Cstr(S1_FIRST_LINE_NUM+1), S1_EVOL_COL_ID & lastQRRow).PasteSpecial(xlPasteFormulas)
		
				logText "reformat last column date"
		'.Range(S1_LAST_COL_ID & S1_FIRST_LINE_NUM, S1_LAST_COL_ID & lastQRRow).Value = 	.Evaluate("LEFT(" & 	.Range(S1_LAST_COL_ID & 'S1_FIRST_LINE_NUM, S1_LAST_COL_ID & lastQRRow).Address & "," & S1_FIRST_LINE_NUM & ")")
    
		logText "loop:Add link to second sheet"
		'Loop on all rows
		i = S1_FIRST_LINE_NUM
		If IsNVSheetEmpty = False Or IsFVSheetEmpty = False Then
				logText "do on all rows : adding link"
			Do While i <= lastQRRow
				'add links to second sheet only to New Violations"
				If 	.Cells(i, S1_NV_COL_NUM).Value>CDbl(0) Then
					refRow = XLAPP.Match(	.Cells(i, S1_RNAME_COL_NUM).Value, secondSheet.range(S2_RNAME_COL_ID & "1", S2_RNAME_COL_ID & lastNVRow), 0)
					If Not XLAPP.WorksheetFunction.IsError(refRow) Then
						.Hyperlinks.Add 	.Cells(i, S1_NV_COL_NUM), "", "'" & tab_violation_list & "'!$" & S2_RNAME_COL_ID & "$" & refRow
					End If
				End If
				'add links to second sheet only to Fixed Violations"
				If 	.Cells(i, S1_FV_COL_NUM).Value>CDbl(0) Then
					refRow = XLAPP.Match(	.Cells(i, S1_RNAME_COL_NUM).Value, secondSheet.range(S2_RNAME_COL_ID & CStr(lastNVRow + 1), S2_RNAME_COL_ID & lastVLRow), 0)
					If Not XLAPP.WorksheetFunction.IsError(refRow) Then
						.Hyperlinks.Add 	.Cells(i, S1_FV_COL_NUM), "", "'" & tab_violation_list & "'!$" & S2_RNAME_COL_ID & "$" & CStr(lastNVRow + refRow)
					End If
				End If
				i = i + 1
			Loop
		End If
		
				logText "add total row:"
		.Cells(lastQRRow+1,S1_RNAME_COL_NUM).Value = "TOTAL"
		.Cells(lastQRRow+1,S1_NV_COL_NUM).Value = 	"=SUM(" & S1_NV_COL_ID & S1_FIRST_LINE_NUM & ":" & S1_NV_COL_ID & lastQRRow & ")"
		.Cells(lastQRRow+1,S1_FV_COL_NUM).Value = 	"=SUM(" & S1_FV_COL_ID & S1_FIRST_LINE_NUM & ":" & S1_FV_COL_ID & lastQRRow & ")"
    
		.UsedRange.Columns.AutoFit
		'.Cells(1,1).Select
		XLAPP.Goto .Cells(1,1), True
	'finalize and point pointer to first cell   
	End With
	Set secondSheet = Nothing
	
	If Err Then logText "ERROR::FORMAT FIRST SHEET : "& Err.Description
On Error Goto 0
End Function

