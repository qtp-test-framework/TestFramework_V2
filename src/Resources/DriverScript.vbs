'========== Declare and initialize the variables ==========
Dim sTestCaseFolder, strScriptPath, sQTPResultsPath, sQTPResultsPathOrig
Dim isCurTestCase_Sel
 
sBatchSheetPath = Wscript.Arguments(0)
sTestCaseFolder = Wscript.Arguments(1)				'"C:\Users\Mohit-Ankit\Documents\Unified Functional Testing\"
selChkBox = Wscript.Arguments(2)						' Values seperated by '*'
sQTPResultsPathOrig = "../Test_Results\"

'Replacing all the $$ occurences in path back to 'spaces'
sBatchSheetPath = Replace(sBatchSheetPath,"$$"," ")
sTestCaseFolder = Replace(sTestCaseFolder,"$$"," ")

'WScript.Quit(99)
'MsgBox "selChkBox = " & selChkBox

'Pipes for writing logs to a text file
Set fs = CreateObject("Scripting.FileSystemObject")
Set log_out = fs.CreateTextFile("Execution_Logs/Execution_Logs.txt", True)

 
'========== Create an object to access QTP's objects, methods and properties ==========
Set qtpApp = CreateObject("QuickTest.Application")
 
'Open QTP if it is already not open
If qtpApp.launched <> True Then
	log_out.WriteLine("Launching QTP....")
	qtpApp.Launch
End If
 
qtpApp.Visible = True
 
'========== Set the run options for QTP ==========
qtpApp.Options.Run.ImageCaptureForTestResults = "OnError"
qtpApp.Options.Run.RunMode = "Fast"
 
'Set ViewResults property to false. This is because if you run many test cases in batch, you would not want QTP to open a separate result window for each of them
qtpApp.Options.Run.ViewResults = False

log_out.WriteLine("Retrieving the selected Testcases")

'splitting the selected checkBox values by '*'
currChkBox_Array = Split(selChkBox,"*")

'Get total Test cases in excel file
totRowsinExcel = getTotalRowsInExcel(sBatchSheetPath)
'MsgBox "totRowsinExcel = " & CStr(totRowsinExcel)


'Loop through all the Rows
'1st row is header and the test case list starts from 2nd row. So, For loop is started from 2nd row
For iR = 2 to totRowsinExcel
	isCurTestCase_Sel = false
	
	'check if current test case has been checked by user
	for each x in currChkBox_Array
		if iR = CInt(x) then
			isCurTestCase_Sel = true
			exit for
		end if
	next
	
	IF isCurTestCase_Sel = TRUE THEN
		' ========== Read test cases from batch excel sheet ==========
		Set excel_Batch = CreateObject("Excel.Application")
		
		'Get the WorkBook Object
		Set workbook_Batch = excel_Batch.WorkBooks.Open(sBatchSheetPath)
		
		'Get the Batch WorkSheet
		Set worksheet_batch = workbook_Batch.Worksheets("Batch")
			
		'Get Test Case Name
		sTestCaseName = worksheet_batch.Cells(iR, 1).Value
		'sTestCaseName = excel_Batch.Cells(iR, 1).Value
		'MsgBox "sTestCaseName = " & sTestCaseName 
		
		log_out.WriteLine("Loading Test Case : "&sTestCaseName)
		
		'Close the excel workbook object
		workbook_Batch.Close
		
		'Close the excel object
		'excel_Batch.Application.Quit
		excel_Batch.Quit

		strScriptPath = ""
		'Get the location where the test case is stored
		strScriptPath = sTestCaseFolder & sTestCaseName
		
		'Open the Test Case in Read-Only mode
		qtpApp.Open strScriptPath, True
		WScript.Sleep 2000

		'Create an object of type QTP Test
		Set qtpTest = qtpApp.Test

		'Instruct QTP to perform next step when error occurs
		'qtpTest.Settings.Run.OnError = "NextStep"
		qtpTest.Settings.Run.OnError = "Stop"

		'Create the Run Results Options object
		Set qtpResult = CreateObject("QuickTest.RunResultsOptions")

		'Set the results location. This result refers to the QTP result
		sQTPResultsPath = sQTPResultsPathOrig
		sQTPResultsPath = sQTPResultsPath & sTestCaseName
		qtpResult.ResultsLocation = sQTPResultsPath
				
		'Run the test. The result will automatically be stored in the location set by you
		WScript.Sleep 2000
		
		log_out.WriteLine("Starting Test Case : "&sTestCaseName)
		
		qtpTest.Run
		'qtpTest.Run ParamSet
	END IF
 
 'Exit For
  'MsgBox "After Loop"
 
Next
log_out.WriteLine("All test cases have been executed....")
WScript.Sleep 2000					'so that above msg is printed before "Mailing Test Case Results'
log_out.WriteLine("stop")

'Function to get total Test Cases in the Excel File
function getTotalRowsInExcel(vBatchSheetPath)
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Workbooks.Open (vBatchSheetPath)
	Set objSheet = objExcel.ActiveWorkbook.Worksheets("Batch")
	'objSheet = objExcel.ActiveWorkbook.Worksheets(1)
	Set objRange = objSheet.UsedRange
	Set objRows = objRange.Rows
	
	getTotalRowsInExcel = objRows.Count
	objExcel.Quit
end function

