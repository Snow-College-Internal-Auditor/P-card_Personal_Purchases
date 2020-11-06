Begin Dialog NewDialog 113,35,150,108,"NewDialog", .NewDialog
  Text 19,18,98,14, "Is there a database to open?", .Text1
  OKButton 17,51,40,14, "Yes", .OKButton1
  CancelButton 80,51,40,14, "No", .CancelButton1
End Dialog
Dim testDialog As NewDialog

Sub Main
	Call CreateOrOpenDatabase()
End Sub

Function CreateOrOpenDatabase()
	Dim button As Integer
	button = Dialog(testDialog)
	If button = -1 Then
		Call OpenPurchaesDatabase()
		Call AppendData()
	ElseIf button = 0 Then
		MsgBox("Hit else")
	End If 
End Function 
 

' File - Import Assistant: Excel
Function OpenPurchaesDatabase()
	Dim task As task 
	Dim obj As obj 
	Dim importedFile As String
	Dim tempFileName As String 
	Set task = Client.GetImportTask("ImportExcel")
	Set obj = client.commondialogs
		importedFile =  obj.fileopen("","","All Files (*.*)|*.*||;")
	task.FileToImport = importedFile
	task.SheetToImport = "Database"
	task.OutputFilePrefix = iSplit(importedFile ,"","\",1,1)
	importedFile =  iSplit(importedFile ,"","\",1,1)
	tempFileName = importedFile
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	importedFile = task.OutputFilePath("Database")
	Set task = Nothing
End Function

Function AppendData()
	Set db = Client.OpenDatabase("On going list.xlsx-Database.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "List of blocked Merchant Category Codes Cleaned.IMD"
	dbName = "On going list " + CStr(Month(Date())) + " " + CStr(Year(Date())) +  ".IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
