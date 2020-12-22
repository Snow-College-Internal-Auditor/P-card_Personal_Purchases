Begin Dialog NewDialog 50,49,150,102,"NewDialog", .NewDialog
  Text 17,10,106,14, "Is there a database to save to?", .Text1
  PushButton 17,38,40,14, "Yes", .PushButton1
  PushButton 82,38,40,14, "No", .PushButton2
End Dialog

Dim db As Object 
Dim subDb As Object
Dim emptyArrayCount As Integer 
Dim NotEmptyArrayCount As Integer 
Dim EmptyDatabaseArray(50) As String 
Dim NotEmptyDatabaseArray(50) As String 

Dim categories(18) As String

Dim dbName As String 
Dim subFilename As String 
Dim customdbName As String 
Dim PrimaryDatabaseName As String 
Dim ApprovedVendors As String

Dim getDatabaseDialog As NewDialog

Sub Main
	PrimaryDatabaseName = "Append Databases.IMD"
	Call SetArrayOfCategorys()
	Call ScriptForPcardStatment()
	MsgBox("File explorer is about to open. Bring in the Approved vendor database.")
	Call OpenApprovedVendorsDatabase()
	Call RemoveApprovedVendors()
	For Each item In categories
  		Call Category(item)
  		Client.RefreshFileExplorer
	Next
	If emptyArrayCount  > 0 Then 
		Call createFolder()
		Call moveDatabase()
	End If
	If NotEmptyArrayCount > 1 Then
		Call AppendAllNoneEmptyDatabases()
	End If 
	Client.Closeall
	Call RemoveUnneededColumns()
	Call IndexByName()
	Client.RefreshFileExplorer
	Call CreateOrOpenDatabase()
	Client.RefreshFileExplorer
End Sub


Function SetArrayOfCategorys()

	 categories(0) = "BEAUTY"
	 categories(1) = "CABLE"
	 categories(2) = "CANDY"
	 categories(3) = "CATALOG MERCHANT"
	 categories(4) = "COMPUTER"
	 categories(5) = "DEPARTMENT"
	 categories(6) = "LARGE DIGITAL"
	 categories(7) = "DRINKING"
	 categories(8) = "FLORISTS"
	 categories(9) = "GIFT"
	 categories(10) = "MEDICAL"
	 categories(11) = "MOTION PICTURE"
	 categories(12) = "PET"
	 categories(13) = "PUBLISHING"
	 categories(14) = "PUBLIC GOLF"
	 categories(15) = "RELIGIOUS"
	 categories(16) = "SPORT"
	 categories(17) = "CONTINUITY"
	 categories(18) = "VIDEO"

End Function 


'This calls a script that will loop through pcard statements and append them together
Function ScriptForPcardStatment
	On Error GoTo ErrorHandler
	'TODO make error check if the file cant be reached. 
	Dim filename As String
	Dim obj As Object
	' Access the CommomDialogs object.
	MsgBox("When File explorere opens locate the Loop and Pull script. It will be located in the Audit internal drive ")
	Set obj = Client.CommonDialogs
	filename = obj.FileOpen("","","All Files (*.*)|*.*||;")
	Client.RunIDEAScriptEx filename, "", "", "", ""
		'TODO fix append error if one already is there
	PrimaryDatabaseName = "Append Databases.IMD"
	Client.OpenDatabase(PrimaryDatabaseName)
	Set obj = Nothing
	Exit Sub
	ErrorHandler:
		MsgBox "Idea script Loop Pull and Join could not be run properly. IDEA script stopping."
		Stop
End Function


' File - Import Assistant: Excel
Function OpenApprovedVendorsDatabase()
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
	ApprovedVendors = task.OutputFilePath("Database")
	Set task = Nothing
End Function


Function RemoveApprovedVendors
	Set db = Client.OpenDatabase(ApprovedVendors)
	Set task = db.JoinDatabase
	task.FileToJoin PrimaryDatabaseName
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "APPROVED_MERCHANT_NAME", "MERCHANT_NAME", "A"
	task.CreateVirtualDatabase = False
	PrimaryDatabaseName = "UnverifiedVendors.IMD"
	task.PerformTask PrimaryDatabaseName, "", WI_JOIN_NOC_PRI_MATCH
	Set task = Nothing
	Set db = Nothing
	Call RemoveFieldCreatedDuringJoin()
	Client.OpenDatabase (PrimaryDatabaseName)
End Function


' Remove Field
Function RemoveFieldCreatedDuringJoin
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.TableManagement
	task.RemoveField "APPROVED_MERCHANT_NAME"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function


Function Category(item)
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = item + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = """ & item & """"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()

End Function


'This keeps an array of all the db's with no data
Function emptyDatabase
	emptyArrayCount = 1 + emptyArrayCount
	EmptyDatabaseArray(emptyArrayCount) = dbName
End Function 


'This keeps an array of all db's with data
Function NotEmptyDatabase
	notEmptyArrayCount = 1 + notEmptyArrayCount
	NotEmptyDatabaseArray(notEmptyArrayCount) = dbName
End Function 


'This creates a folder
Function createFolder
	' Set the task type.
	Set task = Client.ProjectManagement
	
	subFilename = "No Purchaes Found"
	
	' Create a new folder.
	task.CreateFolder subFilename
	Set task = Nothing
End Function



'This uses the EmptyDatabaseArray to move all of the db's in the EmptyDatabaseArray to there own folder
Function moveDatabase
	' Declare variables and objects.
	Dim path As String
	Dim pm As Object
	
	' Access project management object to manage databases/projects on
	' server.
	Set pm = Client.ProjectManagement
	
	For i = 1 To emptyArrayCount 
		' Use path object to get the full path and file name to the specified database.
		Set path = EmptyDatabaseArray(i) 
	
		' Move the file from the server to a different server location.
		pm.MoveDatabase path, subFilename
	Next
	
	' Refresh the File Explorer.
	Client.RefreshFileExplorer
	
	' Clear the path.
	Set pm = Nothing
End Function


'Checks if there is an data and if there is not it calls the emptyDatabase
'function and if there is data it calls the NotEmptyDatabase function 
Function OrganizeDatabase
	Set subDb = Client.OpenDatabase (dbName)
	
	'Checks if column name has any rows
	Set stats = subDb.FieldStats("Name")
	' Sets num equal to Number of rows in column
	Dim num As Integer
	num = stats.NumRecords()
	
	'If num is zero it will close the databse
	If num < 1 Then
		subDb.Close
		Call emptyDatabase()
	ElseIf num >= 1 Then
		Call NotEmptyDatabase() 
	End If 	
	Set subDb = Nothing
End Function  


'This loops through the NotEmptyDatabaseArray and appends all 
'of the databases together into one database
Function AppendAllNoneEmptyDatabases
	' Declare variables and objects.
	Dim path As String
	Dim pm As Object
	
	' Access project management object to manage databases/projects on
	' server.
	Set pm = Client.ProjectManagement
	Dim j As Integer 
	j = 0
	For i = 1 To NotEmptyArrayCount  
		' Use path object to get the full path and file name to the specified database.
		If i = 1 Then 
			Set path = NotEmptyDatabaseArray(i) 
		
			Set db = Client.OpenDatabase(path)
			Set task = db.AppendDatabase
			Set path = NotEmptyDatabaseArray(i + 1)
			task.AddDatabase path
			If j = NotEmptyArrayCount Then
				dbName = "List of blocked Merchant Category Codes"
			ElseIf j < NotEmptyArrayCount Then 
				dbName = "Append Databases " + path
			End If
			task.PerformTask dbName, ""
			i = i + 1
			j = j + 3
			Client.RefreshFileExplorer
		ElseIf i >= 3 Then 
			Set db = Client.OpenDatabase(dbName)
			Set task = db.AppendDatabase
			Set path = NotEmptyDatabaseArray(i)
			task.AddDatabase path
			If j = NotEmptyArrayCount Then
				dbName = "List of blocked Merchant Category Codes"
			ElseIf j < NotEmptyArrayCount Then 
				dbName = "Append Databases " + path
			End If
			task.PerformTask dbName, ""
			j = j + 1
			Client.RefreshFileExplorer

		End If
	Next
	
	' Refresh the File Explorer.
	Client.RefreshFileExplorer
	
	' Clear the path.
	Set pm = Nothing
	Set task = Nothing
	Set db = Nothing
End Function


'RemoveUnneededColumns 
Function RemoveUnneededColumns()
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "SHORT_NAME"
	task.AddFieldToInc "ACCOUNT_NUMBER"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "TRANSACTION_STATUS"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_GROUP_CODE"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_DESCRIPTION"
	task.AddFieldToInc "MERCHANT_NAME"
	task.AddFieldToInc "MERCHANT_CITY"
	task.AddFieldToInc "MERCHANT_STATE_PROVINCE"
	task.AddFieldToInc "MERCHANT_ORDER_NUMBER"
	task.AddFieldToInc "TRANSACTION_COMMENTS"
	task.AddFieldToInc "DEPARTMENT"
	PrimaryDatabaseName = "List of blocked Merchant Category Codes Cleaned.IMD"
	task.AddExtraction PrimaryDatabaseName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (PrimaryDatabaseName)
End Function 


Function IndexByName()
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Index
	task.AddKey "NAME", "A"
	task.Index FALSE
	Set task = Nothing
	Set db = Nothing
End Function 


Function CreateOrOpenDatabase()
	Dim button As Integer
	button = Dialog(getDatabaseDialog)
	If button = 1 Then
		Call OpenPurchaesHistoryDatabase()
		Call AppendData()
	ElseIf button = 2 Then
		Call RenamePrimaryDatabaseNameForExport()
		Call ExportPurchaesHistoryDatabase()
	End If 
End Function 
 

' File - Import Assistant: Excel
Function OpenPurchaesHistoryDatabase()
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


Function RenamePrimaryDatabaseNameForExport()
	Client.Closeall
	Dim newDatabaseName As String 
	newDatabaseName = "On going list " + CStr(Month(Date())) + " " + CStr(Year(Date())) +  ".IMD"
	Set ProjectManagement = client.ProjectManagement
	ProjectManagement.RenameDatabase PrimaryDatabaseName, newDatabaseName
	PrimaryDatabaseName = newDatabaseName
	Set ProjectManagement = Nothing
End Function 


Function ExportPurchaesHistoryDatabase()
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Index
	task.AddKey "NAME", "A"
	task.Index FALSE
	task = db.ExportDatabase
	task.IncludeAllFields
	' Display the setup dialog box before performing the task.
	task.DisplaySetupDialog 0
	Set db = Nothing
	Set task = Nothing
End Function
