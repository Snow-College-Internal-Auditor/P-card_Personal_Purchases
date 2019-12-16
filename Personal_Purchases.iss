Dim db As Object 
Dim subDb As Object
Dim arrayCount As Integer 
Dim MyArray(20) As String 
Dim dbName As String 
Dim subFilename As String 
Dim customdbName As String 


Sub Main
	call Filename()
	Call ExcelImport()
	Call Beauty()
	Call Cable()
	Call Candy_Eating()
	Call Catalog()
	Call Computer()
	Call Department_stores()
	Call Digital()
	Call Drinking()
	Call Florist()
	Call Gift()
	Call Medical()
	Call Motion_Picture()
	Call Pet()
	Call Prints()
	Call Golf()
	Call Religious()
	Call Sport()
	Call Subscription()
	Call Video()
	Call Wholesale_medical_dentail()
	If arrayCount > 0 Then 
		Call createFolder()
		Call moveDatabase()
	End If
	Client.RefreshFileExplorer
End Sub

Function Filename

	subFilename = InputBox("Type The Name of The Month: ", "Name Input", "Month")

End Function

Function emptyDatabase
	arrayCount = 1 + arrayCount
	MyArray(arrayCount) = dbName
End Function 

Function createFolder
	' Set the task type.
	Set task = Client.ProjectManagement
	
	subFilename = InputBox("Type The Name of The Month: ", "Name Input", "IDEATest_" + subFilename)
	
	' Create a new folder.
	task.CreateFolder subFilename
	Set task = Nothing
End Function

Function moveDatabase
	' Declare variables and objects.
	Dim path As String
	Dim pm As Object
	
	' Access project management object to manage databases/projects on
	' server.
	Set pm = Client.ProjectManagement
	
	For i = 1 To arrayCount
		' Use path object to get the full path and file name to the specified database.
		Set path = MyArray(i) 
	
		' Move the file from the server to a different server location.
		pm.MoveDatabase path, subFilename
	Next
	
	' Refresh the File Explorer.
	Client.RefreshFileExplorer
	
	' Clear the path.
	Set pm = Nothing
End Function

'Imports starting database
Function ExcelImport
	Set task = Client.GetImportTask("ImportExcel")
	Set obj = client.commondialogs
		dbName =  obj.fileopen("","","All Files (*.*)|*.*||;")
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = iSplit(dbName ,"","\",1,1)
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Sheet1")
	Set task = Nothing
	Set db = Client.OpenDatabase(dbName)
End Function

' Data: Direct Extraction
Function Beauty
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "BEAUTY_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""BEAUTY"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
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
	End If 	

	Set subDb = Nothing
End Function

Function Cable
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "CABLE_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""CABLE"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
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
	End If 
	
	Set subDb = Nothing
End Function

Function Candy_Eating
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "CANDY_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""CANDY"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 
	
	Set subDb = Nothing
End Function

Function Catalog
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "CATALOG_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""CATALOG MERCHANT"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
		
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Computer
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "COMPUTER_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""COMPUTER"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Department_stores
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "DEPARTMENT_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""DEPARTMENT"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Digital
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "DIGITAL_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""LARGE DIGITAL"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	 
	
	Set subDb = Nothing
End Function

Function Drinking
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "DRINKING_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""DRINKING"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Florist
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "FLORISTS_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""FLORISTS"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Gift
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "GIFT_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""GIFT"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Medical
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Medical_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION = ""MEDICAL"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Motion_Picture
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "MOTION _PICTURE_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""MOTION PICTURE"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
		
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Pet
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "PET_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""PET"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Prints
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "PRINTS_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""PUBLISHING"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Golf
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "PUBLIC_GOLF_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""PUBLIC GOLF"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Religious
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "RELIGIOUS_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""RELIGIOUS"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Sport
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "SPORT_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""SPORT"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	 
	
	Set subDb = Nothing
End Function

Function Subscription
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Subscriptions_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""CONTINUITY"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function

Function Video
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "VIDEO_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""VIDEO"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	 
	
	Set subDb = Nothing
End Function

Function Wholesale_medical_dentail
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "WHOLESALE_MED_DENTAL_" + subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""WHOLESALE MED/DENTAL"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	
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
	End If 	
	
	Set subDb = Nothing
End Function


