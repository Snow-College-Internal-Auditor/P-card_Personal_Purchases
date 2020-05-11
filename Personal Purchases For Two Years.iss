Dim db As Object 
Dim subDb As Object
Dim arrayCount As Integer 
Dim MyArray(20) As String 

Dim dbName As String 
Dim subFilename As String 
Dim customdbName As String 
Dim PrimaryDatabaseName As String 
Sub Main
	Call CallScriptForPcardStatment()
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
	Client.Closeall
	Client.RefreshFileExplorer
End Sub

Function CallScriptForPcardStatment
	Client.RunIDEAScriptEx "Z:\2020 Activities\Data Analytics\Active Scripts\Master Scripts\Loop Pull and Join.iss", "", "", "", ""
	PrimaryDatabaseName = InputBox("Now that we have appended all the Databases together what is the Primary one? ", "Name Input", "[Year][Month]TransactionStatement.xlsx Clean")
	PrimaryDatabaseName = PrimaryDatabaseName + ".IMD"
End Function


Function emptyDatabase
	arrayCount = 1 + arrayCount
	MyArray(arrayCount) = dbName
End Function 


Function createFolder
	' Set the task type.
	Set task = Client.ProjectManagement
	
	subFilename = "No Purchaes Found"
	
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


' Data: Direct Extraction
Function Beauty
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "BEAUTY.IMD"
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
	dbName = "CABLE.IMD"
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
	dbName = "CANDY.IMD"
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
	dbName = "CATALOG.IMD"
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
	dbName = "COMPUTER.IMD"
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
	dbName = "DEPARTMENT.IMD"
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
	dbName = "DIGITAL.IMD"
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
	dbName = "DRINKING.IMD"
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
	dbName = "FLORISTS.IMD"
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
	dbName = "GIFT.IMD"
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
	dbName = "Medical.IMD"
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
	dbName = "MOTION _PICTURE.IMD"
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
	dbName = "PET.IMD"
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
	dbName = "PRINTS.IMD"
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
	dbName = "PUBLIC_GOLF.IMD"
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
	dbName = "RELIGIOUS.IMD"
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
	dbName = "SPORT.IMD"
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
	dbName = "Subscriptions.IMD"
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
	dbName = "VIDEO.IMD"
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
	dbName = "WHOLESALE_MED_DENTAL.IMD"
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


