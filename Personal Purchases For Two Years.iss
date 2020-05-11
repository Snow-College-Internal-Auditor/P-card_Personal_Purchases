Dim db As Object 
Dim subDb As Object
Dim emptyArrayCount As Integer 
Dim NotEmptyArrayCount As Integer 
Dim EmptyDatabaseArray(50) As String 
Dim NotEmptyDatabaseArray(50) As String 

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
	If emptyArrayCount  > 0 Then 
		Call createFolder()
		Call moveDatabase()
	End If
	If NotEmptyArrayCount > 1 Then
		Call AppendAllNoneEmptyDatabases()
	End If 
	Client.Closeall
	Call RemoveUnneededColumns
	Client.RefreshFileExplorer
End Sub


'This calls a script that will loop through pcard statements and append them together
Function CallScriptForPcardStatment
	Client.RunIDEAScriptEx "Z:\2020 Activities\Data Analytics\Active Scripts\Master Scripts\Loop Pull and Join.iss", "", "", "", ""
	PrimaryDatabaseName = "Append Databases.IMD"
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


'This filters the db for the specific merchent code listed in the function name
Function Beauty
	Set db = Client.OpenDatabase(PrimaryDatabaseName)
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "BEAUTY.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""BEAUTY"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name 
Function Cable
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "CABLE.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""CABLE"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Candy_Eating
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "CANDY.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""CANDY"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Catalog
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "CATALOG.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""CATALOG MERCHANT"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Computer
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "COMPUTER.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""COMPUTER"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Department_stores
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "DEPARTMENT.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""DEPARTMENT"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Digital
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "DIGITAL.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""LARGE DIGITAL"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Drinking
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "DRINKING.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""DRINKING"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Florist
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "FLORISTS.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""FLORISTS"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Gift
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "GIFT.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""GIFT"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Medical
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Medical.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION = ""MEDICAL"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Motion_Picture
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "MOTION _PICTURE.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""MOTION PICTURE"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Pet
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "PET.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""PET"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Prints
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "PRINTS.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""PUBLISHING"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Golf
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "PUBLIC_GOLF.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""PUBLIC GOLF"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Religious
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "RELIGIOUS.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""RELIGIOUS"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Sport
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "SPORT.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""SPORT"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Subscription
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Subscriptions.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""CONTINUITY"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name 
Function Video
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "VIDEO.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""VIDEO"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
End Function


'This filters the db for the specific merchent code listed in the function name
Function Wholesale_medical_dentail
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "WHOLESALE_MED_DENTAL.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""WHOLESALE MED/DENTAL"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Call OrganizeDatabase()
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
Function RemoveUnneededColumns 
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
	dbName = "List of blocked Merchant Category Codes Cleaned.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function 
