Sub Main
	Call DirectExtraction()	'EXTRACTION1.IMD
End Sub


' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("EXTRACTION1.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "MOTION _PICTURE.IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_DESCRIPTION = ""MOTION PICTURE"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function