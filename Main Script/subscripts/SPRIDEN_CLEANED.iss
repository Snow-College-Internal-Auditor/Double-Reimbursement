Sub Main
	Call Clean_SPRIDEN_Database()	'SATURN.SPRIDEN.IMD
End Sub


' Data: Direct Extraction
Function Clean_SPRIDEN_Database
	Set db = Client.OpenDatabase("SATURN.SPRIDEN.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "SPRIDEN_CLEANED.IMD"
	task.AddExtraction dbName, "", "SPRIDEN_ID = ""00"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function