Sub Main
	Call FilterFinalTransaction()	'Final Transaction Detail.IMD
End Sub


' Data: Direct Extraction
'This filters out any SPRIDEN ID's that do not have a number. So, gets only vendors with baner id's
Function FilterFinalTransaction
	Set db = Client.OpenDatabase("Final Transaction Detail.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Final Transaction Detail With SPRIDEN_ID.IMD"
	task.AddExtraction dbName, "", "@NoMatch(SPRIDEN_ID, """") .AND. FGBTRND_FIELD_CODE == ""03"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function