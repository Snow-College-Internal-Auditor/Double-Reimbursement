Sub Main
	Call JoinDatabase()
	Call DirectExtraction()
	Client.RefreshFileExplorer
End Sub


' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("PIDM Number_Test.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "SPRIDEN_CLEANED.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "SPRIDEN_ID"
	task.AddMatchKey "FABINVH_VEND_PIDM", "SPRIDEN_PIDM", "A"
	task.CreateVirtualDatabase = False
	dbName = "Final Transaction Detail_Test1.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	db.Close
	Set task = Nothing
	Set db = Nothing
End Function

' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase("Final Transaction Detail_Test1.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Final Transaction Detail With SPRIDEN_ID_Test1.IMD"
	task.AddExtraction dbName, "", "@NoMatch(SPRIDEN_ID, """") .AND. FGBTRND_FIELD_CODE == ""03"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	db.Close
	Set task = Nothing
	Set db = Nothing
End Function

' File: Join Databases
Function JoinDatabase
	Set db = Client.OpenDatabase("Final Transaction Detail With SPRIDEN_ID_Test1.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "FIMSMGR.FTVCARD.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "CARD_LAST_FOUR"
	task.AddMatchKey "FABINVH_VEND_PIDM", "FTVCARD_CARDHOLDER_PIDM", "A"
	task.CreateVirtualDatabase = False
	dbName = "Card Number Added To Final Transaction_Test.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function



