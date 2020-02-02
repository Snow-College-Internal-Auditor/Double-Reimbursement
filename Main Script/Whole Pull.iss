Sub Main
	Call GetVendorNames()	
	Call GetCheckInfo()
	Call GetPIDM_Number()
	Call Clean_SPRIDEN_Database()
	Call Get_SPRIDEN_ID()
	Call FilterFinalTransaction()
	Client.RefreshFileExplorer
End Sub


' File: Join Databases
'Adds vendor names to the end of the database
Function GetVendorNames
	Set db = Client.OpenDatabase("FIMSMGR.FGBTRND3.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "FIMSMGR.FGBTRNH.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "FGBTRNH_TRANS_DESC"
	task.AddMatchKey "FGBTRND_DOC_CODE", "FGBTRNH_DOC_CODE", "A"
	task.AddMatchKey "FGBTRND_SEQ_NUM", "FGBTRNH_SEQ_NUM", "A"
	task.CreateVirtualDatabase = False
	dbName = "Vendor Names.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Client.CloseDatabase "FIMSMGR.FGBTRND3.IMD"
	Set task = Nothing
	Set db = Nothing
End Function

' File: Join Databases
'Adds Check information to the end of the database
Function GetCheckInfo
	Set db = Client.OpenDatabase("Vendor Names.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "FIMSMGR.FABINCK.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "FABINCK_CHECK_NUM"
	task.AddSFieldToInc "FABINCK_ACTIVITY_DATE_DATE"
	task.AddSFieldToInc "FABINCK_CHECK_TYPE_IND"
	task.AddSFieldToInc "FABINCK_NET_AMT"
	task.AddMatchKey "FGBTRND_DOC_CODE", "FABINCK_INVH_CODE", "A"
	task.CreateVirtualDatabase = False
	dbName = "Transaction Detail with Check Information.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Client.CloseDatabase "Vendor Names.IMD"
	Set task = Nothing
	Set db = Nothing
End Function

' File: Join Databases
'Adds the PIDM number to the end of the database
Function GetPIDM_Number
	Set db = Client.OpenDatabase("Transaction Detail with Check Information.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "FIMSMGR.FABINVH.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "FABINVH_VEND_PIDM"
	task.AddMatchKey "FGBTRND_DOC_CODE", "FABINVH_CODE", "A"
	task.CreateVirtualDatabase = False
	dbName = "PIDM Number.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Client.CloseDatabase "Transaction Detail with Check Information.IMD"
	Set task = Nothing
	Set db = Nothing
End Function

' Data: Direct Extraction
Function Clean_SPRIDEN_Database
	Set db = Client.OpenDatabase("SATURN.SPRIDEN.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "SPRIDEN_CLEANED.IMD"
	task.AddExtraction dbName, "", "SPRIDEN_ID = ""00"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Client.CloseDatabase "SATURN.SPRIDEN.IMD"
	Client.CloseDatabase  "SPRIDEN_CLEANED.IMD"
	Set task = Nothing
	Set db = Nothing
End Function

' File: Join Databases
'Add this gets the SPRIDEN number, which is a baner id, and adds it to the end of the database. 
Function Get_SPRIDEN_ID
	Set db = Client.OpenDatabase("PIDM Number.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "SPRIDEN_CLEANED.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "SPRIDEN_ID"
	task.AddMatchKey "FABINVH_VEND_PIDM", "SPRIDEN_PIDM", "A"
	task.CreateVirtualDatabase = False
	dbName = "Final Transaction Detail.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Client.CloseDatabase "PIDM Number.IMD"
	Set task = Nothing
	Set db = Nothing
End Function

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
	Client.CloseDatabase "Final Transaction Detail.IMD"
	Set task = Nothing
	Set db = Nothing
End Function