Sub Main
	Call GetCheckInfo()	'Join Databases.IMD
End Sub


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
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function