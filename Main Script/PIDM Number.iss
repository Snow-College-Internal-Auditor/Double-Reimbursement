Sub Main
	Call GetPIDM_Number()	'Transaction Detail with Check Num.IMD
End Sub


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
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function