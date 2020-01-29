Sub Main
	Call FinalTransactionDatabase()	'Join Databases1.IMD
End Sub


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
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function