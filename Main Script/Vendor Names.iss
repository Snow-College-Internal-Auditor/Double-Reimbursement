Sub Main
	Call GetVendorNames()	'FIMSMGR.FGBTRND3.IMD
End Sub


' File: Join Databases
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
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function