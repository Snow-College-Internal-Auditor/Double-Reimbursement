Dim subFilename As String 
Dim dbName As String

Sub Main
	Call ExcelImport()	'C:\Users\mckinnin.lloyd\Documents\Active Projects\Double Rem\Copy of Year2018July2019JuneTransactionStatement.xlsx
	Call DirectExtraction()	
	Call Order()
End Sub


' File - Import Assistant: Excel
Function ExcelImport
	Set task = Client.GetImportTask("ImportExcel")
	Set obj = client.commondialogs
		dbName =  obj.fileopen("","","All Files (*.*)|*.*||;")
	task.FileToImport = dbName
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = iSplit(dbName ,"","\",1,1)
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	dbName = task.OutputFilePath("Sheet1")
	Set task = Nothing
End Function

' Data: Direct Extraction
Function DirectExtraction
	Set db = Client.OpenDatabase(dbName)
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION"
	task.AddFieldToInc "MERCHANT_NAME"
	subFilename = InputBox("Type The Name of The Month: ", "Name Input", "Database")
	dbName = subFilename + ".IMD"
	task.AddExtraction dbName, "", "MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION = ""AIRLINE""  .OR. MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION =  ""AUTO/RV DEALERS"" .OR. MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION = ""EATING/DRINKING""  .OR.  MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION =  ""HOTELS"" .OR. MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION =  ""OTHER TRAVEL"" .OR. MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION =  ""RENTAL CARS"" .OR. MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION = ""VEHICLE EXPENSE"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

Function Order
Set db = Client.OpenDatabase(dbName)
Set task = db.Index
task.AddKey "MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION", "A"
task.Index FALSE
Set task = Nothing
Set db = Nothing
End Function 

