Dim importedFile As String
Dim Num As Integer
Dim CleanYearDatabase(1) As String


Sub Main
	
	Call NumberOfPulls() 
	i = 0
	Do While i < Num
		Call ExcelImport(i)
		Call CleanYear()
		i = i +1
		Client.RefreshFileExplorer
	Loop
	If Num > 1 Then
		i = 0 
		Do While i < Num 
			MsgBox(CleanYearDatabase(i))
			i = i + 1
			'Will make it so it will pull from the array CleanYearDatabase what databases need to be inner joined together
		Loop
	End If


End Sub

Function NumberOfPulls
	subFileName = InputBox("How many sheets you want to pull: ", "Name Input", "1")
	Num  = Val(subFileName)
End Function 

' File - Import Assistant: Excel
Function ExcelImport(i)
	Set task = Client.GetImportTask("ImportExcel")
	Set obj = client.commondialogs
		importedFile =  obj.fileopen("","","All Files (*.*)|*.*||;")
	task.FileToImport = importedFile
	task.SheetToImport = "Sheet1"
	task.OutputFilePrefix = iSplit(importedFile ,"","\",1,1)
	task.FirstRowIsFieldName = "TRUE"
	task.EmptyNumericFieldAsZero = "TRUE"
	task.PerformTask
	importedFile = task.OutputFilePath("Sheet1")
	'adding the name of the new database into the array
	CleanYearDatabase(i) = importedFile 
	Set task = Nothing
End Function

' Data: Direct Extraction
Function CleanYear
	Set db = Client.OpenDatabase(importedFile)
	Set task = db.Extraction
	task.AddFieldToInc "NAME"
	task.AddFieldToInc "ACCOUNT_NUMBER"
	task.AddFieldToInc "TRANSACTION_DATE"
	task.AddFieldToInc "TRANSACTION_AMOUNT"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_GROUP_CODE"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_GROUP_DESCRIPTION"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE"
	task.AddFieldToInc "MERCHANT_CATEGORY_CODE_DESCRIPTION"
	task.AddFieldToInc "MERCHANT_NAME"
	dbName = importedFile +  " Clean.IMD"
	task.AddExtraction dbName, "", ""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Client.CloseDatabase importedFile
	Set task = Nothing
	Set db = Nothing
End Function


