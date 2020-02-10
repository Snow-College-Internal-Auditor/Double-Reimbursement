Dim importedFile As String
Dim Num As Integer
Dim CleanYearDatabase(50) As String
Dim PrimeDatabase As String
Dim SecondDatabase As String
Dim NewDatabaseName As String

Sub Main
	
	Call NumberOfPulls() 
	i = 0
	Do While i < Num
		Call ExcelImport(i)
		Call CleanYear()
		i = i +1
		Client.RefreshFileExplorer
	Loop
	'If there is only one database we will not need to do a join. If there is then it will run through the join script 
	If Num > 1 Then 
		j = 0
		Do While j + 1 < Num
			Call DatabaseToJoin()
			Call JoinDatabase(PrimeDatabase, SecondDatabase)
			j = j + 1
			Client.RefreshFileExplorer
		Loop
	End If

End Sub

Function NumberOfPulls
	subFileName = InputBox("How many sheets you want to pull: ", "Name Input", "1")
	Num  = Val(subFileName)
End Function 

Function DatabaseToJoin
	PrimeDatabase = InputBox("Enter primary database: ", "Name Input", "Database")
	PrimeDatabase = PrimeDatabase + ".IMD"
	SecondDatabase = InputBox("Enter secondary database: ", "Name Input", "Database")
	SecondDatabase = SecondDatabase + ".IMD"
	NewDatabaseName = InputBox("Enter the neam of the new database: ", "Name Input", "Database")
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

' File: Join Databases
Function JoinDatabase(PrimeDatabase, SecondDatabase)
	Set db = Client.OpenDatabase(PrimeDatabase)
	Set task = db.JoinDatabase
	task.FileToJoin SecondDatabase
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "NAME", "NAME", "A"
	task.CreateVirtualDatabase = False
	dbName = NewDatabaseName + ".IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_REC
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function


