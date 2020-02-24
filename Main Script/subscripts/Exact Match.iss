Dim importedFile As String
Dim Num As Integer
Dim CleanYearDatabase(50) As String
Dim PrimeDatabase As String
Dim SecondDatabase As String
Dim NewDatabaseName As String
Dim singleDatabase As String 
Dim tempFileName As String 


Sub Main
	Call CreateCard()
	Call CleanCardDatabase()
	Call FilterCardDatabase()
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
		Call ExactMatch2()
	ElseIf Num = 1 Then
		Call ExactMatch1()
	End If
	
	'Call CleanYear()
	'Call ExactMatch()
	Call AppendField()
	Call FilterForValidDate()
	Call ExportDatabaseXLSX()
	Client.RefreshFileExplorer
End Sub

' File: Join Databases
Function CreateCard
	Set db = Client.OpenDatabase("Final Transaction Detail With SPRIDEN_ID.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin "FIMSMGR.FTVCARD.IMD"
	task.IncludeAllPFields
	task.AddSFieldToInc "CARD_LAST_FOUR"
	task.AddMatchKey "FABINVH_VEND_PIDM", "FTVCARD_CARDHOLDER_PIDM", "A"
	task.CreateVirtualDatabase = False
	dbName = "Card Number Added To Final Transaction.IMD"
	task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Client.CloseDatabase  "Final Transaction Detail With SPRIDEN_ID.IMD"
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

Function CleanCardDatabase
	Set db = Client.OpenDatabase("Card Number Added To Final Transaction.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Card Number Added To Final Transaction_Clean.IMD"
	task.AddExtraction dbName, "", "@NoMatch(SPRIDEN_ID,  """" )  .AND. @NoMatch(CARD_LAST_FOUR,  """" )    "
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Client.CloseDatabase  "Card Number Added To Final Transaction.IMD"
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

Function FilterCardDatabase
	Set db = Client.OpenDatabase("Card Number Added To Final Transaction_Clean.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "FGBTRND_DOC_SEQ_CODE"
	task.AddFieldToInc "FGBTRND_DOC_CODE"
	task.AddFieldToInc "FGBTRND_SEQ_NUM"
	task.AddFieldToInc "FGBTRND_ACTIVITY_DATE_DATE"
	task.AddFieldToInc "FGBTRND_USER_ID"
	task.AddFieldToInc "FGBTRND_ACCI_CODE"
	task.AddFieldToInc "FGBTRND_FUND_CODE"
	task.AddFieldToInc "FGBTRND_ORGN_CODE"
	task.AddFieldToInc "FGBTRND_ACCT_CODE"
	task.AddFieldToInc "FGBTRND_TRANS_AMT"
	task.AddFieldToInc "FABINCK_CHECK_NUM"
	task.AddFieldToInc "FABINCK_ACTIVITY_DATE_DATE"
	task.AddFieldToInc "FABINCK_NET_AMT"
	task.AddFieldToInc "FABINVH_VEND_PIDM"
	task.AddFieldToInc "SPRIDEN_ID"
	task.AddFieldToInc "CARD_LAST_FOUR"
	dbName = "Card Number Filtered.IMD"
	task.AddExtraction dbName, "", "FGBTRND_DR_CR_IND = ""+""    "
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Client.CloseDatabase "Card Number Added To Final Transaction_Clean.IMD"
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function	

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
	importedFile =  iSplit(importedFile ,"","\",1,1)
	tempFileName = importedFile
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
	singleDatabase = tempFileName + " Clean.IMD"
	task.AddExtraction singleDatabase, "", ""
	MsgBox(singleDatabase)
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

Function ExactMatch1
	Set db = Client.OpenDatabase("Card Number Filtered.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin singleDatabase
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "FGBTRND_TRANS_AMT", "TRANSACTION_AMOUNT", "A"
	task.AddMatchKey "CARD_LAST_FOUR", "ACCOUNT_NUMBER", "A"
	task.CreateVirtualDatabase = False
	dbName = "Exact Match.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Client.CloseDatabase "Card Number Filtered.IMD"
	Client.CloseDatabase  NewDatabaseName  + ".IMD"
	Set task = Nothing
	Set db = Nothing
End Function

Function ExactMatch2
	Set db = Client.OpenDatabase("Card Number Filtered.IMD")
	Set task = db.JoinDatabase
	task.FileToJoin NewDatabaseName  + ".IMD"
	task.IncludeAllPFields
	task.IncludeAllSFields
	task.AddMatchKey "FGBTRND_TRANS_AMT", "TRANSACTION_AMOUNT", "A"
	task.AddMatchKey "CARD_LAST_FOUR", "ACCOUNT_NUMBER", "A"
	task.CreateVirtualDatabase = False
	dbName = "Exact Match.IMD"
		task.PerformTask dbName, "", WI_JOIN_ALL_IN_PRIM
	Client.CloseDatabase "Card Number Filtered.IMD"
	Client.CloseDatabase  NewDatabaseName  + ".IMD"
	Set task = Nothing
	Set db = Nothing
End Function

Function AppendField
	Set db = Client.OpenDatabase("Exact Match.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "TIME_BETWEEN"
	field.Description = "Time between dates"
	field.Type = WI_VIRT_NUM
	field.Equation = "@Age(FGBTRND_ACTIVITY_DATE_DATE, TRANSACTION_DATE)"
	field.Decimals = 0
	task.AppendField field
	task.PerformTask
	Client.CloseDatabase "Exact Match.IMD"
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

Function FilterForValidDate
	Set db = Client.OpenDatabase("Exact Match.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Exact Match Narrow.IMD"
	task.AddExtraction dbName, "", "TRANSACTION_AMOUNT > 0  .AND.  TRANSACTION_DATE < FGBTRND_ACTIVITY_DATE_DATE  .AND. TIME_BETWEEN < 100"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Client.CloseDatabase "Exact Match.IMD"
	Client.CloseDatabase "Exact Match Narrow.IMD"
	Set task = Nothing
	Set db = Nothing
End Function 

' File - Export Database: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("Exact Match Narrow.IMD")
	Set task = db.ExportDatabase
	' Configure the task.
	task.IncludeAllFields
	' Display the setup dialog box before performing the task.
	task.DisplaySetupDialog 0
	' Clear the memory.
	Set db = Nothing
	Set task = Nothing
End Function




