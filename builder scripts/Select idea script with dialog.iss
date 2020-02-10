Begin Dialog NewDialog 50,50,150,150,"NewDialog", .NewDialog
  PushButton 51,49,50,22, "Exact Match", .PushButton1
End Dialog
Sub Main
	Dim dlg As NewDialog
	button = Dialog(dlg)
		If button = 1 Then
			Client.RunIDEAScriptEx "C:\Users\mckinnin.lloyd\Documents\Active Projects\Double-Reimbursement\Main Script\subscripts\Exact Match.iss", "", "", "", ""
		end if
End Sub
