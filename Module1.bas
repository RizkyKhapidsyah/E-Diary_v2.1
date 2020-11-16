Attribute VB_Name = "Module1"
Public Cn As ADODB.Connection
Public Rs As ADODB.Recordset
Public UserName As String
Public dDate As String
Public Sub Connection()
    Set Cn = New ADODB.Connection
    Cn.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Diary\dbDairy.mdb"
End Sub

Public Sub CloseProgram()
    reply = MsgBox("Do You Really Want To Exit?", 4 + 64, "Exit?")
    If reply = vbYes Then
        MsgBox "Thanks For Using This Electronic Diary. For More Free Excelent Softwares, Please Contact Lim Meng Huey(xtremeprohacker@yahoo.com.sg).", , "Thank You."
        End
    Else
        Exit Sub
    End If
End Sub
