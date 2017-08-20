Attribute VB_Name = "Module1"
Public Conn As New ADODB.Connection
Public TempRS As New ADODB.Recordset
Public RSF As New ADODB.Recordset
Public RS As New ADODB.Recordset
Public UserNameVar, StrSql, Str1, Str2, StrVar, RegNoVar As String
Public MsgVar, TranMain As String
Public I, J, K, RNo, NoVar As Long
Public INoVar As Byte
Enum CtrlType
     TextBox = 1
     ComboBox = 2
End Enum
Public Sub ClearTxtControls(frm As Object, ControlType As CtrlType, Optional Tagstr As Variant)
Dim Contrl As Object
For Each Contrl In frm.Controls
         If Not (IsMissing(Tagstr)) Then
         If Trim(UCase(Contrl.Tag)) = Trim(UCase(Tagstr)) Then
            Contrl.Text = ""
            Exit For
          End If
          Else
          Select Case ControlType
                 Case CtrlType.ComboBox
                   If TypeOf Contrl Is ComboBox Then Contrl.Text = ""
                 Case CtrlType.TextBox
                   If TypeOf Contrl Is TextBox Then Contrl.Text = ""
          End Select
          End If
    Next
Set Contrl = Nothing
End Sub
Public Function CheckChar(CharString) As String
Dim L1, Con1, sinchar

L1 = Len(CharString)
For I = 1 To L1
If Con1 = True Then CharString = Mid(CharString, 1, (I - 1)) & UCase(Mid(CharString, I, 1)) & Mid(CharString, I + 1, L1)
sinchar = Mid(CharString, I, 1)
    If sinchar = " " Then
    Con1 = True
    Else
    Con1 = False
    End If
Next I

CheckChar = CharString
End Function

Function CheckNum(KeyNum) As Integer
If KeyNum = 8 Then CheckNum = KeyNum: Exit Function
If KeyNum < 46 Or KeyNum > 57 Then
CheckNum = 0
MsgBox ("Please Enter Numbers Only")
Else
CheckNum = KeyNum
End If
End Function

Function DateFormat(vdate1)
DateFormat = Format(vdate1, "dd/MMM/yyyy")
End Function

Function ForCur(val1)
ForCur = Format(val1, "##,##,###.00")
End Function

