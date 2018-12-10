Attribute VB_Name = "HTScreen"
Sub HTScreen()
'
' HTScreen Макрос
' HTScreen (Hyper Text Screen) - программный модуль (макрос), для автоматического создания гипертекстовой ссылки в электронной таблице Microsoft Office Excel.
'

Dim adress As String
Dim element As String
Dim stlb As String
Dim strok As Integer
Dim strok_1 As Integer
Dim link As String

stlb = InputBox("Введите букву столбца с номерами скриншота")
strok_1 = InputBox("Сколько строк нужно отступить?")
k = Cells(Rows.Count, stlb).End(xlUp).Row
For strok = strok_1 To k
    link = stlb + CStr(strok)
    Range(link).Select
    element = Range(link).Value
        If element <> "" Then
         adress = "[path]\" + element + ".pas"
                ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
            adress
        End If
Next strok

End Sub

