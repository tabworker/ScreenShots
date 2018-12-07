REM HTScreen

Sub HTScreen()
'
' HTScreen Макрос
' HTScreen (Hyper Text Screen) - программный модуль (макрос), для автоматического создания гипертекстовой ссылки в электронной таблице Microsoft Office Excel.
'

Dim adress As String
Dim element As String
Dim stlb As String
Dim strok As Integer
Dim link As String

stlb = InputBox("Введите букву столбца с номерами скриншота")
k = InputBox("Введите количество строк в таблице")

For strok = 2 To k
    link = stlb + CStr(strok)
    Range(link).Select
    element = Range(link).Value
        If element <> "" Then
         adress = "[path]\" + element + ".jpg"
                ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
            adress
        End If
Next strok

End Sub
