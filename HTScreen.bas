Attribute VB_Name = "HTScreen"
Sub HTScreen()
'
' HTScreen ������
' HTScreen (Hyper Text Screen) - ����������� ������ (������), ��� ��������������� �������� �������������� ������ � ����������� ������� Microsoft Office Excel.
'

Dim adress As String
Dim element As String
Dim stlb As String
Dim strok As Integer
Dim strok_1 As Integer
Dim link As String

stlb = InputBox("������� ����� ������� � �������� ���������")
strok_1 = InputBox("������� ����� ����� ���������?")
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

