VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "��Ʈ ����"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7140
   OleObjectBlob   =   "��Ʈ ���� ��.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim i As Integer
Dim j As Integer
Dim myRange As Range
Dim SheetName As String

For j = 0 To Sheets.Count - 1
    If ListBox2.Selected(j) Then
        SheetName = Sheets(i + 1).Name
    End If
    
    Next j

If SheetName = "" Then
    MsgBox "���� ������ ���� ��Ʈ�� �������� �ʾҽ��ϴ�."
    End

End If

Set mySheet = Worksheets(SheetName)

For i = 0 To Sheets.Count - 1
    If ListBox1.Selected(i) Then
        Set myRange = mySheet.Cells(Rows.Count, 1).End(xlUp).Offset(0, 0)
        Sheets(i + 1).UsedRange.Offset(TextBox1.Text).Copy myRange
    End If

    Next i
    
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then

If (KeyAscii = 8) Then '�� �齺���̽�Ű�� ������ �ƹ� ������ ����.

Else

MsgBox "���ڸ� �Է��ϼ���"

KeyAscii = 0

End If

End If

End Sub

Private Sub UserForm_Initialize()

For i = 1 To Sheets.Count
        ListBox1.AddItem (Sheets(i).Name)
        Next i

For i = 1 To Sheets.Count
        ListBox2.AddItem (Sheets(i).Name)
        Next i

End Sub
