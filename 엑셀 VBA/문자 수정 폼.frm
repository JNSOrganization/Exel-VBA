VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���ڼ����� 
   Caption         =   "���� ���� ��"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "���� ���� ��.frx":0000
   StartUpPosition =   1  '������ ���
End
Attribute VB_Name = "���ڼ�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SheetIndex As Integer

Sub example_11()

Dim n As Long
Dim m As Long

sortemp = ListBox1.List(0)

For n = 0 To ListBox1.ListCount - 2        'ù�� ° ����� ��� �ݺ��Ǿ��ϴ��� ���� �� 100����� 99�� �ݺ�
    For m = 0 To ListBox1.ListCount - 2 - n            '�迭�� ������ŭ �ݺ��ϴ� �ݺ���
        If CInt(ListBox1.List(m, 0)) > CInt(ListBox1.List(m + 1, 0)) Then        'm��° ���� m + 1��° �׺��� �۴ٸ�
            SortTemp = ListBox1.List(m + 1)            '�ӽ����庯���� m + 1 ���� ���� �ְ�
            ListBox1.List(m + 1) = ListBox1.List(m)            'm ���� ���� m + 1 �׿� �ִ´�
            ListBox1.List(m) = SortTemp                'm �׿��� �ӽ����庯���� ������ִ� m + 1���� �������� �ִ´�
        End If
    Next m
Next n

End Sub

Private Sub AddButton_Click()

For i = 0 To ListBox1.ListCount - 1
    If ListBox1.List(i, 0) = TextBox2.Text Then
        Exit Sub
    End If
        Next i

ListBox1.AddItem
ListBox1.List(ListBox1.ListCount - 1, 0) = TextBox2.Text
ListBox1.List(ListBox1.ListCount - 1, 1) = TextBox3.Text

Call example_11

End Sub

Private Sub CommandButton1_Click()

Dim AddressRange As Range

If ListBox1.ListCount <> 0 And TextBox1.Text <> "" Then
    Set AddressRange = ActiveSheet.Range(TextBox1.Text)
    For Each Cel In AddressRange
        StartIndex = 1
        Result = ""
        
        For j = 0 To ListBox1.ListCount - 1
            Result = Result & Mid(Cel, StartIndex, ListBox1.List(j, 0) - StartIndex + 1) + ListBox1.List(j, 1)
            StartIndex = ListBox1.List(j, 0) + 1
            Next j
            
        Result = Result & Mid(Cel, StartIndex, Len(Cel) - StartIndex + 1)
        Cel.Value = Result
        Next Cel
End If


End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

If ListBox1.ListCount <> 0 Then
    ListBox1.RemoveItem (ListBox1.ListIndex)
End If
    

End Sub

Private Sub TextBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Dim AddressRange As Range

Set AddressRange = Application.InputBox(prompt:="������ ������ �����ϼ���", Title:="���� ����", Type:=8)

TextBox1.Text = AddressRange.Address

End Sub
