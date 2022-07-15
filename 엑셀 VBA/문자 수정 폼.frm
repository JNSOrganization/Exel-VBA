VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 문자수정폼 
   Caption         =   "문자 수정 폼"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   OleObjectBlob   =   "문자 수정 폼.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "문자수정폼"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SheetIndex As Integer

Sub example_11()

Dim n As Long
Dim m As Long

sortemp = ListBox1.List(0)

For n = 0 To ListBox1.ListCount - 2        '첫번 째 행부터 몇번 반복되야하는지 설정 총 100개라면 99번 반복
    For m = 0 To ListBox1.ListCount - 2 - n            '배열의 개수만큼 반복하는 반복문
        If CInt(ListBox1.List(m, 0)) > CInt(ListBox1.List(m + 1, 0)) Then        'm번째 항이 m + 1번째 항보다 작다면
            SortTemp = ListBox1.List(m + 1)            '임시저장변수에 m + 1 항의 값을 넣고
            ListBox1.List(m + 1) = ListBox1.List(m)            'm 항의 값을 m + 1 항에 넣는다
            ListBox1.List(m) = SortTemp                'm 항에는 임시저장변수에 저장되있던 m + 1항의 원래값을 넣는다
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

Set AddressRange = Application.InputBox(prompt:="수정할 영역을 선택하세요", Title:="영역 선택", Type:=8)

TextBox1.Text = AddressRange.Address

End Sub
