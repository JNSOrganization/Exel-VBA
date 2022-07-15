VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "탐색"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   OleObjectBlob   =   "정보탐색 폼.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub SearchData(ByVal k As Integer, ByVal ColumNumber As Integer, ByVal SheetNumber As Integer, ByVal data_rng As Range)

Dim B As Integer
B = 0

For i = 1 To data_rng.Rows.Count
    If data_rng.Cells(i, ColumNumber) = NameTextBox.Text Then
        ResultList.AddItem
        ResultList.List(ResultList.ListCount() - 1, 0) = ResultList.ListCount()
        ResultList.List(ResultList.ListCount() - 1, 1) = data_rng.Cells(i, ColumNumber)
        
        If Len(data_rng.Cells(i, ColumNumber).Offset(0, 1)) > 6 Then
            ResultList.List(ResultList.ListCount() - 1, 2) = Left(data_rng.Cells(i, ColumNumber).Offset(0, 1), 6) & "-" & Right(data_rng.Cells(i, ColumNumber).Offset(0, 1), 7)
        Else
            ResultList.List(ResultList.ListCount() - 1, 2) = Left(data_rng.Cells(i, ColumNumber).Offset(0, 1), 6)
        End If
        
        ResultList.List(ResultList.ListCount() - 1, 3) = Sheets(SheetNumber).Name
        ResultList.List(ResultList.ListCount() - 1, 4) = data_rng.Cells(i, ColumNumber).Address
        B = B + 1
        
        If B = k Then
            Exit For
        End If
        
    End If
    Next i

End Sub

Private Sub CommandButton1_Click()
    ResultList.Clear

    신청인이름 = NameTextBox.Text

    Dim k As Integer
    For i = 1 To Sheets.Count
            Dim data_rng As Range
            Dim filter_value As String
            Dim ColumNumber As Integer

            Set data_rng = Sheets(Sheets(i).Name).Cells(1, "A").CurrentRegion
            filter_value = 신청인이름
            
            For j = 1 To 8
                k = Application.CountIf(data_rng.Columns(j), "=" & filter_value)
                If k >= 1 Then
                    ColumNumber = j
                    Exit For
                End If
                    Next j
        
            If k >= 1 Then
                Call SearchData(k, ColumNumber, i, data_rng)
                k = 0
                    
            End If
                Next i
End Sub

Private Sub ResultList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim ms As Worksheet

Set ms = Workbooks(ThisWorkbook.Name).Sheets(ResultList.List(ResultList.ListIndex, 3))

Application.Goto reference:=ms.Range(ResultList.List(ResultList.ListIndex, 4)), Scroll:=True
'Sheets(Sheets(ResultList.List(ResultList.ListIndex, 3)).Name).Cells(ResultList.List(ResultList.ListIndex, 4), ResultList.List(ResultList.ListIndex, 5))
End Sub
