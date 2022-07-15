VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "탐색"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   OleObjectBlob   =   "검색 폼.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub SearchData(ByVal k As Integer, ByVal ColumNumber As Integer, ByVal SheetNumber As Integer, ByVal data_rng As Range)

Dim C As Range

Set C = data_rng.Find(NameTextBox.Text)

ResultList.AddItem
ResultList.List(ResultList.ListCount() - 1, 0) = ResultList.ListCount()
ResultList.List(ResultList.ListCount() - 1, 1) = Sheets(Sheets(SheetNumber).Name).Range(C.Address).Offset(0, 0)
ResultList.List(ResultList.ListCount() - 1, 2) = Sheets(Sheets(SheetNumber).Name).Range(C.Address).Offset(0, 1)
    
For i = 1 To k - 1
    Set C = data_rng.FindNext(C)
    ResultList.AddItem
    ResultList.List(ResultList.ListCount() - 1, 0) = ResultList.ListCount()
    ResultList.List(ResultList.ListCount() - 1, 1) = Sheets(Sheets(SheetNumber).Name).Range(C.Address).Offset(0, 0)
    ResultList.List(ResultList.ListCount() - 1, 2) = Sheets(Sheets(SheetNumber).Name).Range(C.Address).Offset(0, 1)
    Next i
    
End Sub

Private Sub CommandButton1_Click()
    ResultList.Clear
    
    Dim excelFile As Workbook
    Set excelFile = GetObject(ActiveWorkbook.Path + "\당진시 상생지원금 지급관리.xlsm")

    신청인이름 = NameTextBox.Text
    '신청인주민번호 = NumberTextBox.Text

    Dim k As Integer
    For i = 1 To Sheets.Count
            Dim data_rng As Range
            Dim filter_value As String
            Dim ColumNumber As Integer

            Set data_rng = Sheets(Sheets(i).Name).Cells(1, "A").CurrentRegion
            'filter_value = 신청인 & 신청인주민번호
            filter_value = 신청인이름
            
            For j = 1 To 8
                k = k + Application.CountIf(data_rng.Columns(j), "=" & filter_value)
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
