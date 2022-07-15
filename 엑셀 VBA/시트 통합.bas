Attribute VB_Name = "Module3"
Sub SheetUnit()

Dim i As Integer
Dim mySheet As Worksheet
Dim myRange As Range

Set AddressRange = Application.InputBox(prompt:="통합 데이터를 저장 할 곳을 입력해주세요.", Title:="시트 이름 입력", Type:=2)

Set mySheet = Worksheets("AddressRange")
For Each sht In ActiveWindow.SelectedSheets
    Set myRange = mySheet.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)
    sht.UsedRange.Copy myRange
    Next sht

End Sub

