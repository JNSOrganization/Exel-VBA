Option Explicit
Option Base 1
 
Sub 우편번호및법정리출력함수()
 
    Dim AddressRange As Range
    Dim v
    Dim addv(), v1 As String
    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sT As Date: sT = Time   '시작시간
    Dim nT As Date
    Dim oT As Date
    
    '속도 향상
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    '영역 선택
    On Error GoTo err
        Set AddressRange = Application.InputBox(prompt:="주소가 있는 영역을 선택하세요", Title:="주소 선택", Type:=8)
    On Error GoTo 0
    
    v = AddressRange
    r = UBound(v)
    ReDim addv(r)
    
    Sheets.Add
    Range("A1:E1") = Array("주소", "지번주소", "도로명주소", "법정리", "우편번호")
        
    For i = 1 To r
        
        '데이터가 많을 때 시간 처리
        nT = Time - sT
        If nT <> oT Then
            DoEvents
            Application.StatusBar = "Progress : " & i & " / " & r & "(" & Format(i / r, "0.00%") & ")" & ", " & Format(nT, "hh:mm:ss")
            oT = nT
        End If
        
        '주소 결과를 배열에 삽입
        v1 = v(i, 1)
        addv(i) = NewAdd(v1)
    Next
    
    '출력
    For i = 1 To r
        For j = 1 To 5
            Cells(i + 1, j) = addv(i)(j)
        Next
    Next
        
    Range("A1").CurrentRegion.Columns.AutoFit
    
err:
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .StatusBar = "Progress : 100%" & ", " & Format(Time - sT, "hh:mm:ss")
    End With
    
End Sub
 
Function NewAdd(MyText As String) As Variant
'juso.go.kr API 에서 데이터 가져오는 함수
 
    Dim sURL As String
    Dim oXMLHTTP As Object
    Dim tmp(5) As String
 
    Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    sURL = "http://www.juso.go.kr/addrlink/addrLinkApi.do?currentPage=1&countPerPage=1&keyword=" & ENDECODingURL(MyText) & "&confmKey=U01TX0FVVEgyMDIwMDYwNDExMzEwOTEwOTgyODk="
    
    With oXMLHTTP
        .Open "GET", sURL, False
        .send
        
        On Error Resume Next
        With .responseXML
            tmp(1) = MyText '원래 입력했던 주소
            tmp(2) = .SelectSingleNode("results/juso/jibunAddr").Text   '지번 주소
            tmp(3) = .SelectSingleNode("results/juso/roadAddr").Text    '도로명 주소
            tmp(4) = Split(.SelectSingleNode("results/juso/jibunAddr").Text)(3)   '법정리
            tmp(5) = .SelectSingleNode("results/juso/zipNo").Text   '우편번호
        End With
        On Error GoTo 0
    End With
    
    NewAdd = tmp
 
End Function
 
 
 
Function ENDECODingURL(varText As String, Optional blnEncode = True)
 
    Static objHtmlfile As Object
    
    If objHtmlfile Is Nothing Then
      Set objHtmlfile = CreateObject("htmlfile")
      
      With objHtmlfile.parentWindow
        .execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
        .execScript "function decode(s) {return decodeURIComponent(s)}", "jscript"
      End With
      
    End If
    
    If blnEncode Then
      ENDECODingURL = objHtmlfile.parentWindow.encode(varText)
      
    Else
      ENDECODingURL = objHtmlfile.parentWindow.decode(varText)
    End If
    
End Function
