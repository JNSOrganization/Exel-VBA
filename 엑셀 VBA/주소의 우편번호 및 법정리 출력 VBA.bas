Attribute VB_Name = "Module1"
Option Explicit
Option Base 1
 
Sub Macro()
 
    Dim AddressRange As Range
    Dim v
    Dim addv(), v1 As String
    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sT As Date: sT = Time   '���۽ð�
    Dim nT As Date
    Dim oT As Date
    
    '�ӵ� ���
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
    End With
    
    '���� ����
    On Error GoTo err
        Set AddressRange = Application.InputBox(prompt:="�ּҰ� �ִ� ������ �����ϼ���", Title:="�ּ� ����", Type:=8)
    On Error GoTo 0
    
    v = AddressRange
    r = UBound(v)
    ReDim addv(r)
    
    Sheets.Add
    Range("A1:E1") = Array("�ּ�", "�����ּ�", "���θ��ּ�", "������", "�����ȣ")
        
    For i = 1 To r
        
        '�����Ͱ� ���� �� �ð� ó��
        nT = Time - sT
        If nT <> oT Then
            DoEvents
            Application.StatusBar = "Progress : " & i & " / " & r & "(" & Format(i / r, "0.00%") & ")" & ", " & Format(nT, "hh:mm:ss")
            oT = nT
        End If
        
        '�ּ� ����� �迭�� ����
        v1 = v(i, 1)
        addv(i) = NewAdd(v1)
    Next
    
    '���
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
'juso.go.kr API ���� ������ �������� �Լ�
 
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
            tmp(1) = MyText '���� �Է��ߴ� �ּ�
            tmp(2) = .SelectSingleNode("results/juso/jibunAddr").Text   '���� �ּ�
            tmp(3) = .SelectSingleNode("results/juso/roadAddr").Text    '���θ� �ּ�
            tmp(4) = Split(.SelectSingleNode("results/juso/jibunAddr").Text)(3)   '������
            tmp(5) = .SelectSingleNode("results/juso/zipNo").Text   '�����ȣ
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
