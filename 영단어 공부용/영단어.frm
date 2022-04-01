VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "영단어"
   ClientHeight    =   2940
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5772
   OleObjectBlob   =   "영단어.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Answer As String

Sub Refresh()

Set Word = Range("a1").CurrentRegion

A = QLable.Caption
i = ((Rnd() * Word.Count) Mod Word.Count * Word.Count) + 1

QLable.Caption = Word(i)

If i Mod 2 = 0 Then
    Answer = Word(i - 1)
Else
    Answer = Word(i + 1)
    End If

While (QLable.Caption = "" Or QLable.Caption = A)
    i = ((Rnd() * Word.Count) Mod Word.Count) + 1
    QLable.Caption = Word(i)
    
    If i Mod 2 = 0 Then
        Answer = Word(i - 1)
    Else
        Answer = Word(i + 1)
        End If
    
    Wend

End Sub

Function Check() As Boolean

If AText.Text = Answer Then
    Check = True
    AText.Text = ""
    AText.SetFocus
Else
    Check = False
    End If

End Function

Private Sub CommandButton1_Click()

If Check Then
    Refresh
Else
    MsgBox "다시 시도!"
    End If

End Sub


Private Sub UserForm_Initialize()

Refresh

End Sub
