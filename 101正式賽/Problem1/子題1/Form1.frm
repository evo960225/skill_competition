VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Open App.Path & "/in1.txt" For Input As #1
Open App.Path & "/out.txt" For Output As #2
Open App.Path & "/in2.txt" For Input As #3
Dim ex As String, st As String, ch As String, stt As String
Dim sum As Long, exAdd As Long
For ii = 1 To 3 Step 2
    Input #ii, ex
    ex = LCase(ex)
    sum = 0: exAdd = 0: st = ""
    Do
        Input #ii, ch
        If ch = "EOF" Then Exit Do
        st = st & LCase(ch) & " "
    Loop
    For i = 1 To Len(st)
        ch = Mid(st, i, 1)
        If ch <> ":" And ch <> "." And ch <> " " Then
            stt = stt & ch
        ElseIf stt <> "" Then
            sum = sum + 1
            If stt = ex Then exAdd = exAdd + 1
            stt = ""
        End If
    Next
    Print #2, CStr(exAdd) & "," & sum & vbCrLf
Next
Close
Unload Me
End Sub
