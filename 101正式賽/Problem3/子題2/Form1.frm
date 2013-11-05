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
Dim n As Long, ch As String, st As String, add As Byte, ans As String
For ii = 1 To 3 Step 2
Input #ii, n
For sti = 1 To n
    Input #ii, st
    ans = ""
    add = 0
    For i = 1 To Len(st)
        ch = Mid(st, i, 1)
        If ch = "0" Or add = 5 Then
            If add <> 5 Then i = i + 1
            If Mid(st, i, 1) = "0" Then
                If add = 0 Then ans = ans & "A" Else ans = ans & (add - 1) * 2
            Else
                If add = 0 Then ans = ans & "B" Else ans = ans & (add - 1) * 2 + 1
            End If
            add = 0
        Else
            add = add + 1
        End If
    Next
    Print #2, Left(ans, 4) & "," & Mid(ans, 5)
Next
Print #2,
Next
Close
Unload Me
End Sub
