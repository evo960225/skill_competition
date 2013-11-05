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
Dim n As Long, st As String
For ii = 1 To 3 Step 2
Input #ii, n
For sti = 1 To n
    Input #ii, st
    Print #2, Pri(st, 4) & "," & Pri(st, 3)
Next
Print #2,
Next
Close
Unload Me
End Sub

Function Pri(ByVal st, o)
Dim ch As String, le As Integer, add As Byte
    le = Len(st)
    st = String(((le + o - 1) \ o) * o - le, "0") & st
    For i = 0 To le - 1 Step o
        For j = 1 To o
            If Mid(st, i + j, 1) = "1" Then add = add + 2 ^ (o - j)
        Next
        If add > 9 Then ans = ans & Chr(add + 55) Else ans = ans & add
        add = 0
    Next
Pri = ans
End Function
