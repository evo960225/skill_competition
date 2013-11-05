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
Dim n As Long, m As Byte, s As Byte, qa As Long, qb As Long, dx As Long
For ii = 1 To 3 Step 2
Input #ii, n
For sti = 1 To n
    Input #ii, m
    m = m - 1
    qa = 0: qb = 0: dx = 0
    ReDim srr(m, m), q(m, 1)
    For i = 1 To m
      Input #ii, s, s
      srr(i, s) = 1: srr(s, i) = 1
    Next
    Do
      dx = q(qa, 0): sum = q(qa, 1): qa = qa + 1
      For i = 0 To m
      If srr(dx, i) > 0 Then qb = qb + 1: q(qb, 0) = i: q(qb, 1) = sum + 1: srr(i, dx) = 0
      Next
    Loop Until qa > qb
    Print #2, CStr(sum + 1)
    If sti <> n Then Input #ii, s
Next
Print #2,
Next
Close
Unload Me
End Sub
