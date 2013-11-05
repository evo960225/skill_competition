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
Dim n As Long, a As Byte, b As Byte, qa As Long, qb As Long, dx As Long, sum As Long
Dim io As Byte, add As Long
For ii = 1 To 3 Step 2
Input #ii, n
For sti = 1 To n
    qa = 0: qb = 0: sum = 0: add = 0
    ReDim srr(20, 20), q(20, 1)
    Do
      Input #ii, a, b
      If a = 0 And b = 0 Then Exit Do
      srr(a, b) = 1: srr(b, a) = 1: add = add + 1
    Loop
   If add <> 0 Then
    Do
      dx = q(qa, 0): qa = qa + 1: io = 0: er = 0
      For i = 0 To 20
      If srr(dx, i) > 0 Then qb = qb + 1: q(qb, 0) = i: srr(i, dx) = -1: io = io + 1
      If srr(dx, i) = -1 Then er = er + 1
      Next
      If io = 0 Then sum = sum + 1
      If er > 1 Then Exit Do
    Loop Until qa > qb
    Print #2, IIf(qa <> add + 1 Or er = 2, "F", CStr(sum))
   Else
    Print #2, "0"
   End If
Next
Print #2,
Next
Close
Unload Me
End Sub
