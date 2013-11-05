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
Dim st As String, arr() As String
Dim dx As Long, MaxIn As Long
For ii = 1 To 3 Step 2
dx = 0
ReDim starr(200), stsum(200)
    Do
        Input #ii, st
        If st = "EOF" Then Exit Do
        st = LCase(st)
        arr = Split(st, " ")
        For i = 0 To UBound(arr)
          For j = 0 To dx - 1
          If starr(j) = arr(i) Then stsum(j) = stsum(j) + 1: Exit For
          Next
          If j = dx Then
            starr(dx) = arr(i): stsum(dx) = 1
            dx = dx + 1
          End If
        Next
    Loop
    For j = 1 To 3
      For i = 0 To dx - 1
        If stsum(i) > stsum(MaxIn) Then MaxIn = i
      Next
      Print #2, CStr(stsum(MaxIn));
      If j <> 3 Then Print #2, ","; Else Print #2, vbCrLf
      stsum(MaxIn) = 0
    Next
Next
Close
Unload Me
End Sub
