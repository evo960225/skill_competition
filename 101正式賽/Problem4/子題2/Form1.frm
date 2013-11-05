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
Dim n As Long, st As String, arr() As String, k As Long, ch As Byte, sum As Byte
For ii = 1 To 3 Step 2
    Input #ii, n
    For sti = 1 To n
        Input #ii, st
        k = 0: sum = 0
        arr = Split(st)
        For i = 0 To UBound(arr)
            ch = Val(arr(i)) Mod 13
            If ch > 10 Or ch = 0 Then ch = 10
            If ch = 1 Then k = 1
            sum = sum + ch
        Next
        If k = 1 And sum + 10 < 22 Then sum = sum + 10
        Print #2, IIf(sum < 22, CStr(sum), "F")
    Next
    Print #2,
Next
Close
Unload Me
End Sub
