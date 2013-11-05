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
Dim trr(26) As Byte, k As Byte, n As Byte, st As String, dx As Long, tmp As String
Dim sum As Integer, a As Byte, b As Byte, c As Byte
For i = 1 To 25
    If i = 9 Or i = 15 Or i = 23 Then k = k + 1
    trr(i) = i + 9 - k
Next
trr(9) = 34: trr(15) = 35: trr(23) = 32: trr(26) = 33
For ii = 1 To 3 Step 2
    Input #ii, n
    ReDim srr(10)
    dx = 0: a = 0: b = 0: c = 0
    For i = 1 To n
      Input #ii, tmp
      For j = 1 To dx
        If srr(j) = tmp Then Exit For
      Next
      If j > dx Then
        st = trr(Asc(Left(tmp, 1)) - 64) & Mid(tmp, 2)
        sum = Val(Left(st, 1)) + Val(Right(st, 1))
        For j = 2 To 10
          sum = sum + Val(Mid(st, j, 1)) * (11 - j)
        Next
        If sum Mod 10 = 0 And Mid(st, 3, 1) < "3" And Left(tmp, 1) >= "A" And Left(tmp, 1) <= "Z" Then
            a = a + 1
            dx = dx + 1: srr(dx) = tmp
        Else
            c = c + 1
        End If
      Else
        b = b + 1
      End If
    Next
    Print #2, a & "," & b & "," & c & vbCr
Next
Close
Unload Me
End Sub
