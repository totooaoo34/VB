VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2388
   ClientLeft      =   84
   ClientTop       =   408
   ClientWidth     =   3672
   LinkTopic       =   "Form1"
   ScaleHeight     =   2388
   ScaleWidth      =   3672
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   1560
   End
   Begin VB.Shape Shape1 
      Height          =   852
      Index           =   3
      Left            =   1440
      Top             =   120
      Width           =   1092
   End
   Begin VB.Shape Shape1 
      Height          =   852
      Index           =   2
      Left            =   2280
      Top             =   720
      Width           =   1092
   End
   Begin VB.Shape Shape1 
      Height          =   852
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   1092
   End
   Begin VB.Shape Shape1 
      Height          =   852
      Index           =   0
      Left            =   840
      Top             =   600
      Width           =   1092
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const hn = 20, wn = 10, boxl = 576 / 2
Private Sub Form_Load()
Timer1.Interval = 1000
Dim x(4) As Integer, y(4) As Integer

RndER x(1), wn
Randomize
RndER y(1), 2
If y(1) = 0 Then movex 2, 1, 1: y(1) = 0 Else movey 2, 1, 1: x(2) = x(1): y(1) = 0
RndER y(3), 2
If y(3) = 0 Then movex 3, 2, 1 Else movey 3, 2, 1: x(3) = x(1)
Select Case x(3)
    Case x(1) + boxl * 2:
        RndER y(4), 2
        If y(4) = 0 Then
        RndER y(4), 2
            If y(4) = 0 Then y(4) = 0: movex 4, 3, 1 Else y(4) = boxl: x(4) = x(3)
        Else
        x(4) = x(2): y(4) = boxl
        End If
    Case x(1) + boxl:
        If x(2) = boxl Then
        RndER y(4), 3
            Select Case y(4)
            Case 0:
            Case 1:
            Case 2:
            End Select
        Else
        RndER y(4), 3
            Select Case y(4)
            Case 0:
            Case 1:
            Case 2:
            End Select
        End If
    Case x(1):
    RndER y(4), 2
    If y(4) = 0 Then movex 4, 3, 1 Else movey 4, 3, 1
End Select
End Sub
Public Sub RndER(a As Integer, b As Integer)
Randomize
a = (Rnd * b)
End Sub
Sub movex(a As Integer, b As Integer, c As Integer)
Dim x(4) As Integer, y(4) As Integer
x(a) = x(b) + c * boxl
End Sub
Sub movey(a As Integer, b As Integer, c As Integer)
Dim x(4) As Integer, y(4) As Integer
y(a) = y(b) + c * boxl
End Sub
Sub xup()
Dim y(4) As Integer
For a = 1 To 4
y(a) = y(a) + boxl
Next a
End Sub
Sub yback()

End Sub

Private Sub Timer1_Timer()
Dim x(4) As Integer, y(4) As Integer
For a = 1 To 4
Shape1(a).Move x(a), y(a)
Next a
End Sub
