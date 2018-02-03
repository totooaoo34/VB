VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2496
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3744
   LinkTopic       =   "Form1"
   ScaleHeight     =   2496
   ScaleWidth      =   3744
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Dim Bl(1 To 10, 1 To 20) As Integer, c As Variant
Bl(2, 3) = 1

If i = 0 Then GoTo R:
W:
Open App.Path & "\savegame.txt" For Output As #1
For a = 1 To 10
For b = 1 To 20
Print #1, Bl(a, b)
Next b
Next a
Close (1)
ib = ib + 1
MsgBox "w"
R:
On Error GoTo W:
Open App.Path & "\savegame.txt" For Input As #1
For a = 1 To 10
For b = 1 To 20
Line Input #1, c
If Val(c) = 1 Then MsgBox a: MsgBox b
Next b
Next a
Close (1)
MsgBox "r"

MsgBox "done"
End
End Sub

