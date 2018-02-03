Attribute VB_Name = "Module1"
'Totoo作品，中国智造

Public x(1 To 4), y(1 To 4) As Integer
Public BF(1 To 10, 1 To 20) As Integer, mX(1 To 4), mY(1 To 4) As Integer
Public retY(1 To 20), Toper(1 To 20) As Integer, Saver(1 To 10) As String
Public socre, Iret, MarkNum As Integer, TimeLen As Integer, EntI As Integer
Public SystemAsk, Cant, BottomAsk As Boolean, ret As String
Public Repeat As Boolean, Clean As Boolean


Public Sub MoveLeft()
    'Totoo作品
    On Error GoTo Pass:
    For a = 1 To 4
    mX(a) = x(a) - 1
    If BF(mX(a) + 1, y(a) + 1) = 1 Then GoTo Pass:
    Next a
    For a = 1 To 4
    x(a) = mX(a)
    Next a
Pass:
End Sub

Public Sub MoveRight()
    On Error GoTo Pass:
    For a = 1 To 4
    mX(a) = x(a) + 1
    If BF(mX(a) + 1, y(a) + 1) = 1 Then GoTo Pass:
    Next a
    For a = 1 To 4
    x(a) = mX(a)
    Next a
Pass:
End Sub

Public Sub ClsBox()
For a = 1 To 10
    For b = 1 To 20
    BF(a, b) = 0
    Next b
Next a
End Sub

Public Sub NextBox(a As Integer, b As Integer, c As Integer, d As Integer)
If d = 0 Then
x(a) = x(b): y(a) = y(b) + c
Else
x(a) = x(b) + c: y(a) = y(b)
End If
End Sub

Public Sub Rndget(a, b As Integer)
Randomize
a = Fix(Rnd * b)
End Sub

Public Sub ComeTure()
For a = 1 To 4
x(a) = mX(a): y(a) = mY(a)
Next a
End Sub
