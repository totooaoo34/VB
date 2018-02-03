VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9744
   ClientLeft      =   1320
   ClientTop       =   1752
   ClientWidth     =   5520
   LinkTopic       =   "MDIForm1"
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VB语言版俄罗斯方块
'Totoo、Aoo34智造，一些方块，很多计算

Const WN As Integer = 10, HN As Integer = 20
Const Boxl As Integer = 372, BoxNum As Integer = 200


Private Sub Timer1_Timer()
Timer1.Interval = TimeLen
CheckTop
Down
Cleaner
XFull
End Sub

Private Sub Form_Load()
    Call Load
    Call givebox
    For a = 0 To 3
    Label1(a).Caption = "开发者研究重地！点击使用键盘w\a\s\d\控制"
    Next a
With Label2
.Move 0, 20 * Boxl
.Caption = "使用滑竿调整速度！                              点击使用键盘控制"
End With
Form1.Caption = "w,a,s,d分别为变形、左、右及降落"
    TimeLen = 500
Timer1.Interval = 1000
Call ClearUpEr
ShapeAdd
    For a = 0 To 3
With Shape2(a)
.Width = Boxl
.Height = Boxl
End With
    Next a
End Sub
 
Private Sub ClearUpEr()
'Totoo作品
With Form1
.Width = WN * 372 / 2 * 3
.Height = 27 * Boxl
End With
    Dim Ia As Integer, ib As Integer
    Dim x(BoxNum) As Integer, y(BoxNum) As Integer
    x(1) = 0
    y(1) = 0
        For a = 0 To 199
With Shape1(a)
.Width = Boxl * (Iret + 1)
.Height = Boxl * (Iret + 1)
End With
    Ia = Ia + 1
        If (Ia <> 0) And (a Mod WN = 0) Then Ia = 0: ib = ib + 1
    x(a) = Boxl * Ia
    y(a) = Boxl * (ib - 1)
    Shape1(a).Move x(a), y(a)
        Next a
'Totoo作品
End Sub

Sub ShapeAdd()
'Totoo作品
    Dim ret As Integer
    RndGet ret, 6
    Select Case ret
    Case 0
    x(1) = 0: x(2) = x(1): x(3) = 1: x(4) = x(3)
    y(1) = 0: y(2) = 1: y(3) = y(1): y(4) = y(2)
    MarkNum = 1 'Totoo作品，中国智造
    'MsgBox "0"
    Case 1
    x(1) = 0: x(2) = 1: x(3) = 2: x(4) = 3
    y(1) = 0: y(2) = y(1): y(3) = y(2): y(4) = y(3)
    'MsgBox "1"
    MarkNum = 2
    Case 2
    x(1) = 0: x(2) = 1: x(3) = 1: x(4) = 2
    y(1) = 0: y(2) = 0: y(3) = 1: y(4) = 1
    'MsgBox "2"
    MarkNum = 3
    Case 3
    x(1) = 1: x(2) = 0: x(3) = 1: x(4) = 1
    y(1) = 0: y(2) = 1: y(3) = 1: y(4) = 2
    'MsgBox "3,1"
    MarkNum = 4
    Case 4
    x(1) = 1: x(2) = x(1): x(3) = x(1): x(4) = 0
    y(1) = 0: y(2) = 1: y(3) = 2: y(4) = y(3)
    'MsgBox "4"
    MarkNum = 5
    Case 5
    x(1) = 1: x(2) = 1: x(3) = 0: x(4) = 0
    y(1) = 0: y(2) = 1: y(3) = 1: y(4) = 2
    'MsgBox "22"
    MarkNum = 6
    Case 6
    x(1) = 0: x(2) = 0: x(3) = 0: x(4) = 1
    y(1) = 0: y(2) = 1: y(3) = 2: y(4) = 2
    'MsgBox "6"
    MarkNum = 7
    End Select
        For a = 1 To 4
With Shape2(a - 1)
.Move x(a) * Boxl, y(a) * Boxl
.Width = Boxl
.Height = Boxl
End With
        Next a
    Dim reta3, reta4 As Integer
        For a = 1 To 4
    reta3 = x(a)
        If reta3 > reta4 Then: reta4 = reta3
        Next a
    Randomize
    reta3 = Fix(Rnd * (9 - reta4)) + 1
        For a = 1 To 4
    x(a) = x(a) + reta3
        Next a
'Totoo作品
End Sub

Sub Cleaner()
'Totoo作品，中国智造
    For a = 1 To 10
        For b = 1 To 20
            If BF(a, b) = 1 Then
Shape1(a + (b - 1) * 10 - 1).FillStyle = 0
            Else
Shape1(a + (b - 1) * 10 - 1).FillStyle = 1
            End If
        Next b
    Next a

End Sub


Sub CheckTop()
    'Totoo作品，中国智造
On Error GoTo done:
        For a = 1 To 4
    If x(a) + 1 < 19 Then On Error Resume Next
    If y(a) > 18 Then GoTo done:
    If BF(x(a) + 1, y(a) + 2) = 1 Then GoTo done:

On Error GoTo Over:
    If x(a) + 1 > 20 Or x(a) + 1 < 1 Then GoTo Over:
        Next a
    If 1 = 2 Then
Over:
    Call ClsBox
        'Timelen = 500
        Call ShapeAdd
        'MsgBox "GameOver!": End
    End If
    If 1 = 2 Then
done:
        For a = 1 To 4
            If BF(x(a) + 1, y(a) + 1) = 1 Then GoTo Over:
        Next a
        For a = 1 To 4
    BF(x(a) + 1, y(a) + 1) = 1
        Next a
    Call ShapeAdd: If BottomAsk = True Then TimeLen = 500: BottomAsk = False
    End If
Pass:
End Sub

Private Sub Turn()
    If MarkNum <> 1 Then
    Dim castX(1 To 4), castY(1 To 4) As Integer
        For a = 1 To 4
    mX(a) = x(a) - x(3): mY(a) = y(a) - y(3)
Label1(a - 1).Caption = Str(mX(a)) & Str(mY(a))
    If mY(a) = 0 Or mX(a) = 0 Then
        If mY(a) <> 0 Then castX(a) = x(3) - mY(a): castY(a) = y(3)
        If mX(a) <> 0 Then castX(a) = x(3): castY(a) = x(3) - mX(a)
    Else
        castX(a) = x(3) - mY(a): castY(a) = y(3) + mX(a)
    End If
    On Error GoTo Pass:
            If BF(castX(a) + 1, castY(a) + 1) = 1 Then GoTo Pass:
        Next a
        
        For a = 1 To 4
    x(a) = castX(a): y(a) = castY(a)
        'Shape2(a - 1).Move x(a) * Boxl, y(a) * Boxl
        Next a
Pass:
    End If 'Totoo作品，中国智造
End Sub

Sub XFull() 'Totoo作品，中国智造
    Dim Ia As Integer, I As Integer
    Dim mY As Integer, BfRet(1 To 10, 1 To 20) As Integer
    Dim Cleanit As Boolean
        For b = 1 To 20
            For a = 1 To 10
                If BF(a, b) = 1 Then Ia = Ia + 1
            Next a
                If Ia = 10 Then I = I + 1: Toper(I) = b:  '记录满格
    Ia = 0
        Next b
    If I <> 0 Then
        For b = 1 To I
            For a = 1 To 10
        BF(a, Toper(b)) = 0
            Next a
socre = socre + 200
            Next b
Label2.Caption = "得分：" & Str(socre)
    End If
    If (Clean = True) Then
        For a = 1 To 10
    Cleanit = False
            For b = 1 To 20
        mY = 0
        mY = BF(a, b)
        If BF(a, b) = 1 Then
                For c = 1 To I
            If Toper(c) <> 0 Then
                If b < Toper(c) Then
                mY = mY + 1
                Cleanit = True
                End If
            End If
            If c = I Then
                If b + mY > 20 Then GoTo Pass:
            BfRet(a, b + mY - 1) = 1
                If 1 = 2 Then
Pass:
                For d = 1 To 10
                BfRet(a, 20) = 1
                Next d
                End If
        End If
    Next c
    End If
    mY = 0
    Next b
    If Cleanit = True Then
    For b = 1 To 20
    BF(a, b) = BfRet(a, b)
    BfRet(a, b) = 0
    Next b
    End If
Next a
End If
    For L = 1 To I
    Toper(L) = 0
    Next L
End Sub

Public Function RndGet(a As Integer, b As Integer) As Integer
    Randomize
    a = Fix(Rnd * b)
End Function

Private Sub Save()
    Dim SFN As String
    CommonDialog1.ShowOpen
    SFN = CommonDialog1.FileName
    If SFN <> "" Then
    Open SFN & ".totooDat" For Output As #1
    For a = 1 To 10
    For b = 1 To 20
    Print #1, BF(a, b)
    Next b, a
    Print socre
    Close #1
    End If
End Sub

Private Sub Down()
    Clean = True
        For a = 1 To 4
    y(a) = y(a) + 1
Shape2(a - 1).Move x(a) * Boxl, y(a) * Boxl
        Next a
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 65, 37: mLeft   '左
    Case 68, 39: mRight  '右
    Case 87, 38: Turn
    Case 83, 40: TimeLen = 20: BottomAsk = True
    End Select
    If KeyCode = 13 Then
    EntI = EntI + 1
        If EntI Mod 2 = 1 Then
            TimeLen = 10
        Else
            TimeLen = 1000
        End If
    End If
End Sub

Private Sub Combo1_Change()
Dim ret As Integer
For c = 1 To 10
If Combo1.Text = Saver(c) Then
On Error GoTo Bad:
Open App.Path & "\" & Saver(c) For Input As #4
For a = 1 To 10
For b = 1 To 20
Input #4, ret
BF(a, b) = ret
Next b
Next a
Close (4)
End If
Next c
Bad:
Combo1.Text = "无效的存档！"
End Sub

Private Sub Combo1_DropDown()
Save
End Sub
