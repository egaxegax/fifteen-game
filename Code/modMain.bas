Attribute VB_Name = "modMain"
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Public i_Ind As Integer   '������ ����� ��������������� �����

Dim i_LastC As Integer    '������ ����� ������� �����
Dim i_LastF As Integer    '������ ����� ������� ������

Sub Main()
    InitCommonControls
    frmGame.Show
End Sub
'������� ����������� ������������ ����������� ����� �� ����� � ������� i_pC, ������� ��
Function IsPositionPart( _
    i_P As Integer _
  ) As Boolean
Dim i As Integer

  '���� ������������������ �� ������ i_P
    If i_P > 15 Then i_P = 15
    For i = 1 To i_P
        If frmGame.pC(i).Container.Index <> i Then Exit Function
    Next i
    IsPositionPart = True
End Function
'��������� ������/����������� ����
Function Game() As Boolean
    If Not IsPositionPart(16) Then Exit Function
    Beep
    Select Case MsgBox( _
                    "��� ���?", _
                    vbYesNo Or vbInformation, _
                    "�����������!" _
                  )
    Case vbNo
        Unload frmGame
    Case vbYes
        frmGame.mnuNew_Click
    End Select
    Game = True
End Function
'��������� ��������� ����������� �����
Sub Random()

Dim i As Integer, j As Integer, X As Integer
Dim i_Rnd(1 To 15) As Integer
    
  '������� ������ ��������� �����
    For j = 1 To frmGame.P.UBound - 1
        Randomize
Begin:
        i_Rnd(j) = Int((frmGame.P.UBound * Rnd) + 1)
        If j > 1 Then
            For X = 1 To j - 1
                If i_Rnd(X) = i_Rnd(j) Then GoTo Begin:
           Next X
        End If
    Next j
  '���������� ������ �� ��������� ������ �� �������
    For i = 1 To frmGame.pC.UBound
        Set frmGame.pC(i).Container = frmGame.P(i_Rnd(i))
    Next i
  '���� ����� ����������� �� �������
    If IsPositionPart(16) Then Random
End Sub
'������� ����������� ����������� ����
Function IsMove( _
    i_C As Integer, _
    i_P As Integer _
  ) As Boolean
Dim C_Ctl As Object '������ ������� ������������� �����
Dim P_Ctl As Object '������ ������� ��������� ������

    Set C_Ctl = frmGame.pC(i_C).Container
    Set P_Ctl = frmGame.P(i_P)
  On Error Resume Next
    IsMove = ( _
        (C_Ctl.Index = P_Ctl.Index - 1 And C_Ctl.Top = P_Ctl.Top) Or _
        (C_Ctl.Index = P_Ctl.Index + 1 And C_Ctl.Top = P_Ctl.Top) Or _
        (C_Ctl.Index = P_Ctl.Index - 4 And C_Ctl.Left = P_Ctl.Left) Or _
        (C_Ctl.Index = P_Ctl.Index + 4 And C_Ctl.Left = P_Ctl.Left) _
      )
End Function
'������ �����
Sub Master()

Dim i As Integer, j As Integer
Dim i_X As Integer          '������ ����� ��������� ������
Dim i_IsMove() As Integer   '������ ������ �����, ������� ����� ������� �� i_X
Dim i_LastPartC             '������ ����� ��������� ����� � ��������� ������������������ �� �������
  
  '��������� ������ ������� �����
    For i = 1 To frmGame.pC.UBound
        i_X = i_X + frmGame.pC(i).Container.Index
    Next i
  '136 - ����� ���������� ������� ���� ����� 1+2+..+16
  '������� �� 136 ����� ������� ����� �������� ����� ������ ������
    i_X = 136 - i_X
  '������ ����� - ������ ���� ��������� ����� (�����, ������, ������, �����)
    ReDim i_IsMove(3)
    For i = 1 To frmGame.pC.UBound
      '���������� � ������ ����� ������ ��������� � ���� �����
        If IsMove(i, i_X) Then
            If IsPositionPart(frmGame.pC(i).Container.Index) Then i_LastPartC = i Else i_IsMove(j) = i: j = j + 1
        End If
    Next i
    If j = 0 Then Exit Sub          '���� ������ ����� ���� - ����� ����, ����� �� ���������
    ReDim Preserve i_IsMove(j - 1)  '������ ����������� ������� ����� �� ����� ��������� �����
    Randomize
More:
  '���� ������� 1 ��������� ���, ��  ������ ��������� � ���� ��������� ����� � ��������� ������������������
    If UBound(i_IsMove) = 0 Then
        ReDim Preserve i_IsMove(UBound(i_IsMove) + 1)
        i_IsMove(UBound(i_IsMove)) = i_LastPartC
    End If
    
  '�������� ��������� �����
    i = Int((UBound(i_IsMove) + 1) * Rnd + 1) - 1
    
  '���� �����, ��������������� ���e�� ��������� ������ � ,����
  '�� ��������� ������, �� ������� �� ����� ��������� ����� ����
  '������������������, �� ������� �����, ����� ���
    For j = 0 To UBound(i_IsMove)
        If i_IsMove(j) = i_X Then
            If IsPositionPart(i_X - 1) Then i = j: Exit For
        End If
    Next j
  '���� ����� ����� ���, �� ����� ����� � �������, ������� ������ ������ �
  '� ������� �� ������ � ������� ������� (�������� �����)
  '���� ��� ������������������ �� 6 ����� (����� �� ��������� �����)
    If Not IsPositionPart(6) Then
MoveUp:
        If j - 1 = UBound(i_IsMove) Then
            For j = 0 To UBound(i_IsMove)
                If i_IsMove(j) < i_X And frmGame.pC(i_IsMove(j)).Container.Index > i_X Then i = j: Exit For
            Next j
        End If
    Else
  '���� ���� ������������������ �� 6 ����� ����� ��������, ���� �����
        If Int(10 * Rnd + 1) Mod 2 = 0 Then GoTo MoveUp
    End If
  '���� ����� ����� � ����� ������ ����� ��, ��� �� ���������� ����, ��
  '����� ��������� �����
    If i_LastC > 0 Then
        If i_LastC = i_IsMove(i) And i_LastF = i_X Then
            j = i           '���������� �����
            Do While j = i  '�������� ���� �� ������ ������ ����� �����
                i = Int((UBound(i_IsMove) + 1) * Rnd + 1) - 1
            Loop
         End If
    End If
  '���������� ����� ������� �����, ����� �� ��������� ��� �� ��������� ����
    i_LastC = i_IsMove(i)
    i_LastF = frmGame.pC(i_IsMove(i)).Container.Index  '���������� ����� ������� ������. ����� �� �����������
    Set frmGame.pC(i_IsMove(i)).Container = frmGame.P(i_X) '������������ ��������� �����
End Sub
'��������� ��������� ��� ���� �����
Sub PaintPossible( _
    Optional bPaint As Boolean _
  )
Dim i As Integer
    
    With frmGame.pC
        For i = 1 To .UBound
            .Item(i).BackColor = CLng(&H4000&)
            If bPaint Then
                'If .Item(i).Index = .Item(i).Container.Index Then .Item(i).BackColor = vbWhite
                If IsMove(i, GetFreeCell()) Then .Item(i).BackColor = vbYellow
            End If
        Next
    End With
    PaintNumbers
End Sub
'���������� ����� ��������� ������
Function GetFreeCell() As Integer
  
Dim i As Integer
  
    For i = 1 To frmGame.pC.UBound
        GetFreeCell = GetFreeCell + frmGame.pC(i).Container.Index
    Next
  '1+...+16 - ������ ��������� �������� P
    GetFreeCell = 136 - GetFreeCell
End Function
'�����������
Sub MoveIt( _
    i_Free As Integer _
  )
    If IsMove(i_Ind, i_Free) Then
        Set frmGame.pC(i_Ind).Container = frmGame.P(i_Free) '������������
        If frmGame.mnuPaint.Checked Then PaintPossible frmGame.mnuPaint.Checked
        Game                                               '��������� �� ������������
    End If
    i_Ind = 0       '���������� ����� ��������������� �����
End Sub
'������ ������
Sub PaintNumbers()

Dim i As Integer

    With frmGame
        For i = 1 To .pC.UBound
            .pC(i).FontName = "Trebuchet MS"
            .pC(i).FontSize = 14
            .pC(i).FontBold = True
            .pC(i).CurrentX = .pC(i).ScaleWidth / IIf(i < 10, 4, 6)
            .pC(i).CurrentY = .pC(i).ScaleHeight / 4
            .pC(i).Print i
        Next
    End With
End Sub
