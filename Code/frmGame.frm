VERSION 5.00
Begin VB.Form frmGame 
   BackColor       =   &H00004000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   5256
   ClientLeft      =   48
   ClientTop       =   732
   ClientWidth     =   3708
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5256
   ScaleWidth      =   3708
   StartUpPosition =   1  'CenterOwner
   Tag             =   "                         ��������"
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Index           =   0
      Left            =   3120
      Top             =   4200
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Index           =   1
      Left            =   3120
      Top             =   4680
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Index           =   2
      Left            =   3120
      Top             =   3720
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00004000&
      BorderStyle     =   0  '���
      Height          =   3615
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      Begin VB.PictureBox P 
         Appearance      =   0  '������
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   1
         Left            =   120
         ScaleHeight     =   60
         ScaleMode       =   3  '�������
         ScaleWidth      =   60
         TabIndex        =   16
         Top             =   240
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   1
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   30
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   15
         Left            =   1800
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   15
         Top             =   2760
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   15
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   17
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   14
         Left            =   960
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   14
         Top             =   2760
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   14
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   29
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   13
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   13
         Top             =   2760
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   13
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   28
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   12
         Left            =   2640
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   12
         Top             =   1920
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   12
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   27
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   11
         Left            =   1800
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   11
         Top             =   1920
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   11
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   26
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         Appearance      =   0  '������
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         ForeColor       =   &H80000008&
         Height          =   720
         Index           =   10
         Left            =   960
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   10
         Top             =   1920
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   10
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   25
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   9
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   9
         Top             =   1920
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   9
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   24
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   8
         Left            =   2640
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   8
         Top             =   1080
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   8
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   23
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   7
         Left            =   1800
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   7
         Top             =   1080
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   7
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   22
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   6
         Left            =   960
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   6
         Top             =   1080
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   6
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   21
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   5
         Left            =   120
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   5
         Top             =   1080
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   5
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   20
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   4
         Left            =   2640
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   4
         Top             =   240
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   720
            Index           =   4
            Left            =   0
            ScaleHeight     =   720
            ScaleWidth      =   720
            TabIndex        =   19
            Top             =   0
            Width           =   720
         End
      End
      Begin VB.PictureBox P 
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   3
         Left            =   1800
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   3
         Top             =   240
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   760
            Index           =   3
            Left            =   0
            ScaleHeight     =   756
            ScaleWidth      =   756
            TabIndex        =   18
            Top             =   0
            Width           =   760
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   2
         Left            =   960
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   2
         Top             =   240
         Width           =   720
         Begin VB.PictureBox pC 
            Appearance      =   0  '������
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00004000&
            BorderStyle     =   0  '���
            DragMode        =   1  '����
            ForeColor       =   &H80000008&
            Height          =   760
            Index           =   2
            Left            =   0
            ScaleHeight     =   756
            ScaleWidth      =   756
            TabIndex        =   35
            Top             =   0
            Width           =   760
         End
      End
      Begin VB.PictureBox P 
         BackColor       =   &H00004000&
         BorderStyle     =   0  '���
         Height          =   720
         Index           =   16
         Left            =   2640
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   1
         Top             =   2760
         Width           =   720
      End
   End
   Begin VB.Label lblHodCap 
      BackColor       =   &H00004000&
      Caption         =   "���:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1200
      TabIndex        =   34
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblTimeCap 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      Caption         =   "�����:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   33
      Top             =   4680
      Width           =   540
   End
   Begin VB.Label lblHod 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   32
      Top             =   4920
      Width           =   405
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   31
      Top             =   4680
      Width           =   405
   End
   Begin VB.Menu mnuMain 
      Caption         =   "0"
      Index           =   0
      Begin VB.Menu mnuGame 
         Caption         =   "1"
         Index           =   0
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuGame 
         Caption         =   "2"
         Index           =   2
         Begin VB.Menu mnuDrag 
            Caption         =   "3"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuDrag 
            Caption         =   "4"
            Index           =   1
         End
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuGame 
         Caption         =   "5"
         Index           =   4
      End
      Begin VB.Menu mnuGame 
         Caption         =   "6"
         Index           =   5
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuGame 
         Caption         =   "7"
         Enabled         =   0   'False
         Index           =   7
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuGame 
         Caption         =   "8"
         Index           =   9
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "9"
      Index           =   1
      Begin VB.Menu mnuF1 
         Caption         =   "10"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuF1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuF1 
         Caption         =   "11"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Dim b_Init As Boolean
Dim i_Ind As Integer      '������ ����� ��������������� �����
Dim i_LastC As Integer    '������ ����� ������� �����
Dim i_LastF As Integer    '������ ����� ������� ������

Private Sub mnuGame_Click( _
    Index As Integer _
  )
    With mnuGame(Index)
        Select Case Index
        Case 0  '����� ����
            If Not b_Init Then
                Frame(0).Visible = True
                mnuGame(7).Enabled = True
              '������ ��-���������
                Load frmShirt
                For Index = pC.LBound To pC.UBound
                    Set pC(Index).Picture = frmShirt.img(0).Picture
                Next
                Unload frmShirt
                PaintNumbers
                b_Init = True
            End If
            If mnuGame(7).Checked Then mnuGame_Click 7
            Random                      '�������� ����������� �����
            If mnuGame(4).Checked Then PaintPossible mnuGame(4).Checked
        Case 4  '���������
            .Checked = Not .Checked
            PaintPossible .Checked
        Case 5  '�������
            frmShirt.Show vbModal, Me
        Case 7  '������ �����
            .Checked = (.Checked = False)
            If Not .Checked Then
              '������������� ������� ���� (��������)
                lblTime.Caption = ""
                lblHod.Caption = ""
                tmr(2).Enabled = False
            Else
              '��������� ������� ���� (��������)
                tmr(2).Interval = 100
                lblTime.Tag = Time
                tmr(2).Enabled = True   '��������� ������� ����
            End If
        Case 9  '�����
            Unload Me '�������� �����
        End Select
    End With
End Sub

Private Sub mnuDrag_Click( _
    Index As Integer _
  )
Dim i As Integer
    
    If (Not mnuDrag(Index).Checked) Then
        mnuDrag(Index).Checked = True
        mnuDrag(IIf(Index = 0, 1, 0)).Checked = False
    End If
    For i = 1 To pC.UBound
        pC(i).DragMode = Index
    Next
End Sub

Private Sub mnuF1_Click( _
    Index As Integer _
  )
    Select Case Index
    Case 0
        If (Dir$("help.txt") = "") Then
            MsgBox LoadResString(30), , App.Title
        Else
            Shell LoadResString(31), vbNormalFocus
         End If
    Case 2
        frmAbout.Show vbModal, Me                   '������ ���� � ���������
    End Select
End Sub

Private Sub pC_Click( _
    Index As Integer _
  )
  '���������� ��������� ������ � ������� �����
    If i_Ind = 0 Then i_Ind = Index
    MoveIt GetFreeCell()
End Sub

Private Sub Form_Load()
    Dim i As Integer
    InitCommonControls
    GetMenuCaption mnuMain
    GetMenuCaption mnuGame
    GetMenuCaption mnuDrag
    GetMenuCaption mnuF1
    tmr(0).Interval = 100
    tmr(1).Interval = 10000
    tmr(1).Enabled = True
    For i = 1 To pC.UBound
        pC(i).DragMode = 0
    Next
    mnuGame_Click 0
End Sub

Private Sub pC_DragDrop( _
    Index As Integer, _
    Source As Control, _
    X As Single, _
    Y As Single _
  )
    i_Ind = 0       '���������� ����� ��������������� �����
End Sub

Private Sub pC_DragOver( _
    Index As Integer, _
    Source As Control, _
    X As Single, _
    Y As Single, _
    State As Integer _
  )
    If i_Ind = 0 Then i_Ind = Index
End Sub

Private Sub P_DragDrop( _
    Index As Integer, _
    Source As Control, _
    X As Single, _
    Y As Single _
  )
    MoveIt Index
End Sub

'������� ����������� ������������ ����������� ����� �� ����� � ������� i_pC, ������� ��
Function IsPositionPart( _
    i_P As Integer _
  ) As Boolean
Dim i As Integer

  '���� ������������������ �� ������ i_P
    If i_P > 15 Then i_P = 15
    For i = 1 To i_P
        If pC(i).Container.Index <> i Then Exit Function
    Next i
    IsPositionPart = True
End Function

'��������� ������/����������� ����
Function Game() As Boolean
    If Not IsPositionPart(16) Then Exit Function
    Beep
    Select Case MsgBox( _
                    LoadResString(20), _
                    vbYesNo Or vbInformation, _
                    LoadResString(21) _
                  )
    Case vbNo
        Unload frmGame
    Case vbYes
        mnuGame_Click 0
    End Select
    Game = True
End Function

'��������� ��������� ����������� �����
Sub Random()

Dim i As Integer, j As Integer, X As Integer
Dim i_Rnd(1 To 15) As Integer
    
  '������� ������ ��������� �����
    For j = 1 To P.UBound - 1
        Randomize
Begin:
        i_Rnd(j) = Int((P.UBound * Rnd) + 1)
        If (j > 1) Then
            For X = 1 To j - 1
                If i_Rnd(X) = i_Rnd(j) Then GoTo Begin:
           Next X
        End If
    Next j
  '���������� ������ �� ��������� ������ �� �������
    For i = 1 To pC.UBound
        Set pC(i).Container = P(i_Rnd(i))
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

    Set C_Ctl = pC(i_C).Container
    Set P_Ctl = P(i_P)
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
    For i = 1 To pC.UBound
        i_X = i_X + pC(i).Container.Index
    Next i
  '136 - ����� ���������� ������� ���� ����� 1+2+..+16
  '������� �� 136 ����� ������� ����� �������� ����� ������ ������
    i_X = 136 - i_X
  '������ ����� - ������ ���� ��������� ����� (�����, ������, ������, �����)
    ReDim i_IsMove(3)
    For i = 1 To pC.UBound
      '���������� � ������ ����� ������ ��������� � ���� �����
        If IsMove(i, i_X) Then
            If IsPositionPart(pC(i).Container.Index) Then i_LastPartC = i Else i_IsMove(j) = i: j = j + 1
        End If
    Next i
    If j = 0 Then Exit Sub          '���� ������ ����� ���� - ����� ����, ����� �� ���������
    ReDim Preserve i_IsMove(j - 1)  '������ ����������� ������� ����� �� ����� ��������� �����
    Randomize
More:
  '���� ������� 1 ��������� ���, ��  ������ ��������� � ���� ��������� ����� � ��������� ������������������
    If (UBound(i_IsMove) = 0) Then
        ReDim Preserve i_IsMove(UBound(i_IsMove) + 1)
        i_IsMove(UBound(i_IsMove)) = i_LastPartC
    End If
    
  '�������� ��������� �����
    i = Int((UBound(i_IsMove) + 1) * Rnd + 1) - 1
    
  '���� �����, ��������������� ���e�� ��������� ������ � ,����
  '�� ��������� ������, �� ������� �� ����� ��������� ����� ����
  '������������������, �� ������� �����, ����� ���
    For j = 0 To UBound(i_IsMove)
        If (i_IsMove(j) = i_X) Then
            If IsPositionPart(i_X - 1) Then i = j: Exit For
        End If
    Next j
  '���� ����� ����� ���, �� ����� ����� � �������, ������� ������ ������ �
  '� ������� �� ������ � ������� ������� (�������� �����)
  '���� ��� ������������������ �� 6 ����� (����� �� ��������� �����)
    If (Not IsPositionPart(6)) Then
MoveUp:
        If j - 1 = UBound(i_IsMove) Then
            For j = 0 To UBound(i_IsMove)
                If i_IsMove(j) < i_X And pC(i_IsMove(j)).Container.Index > i_X Then i = j: Exit For
            Next j
        End If
    Else
  '���� ���� ������������������ �� 6 ����� ����� ��������, ���� �����
        If Int(10 * Rnd + 1) Mod 2 = 0 Then GoTo MoveUp
    End If
  '���� ����� ����� � ����� ������ ����� ��, ��� �� ���������� ����, ��
  '����� ��������� �����
    If (i_LastC > 0) Then
        If i_LastC = i_IsMove(i) And i_LastF = i_X Then
            j = i           '���������� �����
            Do While j = i  '�������� ���� �� ������ ������ ����� �����
                i = Int((UBound(i_IsMove) + 1) * Rnd + 1) - 1
            Loop
         End If
    End If
  '���������� ����� ������� �����, ����� �� ��������� ��� �� ��������� ����
    i_LastC = i_IsMove(i)
    i_LastF = pC(i_IsMove(i)).Container.Index '���������� ����� ������� ������. ����� �� �����������
    Set pC(i_IsMove(i)).Container = P(i_X)    '������������ ��������� �����
End Sub

'��������� ��������� ��� ���� �����
Sub PaintPossible( _
    Optional bPaint As Boolean _
  )
Dim i As Integer
    
    With pC
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
  
    For i = 1 To pC.UBound
        GetFreeCell = GetFreeCell + pC(i).Container.Index
    Next
  '������ ��������� �������� P
    GetFreeCell = 136 - GetFreeCell
End Function

'�����������
Sub MoveIt( _
    i_Free As Integer _
  )
    If IsMove(i_Ind, i_Free) Then
        Set pC(i_Ind).Container = P(i_Free) '������������
        If mnuGame(4).Checked Then PaintPossible mnuGame(4).Checked
        Game                                                '��������� �� ������������
    End If
    i_Ind = 0       '���������� ����� ��������������� �����
End Sub

'������ ������
Sub PaintNumbers()

Dim i As Integer

    With frmGame
        For i = .pC.LBound To .pC.UBound
            .pC(i).FontName = "Trebuchet MS"
            .pC(i).FontSize = 14
            .pC(i).FontBold = True
            .pC(i).CurrentX = .pC(i).ScaleWidth / IIf(i < 10, 4, 8)
            .pC(i).CurrentY = .pC(i).ScaleHeight / 4
            .pC(i).Print i
        Next
    End With
End Sub

Sub GetMenuCaption( _
    m As Object _
  )
Dim i As Integer

    For i = m.LBound To m.UBound
        With m(i)
            If .Caption <> "-" Then
                .Caption = LoadResString(CLng(.Caption))
            End If
        End With
    Next
End Sub

Private Sub tmr_Timer( _
    Index As Integer _
  )
    Select Case Index
    Case 0
        Caption = Right$(Caption, Len(Caption) - 1) '������� ��������� ����� ����
        If Len(Caption) = 0 Then Caption = Tag
        If Len(Caption) = 8 Then tmr(0).Enabled = False
    Case 1
        tmr(0).Enabled = True                       '������������ tmr(0)
    Case 2
        lblTime = CDate(Time - CDate(lblTime.Tag))  '���������� ������ ��������
        lblHod.Caption = Val(lblHod.Caption) + 1    '����������  ���������� ����� � ��������
        Master                                      '��� ������� ����
        If mnuGame(4).Checked Then PaintPossible mnuGame(4).Checked
        Game
    End Select
End Sub
