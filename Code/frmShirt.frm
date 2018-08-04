VERSION 5.00
Begin VB.Form frmShirt 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4620
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   3552
   Icon            =   "frmShirt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3552
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk 
      Caption         =   "60"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   3600
      Width           =   1695
   End
   Begin VB.CommandButton cmd 
      Caption         =   "50"
      Default         =   -1  'True
      Height          =   420
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "51"
      Height          =   429
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   14
      Left            =   1920
      Picture         =   "frmShirt.frx":000C
      Top             =   2760
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   13
      Left            =   1080
      Picture         =   "frmShirt.frx":0ED6
      Top             =   2760
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   12
      Left            =   240
      Picture         =   "frmShirt.frx":1DA0
      Top             =   2760
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   11
      Left            =   2760
      Picture         =   "frmShirt.frx":2C6A
      Top             =   1920
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   10
      Left            =   1920
      Picture         =   "frmShirt.frx":3B34
      Top             =   1920
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   9
      Left            =   1080
      Picture         =   "frmShirt.frx":49FE
      Top             =   1920
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   8
      Left            =   240
      Picture         =   "frmShirt.frx":58C8
      Top             =   1920
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   7
      Left            =   2760
      Picture         =   "frmShirt.frx":6792
      Top             =   1080
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   6
      Left            =   1920
      Picture         =   "frmShirt.frx":765C
      Top             =   1080
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   5
      Left            =   1080
      Picture         =   "frmShirt.frx":8526
      Top             =   1080
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   4
      Left            =   240
      Picture         =   "frmShirt.frx":93F0
      Top             =   1080
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   3
      Left            =   2760
      Picture         =   "frmShirt.frx":A2BA
      Top             =   240
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   2
      Left            =   1920
      Picture         =   "frmShirt.frx":B184
      Top             =   240
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      Height          =   576
      Index           =   1
      Left            =   1080
      Picture         =   "frmShirt.frx":C04E
      Top             =   240
      Width           =   576
   End
   Begin VB.Image img 
      Appearance      =   0  'Плоска
      BorderStyle     =   1  'Фиксировано один
      Height          =   600
      Index           =   0
      Left            =   240
      Picture         =   "frmShirt.frx":CF18
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "frmShirt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i_Ico As Integer, bAssort As Boolean

Private Sub chk_Click()
    bAssort = (chk.Value = 1)
End Sub

Private Sub cmd_Click( _
    Index As Integer _
  )
    Select Case Index
    Case 0
        With frmGame
            For Index = .pC.LBound To .pC.UBound
                Set .pC(Index).Picture = img(IIf(bAssort, Index - 1, i_Ico)).Picture
            Next
            .PaintNumbers
        End With
        Unload Me
    Case 1
        Unload Me
    End Select
End Sub

Private Sub Form_Load()

Dim i As Integer
  
    For i = cmd.LBound To cmd.UBound
        cmd(i).Caption = LoadResString(CLng(cmd(i).Caption))
    Next
    chk.Caption = LoadResString(CLng(chk.Caption))
    chk.Value = Abs(bAssort)
    img_Click i_Ico
End Sub

Private Sub img_Click( _
    Index As Integer _
  )
Dim i As Integer

    For i = img.LBound To img.UBound
        img(i).BorderStyle = 0
    Next
    img(Index).BorderStyle = 1
    i_Ico = Index
End Sub
