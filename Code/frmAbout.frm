VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1848
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   4572
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1848
   ScaleWidth      =   4572
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "50"
      Default         =   -1  'True
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Плоска
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'Нет
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   384
      Left            =   120
      Picture         =   "frmAbout.frx":000C
      ScaleHeight     =   269.696
      ScaleMode       =   0  'Пользовательское
      ScaleWidth      =   269.696
      TabIndex        =   1
      Top             =   120
      Width           =   384
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   792
      Left            =   720
      TabIndex        =   3
      Top             =   240
      Width           =   3492
   End
   Begin VB.Label lblAuthor 
      Caption         =   "42"
      Height          =   432
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   2388
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Caption = LoadResString(40) & App.ProductName
    lblAuthor.Caption = LoadResString(CLng(lblAuthor.Caption))
    cmdOK.Caption = LoadResString(CLng(cmdOK.Caption))
    lblAuthor.Caption = lblAuthor.Caption & App.CompanyName
    lblDescription.Caption = App.FileDescription
    Set picIcon.Picture = frmGame.Icon
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub
