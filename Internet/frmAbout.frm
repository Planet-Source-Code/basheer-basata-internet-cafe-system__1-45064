VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2445
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4785
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1687.583
   ScaleMode       =   0  'User
   ScaleWidth      =   4493.362
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   45
      Picture         =   "frmAbout.frx":0442
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   45
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000018&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2025
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   -253.544
      X2              =   4971.339
      Y1              =   911.088
      Y2              =   911.088
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Caption         =   "This program is written to cafe internet to calculate the cost of the using internet "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   495
      TabIndex        =   2
      Top             =   720
      Width           =   3840
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Caption         =   "Internet Cafe Program 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   675
      TabIndex        =   4
      Top             =   135
      Width           =   4065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   -281.715
      X2              =   4929.082
      Y1              =   921.441
      Y2              =   921.441
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      BackColor       =   &H80000016&
      Caption         =   "All Right Reserved © 2002 Basheer Basata  basheer_b@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   495
      TabIndex        =   3
      Top             =   1440
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdOk_Click()
Unload Me
FrmMain.SetFocus
End Sub

