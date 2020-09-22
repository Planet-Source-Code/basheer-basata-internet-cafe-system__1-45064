VERSION 5.00
Begin VB.Form FrmAdd 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4185
   Icon            =   "FrmAdd.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Dat1 
      Caption         =   "Dat1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Basheer Basata\Desktop\Internet2\Internet.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3150
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Internet"
      Top             =   1845
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   " Add New "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   45
      TabIndex        =   4
      Top             =   135
      Width           =   4020
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H80000018&
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   2205
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1575
         Width           =   825
      End
      Begin VB.ComboBox Cbo1 
         Height          =   315
         ItemData        =   "FrmAdd.frx":0442
         Left            =   1530
         List            =   "FrmAdd.frx":0452
         MouseIcon       =   "FrmAdd.frx":0466
         TabIndex        =   2
         Top             =   1035
         Width           =   1680
      End
      Begin VB.CommandButton CmdOK 
         BackColor       =   &H80000018&
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1575
         Width           =   825
      End
      Begin VB.TextBox TxtAccount 
         Height          =   330
         Left            =   1530
         TabIndex        =   1
         Top             =   675
         Width           =   1680
      End
      Begin VB.TextBox TxtName 
         Height          =   330
         Left            =   1530
         TabIndex        =   0
         Top             =   315
         Width           =   1680
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000016&
         Caption         =   "The hours"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   630
         TabIndex        =   7
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000016&
         Caption         =   "Account Number"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "The name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   780
      End
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cbo1_GotFocus()
If TxtName.Text <> "" And TxtAccount.Text <> "" Then CmdOk.Enabled = True
End Sub

Private Sub CmdOk_Click()
Dim bool As Boolean
bool = False
With Dat1.Recordset
.MoveFirst
Do While .EOF = False And bool = False
   If Val(TxtAccount.Text) = !Account_Number Or TxtName.Text = !The_name Then bool = True
Dat1.Recordset.MoveNext
Loop
.AddNew
If bool = False Then
!The_name = TxtName.Text
!Account_Number = Val(TxtAccount.Text)
!The_hours = Val(Cbo1.Text)
.Update
Unload Me
ElseIf bool = True Then
  MsgBox "Repeat data .. write new one ", vbCritical, "Error"
  TxtName = ""
  TxtAccount = ""
  Cbo1.ListIndex = -1
  .CancelUpdate
End If
End With
End Sub

Private Sub Cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dat1.DatabaseName = (App.Path & "\internet.mdb")
    Dat1.RecordSource = "internet"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
