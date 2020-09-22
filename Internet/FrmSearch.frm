VERSION 5.00
Begin VB.Form FrmSearch 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   Icon            =   "FrmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCancel2 
      BackColor       =   &H80000018&
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3600
      Width           =   735
   End
   Begin VB.Timer Tim1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3735
      Top             =   990
   End
   Begin VB.CommandButton CmdOK2 
      BackColor       =   &H80000018&
      Caption         =   "OK"
      Height          =   330
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000016&
      Caption         =   " Result "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   90
      TabIndex        =   11
      Top             =   1890
      Width           =   4020
      Begin VB.TextBox TxtName2 
         DataField       =   "The_name"
         DataSource      =   "Dat1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1530
         TabIndex        =   5
         Top             =   315
         Width           =   1680
      End
      Begin VB.TextBox TxtAccount2 
         DataField       =   "Account_Number"
         DataSource      =   "Dat1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1530
         TabIndex        =   6
         Top             =   675
         Width           =   1680
      End
      Begin VB.TextBox TxtHours 
         DataField       =   "The_hours"
         DataSource      =   "Dat1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1530
         MaxLength       =   5
         TabIndex        =   7
         Top             =   1035
         Width           =   1680
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
         TabIndex        =   14
         Top             =   360
         Width           =   780
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
         TabIndex        =   13
         Top             =   720
         Width           =   1275
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
         TabIndex        =   12
         Top             =   1080
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   " Search "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   4020
      Begin VB.CommandButton Cmdcancel 
         BackColor       =   &H80000018&
         Caption         =   "Cancel"
         Height          =   330
         Left            =   2115
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1170
         Width           =   735
      End
      Begin VB.Data Dat1 
         Caption         =   "Dat1"
         Connect         =   "Access"
         DatabaseName    =   "E:\VB\Blood Bank\BloodBank.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3555
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1305
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TxtAccount 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1575
         TabIndex        =   1
         Top             =   675
         Width           =   1545
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000016&
         Caption         =   "By number"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   675
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000016&
         Caption         =   "By name"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   315
         Width           =   1185
      End
      Begin VB.TextBox TxtName 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1575
         TabIndex        =   0
         Top             =   315
         Width           =   1545
      End
      Begin VB.CommandButton CmdOK 
         BackColor       =   &H80000018&
         Caption         =   "OK"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1170
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdcancel_Click()
    Unload Me
End Sub

Private Sub CmdCancel2_Click()
    Dat1.Recordset.CancelUpdate
    Unload Me
End Sub

Private Sub CmdOK_Click()
  Dim Number As Long, Name As String
    If Option1.Value = True Then
    Name = TxtName.Text
    Dat1.Refresh
    Dat1.Recordset.FindFirst "The_name like '" & Name & "'"
    If Not Dat1.Recordset.NoMatch Then
    Tim1_Timer
    Else: MsgBox "There is no body has this name ", vbCritical, "Error"
    End If
    ElseIf Option2.Value = True Then
    Number = Val(TxtAccount.Text)
    Dat1.Recordset.FindFirst "Account_Number = " & Number & ""
    If Not Dat1.Recordset.NoMatch Then
    Tim1_Timer
    Else: MsgBox "There is no body has this name ", vbCritical, "Error"
    End If
    TxtName = ""
    TxtAccount = ""
    End If
Dat1.Recordset.Edit
End Sub

Private Sub CmdOK2_Click()
    Dat1.Recordset.Edit
    Dat1.Recordset.Update
    Height = 2385
    Unload Me
End Sub


Private Sub Form_Load()
 Dat1.DatabaseName = (App.Path & "\internet.mdb")
 Dat1.RecordSource = "internet"
 End Sub


Private Sub Option1_Click()
    If Option1.Value = True Then
    TxtName.Enabled = True
    TxtAccount.Enabled = False
    ElseIf Option1.Value = False Then
    TxtName.Enabled = False
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
    TxtAccount.Enabled = True
    TxtName.Enabled = False
    ElseIf Option2.Value = False Then
    TxtAccount.Enabled = False
    End If
End Sub

Private Sub Tim1_Timer()
  If Height < 4544 Then
  Tim1.Enabled = True
  Height = Height + 30
  ElseIf Height >= 4544 Then
  Tim1.Enabled = False
  End If
End Sub

Private Sub TxtAccount_Change()
    If TxtAccount.Text <> "" Then CmdOK.Enabled = True
End Sub


Private Sub TxtName_Change()
  If TxtName.Text <> "" Then CmdOK.Enabled = True
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 48 And KeyAscii <= 57 And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
  If (KeyAscii < 48 Or KeyAscii > 57) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
