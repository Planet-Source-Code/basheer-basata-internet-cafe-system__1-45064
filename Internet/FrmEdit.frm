VERSION 5.00
Begin VB.Form FrmEdit 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3765
   Icon            =   "FrmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H80000018&
      Caption         =   "&OK"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1035
      Width           =   870
   End
   Begin VB.TextBox TxtCurr 
      BackColor       =   &H80000014&
      Height          =   330
      Left            =   2520
      TabIndex        =   3
      Top             =   585
      Width           =   1140
   End
   Begin VB.TextBox TxtPrice 
      BackColor       =   &H80000014&
      Height          =   330
      Left            =   2520
      TabIndex        =   2
      Top             =   135
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000016&
      Caption         =   "Enter the new currency : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   630
      TabIndex        =   1
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      Caption         =   "Enter the new price per minute : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   2355
   End
End
Attribute VB_Name = "FrmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdOk_Click()

Open "price.txt" For Output As #1
Print #1, (TxtPrice.Text)
Close #1

Open "curr.txt" For Output As #2
Print #2, (TxtCurr.Text)
Close #2

Open "price.txt" For Input As #1
price = Input(LOF(1), #1)
Close #1

Open "curr.txt" For Input As #2
curr = Input(LOF(2), #2)
Close #2


Unload Me

End Sub

Private Sub TxtCurr_Change()
    If TxtPrice <> "" And TxtCurr <> "" Then CmdOk.Enabled = True
End Sub

Private Sub TxtPrice_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And Not KeyAscii = 8 And Not KeyAscii = 46 Then KeyAscii = 0
End Sub

Private Sub Txtcurr_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
