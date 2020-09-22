VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form FrmBrowse 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Browse"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   Icon            =   "FrmBrowse.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Dat1 
      Caption         =   "Dat1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Basheer Basata\Desktop\Internet\Internet.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1755
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Internet"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FrmBrowse.frx":0442
      Height          =   4875
      Left            =   0
      OleObjectBlob   =   "FrmBrowse.frx":0455
      TabIndex        =   0
      Top             =   0
      Width           =   4830
   End
End
Attribute VB_Name = "FrmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Dat1.DatabaseName = (App.Path & "\internet.mdb")
 Dat1.RecordSource = "internet"
End Sub
