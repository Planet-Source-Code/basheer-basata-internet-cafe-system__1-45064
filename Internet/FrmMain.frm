VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Internet Cafe Program"
   ClientHeight    =   3810
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   7260
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Dat1 
      Caption         =   "Dat1"
      Connect         =   "Access"
      DatabaseName    =   "E:\VB\Internet2\Internet.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6030
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Internet"
      Top             =   3690
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Txt2m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3330
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   87
      Top             =   135
      Width           =   510
   End
   Begin VB.TextBox Txt2m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3330
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   86
      Top             =   585
      Width           =   510
   End
   Begin VB.TextBox Txt2m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3330
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   85
      Top             =   1035
      Width           =   510
   End
   Begin VB.TextBox Txt2m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3330
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   84
      Top             =   1485
      Width           =   510
   End
   Begin VB.TextBox Txt2m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3330
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   83
      Top             =   1935
      Width           =   510
   End
   Begin VB.TextBox Txt2m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3330
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   82
      Top             =   2385
      Width           =   510
   End
   Begin VB.TextBox Txt2m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3330
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   81
      Top             =   2835
      Width           =   510
   End
   Begin VB.TextBox Txt2m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3330
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   80
      Top             =   3285
      Width           =   510
   End
   Begin VB.TextBox Txt1m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1710
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   79
      Top             =   135
      Width           =   510
   End
   Begin VB.TextBox Txt1m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1710
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   78
      Top             =   585
      Width           =   510
   End
   Begin VB.TextBox Txt1m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1710
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   77
      Top             =   1035
      Width           =   510
   End
   Begin VB.TextBox Txt1m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1710
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   76
      Top             =   1485
      Width           =   510
   End
   Begin VB.TextBox Txt1m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1710
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   75
      Top             =   1935
      Width           =   510
   End
   Begin VB.TextBox Txt1m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1710
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   74
      Top             =   2385
      Width           =   510
   End
   Begin VB.TextBox Txt1m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1710
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   73
      Top             =   2835
      Width           =   510
   End
   Begin VB.TextBox Txt1m 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1710
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   72
      Top             =   3285
      Width           =   510
   End
   Begin VB.CommandButton Cmd5 
      BackColor       =   &H80000018&
      Caption         =   "Account"
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
      Index           =   0
      Left            =   6390
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   180
      Width           =   780
   End
   Begin VB.CommandButton Cmd5 
      BackColor       =   &H80000018&
      Caption         =   "Account"
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
      Index           =   1
      Left            =   6390
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   630
      Width           =   780
   End
   Begin VB.CommandButton Cmd5 
      BackColor       =   &H80000018&
      Caption         =   "Account"
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
      Index           =   2
      Left            =   6390
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   780
   End
   Begin VB.CommandButton Cmd5 
      BackColor       =   &H80000018&
      Caption         =   "Account"
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
      Index           =   3
      Left            =   6390
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1530
      Width           =   780
   End
   Begin VB.CommandButton Cmd5 
      BackColor       =   &H80000018&
      Caption         =   "Account"
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
      Index           =   4
      Left            =   6390
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   1980
      Width           =   780
   End
   Begin VB.CommandButton Cmd5 
      BackColor       =   &H80000018&
      Caption         =   "Account"
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
      Index           =   5
      Left            =   6390
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   2430
      Width           =   780
   End
   Begin VB.CommandButton Cmd5 
      BackColor       =   &H80000018&
      Caption         =   "Account"
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
      Index           =   6
      Left            =   6390
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   2880
      Width           =   780
   End
   Begin VB.CommandButton Cmd5 
      BackColor       =   &H80000018&
      Caption         =   "Account"
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
      Index           =   7
      Left            =   6390
      MaskColor       =   &H0000C0C0&
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3330
      Width           =   780
   End
   Begin VB.CommandButton Cmd4 
      BackColor       =   &H80000018&
      Caption         =   "Cash"
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
      Index           =   7
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   3330
      Width           =   735
   End
   Begin VB.CommandButton Cmd4 
      BackColor       =   &H80000018&
      Caption         =   "Cash"
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
      Index           =   6
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Cmd4 
      BackColor       =   &H80000018&
      Caption         =   "Cash"
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
      Index           =   5
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   2430
      Width           =   735
   End
   Begin VB.CommandButton Cmd4 
      BackColor       =   &H80000018&
      Caption         =   "Cash"
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
      Index           =   4
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   1980
      Width           =   735
   End
   Begin VB.CommandButton Cmd4 
      BackColor       =   &H80000018&
      Caption         =   "Cash"
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
      Index           =   3
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1530
      Width           =   735
   End
   Begin VB.CommandButton Cmd4 
      BackColor       =   &H80000018&
      Caption         =   "Cash"
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
      Index           =   2
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Cmd4 
      BackColor       =   &H80000018&
      Caption         =   "Cash"
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
      Index           =   1
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   630
      Width           =   735
   End
   Begin VB.CommandButton Cmd4 
      BackColor       =   &H80000018&
      Caption         =   "Cash"
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
      Index           =   0
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   735
   End
   Begin VB.TextBox Txt3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4455
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   3285
      Width           =   960
   End
   Begin VB.TextBox Txt3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4455
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   2835
      Width           =   960
   End
   Begin VB.TextBox Txt3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4455
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   2385
      Width           =   960
   End
   Begin VB.TextBox Txt3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4455
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   1935
      Width           =   960
   End
   Begin VB.TextBox Txt3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4455
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   1485
      Width           =   960
   End
   Begin VB.TextBox Txt3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4455
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   1035
      Width           =   960
   End
   Begin VB.TextBox Txt3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4455
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   585
      Width           =   960
   End
   Begin VB.TextBox Txt3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4455
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   135
      Width           =   960
   End
   Begin VB.CommandButton Cmd3 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   7
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   3330
      Width           =   420
   End
   Begin VB.CommandButton Cmd3 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   6
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2880
      Width           =   420
   End
   Begin VB.CommandButton Cmd3 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   5
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   2430
      Width           =   420
   End
   Begin VB.CommandButton Cmd3 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   4
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1980
      Width           =   420
   End
   Begin VB.CommandButton Cmd3 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   3
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1530
      Width           =   420
   End
   Begin VB.CommandButton Cmd3 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   2
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   420
   End
   Begin VB.CommandButton Cmd3 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   1
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   630
      Width           =   420
   End
   Begin VB.CommandButton Cmd3 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   0
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   180
      Width           =   420
   End
   Begin VB.TextBox Txt2h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   2835
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   51
      Top             =   3285
      Width           =   510
   End
   Begin VB.TextBox Txt2h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2835
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   50
      Top             =   2835
      Width           =   510
   End
   Begin VB.TextBox Txt2h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2835
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   49
      Top             =   2385
      Width           =   510
   End
   Begin VB.TextBox Txt2h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2835
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   48
      Top             =   1935
      Width           =   510
   End
   Begin VB.TextBox Txt2h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2835
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   47
      Top             =   1485
      Width           =   510
   End
   Begin VB.TextBox Txt2h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2835
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   46
      Top             =   1035
      Width           =   510
   End
   Begin VB.TextBox Txt2h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2835
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   45
      Top             =   585
      Width           =   510
   End
   Begin VB.TextBox Txt2h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2835
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   44
      Top             =   135
      Width           =   510
   End
   Begin VB.CommandButton Cmd2 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   7
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3330
      Width           =   420
   End
   Begin VB.CommandButton Cmd2 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   6
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   2880
      Width           =   420
   End
   Begin VB.CommandButton Cmd2 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   5
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2430
      Width           =   420
   End
   Begin VB.CommandButton Cmd2 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   4
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   1980
      Width           =   420
   End
   Begin VB.CommandButton Cmd2 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   3
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1530
      Width           =   420
   End
   Begin VB.CommandButton Cmd2 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   2
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1080
      Width           =   420
   End
   Begin VB.CommandButton Cmd2 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   1
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   630
      Width           =   420
   End
   Begin VB.CommandButton Cmd2 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Height          =   285
      Index           =   0
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   180
      Width           =   420
   End
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Index           =   7
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3330
      Width           =   420
   End
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Index           =   6
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2880
      Width           =   420
   End
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Index           =   5
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2430
      Width           =   420
   End
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Index           =   4
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1980
      Width           =   420
   End
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Index           =   3
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1530
      Width           =   420
   End
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1080
      Width           =   420
   End
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Index           =   1
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   630
      Width           =   420
   End
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H80000018&
      Caption         =   "-->"
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
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   420
   End
   Begin VB.TextBox Txt1h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   35
      Top             =   3285
      Width           =   510
   End
   Begin VB.TextBox Txt1h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   34
      Top             =   2835
      Width           =   510
   End
   Begin VB.TextBox Txt1h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   33
      Top             =   2385
      Width           =   510
   End
   Begin VB.TextBox Txt1h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   32
      Top             =   1935
      Width           =   510
   End
   Begin VB.TextBox Txt1h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   31
      Top             =   1485
      Width           =   510
   End
   Begin VB.TextBox Txt1h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   30
      Top             =   1035
      Width           =   510
   End
   Begin VB.TextBox Txt1h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   29
      Top             =   585
      Width           =   510
   End
   Begin VB.TextBox Txt1h 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   20
      Top             =   135
      Width           =   510
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PC8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   135
      TabIndex        =   28
      Top             =   3330
      Width           =   420
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PC7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   135
      TabIndex        =   27
      Top             =   2880
      Width           =   420
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PC6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   135
      TabIndex        =   26
      Top             =   2430
      Width           =   420
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PC5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   135
      TabIndex        =   25
      Top             =   1980
      Width           =   420
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PC4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   135
      TabIndex        =   24
      Top             =   1530
      Width           =   420
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PC2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   135
      TabIndex        =   23
      Top             =   630
      Width           =   420
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PC3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   135
      TabIndex        =   22
      Top             =   1080
      Width           =   420
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PC1 "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   21
      Top             =   180
      Width           =   420
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&Accounts"
      Begin VB.Menu MnuAdd 
         Caption         =   "&Add New"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuBrowse 
         Caption         =   "&Browse"
         Shortcut        =   ^B
      End
      Begin VB.Menu MnuSearch 
         Caption         =   "&Search"
         Shortcut        =   ^S
      End
      Begin VB.Menu MnuUpdate 
         Caption         =   "&Update"
         Shortcut        =   ^U
      End
      Begin VB.Menu MnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuCurrencyAndPrice 
         Caption         =   "&Currency and Price"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "A&bout"
   End
   Begin VB.Menu MnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd1_Click(Index As Integer)
    Txt1h(Index).Text = Hour(Time)
    Txt1m(Index).Text = Minute(Time)
End Sub

Private Sub Cmd2_Click(Index As Integer)
    Txt2h(Index).Text = Hour(Time)
    Txt2m(Index).Text = Minute(Time)
End Sub

Private Sub Cmd3_Click(Index As Integer)
   Dim X As Double, Y As Double
    X = (Val(Txt1h(Index).Text) * 60) + (Val(Txt1m(Index).Text))
    Y = (Val(Txt2h(Index).Text) * 60) + (Val(Txt2m(Index).Text))
    Hours = (Y - X) / 60
    Result = (Y - X) * price
    Txt3(Index).Text = Result & " " & curr
End Sub

Private Sub Cmd4_Click(Index As Integer)
    Txt1h(Index).Text = ""
    Txt1m(Index).Text = ""
    Txt2h(Index).Text = ""
    Txt2m(Index).Text = ""
    Txt3(Index).Text = ""
    Txt1h(Index).Enabled = False
    Txt1m(Index).Enabled = False
    Txt2h(Index).Enabled = False
    Txt2m(Index).Enabled = False
    Txt3(Index).Enabled = False
    Cmd2(Index).Enabled = False
    Cmd3(Index).Enabled = False
    Cmd4(Index).Enabled = False
    Cmd5(Index).Enabled = False
End Sub

Private Sub Cmd5_Click(Index As Integer)
    Dim Msg As Double, Msg2 As Variant, Msg3 As Variant, bool As Boolean
    bool = False
    On Error GoTo handle
    Msg = InputBox("Enter the account number .. please :", "Input")
    Dat1.Refresh
    Do While Dat1.Recordset.EOF = False And bool = False
    If Msg = Dat1.Recordset!Account_Number And Dat1.Recordset!The_hours > 0 Then
    Dat1.Recordset.Edit
    Dat1.Recordset!The_hours = Dat1.Recordset!The_hours - Hours
    Dat1.Recordset.Update
    Msg2 = "The account (" & Msg & ") , has now " & Dat1.Recordset!The_hours & " hours"
    Msg3 = "Account " & "(" & Msg & ")"
    MsgBox Msg2, vbQuestion, Msg3
    Txt1h(Index).Text = ""
    Txt1m(Index).Text = ""
    Txt2h(Index).Text = ""
    Txt2m(Index).Text = ""
    Txt3(Index).Text = ""
    Txt1h(Index).Enabled = False
    Txt1m(Index).Enabled = False
    Txt2h(Index).Enabled = False
    Txt2m(Index).Enabled = False
    Txt3(Index).Enabled = False
    Cmd2(Index).Enabled = False
    Cmd3(Index).Enabled = False
    Cmd4(Index).Enabled = False
    Cmd5(Index).Enabled = False
    bool = True
    End If
    Dat1.Recordset.MoveNext
    Loop
     If bool = False Then
     MsgBox "There is no account number as this number ", vbCritical, "Error"
    End If
handle:
End Sub

Private Sub Form_Load()
   
   Dat1.DatabaseName = (App.Path & "\internet.mdb")
 Dat1.RecordSource = "internet"
   
   On Error Resume Next
Open "price.txt" For Input As #1
price = Input(LOF(1), #1)
Close #1

Open "curr.txt" For Input As #2
curr = Input(LOF(2), #2)
Close #2
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Msg
    Msg = MsgBox("Are you sure to exit ?", vbQuestion + vbOKCancel, "Exit")
    If Msg = vbOK Then
    End
    Else
    Cancel = True
    End If
End Sub



Private Sub MnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub MnuAdd_Click()
    FrmAdd.Show
End Sub

Private Sub MnuBrowse_Click()
    FrmBrowse.Show
End Sub


Private Sub MnuCurrencyAndPrice_Click()
    FrmEdit.Show
End Sub

Private Sub MnuDelete_Click()
    Dim Msg As Long, bool As Boolean
    On Error GoTo h
    Msg = InputBox("Enter thee account number you wnat to delete it               (the account must be its hours = 0 to can delete it):", "Delete")
    Do While Dat1.Recordset.EOF = False And bool = False
    If Msg = Dat1.Recordset!Account_Number And Dat1.Recordset!The_hours = 0 Then
    Dat1.Recordset.Delete
    bool = True
    End If
    Dat1.Recordset.MoveNext
    Loop
    If bool = False Then MsgBox "You can not delete this account ", vbCritical, "Error"
h:
End Sub

Private Sub MnuExit_Click()
    Unload Me
End Sub

Private Sub MnuSearch_Click()
    FrmSearch.Caption = "Search"
    FrmSearch.Show
    FrmSearch.TxtAccount2.Locked = True
    FrmSearch.TxtName2.Locked = True
    FrmSearch.TxtHours.Locked = True
End Sub

Private Sub MnuUpdate_Click()
    FrmSearch.Caption = "Update"
    FrmSearch.Show 1
    FrmSearch.TxtAccount2.Locked = False
    FrmSearch.TxtName2.Locked = False
    FrmSearch.TxtHours.Locked = False
End Sub

Private Sub Txt1m_Change(Index As Integer)
    If Txt1h(Index).Text <> "" Then
    Cmd2(Index).Enabled = True
    End If
End Sub

Private Sub Txt2m_Change(Index As Integer)
    If Txt2m(Index).Text <> "" Then
    Cmd3(Index).Enabled = True
    End If
End Sub

Private Sub Txt3_Change(Index As Integer)
    If Txt3(Index).Text <> "" Then
    Cmd4(Index).Enabled = True
    Cmd5(Index).Enabled = True
    End If
End Sub
