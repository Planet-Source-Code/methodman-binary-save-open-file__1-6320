VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Save/Open Binary File"
   ClientHeight    =   4185
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"binary.frx":0000
   End
   Begin VB.ListBox List4 
      Height          =   645
      Left            =   1560
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   645
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"binary.frx":00C9
   End
   Begin VB.ListBox List2 
      Height          =   645
      ItemData        =   "binary.frx":020B
      Left            =   1560
      List            =   "binary.frx":0236
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "binary.frx":028C
      Left            =   120
      List            =   "binary.frx":02B7
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read from file (Open)"
      Height          =   615
      Left            =   4080
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Write to file (Save)"
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "These are the fields where the file will be opened. (2) Listboxes, (1) Label, and (1) RTB"
      Height          =   855
      Left            =   3960
      TabIndex        =   11
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "These are the fields to be saved. (2) Listboxes, (1) Label, and (1) RTB"
      Height          =   855
      Left            =   3960
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "0"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "77699"
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
Call ShowSave("C:\Windows\Desktop", "Results") ''CALLS ShowSave IN MODULE
End Sub


Private Sub Command2_Click()
Call ShowOpen("C:\Windows\Desktop", "Results") ''CALLS ShowOpen IN MODULE
End Sub


Private Sub mnuabout_Click()
MsgBox "        By: Michael Schmidt" & Chr(10) & "   Contact Me: mds@vci.net" & Chr(10) & "               2-28-2000" ''SETS ABOUT MESSAGE
End Sub
