VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView TV 
      Height          =   4095
      Left            =   6720
      TabIndex        =   5
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7223
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   6375
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmMain.frx":0000
      Top             =   120
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RUN"
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ZScript by Joseph Borg (j23ld@hotmail.com)"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   6495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call Command2_Click
List1.Clear

RunCode (Text1)
End Sub




Private Sub Command2_Click()
For i = 0 To vControlCount + 1
    Unload frmCloned(i)
Next i
List1.Clear
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub
