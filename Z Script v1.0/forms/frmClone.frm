VERSION 5.00
Begin VB.Form frmClone 
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3120
      Top             =   2520
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1680
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   450
      Index           =   0
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   1680
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmClone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Check1(Index).Tag & ":lclick {") > 0 Then
    If Button = 1 Then
        'on jos78de:ole:lclick { -----> left click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Check1(Index).Tag & ":lclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If

If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Check1(Index).Tag & ":rclick {") > 0 Then
    If Button = 2 Then
        'on jos78de:ole:rclick {    ----> right click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Check1(Index).Tag & ":rclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If
End Sub

Private Sub Check1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    vFormName = Me.Tag
    'on jos78de:ole:mmove { -----> mouse move
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Check1(Index).Tag & ":mmove {") > 0 Then

    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Check1(Index).Tag & ":mmove {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Combo1_Change(Index As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Combo1(Index).Tag & ":change {") > 0 Then
    
    'on jos78de:ole:change { -----> change
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Combo1(Index).Tag & ":change {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Combo1_Click(Index As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Command1(Index).Tag & ":lclick {") > 0 Then
    'on jos78de:ole:lclick { -----> left click
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Command1(Index).Tag & ":lclick {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Combo1_DblClick(Index As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Combo1(Index).Tag & ":dclick {") > 0 Then
    
    'on jos78de:ole:dclick { -----> double click
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Combo1(Index).Tag & ":dclick {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Combo1_Scroll(Index As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Combo1(Index).Tag & ":scroll {") > 0 Then
    'on jos78de:ole:lclick { -----> left click
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Combo1(Index).Tag & ":lclick {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Command1(Index).Tag & ":lclick {") > 0 Then
    If Button = 1 Then
        'on jos78de:ole:lclick { -----> left click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Command1(Index).Tag & ":lclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)

        vRunSub = False
        vEvent = False
    End If
End If

If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Command1(Index).Tag & ":rclick {") > 0 Then
    If Button = 2 Then
        'on jos78de:ole:rclick {    ----> right click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Command1(Index).Tag & ":rclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)

        vRunSub = False
        vEvent = False
    End If
End If
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
    'on jos78de:ole:mmove { -----> mouse move
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Command1(Index).Tag & ":mmove {") > 0 Then

    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Command1(Index).Tag & ":mmove {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub



Private Sub Form_Load()
'not implimented yet
vFormName = Me.Tag
End Sub

Private Sub Form_Unload(Cancel As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":unload {") > 0 Then
    
    'on jos78de:unload { -----> on unload
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":unload {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Frame1_DblClick(Index As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Frame1(Index).Tag & ":dclick {") > 0 Then
    
    'on jos78de:ole:dclick { -----> double click
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Frame1(Index).Tag & ":dclick {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Frame1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Frame1(Index).Tag & ":lclick {") > 0 Then
   If Button = 1 Then
        'on jos78de:ole:lclick { -----> left click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Frame1(Index).Tag & ":lclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)

        vRunSub = False
        vEvent = False
    End If
End If

If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Frame1(Index).Tag & ":rclick {") > 0 Then
    If Button = 2 Then
        'on jos78de:ole:rclick {    ----> right click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Frame1(Index).Tag & ":rclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If
End Sub

Private Sub Frame1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    vFormName = Me.Tag
    'on jos78de:ole:mmove { -----> mouse move
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Frame1(Index).Tag & ":mmove {") > 0 Then

    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Frame1(Index).Tag & ":mmove {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Image1_DblClick(Index As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Image1(Index).Tag & ":dclick {") > 0 Then
    
    'on jos78de:ole:dclick { -----> double click
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Image1(Index).Tag & ":dclick {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Image1(Index).Tag & ":lclick {") > 0 Then
   If Button = 1 Then
        'on jos78de:ole:lclick { -----> left click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Image1(Index).Tag & ":lclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)

        vRunSub = False
        vEvent = False
    End If
End If

If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Image1(Index).Tag & ":rclick {") > 0 Then
    If Button = 2 Then
        'on jos78de:ole:rclick {    ----> right click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Image1(Index).Tag & ":rclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Image1(Index).Tag & ":mmove {") > 0 Then

    'on jos78de:ole:mmove { -----> mouse move
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Image1(Index).Tag & ":mmove {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
    
End If
End Sub

Private Sub Label1_DblClick(Index As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Label1(Index).Tag & ":dclick {") > 0 Then
    
    'on jos78de:ole:dclick { -----> double click
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Label1(Index).Tag & ":dclick {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Label1(Index).Tag & ":lclick {") > 0 Then
   If Button = 1 Then
        'on jos78de:ole:lclick { -----> left click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Label1(Index).Tag & ":lclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)

        vRunSub = False
        vEvent = False
    End If
End If

If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Label1(Index).Tag & ":rclick {") > 0 Then
    If Button = 2 Then
        'on jos78de:ole:rclick {    ----> right click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Label1(Index).Tag & ":rclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Label1(Index).Tag & ":mmove {") > 0 Then

    'on jos78de:ole:mmove { -----> mouse move
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Label1(Index).Tag & ":mmove {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
    
End If
End Sub

Private Sub Label2_Click(Index As Integer)
vFormName = Me.Tag
Dim Link
    Link = ShellExecute(hWnd, "Open", Label2(Index).Caption, &O0, &O0, SW_NORMAL)

End Sub

Private Sub List1_DblClick(Index As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & List1(Index).Tag & ":dclick {") > 0 Then
    
    'on jos78de:ole:dclick { -----> double click
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & List1(Index).Tag & ":dclick {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub List1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & List1(Index).Tag & ":lclick {") > 0 Then
    If Button = 1 Then
        'on jos78de:ole:lclick { -----> left click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & List1(Index).Tag & ":lclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If

If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & List1(Index).Tag & ":rclick {") > 0 Then
    If Button = 2 Then
        'on jos78de:ole:rclick {    ----> right click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & List1(Index).Tag & ":rclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If
End Sub

Private Sub List1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
    'on jos78de:ole:mmove { -----> mouse move
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & List1(Index).Tag & ":mmove {") > 0 Then

    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & List1(Index).Tag & ":mmove {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Option1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Option1(Index).Tag & ":lclick {") > 0 Then
    If Button = 1 Then
        'on jos78de:ole:lclick { -----> left click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Option1(Index).Tag & ":lclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If

If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Option1(Index).Tag & ":rclick {") > 0 Then
    If Button = 2 Then
        'on jos78de:ole:rclick {    ----> right click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Option1(Index).Tag & ":rclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If
End Sub

Private Sub Option1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
'on jos78de:ole:mmove { -----> mouse move
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Option1(Index).Tag & ":mmove {") > 0 Then

    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Option1(Index).Tag & ":mmove {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Text1_Change(Index As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Text1(Index).Tag & ":change {") > 0 Then
    
    'on jos78de:ole:change { -----> change
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Text1(Index).Tag & ":change {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Text1_DblClick(Index As Integer)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Text1(Index).Tag & ":dclick {") > 0 Then
    
    'on jos78de:ole:dclick { -----> double click
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Text1(Index).Tag & ":dclick {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub

Private Sub Text1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Text1(Index).Tag & ":lclick {") > 0 Then
    If Button = 1 Then
        'on jos78de:ole:lclick { -----> left click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Text1(Index).Tag & ":lclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If

If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Text1(Index).Tag & ":rclick {") > 0 Then
    If Button = 2 Then
        'on jos78de:ole:rclick {    ----> right click
        vRunSub = False
        vEvent = True
        nameofsub = "on " & Me.Tag & ":" & Text1(Index).Tag & ":rclick {"
        'MsgBox nameofsub
        codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
        'MsgBox codex
        codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
        RunCode (codex)
    
        vRunSub = False
        vEvent = False
    End If
End If
End Sub

Private Sub Text1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
vFormName = Me.Tag
If InStr(1, frmMain.Text1.Text, "on " & Me.Tag & ":" & Text1(Index).Tag & ":mmove {") > 0 Then
    
    'on jos78de:ole:mmove { -----> mouse move
    vRunSub = False
    vEvent = True
    nameofsub = "on " & Me.Tag & ":" & Text1(Index).Tag & ":mmove {"
    'MsgBox nameofsub
    codex = Mid(frmMain.Text1.Text, InStr(1, frmMain.Text1.Text, nameofsub) + Len(nameofsub))
    'MsgBox codex
    codex = Mid(codex, 1, InStr(1, codex, "}") - 1)
    RunCode (codex)

    vRunSub = False
    vEvent = False
End If
End Sub
