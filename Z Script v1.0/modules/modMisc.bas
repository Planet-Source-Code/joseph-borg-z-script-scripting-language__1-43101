Attribute VB_Name = "modMisc"
'open a link
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_NORMAL = 1


'API declares
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'API constants
'Textbox
Private Const ES_NUMBER = &H2000&
Private Const ES_LOWERCASE = &H10&
Private Const ES_UPPERCASE = &H8&
'Listview
Private Const HDS_BUTTONS As Long = &H2
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_GETHEADER As Long = (LVM_FIRST + 31)
'Treeview
Private Const TVS_NOTOOLTIPS = &H80
'Commandbutton
Private Const BS_FLAT = &H8000&
Private Const BS_NULL = 1
'Progressbar
Private Const WM_USER = &H400
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
'Type of style to change - normal
Private Const GWL_STYLE = (-16)
'variables
Public mHover As Boolean
Dim InitTBStyle As Long, InitLVStyle As Long, InitTVStyle As Long
Dim InitBTStyle As Long, InitPBStyle As Long, hHeader As Long
Public Sub NumberOnly(Tbox As TextBox)
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong& Tbox.hwnd, GWL_STYLE, InitTBStyle Or ES_NUMBER
End Sub
Public Sub LowercaseOnly(Tbox As TextBox)
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong& Tbox.hwnd, GWL_STYLE, InitTBStyle Or ES_LOWERCASE
End Sub
Public Sub UppercaseOnly(Tbox As TextBox)
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong& Tbox.hwnd, GWL_STYLE, InitTBStyle Or ES_UPPERCASE
End Sub
Public Sub SetInitialTBStyle(Tbox As TextBox)
    'Set the style, which window?, what style - normal or extended?, original style
    SetWindowLong& Tbox.hwnd, GWL_STYLE, InitTBStyle
End Sub
Public Sub GetInitialTBStyle(Tbox As TextBox)
    'variable = Get the style, which window?, what style - normal or extended?
    InitTBStyle = GetWindowLong&(Tbox.hwnd, GWL_STYLE)
End Sub

Public Sub SetInitialBTStyle(BT As CommandButton)
    'if the style is already the original then dont do it again, may cause some flashing
    If GetWindowLong&(BT.hwnd, GWL_STYLE) = InitBTStyle Then Exit Sub
    'Set the style, which window?, what style - normal or extended?, original style
    SetWindowLong& BT.hwnd, GWL_STYLE, InitBTStyle
    BT.Refresh
End Sub
Public Sub GetInitialBTStyle(BT As CommandButton)
    'variable = Get the style, which window?, what style - normal or extended?
    InitBTStyle = GetWindowLong&(BT.hwnd, GWL_STYLE)
End Sub
Public Sub BTFlat(BT As CommandButton)
    'if the style is already the BS_FLAT then dont do it again, may cause some flashing
    If GetWindowLong&(BT.hwnd, GWL_STYLE) And BS_FLAT Then Exit Sub
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong BT.hwnd, GWL_STYLE, InitBTStyle Or BS_FLAT
    BT.Refresh
End Sub
Public Sub BTThick(BT As CommandButton)
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong BT.hwnd, GWL_STYLE, InitBTStyle Or BS_NULL
    BT.Refresh
End Sub
