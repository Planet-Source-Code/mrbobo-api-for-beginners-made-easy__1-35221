Attribute VB_Name = "ModVBControls"
'API declares
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
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
    SetWindowLong& Tbox.hWnd, GWL_STYLE, InitTBStyle Or ES_NUMBER
End Sub
Public Sub LowercaseOnly(Tbox As TextBox)
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong& Tbox.hWnd, GWL_STYLE, InitTBStyle Or ES_LOWERCASE
End Sub
Public Sub UppercaseOnly(Tbox As TextBox)
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong& Tbox.hWnd, GWL_STYLE, InitTBStyle Or ES_UPPERCASE
End Sub
Public Sub SetInitialTBStyle(Tbox As TextBox)
    'Set the style, which window?, what style - normal or extended?, original style
    SetWindowLong& Tbox.hWnd, GWL_STYLE, InitTBStyle
End Sub
Public Sub GetInitialTBStyle(Tbox As TextBox)
    'variable = Get the style, which window?, what style - normal or extended?
    InitTBStyle = GetWindowLong&(Tbox.hWnd, GWL_STYLE)
End Sub
Public Sub SetInitialLVStyle(LV As ListView)
    'Set the style, which window?, what style - normal or extended?, original style
    SetWindowLong& hHeader, GWL_STYLE, InitLVStyle
End Sub
Public Sub GetInitialLVStyle(LV As ListView)
    hHeader = SendMessage(LV.hWnd, LVM_GETHEADER, 0, ByVal 0&) 'handle to column header
    'variable = Get the style, which window?, what style - normal or extended?
    InitLVStyle = GetWindowLong&(hHeader, GWL_STYLE)
End Sub
Public Sub LVFlatColumnHeaders(LV As ListView)
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong hHeader, GWL_STYLE, InitLVStyle Xor HDS_BUTTONS
End Sub
Public Sub SetInitialTVStyle(TV As TreeView)
    'Set the style, which window?, what style - normal or extended?, original style
    SetWindowLong& TV.hWnd, GWL_STYLE, InitTVStyle
End Sub
Public Sub GetInitialTVStyle(TV As TreeView)
    'variable = Get the style, which window?, what style - normal or extended?
    InitTVStyle = GetWindowLong&(TV.hWnd, GWL_STYLE)
End Sub
Public Sub TVNoTooltips(TV As TreeView)
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong TV.hWnd, GWL_STYLE, InitTVStyle Or TVS_NOTOOLTIPS
End Sub
Public Sub SetInitialBTStyle(BT As CommandButton)
    'if the style is already the original then dont do it again, may cause some flashing
    If GetWindowLong&(BT.hWnd, GWL_STYLE) = InitBTStyle Then Exit Sub
    'Set the style, which window?, what style - normal or extended?, original style
    SetWindowLong& BT.hWnd, GWL_STYLE, InitBTStyle
    BT.Refresh
End Sub
Public Sub GetInitialBTStyle(BT As CommandButton)
    'variable = Get the style, which window?, what style - normal or extended?
    InitBTStyle = GetWindowLong&(BT.hWnd, GWL_STYLE)
End Sub
Public Sub BTFlat(BT As CommandButton)
    'if the style is already the BS_FLAT then dont do it again, may cause some flashing
    If GetWindowLong&(BT.hWnd, GWL_STYLE) And BS_FLAT Then Exit Sub
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong BT.hWnd, GWL_STYLE, InitBTStyle Or BS_FLAT
    BT.Refresh
End Sub
Public Sub BTThick(BT As CommandButton)
    'Set the style, which window?, what style - normal or extended?, new style
    SetWindowLong BT.hWnd, GWL_STYLE, InitBTStyle Or BS_NULL
    BT.Refresh
End Sub
Public Sub PBcolor(PB As ProgressBar, Backcolor As Long, Forecolor As Long)
    'Send a message, which window?, what type of message, message value
    SendMessage PB.hWnd, CCM_SETBKCOLOR, 0, ByVal Backcolor
    SendMessage PB.hWnd, PBM_SETBARCOLOR, 0, ByVal Forecolor
End Sub
