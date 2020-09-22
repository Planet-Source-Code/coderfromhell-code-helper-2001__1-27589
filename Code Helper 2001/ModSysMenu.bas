Attribute VB_Name = "ModSysMenu"
Public Declare Function GetSystemMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal bRevert _
    As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_ENABLED = &H0
Public Const MFS_CHECKED = &H8
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal _
    hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii _
    As MENUITEMINFO) As Long
Public Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal _
    hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii _
    As MENUITEMINFO) As Long

Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal _
    hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal _
    cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd _
    As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = -4
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal _
    lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam _
    As Long, ByVal lParam As Long) As Long
Public Const WM_SYSCOMMAND = &H112
Public Const WM_INITMENU = &H116

Public pOldProc As Long
Public ontop As Boolean

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam _
        As Long, ByVal lParam As Long) As Long
    Dim hSysMenu As Long
    Dim mii As MENUITEMINFO
    Dim retval As Long
    
    Select Case uMsg
    Case WM_INITMENU
        hSysMenu = GetSystemMenu(hwnd, 0)
        With mii
            .cbSize = Len(mii)
            .fMask = MIIM_STATE
            .fState = MFS_ENABLED Or IIf(ontop, MFS_CHECKED, 0)
        End With
        retval = SetMenuItemInfo(hSysMenu, 1, 0, mii)
        WindowProc = 0
    Case WM_SYSCOMMAND
        If wParam = 1 Then
            ontop = Not ontop
            If ontop = True Then
                ColorBox frmView.Text1
            Else
                frmView.Text1.SelStart = 0
                frmView.Text1.SelLength = Len(frmView.Text1.Text)
                frmView.Text1.SelColor = vbBlack
                frmView.Text1.SelLength = 0
            End If
            WindowProc = 0
        Else
            WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)
        End If
    Case Else
        WindowProc = CallWindowProc(pOldProc, hwnd, uMsg, wParam, lParam)
    End Select
End Function





