VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmView 
   Caption         =   "View"
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5715
   ScaleWidth      =   9510
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3201
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      RightMargin     =   64000
      TextRTF         =   $"frmView.frx":014A
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long

Private F4Down As Boolean

Private Sub Form_Resize()
Text1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Text1_Change()
HideCaret Text1.hwnd
End Sub

Private Sub Text1_GotFocus()
HideCaret Text1.hwnd
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 And Shift = 0 And Me.Caption = "Information" Then
    F4Down = True
End If
HideCaret Text1.hwnd
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 And Shift = 0 And Me.Caption = "Information" And F4Down Then
    Unload Me
End If
F4Down = False
HideCaret Text1.hwnd
End Sub


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'__________________________________________________________________________________

Private Sub Form_Load()

    If InfoBox Then Exit Sub
    Text1.AutoVerbMenu = True
    Dim hSysMenu As Long
    Dim count As Long
    Dim mii As MENUITEMINFO
    Dim retval As Long
    
    hSysMenu = GetSystemMenu(Me.hwnd, 0)
    count = GetMenuItemCount(hSysMenu)
    
    With mii
        .cbSize = Len(mii)
        .fMask = MIIM_ID Or MIIM_TYPE
        .fType = MFT_SEPARATOR
        .wID = 0
    End With
    retval = InsertMenuItem(hSysMenu, count, 1, mii)
    
    With mii
        .fMask = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
        .fType = MFT_STRING
        .fState = MFS_ENABLED
        .wID = 1
        .dwTypeData = "Syntax &Coloring"
        .cch = Len(.dwTypeData)
    End With
    retval = InsertMenuItem(hSysMenu, count + 1, 1, mii)
    
    ontop = True
    pOldProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If InfoBox Then Exit Sub
    Dim retval As Long
    
    retval = SetWindowLong(Me.hwnd, GWL_WNDPROC, pOldProc)
    retval = GetSystemMenu(Me.hwnd, 1)
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
HideCaret Text1.hwnd
End Sub

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
HideCaret Text1.hwnd
End Sub
