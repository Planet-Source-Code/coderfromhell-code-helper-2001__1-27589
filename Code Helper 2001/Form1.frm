VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Code Helper 2001"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6600
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":02A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0402
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":055E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1032
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":114A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1266
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1382
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSep 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   2880
      ScaleHeight     =   3375
      ScaleWidth      =   135
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox List1 
      Height          =   3180
      IntegralHeight  =   0   'False
      ItemData        =   "Form1.frx":200A
      Left            =   0
      List            =   "Form1.frx":200C
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   6600
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "News Gothic MT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   45
         Width           =   6015
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   1005
      ButtonWidth     =   1323
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Select all"
            Key             =   "Select"
            Object.ToolTipText     =   "Highlight/Select All Text"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy"
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy Highlighted/Selected Text"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Copy all"
            Key             =   "Copy all"
            Object.ToolTipText     =   "Copy All Text"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "Search"
            Object.ToolTipText     =   "Search For Keyword"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Info"
            Key             =   "Info"
            Object.ToolTipText     =   "Code Information"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Pic"
            Key             =   "Pic"
            Object.ToolTipText     =   "Code Picture"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Color"
            Key             =   "Color"
            Object.ToolTipText     =   "Syntax Coloring"
            ImageIndex      =   14
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2655
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
      _Version        =   393217
      BackColor       =   -2147483639
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      RightMargin     =   64000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"Form1.frx":200E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image ImgSep 
      Height          =   3255
      Left            =   2400
      MousePointer    =   9  'Size W E
      Top             =   960
      Width           =   135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCopyall 
         Caption         =   "Copy all"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
         Visible         =   0   'False
      End
      Begin VB.Menu Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "&Select all"
         Shortcut        =   ^A
      End
      Begin VB.Menu Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuInfo 
         Caption         =   "&Information"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuViewPic 
         Caption         =   "&Picture"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Bar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSynCol 
         Caption         =   "&Syntax Coloring"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu Bar3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
         Shortcut        =   {F1}
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurX As Single, CurY As Single
Private Const MinWidth = 6720
Private Const MinHeight = 4110
Private Loading As Boolean

Private Sub Form_Load()
On Local Error GoTo XS
Dim Num As Integer, WinState As Integer
Loading = True
WinState = GetSetting("CH2001", "Main", "WinState", vbNormal)
    Me.Left = GetSetting("CH2001", "Main", "Left", 1000)
    Me.Top = GetSetting("CH2001", "Main", "Top", 1000)
    Me.Width = GetSetting("CH2001", "Main", "Width", MinWidth)
    Me.Height = GetSetting("CH2001", "Main", "Height", MinHeight)
If WinState = vbMaximized + 1 Then
    Me.WindowState = WinState - 1
End If
List1.Width = GetSetting("CH2001", "Main", "SepWid", 2000)
Loading = False
Form_Resize
List1.Clear
Show
SetInfo
DoEvents
SetScriptKeywords
OpenFile
DoEvents
For Num = 0 To UBound(MyCode)
    List1.AddItem MyCode(Num).EntryName
Next Num
If List1.ListCount > 0 Then List1.ListIndex = 0
XS:
End Sub

Private Sub Form_Resize()
If Loading = True Then Exit Sub
If Me.WindowState <> vbMinimized Then
    If Me.Width < MinWidth Then Me.Width = MinWidth
    If Me.Height < MinHeight Then Me.Height = MinHeight
    List1.Move 0, Picture1.Top + Picture1.Height + 40, List1.Width, ScaleHeight - (Picture1.Top + Picture1.Height + 40)
    Text1.Move List1.Left + List1.Width + 40, List1.Top, ScaleWidth - (List1.Left + List1.Width + 40), List1.Height
    ImgSep.Left = List1.Left + List1.Width
    ImgSep.Height = ScaleHeight - ImgSep.Top
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.WindowState <> vbMaximized Then
    SaveSetting "CH2001", "Main", "Left", Me.Left
    SaveSetting "CH2001", "Main", "Top", Me.Top
    SaveSetting "CH2001", "Main", "Width", Me.Width
    SaveSetting "CH2001", "Main", "Height", Me.Height
    SaveSetting "CH2001", "Main", "SepWid", ImgSep.Left
End If
SaveSetting "CH2001", "Main", "WinState", Me.WindowState + 1
End Sub

Private Sub ImgSep_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    CurX = x: CurY = y
    picSep.Left = ImgSep.Left
    picSep.Height = ImgSep.Height
    picSep.Width = 40
    picSep.Visible = True
End If
End Sub

Private Sub ImgSep_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim L As Single
If Button = 1 Then
    L = ImgSep.Left + x - CurX
    If L > 1000 And L < (ScaleWidth - 2000) Then
        picSep.Left = L
    End If
End If
End Sub

Private Sub ImgSep_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
    picSep.Visible = False
    List1.Width = (picSep.Left - List1.Left)
    Form_Resize
End If
End Sub

Private Sub List1_Click()
lblTitle.Caption = List1.List(List1.ListIndex)
Text1.Text = MyCode(GetIndex(List1.Text)).EntryValue
If Text1.Text = "" Then
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(3).Enabled = False
    Toolbar1.Buttons(1).Enabled = False
    Me.mnuCopy.Enabled = False
    Me.mnuCopyall.Enabled = False
    Me.mnuSelect.Enabled = False
Else
    Toolbar1.Buttons(2).Enabled = True
    Toolbar1.Buttons(3).Enabled = True
    Toolbar1.Buttons(1).Enabled = True
    Me.mnuCopy.Enabled = True
    Me.mnuCopyall.Enabled = True
    Me.mnuSelect.Enabled = True
End If
If MyCode(GetIndex(List1.Text)).EntryInfo = "" Then
    Toolbar1.Buttons(7).Enabled = False
    mnuInfo.Enabled = False
Else
    Toolbar1.Buttons(7).Enabled = True
    mnuInfo.Enabled = True
End If
If MyCode(GetIndex(List1.Text)).EntryPicName = "" Or Dir(Apath & "Pictures\" & MyCode(GetIndex(List1.Text)).EntryPicName) = "" Then
    Toolbar1.Buttons(8).Enabled = False
    mnuViewPic.Enabled = False
Else
    Toolbar1.Buttons(8).Enabled = True
    mnuViewPic.Enabled = True
End If
If Text1.Text <> "" And Toolbar1.Buttons(10).Value = tbrPressed Then
    ColorBox Text1
End If
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim maxlen As Single, tmp As Single, i As Integer
If KeyCode = vbKeySpace Then
    If List1.ListCount > 0 Then
        For i = 0 To List1.ListCount - 1
            tmp = Me.TextWidth(List1.List(i))
            If tmp > maxlen Then maxlen = tmp
            tmp = 0
        Next i
    End If
    List1.Width = maxlen + 160
    Form_Resize
    KeyCode = 0
End If
End Sub


Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuCopy_Click()
If Text1.SelText <> "" Then
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
End If
End Sub

Private Sub mnuCopyall_Click()
If Text1.Text <> "" Then
    Clipboard.Clear
    Clipboard.SetText Text1.Text
End If
End Sub

Private Sub mnuCut_Click()
If Text1.SelText <> "" Then
    Clipboard.Clear
    Clipboard.SetText Text1.SelText
    Text1.SelText = ""
End If
End Sub

Private Sub mnuDelete_Click()
Text1.SelText = ""
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuInfo_Click()
If MyCode(GetIndex(List1.Text)).EntryInfo = "" Then Exit Sub
Dim FM As frmView
Set FM = frmView
InfoBox = True
Load FM
FM.Caption = "Information"
FM.Text1.BackColor = vbButtonFace '&HC0FFFF
'FM.Text1.ForeColor = &HC00000
FM.Text1.Font.Size = 10
FM.Text1.Font.Name = "Courier New"
FM.Text1.Text = MyCode(GetIndex(List1.Text)).EntryInfo
FM.Show , Me
End Sub

Private Sub mnuPaste_Click()
If Clipboard.GetText <> "" Then
    Text1.SelText = Clipboard.GetText
End If
End Sub

Private Sub mnuSearch_Click()
frmSearch.Show , Me
End Sub

Private Sub mnuSelect_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SetFocus
End Sub

Private Sub mnuSynCol_Click()
If mnuSynCol.Checked = True Then
    Me.Text1.SelStart = 0
    Me.Text1.SelLength = Len(Me.Text1.Text)
    Me.Text1.SelColor = vbBlack
    Me.Text1.SelLength = 0
    mnuSynCol.Checked = False
    Toolbar1.Buttons(10).Value = tbrUnpressed
Else
    mnuSynCol.Checked = True
    Toolbar1.Buttons(10).Value = tbrPressed
    ColorBox Me.Text1
End If
End Sub

Private Sub mnuViewPic_Click()
On Error GoTo XS
If Dir(Apath & "Pictures\" & MyCode(GetIndex(List1.Text)).EntryPicName) <> "" Then
    Load frmPic
    'frmPic.PaintPicture LoadPicture(Apath & "Pictures\" & MyCode(List1.ListIndex).EntryPicName), 0, 0
    frmPic.Picture1.Picture = LoadPicture(Apath & "Pictures\" & MyCode(GetIndex(List1.Text)).EntryPicName)
    frmPic.Width = frmPic.Picture1.Width + (frmPic.Width - frmPic.ScaleWidth)
    frmPic.Height = frmPic.Picture1.Height + (frmPic.Height - frmPic.ScaleHeight)
    frmPic.Show , Me
End If
XS:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "Select":   mnuSelect_Click
    Case "Cut":      mnuCut_Click
    Case "Copy":     mnuCopy_Click
    Case "Copyall":  mnuCopyall_Click
    Case "Paste":    mnuPaste_Click
    Case "Search":   mnuSearch_Click
    Case "Info":     mnuInfo_Click
    Case "Pic":      mnuViewPic_Click
    Case "Color":    mnuSynCol_Click
End Select

End Sub

