VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Results:"
      Enabled         =   0   'False
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
      Begin VB.CommandButton cmdCopy 
         Caption         =   "&Copy List"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdView 
         Caption         =   "&View"
         Height          =   375
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   1620
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox MC 
      Caption         =   "&Match Case"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "&Search keyword:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MinHeight = 1440
Private Const MaxHeight = 3870
Private CodeIndex() As Integer

Private Sub Check1_GotFocus()
cmdSearch.Default = True
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCopy_Click()
Dim Stg As String, i As Integer
For i = 0 To List1.ListCount - 1
    If i <> 0 Then Stg = Stg & vbCrLf
    Stg = Stg & List1.List(i)
Next i
If Stg <> "" Then
    Clipboard.Clear
    Clipboard.SetText Stg
End If
End Sub

Private Sub cmdSearch_Click()
Dim i As Integer, j As Integer, MyCount As Integer
If Combo1.Text = "" Then
    Beep
    MsgBox "You have to type a keyword!", vbInformation, "Info!"
    Exit Sub
End If
Me.Frame1.Enabled = True
'DoSearch
Combo1.Locked = True
List1.Clear
For i = 0 To Form1.List1.ListCount - 1
    If MC.Value = 0 Then
        For j = 0 To UBound(MyCode(i).SearchTags)
            If LCase(Combo1.Text) = LCase(MyCode(i).SearchTags(j)) Then
                List1.AddItem MyCode(i).EntryName
                ReDim Preserve CodeIndex(MyCount)
                CodeIndex(MyCount) = i
                MyCount = MyCount + 1
            End If
        Next j
    Else
        For j = 0 To UBound(MyCode(i).SearchTags)
            If Combo1.Text = MyCode(i).SearchTags(j) Then
                List1.AddItem MyCode(i).EntryName
                ReDim Preserve CodeIndex(MyCount)
                CodeIndex(MyCount) = i
                MyCount = MyCount + 1
            End If
        Next j
    End If
    DoEvents
Next i
'End Search
Me.Height = MaxHeight
If List1.ListCount = 0 Then
    cmdCopy.Enabled = False
    cmdView.Enabled = False
Else
    cmdCopy.Enabled = True
    cmdView.Enabled = True
End If
For j = 1 To OldSearches.count
    If Combo1.Text = OldSearches(j) Then
        GoTo XS
    End If
Next j
OldSearches.Add Combo1.Text
XS:
AddSearches
Combo1.Locked = False
End Sub

Private Sub cmdView_Click()
'Dim A As Long, H As Long
InfoBox = False
Load frmView
frmView.Text1.Text = MyCode(CodeIndex(List1.ListIndex)).EntryValue
ColorBox frmView.Text1
frmView.Show 1
'OpenNotePad
'DoEvents
'Sleep 100
'A = GetActiveWindow
'Sleep 500
'Beep
'H = GetWindowDC(A)
'SendText A, MyCode(CodeIndex(List1.ListIndex)).EntryValue
End Sub

Private Sub Combo1_GotFocus()
Combo1.SelStart = 0
Combo1.SelLength = Len(Combo1.Text)
cmdSearch.Default = True
End Sub

Private Sub Form_Load()
Me.Height = MinHeight
AddSearches
End Sub

Private Sub Form_Unload(Cancel As Integer)
SetActiveWindow Form1.hwnd
End Sub

Private Sub Label1_Click()
Combo1.SetFocus
End Sub

Private Sub List1_GotFocus()
cmdView.Default = True
End Sub

Sub AddSearches()
Dim i As Integer, Stg As String
Stg = Combo1.Text
Combo1.Clear
If OldSearches.count > 0 Then
    For i = 0 To OldSearches.count - 1
        Combo1.AddItem OldSearches(i + 1)
    Next i
End If
Combo1.Text = Stg
End Sub
