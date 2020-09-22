VERSION 5.00
Begin VB.Form frmPic 
   AutoRedraw      =   -1  'True
   Caption         =   "Picture Preview"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9450
   Icon            =   "frmPic.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6795
   ScaleWidth      =   9450
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private F5Down As Boolean

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 And Shift = 0 Then
    F5Down = True
End If
End Sub

Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 And Shift = 0 And F5Down Then
    Unload Me
End If
F5Down = False
End Sub

