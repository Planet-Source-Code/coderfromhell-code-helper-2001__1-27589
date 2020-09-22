Attribute VB_Name = "ModFile"
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long

'GetSystemFolders
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public DataFile As String
Public Apath As String
Public MyCode() As MCD

Public OldSearches As New Collection

Public Type MCD
    EntryName As String
    EntryValue As String ' This is the text
    SearchTags() As String
    EntryInfo As String
    EntryPicName As String
End Type

Public Enum SystemDirs
    dirWindows = 0
    dirSystem = 1
    dirTemp = 2
End Enum

Public Sub SetInfo()
DataFile = "Win32Code2001.CDE"
Apath = App.Path
If Right(Apath, 1) <> "\" Then Apath = Apath & "\"
End Sub

Function FileExist(File As String) As Boolean
If Dir(File) <> "" Then
    FileExist = True
Else
    FileExist = False
End If
End Function

'Public Function SaveFile() As Boolean

'End Function

Public Function OpenFile() As Boolean

On Local Error GoTo XF

Dim CurLine As String, i As Integer, j As Integer
Dim TempVal As Integer, tmpArray() As String, TmpStg As String

If FileExist(Apath & DataFile) = True Then
    If GetAttr(Apath & DataFile) = vbReadOnly Then
        SetAttr Apath & DataFile, vbNormal
    End If
    
    'ReDim MyCode(0)
    
    Close #1
    Open Apath & DataFile For Input As #1
        Do Until EOF(1)
            Line Input #1, CurLine
            If Trim(CurLine) <> "" And Left(CurLine, 2) <> "//" Then
                If Left(MyTrim(CurLine), 15) = "<entry number>=" Then
                    ReDim Preserve MyCode(i)
                    MyCode(i).EntryName = ""
                    MyCode(i).EntryValue = ""
                    i = i + 1
                ElseIf Left(MyTrim(CurLine), 13) = "<entry name>=" Then
                    MyCode(i - 1).EntryName = Mid(MyTrim(CurLine, False), 14)
                ElseIf Left(MyTrim(CurLine), 6) = "<txt>=" Then
                    MyCode(i - 1).EntryValue = MyCode(i - 1).EntryValue & Mid(MyTrim(CurLine, False), 7)
                ElseIf Left(MyTrim(CurLine), 6) = "<ret>=" Then
                    TempVal = CInt(Mid(MyTrim(CurLine), 7))
                    If TempVal <> 0 Then
                        For j = 1 To TempVal
                            MyCode(i - 1).EntryValue = MyCode(i - 1).EntryValue & vbCrLf
                        Next j
                    End If
                ElseIf Left(MyTrim(CurLine), 14) = "<search tags>=" Then
                    If Len(MyTrim(CurLine)) > 15 Then
                        Call Split(Mid(MyTrim(CurLine, False), 15), " ", MyCode(i - 1).SearchTags(), 0)
                    End If
                ElseIf Left(MyTrim(CurLine), 7) = "<info>=" Then
                    MyCode(i - 1).EntryInfo = MyCode(i - 1).EntryInfo & Mid(CurLine, 8)
                ElseIf Left(MyTrim(CurLine), 7) = "<reti>=" Then
                    TempVal = CInt(Mid(MyTrim(CurLine), 8))
                    If TempVal <> 0 Then
                        For j = 1 To TempVal
                            MyCode(i - 1).EntryInfo = MyCode(i - 1).EntryInfo & vbCrLf
                        Next j
                    End If
                ElseIf Left(MyTrim(CurLine), 11) = "<pic name>=" Then
                    If Len(CurLine) > 11 Then
                        MyCode(i - 1).EntryPicName = MyTrim(Mid(CurLine, 12), False)
                    End If
                End If
            End If
            On Local Error Resume Next
            If i <> 0 Then
                TmpStg = MyCode(i - 1).SearchTags(0)
                If Err.Number = 9 Then
                    ReDim Preserve MyCode(i - 1).SearchTags(0)
                    Err.Clear
                End If
                TmpStg = ""
                On Local Error GoTo XF
            End If
            DoEvents
        Loop
    Close #1
    
End If
Exit Function
XF:
MsgBox Err.Number & vbTab & Err.Description, vbCritical, "Error"
End Function



Function MyTrim(Stg As String, Optional mkLowerCase As Boolean = True) As String
Dim Strga As String
Strga = Trim(Stg)
If mkLowerCase = True Then Strga = LCase(Strga)
Do Until Left(Strga, 1) <> vbTab
    'If Left(Strga, 1) = vbTab Then
        Strga = Mid(Strga, 2)
    'End If
Loop
MyTrim = Strga
End Function



Public Function Split(chaine As String, Separator As String, destArray() As String, ArrayStartNum As Integer) As Integer
    On Error GoTo erreur
    Dim pos_act As Integer, pos_occur As Integer
    If Right(chaine, 1) <> Separator Then chaine = chaine & Separator
    Do
        pos_act = pos_occur + Len(Separator)
        pos_occur = InStr(pos_act, chaine, Separator)
        If pos_occur <> 0 Then
            ReDim Preserve destArray(ArrayStartNum)
            destArray(ArrayStartNum) = Mid(chaine, pos_act, pos_occur - pos_act)
            ArrayStartNum = ArrayStartNum + 1
        End If
    Loop Until pos_occur = 0
    Split = 0
Exit Function

erreur:
    Split = Err.Number
End Function

Sub OpenNotePad()
Dim TmpStg As String

    TmpStg = GetSystemFolders(dirWindows)
    If Right(TmpStg, 1) <> "\" Then TmpStg = TmpStg & "\"
    TmpStg = TmpStg & "Notepad.exe"
    Call Shell(TmpStg, vbNormalFocus)

End Sub

Public Sub SendText(WindowHandle As Long, Text As String)
    ' Sends the Text to the given window han
    '     dle
    Dim ReturnValue As Long
    ReturnValue = SendMessageByString(WindowHandle, WM_SETTEXT, 0&, Text)
End Sub

Public Function GetSystemFolders(func As SystemDirs)
Dim r, nSize As Long, tmp As String
 tmp = Space$(256):    nSize = Len(tmp)
Select Case func
   Case 0
      r = GetWindowsDirectory(tmp, nSize):     GetSystemFolders = TrimNull(tmp)
    Case 1
      r = GetSystemDirectory(tmp, nSize):      GetSystemFolders = TrimNull(tmp)
    Case 2
       r = GetTempPath(nSize, tmp):       GetSystemFolders = TrimNull(tmp)
    End Select
End Function

Private Function TrimNull(Item As String)
    Dim pos As Integer:    pos = InStr(Item, Chr$(0))
    If pos Then
          TrimNull = Left$(Item, pos - 1)
    Else: TrimNull = Item
    End If
End Function

Function GetIndex(cName As String) As Integer
    On Error GoTo XF
    For i = 0 To UBound(MyCode)
        If cName = MyCode(i).EntryName Then
            GetIndex = i
            Exit Function
        End If
    Next i
XF:
End Function

