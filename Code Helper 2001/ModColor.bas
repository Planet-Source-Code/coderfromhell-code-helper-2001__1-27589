Attribute VB_Name = "ModColor"
Private Const NormalColor = vbBlack
Private Const CommentColor = &H8000&
Private Const KeywordColor = &HC00000   '&H800000


Private Keywords As String

'___________________________________________________________________________

Private Declare Function apiLockWindowUpdate Lib "user32" _
                         Alias "LockWindowUpdate" _
                        (ByVal hwndLock As Long) As Long
Private Declare Function apiSendMessage Lib "user32" _
                         Alias "SendMessageA" _
                        (ByVal hwnd As Long, _
                         ByVal wMsg As Long, _
                         ByVal wParam As Long, _
                         lParam As Any) As Long

Private Const WM_USER = &H400
Private Const EM_GETTEXTRANGE = (WM_USER + 75)
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_GETFIRSTVISIBLELINE = &HCE
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_UNDO = &HC7
Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Public InfoBox As Boolean


Public Sub ColorBox(Box As RichTextBox)
Dim Stg As String
Dim Word As String
Dim CurLine As Long
Dim LineCount As Long
Dim i As Long, j As Long
Dim wrdStart As Long, wrdEnd As Long
Dim tmp As Integer, tmpa As Long, TempStg As String
    
    LockWindowUpdate Box.hwnd
    
    Stg = Box.Text
    If MyTrim(Stg, False) = "" Then Exit Sub
    'LineCount = GetLineCount(Box)
    
    Stg = ""
        
    With Box
        
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelColor = NormalColor
        .SelStart = 0
        .SelLength = 0
        
        'Get the first word
        wrdStart = 1
        .Text = .Text & vbNewLine
        CurLine = 1
        j = 1

        Do Until wrdStart >= Len(.Text)
            Do Until j = Len(.Text)
                If j >= Len(.Text) Then wrdEnd = Len(.Text) + 1: Exit Do
                tmp = Asc(UCase(Mid(.Text, j, 1)))
                If tmp >= 65 And tmp <= 90 Then ' its a letter
                    If Stg = "" Then wrdStart = j ' start of word
                    Stg = Stg & Mid(.Text, j, 1)
                ElseIf tmp = 32 Or tmp = 41 Then ' "32 = Space"     "41 = )"
                    wrdEnd = j - 1
                    Exit Do
                ElseIf tmp = Asc(vbCr) Then
                    CurLine = CurLine + 1
                    wrdEnd = j - 1
                    Exit Do
                ElseIf tmp = Asc("'") And Stg = "" Or Left(Stg, 1) = "'" Then ' its a comment
                    If Stg = "" Then
                        wrdStart = j
                    Else
                        wrdStart = j - Len(Stg)
                    End If
                    wrdEnd = InStr(j, .Text, vbCr)
                    j = wrdEnd + 1
                    .SelStart = wrdStart - 1
                    .SelLength = wrdEnd - (wrdStart - 1)
                    .SelColor = CommentColor
                    .SelStart = j
                    .SelLength = 0
                    .SelColor = NormalColor
                End If
                j = j + 1
            Loop
            
            If wrdStart > wrdEnd Then wrdEnd = Len(.Text)
            
            'If LCase(Stg) = "public" Or LCase(Stg) = "private" Then
            tmpa = InStr(1, Keywords, "|" & Stg & "|", vbTextCompare)
            If tmpa <> 0 Then
                .SelStart = wrdStart - 1
                .SelLength = wrdEnd - (wrdStart - 1)
                TempStg = Mid(Keywords, tmpa + 1, Len(Stg))
                If TempStg <> Stg Then
                    .SelText = TempStg
                    .SelStart = wrdStart - 1
                    .SelLength = wrdEnd - (wrdStart - 1)
                End If
                .SelColor = KeywordColor
                .SelStart = wrdEnd
                .SelLength = 1
                .SelColor = NormalColor
            End If
            'Debug.Print Stg
            
            wrdStart = wrdEnd + 1
            Stg = ""
            j = j + 1
        Loop
        
        .SelStart = 0
        .SelLength = 0
        .SelColor = NormalColor
        '.Text = Mid(.Text, 1, Len(.Text) - 1)
    End With
    
    LockWindowUpdate 0
    
End Sub

Private Function GetLineCount(Box As RichTextBox) As Long

    GetLineCount = apiSendMessage(Box.hwnd, EM_GETLINECOUNT, 0&, 0&)

End Function


Function GetLineOfText(Stg As String, LineNo As Long) As String
    Dim SP As Long, EP As Long, LL As Long
    
    
    
    
    
    SP = Text1.SelStart
    
    On Local Error Resume Next
    
    If SP = 0 Then Exit Function
    'If SP = Len(Text1.Text) Then Exit Function
    Do Until Mid(Text1.Text, SP - 1, 1) = vbCrLf
        'Debug.Print """" & Mid(Text1.Text, SP - 1, 1) & """"
        If Mid(Text1.Text, SP - 1, 1) = vbCr Then Exit Do
        If Mid(Text1.Text, SP - 1, 1) = vbLf Then Exit Do
        If Mid(Text1.Text, SP - 1, 1) = vbCrLf Then Exit Do
        SP = SP - 1
        If SP = 0 Or SP = 1 Then Exit Do
    Loop
    EP = SP
    Do Until EP = Len(Text1.Text) 'Or Mid(Text1.Text, EP, 1) = vbCr
        EP = EP + 1
    Loop
    'EP = EP + 2
        If Mid(Text1.Text, SP, EP - (SP - 1)) = vbCr Then Exit Function
        If Mid(Text1.Text, SP, EP - (SP - 1)) = vbLf Then Exit Function
        If Mid(Text1.Text, SP, EP - (SP - 1)) = vbCrLf Then Exit Function
    
    GetLineOfText = """" & Mid(Text1.Text, SP, EP - (SP - 1)) & """"
End Function

'*****************************************************************************

Public Sub Colorize(RTFBox As RichTextBox, CommentColor, _
                    StringColor, KeysColor, KeyCode)

'// SETUP LOCAL VARIABLES                                   //
    Dim lTextSelPos As Long
    Dim lTextSelLen As Long
    Dim thisLine As Integer
    Dim cStart As Integer
    Dim cEnd As Integer
    Dim i As Long
    Dim sBuffer As String
    Dim lBufferLen As Long
    Dim lSelPos As Long
    Dim lSelLen As Long
    Dim sTempBuffer As String
    Dim sSearchChar As String
    Dim lSearchCharLen As Long
    Dim StartText As Integer
    Dim RepText As String
    'Dim BLD As Variant
    
    
'// HANDLE ERRORS                                           //
    On Error GoTo ErrHandler

    RTFBox.SelStart = 0
    RTFBox.SelLength = Len(RTFBox.Text)
    RTFBox.SelColor = vbBlack
    RTFBox.SelStart = 0
    RTFBox.SelLength = 0

'// Save the carot position                                //
    lTextSelPos = RTFBox.SelStart
    lTextSelLen = RTFBox.SelLength
    'BLD = RTFBox.SelBold

'// MAKE ENTIRE TEXT BLACK (OR STRINGCOLOR DEFINED)         //
'// SO WHEN USER CHANGES A KEYWORD TO A NON-KEYWORD,        //
'// IT WILL NOT STILL BE BLUE (OR DEFINED KEYWORD COLOR)    //
    With RTFBox
        
    '// ONLY CHANGE CHARS ON THIS LINE                      //
        cStart% = .SelStart     ' CURRENT POSITION OF CURSOR//
        cEnd% = .SelStart       ' AGAIN CURR POS OF CURSOR  //
        
    '// SET "thisLine%" TO THE LINE CURSOR IS ON            //
        thisLine = .GetLineFromChar(.SelStart)
        
    '// IF ENTER WAS KEY PRESSED, COLORIZE LINE ABOVE SINCE //
    '// OUR LOCATION IS NOW THE NEXT LINE                   //
    '// (KEYCODE 13 = [ENTER] KEY)                          //
        If KeyCode = 13 Then
            thisLine = thisLine - 1
            cStart% = cStart% - 1
        End If
        
    '// DETERMINE "cStart" or STARTING CHARACTER TO         //
    '// EVALUATE FOR COLORIZATION PROCESS                   //
    '// WE ARE GOING TO DO THIS BY COUNTING FROM THE        //
    '// CURRENT CURSOR POSITION BACKWARDS TO BEGINNING OF   //
    '// THE CURRENT LINE , OR TO THE BEGINNING OF THE FILE, //
    '// WHICHEVER COMES FIRST                               //
        Do Until .GetLineFromChar(cStart%) <> thisLine
            cStart% = cStart% - 1
            If cStart% < 0 Then
                cStart% = 0
                Exit Do
            End If
        Loop
    '// NOW WE ARE GOING TO DETERMINE THE "cEnd" OR ENDING  //
    '// CHARACTER OF OUR EVALUATION STRING TO COLORIZE.     //
    '// WE DO THIS BY COUNTING FROM CURSOR POSITION TO THE  //
    '// END OF CURRENT LINE OR END OF THE FILE, WHICHEVER   //
    '// COMES FIRST.                                        //
    '//                                                     //
    '// THIS ROUTINE IS NECESSARY SINCE WE MAY BE INSERING  //
    '// TEXT IN THE MIDDLE OF A LINE, BUT WE STILL WANT TO  //
    '// EVALUATE ENTIRE LINE IN CASE WE ARE CHANGING A      //
    '// KEYWORD TO A NON-KEYWORD, OR VICE VERSA             //
        Do Until .GetLineFromChar(cEnd%) <> thisLine
            cEnd% = cEnd% + 1
            If cEnd% > Len(.Text) Then
                cEnd = Len(.Text)
                Exit Do
            End If
        Loop
        
        'cEnd% = Len(.Text)
        
    '// SET COLOR OF TEXT WE ARE WORKING WITH BACK TO       //
    '// ORIGINAL COLOR FOR NOW SINCE IT MAY BE A KEYWORK    //
    '// THAT IS GETTING CHANGED TO A NON-KEYWORD. THE       //
    '// NEXT ROUTINE WILL COLORIZE IT IF IT FINDS KEYWORDS  //
        .SelStart = cStart%
        .SelLength = cEnd% - cStart%
        .SelColor = StringColor
        .SelLength = 0
        
    End With
    


'// BEGIN EVALUAINTG AND CHANGING COLORS OF WORDS           //
    With RTFBox
    '// INSURE "WHOLE WORDS" ARE COLORIZED AND NOT          //
    '// PARTIAL WORDS (EG: the word "If", and not "Gift"    //
    '// WHERE "IF" EXISTS IN "GIFT"                         //
        sBuffer = .Text & " "
        lBufferLen = Len(sBuffer)
        sTempBuffer = ""
        If cStart = 0 Then cStart = 1
        
    '// LOOP THROUGH CHARACTERS USING RANGE WE DEFINED      //
    '// EARLIER IN THIS SUB                                 //
        For i = cStart% To cEnd%
        
        Select Case Asc(Mid(sBuffer, i, 1))
        
    '// COMMENTS - ENTIRE LINE IS COLORIZED REGARDLESS OF   //
    '// CONTENT. COMMENT PREFIXES ARE HARD-CODED HERE, BUT  //
    '// YOU CAN MODIFY/ADD/REMOVE FROM HERE IF YOU WANT     //
    '// BY FIRST INCLUDING PREFIX ASC CODE IN THIS "CASE"   //
    '// STATEMENT, AND THEN WRITING AN ElseIf EVALUATION    //
    '// AGAINST THE CHARACTER(S) THAT MAKE UP YOUR REMARK   //
    '// PREFIX                                              //
    '// (CHR$(47) = "/", AND CHR$(39) = "'"                 //
        Case 47, 39
          
    '// C/C++ STYLE COMMENT                                 //
            If Mid(sBuffer, i, 2) = "//" Then
        '// COLORIZE FROM PREFIX TO "sSearchChar"           //
                sSearchChar = vbCrLf
                lSearchCharLen = 0
    '// VISUAL BASIC STYLE COMMENT                          //
            ElseIf Mid(sBuffer, i, 1) = "'" Then
        '// COLORIZE FROM PREFIX TO "sSearchChar"           //
                sSearchChar = vbCrLf
                lSearchCharLen = 0
    '// IF NOT A COMMENT, GOTO THE "EXITCOMMENT" ROUTINE    //
    '// TO BYPASS COMMENT COLORIZATION ROUTINES             //
            Else
                GoTo ExitComment
            End If
          
    '// SET TEMPBUFFER (sTempBuffer" to NOTHING             //
            sTempBuffer = ""
          
    '// COLORIZE THE COMMENT STRING                         //
    '// "i" IS CURRENT COUNT OF LOOP                        //
            .SelStart = i - 1
            lSelLen = InStr(i, sBuffer, sSearchChar) _
                    + lSearchCharLen
                
            If lSelLen <> lSearchCharLen Then   '// FileEnd?//
                lSelLen = lSelLen - i
            Else
                lSelLen = lBufferLen - i
            End If
                
            .SelLength = lSelLen
            .SelColor = CommentColor
            i = .SelStart + .SelLength
          
'// QUOTE COLORIZE ROUTINE                                  //
ExitComment:
        Case 34
            If Mid(sBuffer, i, 1) = Chr$(34) Then
        '// COLORIZE FROM PREFIX TO "sSearchChar"           //
                sSearchChar = Chr$(34)
                lSearchCharLen = 0
            Else
                GoTo ExitQuote
            End If
          
    '// SET TEMPBUFFER (sTempBuffer" to NOTHING             //
            sTempBuffer = ""
          
    '// COLORIZE THE QUOTE STRING                           //
    '// "i" IS CURRENT COUNT OF LOOP                        //
            .SelStart = i - 1
            lSelLen = InStr(i + 1, sBuffer, sSearchChar) _
                    + lSearchCharLen
                
            If lSelLen <> lSearchCharLen Then   '// FileEnd?//
                lSelLen = lSelLen - i
            ElseIf lSelLen < 1 Then
            '// SET CUR POSITION BACK AND DONT COLORIZE     //
            '// ANYTHING SINCE "END QUOTE" HAS NOT BEEN     //
            '// ENTERED YET                                 //
                GoTo ErrHandler
            Else
                lSelLen = lBufferLen - i
            End If
                
            .SelLength = lSelLen
            .SelColor = StringColor
            i = .SelStart + .SelLength

'// COLORIZE KEYWORDS ROUTINE                               //
ExitQuote:

        '// THE FOLLOWING "CASE" STATEMENT SETS THESE       //
        '// CHARACTERS AS VALID PARTS OF A COLORIZATION     //
        '// STRING. IN OTHER WORDS, ANY KEYWORDS YOU DEFINE //
        '// CAN HAVE THESE ASCII CHARACTERS, AND IF THE     //
        '// DONT, THEY WILL NOT QUALIFY.                    //
        '// EXAMPLE: IF YOUR KEYWORD IS SOMETHING LIKE      //
        '// THIS:  My_ROUTINE (with the underscore), YOU    //
        '// NEED TO MAKE SURE THIS CASE STATEMENT INCLUDES  //
        '// THE ASCII CODE FOR THAT CHARACTER (UNDERSCORE)  //
        '// AS WELL AS ALL ALPHANUMERICK CHARACTERS         //
        '// ASCII 33 = "!", 35 to 38 = #, $, %, &           //
        '// ASCII 46 = . (dot)                              //
        '// ASCII 60 = "<" and 62 = ">"                     //
        '// ASCII 49 to 57 = Numbers 1,2,3,4,5,6,7,8,9,0    //
        '// ASCII 97 to 122 = lowercase a to z              //
        '// ASCII 65 to 90 = UPPERCASE A to Z               //
             Case 33, 35 To 38, 46, 60, 62, _
                  49 To 57, 97 To 122, 65 To 90
                  
                If sTempBuffer = "" Then lSelPos = i
                sTempBuffer = sTempBuffer & Mid(sBuffer, i, 1)
             
             Case Else
                
                If Trim(sTempBuffer) <> "" Then
                    .SelStart = lSelPos - 1
                    .SelLength = Len(sTempBuffer)
                    StartText% = InStr(1, Keywords, _
                                 "|" & sTempBuffer & "|", 1)
                    If StartText% <> 0 Then
                '// ALTER COLOR                             //
                        .SelColor = KeysColor
                        
                
                '// CHANGE FOUND MATCH TO BE THE SAME       //
                '// CASE AS WORD IN LIBRARY                 //
                '// (EG: "print" would change to "Print"    //
                '//  with the CAPITAL "P")                  //
                        RepText$ = _
                        Mid$(Keywords, StartText% + 1, _
                        Len(sTempBuffer))
                        .SelText = RepText$
                    End If
                
                End If
                
                sTempBuffer = ""
        
        
        End Select
      
        Next

        End With

ErrHandler:

    '// Set the Cursor to the old position                  //
    RTFBox.SelStart = lTextSelPos
    RTFBox.SelLength = lTextSelLen
    'RTFBox.SelBold = BLD

End Sub

Public Sub SetScriptKeywords()


    Keywords = "|" & "If" & "|" & "Then" & "|" & "Else" & "|" & "Sub" & "|" & "End" & "|" & "Goto" & "|" & "Do" & "|" & "Loop" & "|" & _
    "Save" & "|" & "Dim" & "|" & "Public" & "|" & "Private" & "|" & "Function" & "|" & "Get" & "|" & "Set" & "|" & "Let" & "|" & _
    "For" & "|" & "Next" & "|" & "To" & "|" & "Step" & "|" & "ElseIf" & "|" & "Print" & "|" & "And" & "|" & "Exit" & "|" & _
    "Select" & "|" & "Case" & "|" & "As" & "|" & "Integer" & "|" & "Long" & "|" & "Single" & "|" & "Double" & "|" & _
    "Variant" & "|" & "Date" & "|" & "Optional" & "|" & "Option" & "|" & "Explicit" & "|" & "Declare" & "|" & "Type" & "|" & "Object" & "|" & "Const" & "|" & _
    "Boolean" & "|" & "String" & "|" & "With" & "|" & "Lib" & "|" & "Alias" & "|" & "ByVal" & "|" & "ReDim" & "|" & "Lbound" & "|" & "Ubound" & "|" & "Friend" & _
    "|" & "Static" & "|" & "Property" & "|" & "Get" & "|" & "Let" & "|" & "Set" & "|" & "Cstr" & "|" & "Cint" & "|" & "Csng" & "|" & _
    "Error" & "|" & "Open" & "|" & "Close" & "|" & "Input" & "|" & "Output" & "|" & "Preserve" & "|"


End Sub

