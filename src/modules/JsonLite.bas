Attribute VB_Name = "JsonLite"
'===============================================================================
' JsonLite - minimal JSON reader/writer for QuickPhrase.
'
' Deliberately NOT a general purpose JSON library. It understands exactly the
' one shape QuickPhrase stores:
'
'   [ { "label": "...", "text": "..." }, ... ]
'
' Unknown keys are skipped, so the format can grow later without breaking
' older builds. Everything is done with plain VBA strings so the add-in has
' no external references beyond ADODB (used in SnippetStore for UTF-8 I/O).
'===============================================================================
Option Explicit

Private mJson As String
Private mPos As Long
Private mLen As Long

'--- Public API ----------------------------------------------------------------

' Parses a QuickPhrase JSON document into two parallel arrays.
' Returns True on success. On failure returns False and fills errMsg.
Public Function ParsePhrases(ByVal json As String, _
                             ByRef labels() As String, _
                             ByRef texts() As String, _
                             ByRef newlines() As Boolean, _
                             ByRef count As Long, _
                             ByRef errMsg As String) As Boolean
    On Error GoTo Fail

    mJson = json
    mLen = Len(json)
    mPos = 1
    count = 0
    ReDim labels(0 To 0)
    ReDim texts(0 To 0)
    ReDim newlines(0 To 0)

    SkipWhitespace
    If Peek() <> "[" Then
        errMsg = "Expected '[' at position " & mPos & "."
        ParsePhrases = False
        Exit Function
    End If
    mPos = mPos + 1

    SkipWhitespace
    If Peek() = "]" Then
        ' Empty list is valid.
        ParsePhrases = True
        Exit Function
    End If

    Dim capacity As Long
    capacity = 16
    ReDim labels(0 To capacity - 1)
    ReDim texts(0 To capacity - 1)
    ReDim newlines(0 To capacity - 1)

    Do
        SkipWhitespace
        If Peek() <> "{" Then
            errMsg = "Expected '{' at position " & mPos & "."
            ParsePhrases = False
            Exit Function
        End If
        mPos = mPos + 1

        Dim lbl As String, txt As String, nl As Boolean
        lbl = ""
        txt = ""
        nl = False

        SkipWhitespace
        If Peek() <> "}" Then
            Do
                SkipWhitespace
                If Peek() <> """" Then
                    errMsg = "Expected key string at position " & mPos & "."
                    ParsePhrases = False
                    Exit Function
                End If

                Dim key As String
                key = ReadString()

                SkipWhitespace
                If Peek() <> ":" Then
                    errMsg = "Expected ':' at position " & mPos & "."
                    ParsePhrases = False
                    Exit Function
                End If
                mPos = mPos + 1

                SkipWhitespace
                Select Case LCase$(key)
                    Case "label":   lbl = ReadString()
                    Case "text":    txt = ReadString()
                    Case "newline": nl = ReadBool()
                    Case Else:      SkipValue
                End Select

                SkipWhitespace
                If Peek() = "," Then
                    mPos = mPos + 1
                ElseIf Peek() = "}" Then
                    Exit Do
                Else
                    errMsg = "Expected ',' or '}' at position " & mPos & "."
                    ParsePhrases = False
                    Exit Function
                End If
            Loop
        End If
        mPos = mPos + 1 ' consume '}'

        ' A phrase with no label is useless on a button; fall back to its text.
        If Len(Trim$(lbl)) = 0 Then lbl = Left$(Replace(txt, vbLf, " "), 24)

        If count >= capacity Then
            capacity = capacity * 2
            ReDim Preserve labels(0 To capacity - 1)
            ReDim Preserve texts(0 To capacity - 1)
            ReDim Preserve newlines(0 To capacity - 1)
        End If
        labels(count) = lbl
        texts(count) = txt
        newlines(count) = nl
        count = count + 1

        SkipWhitespace
        If Peek() = "," Then
            mPos = mPos + 1
        ElseIf Peek() = "]" Then
            Exit Do
        Else
            errMsg = "Expected ',' or ']' at position " & mPos & "."
            ParsePhrases = False
            Exit Function
        End If
    Loop

    ParsePhrases = True
    Exit Function

Fail:
    errMsg = "Parse error: " & Err.Description
    ParsePhrases = False
End Function

' Serialises parallel label/text arrays into pretty-printed JSON.
Public Function SerializePhrases(ByRef labels() As String, _
                                 ByRef texts() As String, _
                                 ByRef newlines() As Boolean, _
                                 ByVal count As Long) As String
    Dim sb As String
    Dim i As Long

    If count <= 0 Then
        SerializePhrases = "[]"
        Exit Function
    End If

    sb = "[" & vbCrLf
    For i = 0 To count - 1
        sb = sb & "  {" & vbCrLf
        sb = sb & "    ""label"": " & QuoteString(labels(i)) & "," & vbCrLf
        sb = sb & "    ""text"": " & QuoteString(texts(i)) & "," & vbCrLf
        sb = sb & "    ""newline"": " & LCase$(CStr(newlines(i))) & vbCrLf
        sb = sb & "  }"
        If i < count - 1 Then sb = sb & ","
        sb = sb & vbCrLf
    Next i
    sb = sb & "]" & vbCrLf

    SerializePhrases = sb
End Function

' Wraps a VBA string as a JSON string literal, escapes included.
Public Function QuoteString(ByVal s As String) As String
    Dim sb As String
    Dim i As Long
    Dim ch As String
    Dim code As Long

    sb = """"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536

        Select Case ch
            Case """": sb = sb & "\"""
            Case "\":  sb = sb & "\\"
            Case vbCr: sb = sb & "\r"
            Case vbLf: sb = sb & "\n"
            Case vbTab: sb = sb & "\t"
            Case Else
                If code < 32 Then
                    sb = sb & "\u" & Right$("000" & LCase$(Hex$(code)), 4)
                Else
                    ' Non-ASCII (Persian, Arabic, emoji, ...) is written as-is.
                    ' The file is UTF-8, so this stays readable in any editor.
                    sb = sb & ch
                End If
        End Select
    Next i

    QuoteString = sb & """"
End Function

'--- Scanner internals ---------------------------------------------------------

Private Function Peek() As String
    If mPos > mLen Then
        Peek = ""
    Else
        Peek = Mid$(mJson, mPos, 1)
    End If
End Function

Private Sub SkipWhitespace()
    Dim ch As String
    Do While mPos <= mLen
        ch = Mid$(mJson, mPos, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            mPos = mPos + 1
        Else
            Exit Do
        End If
    Loop
End Sub

' Reads a JSON string literal starting at the opening quote.
Private Function ReadString() As String
    Dim sb As String
    Dim ch As String

    If Peek() <> """" Then
        Err.Raise vbObjectError + 1, "JsonLite", _
                  "Expected '""' at position " & mPos & "."
    End If
    mPos = mPos + 1

    Do While mPos <= mLen
        ch = Mid$(mJson, mPos, 1)
        mPos = mPos + 1

        If ch = """" Then
            ReadString = sb
            Exit Function
        ElseIf ch = "\" Then
            Dim esc As String
            esc = Mid$(mJson, mPos, 1)
            mPos = mPos + 1
            Select Case esc
                Case """": sb = sb & """"
                Case "\":  sb = sb & "\"
                Case "/":  sb = sb & "/"
                Case "b":  sb = sb & Chr$(8)
                Case "f":  sb = sb & Chr$(12)
                Case "n":  sb = sb & vbLf
                Case "r":  sb = sb & vbCr
                Case "t":  sb = sb & vbTab
                Case "u"
                    sb = sb & ChrW$(CLng("&H" & Mid$(mJson, mPos, 4)))
                    mPos = mPos + 4
                Case Else
                    sb = sb & esc
            End Select
        Else
            sb = sb & ch
        End If
    Loop

    Err.Raise vbObjectError + 2, "JsonLite", "Unterminated string literal."
End Function

' Reads a JSON boolean. Anything else is skipped and reported as False, so a
' hand-edited file with "newline": 1 degrades rather than failing to load.
Private Function ReadBool() As Boolean
    SkipWhitespace

    If LCase$(Mid$(mJson, mPos, 4)) = "true" Then
        mPos = mPos + 4
        ReadBool = True
    ElseIf LCase$(Mid$(mJson, mPos, 5)) = "false" Then
        mPos = mPos + 5
        ReadBool = False
    Else
        SkipValue
        ReadBool = False
    End If
End Function

' Consumes and discards any value, so unknown keys never break parsing.
Private Sub SkipValue()
    Dim ch As String
    Dim depth As Long

    SkipWhitespace
    ch = Peek()

    If ch = """" Then
        ReadString
        Exit Sub
    End If

    If ch = "{" Or ch = "[" Then
        depth = 0
        Do While mPos <= mLen
            ch = Mid$(mJson, mPos, 1)
            If ch = """" Then
                ReadString
            Else
                mPos = mPos + 1
                If ch = "{" Or ch = "[" Then
                    depth = depth + 1
                ElseIf ch = "}" Or ch = "]" Then
                    depth = depth - 1
                    If depth = 0 Then Exit Sub
                End If
            End If
        Loop
        Exit Sub
    End If

    ' Bare literal: number, true, false, null.
    Do While mPos <= mLen
        ch = Mid$(mJson, mPos, 1)
        If ch = "," Or ch = "}" Or ch = "]" Or ch = " " Or _
           ch = vbTab Or ch = vbCr Or ch = vbLf Then Exit Do
        mPos = mPos + 1
    Loop
End Sub
