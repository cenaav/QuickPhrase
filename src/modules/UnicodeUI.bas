Attribute VB_Name = "UnicodeUI"
'===============================================================================
' UnicodeUI - message boxes that can display Persian, Arabic and emoji.
'
' VBA's built-in MsgBox converts its text through the system ANSI codepage, so
' any character outside that codepage arrives as a question mark. On an English
' Windows install "Delete سلام?" renders as "Delete ??????".
'
' The Windows API has always had a Unicode entry point, so this module calls
' MessageBoxW directly and passes UTF-16 string pointers. Return values match
' the VbMsgBoxResult constants exactly (IDOK = 1, IDCANCEL = 2, IDYES = 6,
' IDNO = 7), so call sites read the same as before.
'
' Use MsgBoxW anywhere a message can contain text the user typed.
'===============================================================================
Option Explicit

' VBA7 (Office 2010+) needs PtrSafe and LongPtr; Office 2007 has neither, and
' rejects the file if it sees them. Conditional compilation keeps one source
' file working on both.
#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
        ByVal hWnd As LongPtr, _
        ByVal lpText As LongPtr, _
        ByVal lpCaption As LongPtr, _
        ByVal uType As Long) As Long
#Else
    Private Declare Function MessageBoxW Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal lpText As Long, _
        ByVal lpCaption As Long, _
        ByVal uType As Long) As Long
#End If

' Disables the thread's other windows while the box is up, and brings it to the
' front. Without this, passing no owner window can leave the box behind Word.
Private Const MB_TASKMODAL As Long = &H2000
Private Const MB_SETFOREGROUND As Long = &H10000

' Unicode-safe replacement for MsgBox.
'
' Falls back to the built-in MsgBox if the API call fails, so a message is
' always shown even if it loses characters along the way.
Public Function MsgBoxW(ByVal prompt As String, _
                        Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
                        Optional ByVal title As String = "QuickPhrase") As VbMsgBoxResult
    Dim result As Long

    On Error GoTo Fallback

    result = MessageBoxW(0, StrPtr(prompt), StrPtr(title), _
                         CLng(buttons) Or MB_TASKMODAL Or MB_SETFOREGROUND)

    If result = 0 Then GoTo Fallback

    MsgBoxW = result
    Exit Function

Fallback:
    MsgBoxW = MsgBox(prompt, buttons, title)
End Function
