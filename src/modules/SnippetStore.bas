Attribute VB_Name = "SnippetStore"
'===============================================================================
' SnippetStore - owns the phrase list and its on-disk file.
'
' Phrases live in two parallel arrays kept in this module. The file is UTF-8
' JSON written without a BOM, so it opens cleanly in any editor and keeps
' Persian / Arabic / emoji intact.
'
'   %APPDATA%\QuickPhrase\snippets.json
'===============================================================================
Option Explicit

Private Const APP_FOLDER As String = "QuickPhrase"
Private Const STORE_FILE As String = "snippets.json"

' Where a space is added around an inserted phrase, so clicking two buttons in a
' row does not run the words together.
Public Enum QpSpaceMode
    qpSpaceNone = 0
    qpSpaceBefore = 1
    qpSpaceAfter = 2
End Enum

' The spacing preference is personal rather than part of a phrase set, so it
' lives in the registry instead of snippets.json - exporting phrases to a
' colleague should not change how their Word behaves.
Private Const SETTINGS_APP As String = "QuickPhrase"
Private Const SETTINGS_SECTION As String = "Options"

Private mLabels() As String
Private mTexts() As String
Private mNewlines() As Boolean
Private mCount As Long
Private mLoaded As Boolean

'--- Paths ---------------------------------------------------------------------

Public Function StoreFolder() As String
    StoreFolder = Environ$("APPDATA") & Application.PathSeparator & APP_FOLDER
End Function

Public Function StorePath() As String
    StorePath = StoreFolder() & Application.PathSeparator & STORE_FILE
End Function

'--- Settings ------------------------------------------------------------------

' Defaults to qpSpaceBefore: a leading space is what keeps a clicked phrase from
' sticking to the word already typed, which is the common case.
Public Property Get SpaceMode() As QpSpaceMode
    Dim raw As String

    On Error Resume Next
    raw = GetSetting(SETTINGS_APP, SETTINGS_SECTION, "SpaceMode", "")
    On Error GoTo 0

    Select Case LCase$(Trim$(raw))
        Case "none":  SpaceMode = qpSpaceNone
        Case "after": SpaceMode = qpSpaceAfter
        Case Else:    SpaceMode = qpSpaceBefore
    End Select
End Property

Public Property Let SpaceMode(ByVal value As QpSpaceMode)
    Dim raw As String

    Select Case value
        Case qpSpaceNone:  raw = "none"
        Case qpSpaceAfter: raw = "after"
        Case Else:         raw = "before"
    End Select

    On Error Resume Next
    SaveSetting SETTINGS_APP, SETTINGS_SECTION, "SpaceMode", raw
End Property

'--- Accessors -----------------------------------------------------------------

Public Property Get PhraseCount() As Long
    EnsureLoaded
    PhraseCount = mCount
End Property

' Index is 0-based.
Public Function PhraseLabel(ByVal index As Long) As String
    EnsureLoaded
    If index < 0 Or index >= mCount Then Exit Function
    PhraseLabel = mLabels(index)
End Function

Public Function PhraseText(ByVal index As Long) As String
    EnsureLoaded
    If index < 0 Or index >= mCount Then Exit Function
    PhraseText = mTexts(index)
End Function

' True when inserting this phrase should also end the paragraph.
Public Function PhraseNewline(ByVal index As Long) As Boolean
    EnsureLoaded
    If index < 0 Or index >= mCount Then Exit Function
    PhraseNewline = mNewlines(index)
End Function

'--- Mutations -----------------------------------------------------------------
' None of these write to disk; call SaveToDisk when the user commits.

Public Sub AddPhrase(ByVal label As String, ByVal text As String, _
                     Optional ByVal newline As Boolean = False)
    EnsureLoaded
    Grow mCount + 1
    mLabels(mCount) = label
    mTexts(mCount) = text
    mNewlines(mCount) = newline
    mCount = mCount + 1
End Sub

Public Sub UpdatePhrase(ByVal index As Long, ByVal label As String, _
                        ByVal text As String, _
                        Optional ByVal newline As Boolean = False)
    EnsureLoaded
    If index < 0 Or index >= mCount Then Exit Sub
    mLabels(index) = label
    mTexts(index) = text
    mNewlines(index) = newline
End Sub

Public Sub DeletePhrase(ByVal index As Long)
    EnsureLoaded
    If index < 0 Or index >= mCount Then Exit Sub

    Dim i As Long
    For i = index To mCount - 2
        mLabels(i) = mLabels(i + 1)
        mTexts(i) = mTexts(i + 1)
        mNewlines(i) = mNewlines(i + 1)
    Next i
    mCount = mCount - 1
End Sub

' Moves a phrase by delta (-1 up, +1 down). Returns the new index, or the
' original index if the move was not possible.
Public Function MovePhrase(ByVal index As Long, ByVal delta As Long) As Long
    EnsureLoaded
    MovePhrase = index

    Dim target As Long
    target = index + delta
    If index < 0 Or index >= mCount Then Exit Function
    If target < 0 Or target >= mCount Then Exit Function

    Dim tmpLabel As String, tmpText As String, tmpNewline As Boolean
    tmpLabel = mLabels(index): tmpText = mTexts(index): tmpNewline = mNewlines(index)
    mLabels(index) = mLabels(target): mTexts(index) = mTexts(target): mNewlines(index) = mNewlines(target)
    mLabels(target) = tmpLabel: mTexts(target) = tmpText: mNewlines(target) = tmpNewline

    MovePhrase = target
End Function

Public Sub ClearAll()
    mCount = 0
    mLoaded = True
End Sub

'--- Disk I/O ------------------------------------------------------------------

Public Sub EnsureLoaded()
    If Not mLoaded Then LoadFromDisk
End Sub

' Re-reads the file. Safe to call at any time; on a missing file it seeds the
' default starter phrases and writes them out.
Public Sub LoadFromDisk()
    mLoaded = True
    mCount = 0

    If Len(Dir$(StorePath())) = 0 Then
        SeedDefaults
        SaveToDisk
        Exit Sub
    End If

    Dim json As String
    json = ReadUtf8(StorePath())

    Dim errMsg As String
    If Not JsonLite.ParsePhrases(json, mLabels, mTexts, mNewlines, mCount, errMsg) Then
        mCount = 0
        UnicodeUI.MsgBoxW "QuickPhrase could not read your phrase file:" & vbCrLf & vbCrLf & _
               StorePath() & vbCrLf & vbCrLf & errMsg & vbCrLf & vbCrLf & _
               "The file was left untouched. Fix it, or use Manage Phrases " & _
               "to start a new list.", vbExclamation, "QuickPhrase"
    End If
End Sub

Public Function SaveToDisk() As Boolean
    On Error GoTo Fail

    EnsureFolder StoreFolder()
    WriteUtf8 StorePath(), JsonLite.SerializePhrases(mLabels, mTexts, mNewlines, mCount)
    SaveToDisk = True
    Exit Function

Fail:
    UnicodeUI.MsgBoxW "QuickPhrase could not save your phrases to:" & vbCrLf & vbCrLf & _
           StorePath() & vbCrLf & vbCrLf & Err.Description, _
           vbExclamation, "QuickPhrase"
    SaveToDisk = False
End Function

' Loads phrases from an arbitrary file. When replaceAll is False the imported
' phrases are appended to the current list.
Public Function ImportFrom(ByVal path As String, ByVal replaceAll As Boolean, _
                           ByRef errMsg As String) As Boolean
    On Error GoTo Fail

    Dim inLabels() As String, inTexts() As String, inCount As Long
    Dim inNewlines() As Boolean
    If Not JsonLite.ParsePhrases(ReadUtf8(path), inLabels, inTexts, inNewlines, inCount, errMsg) Then
        ImportFrom = False
        Exit Function
    End If

    EnsureLoaded
    If replaceAll Then mCount = 0

    Dim i As Long
    For i = 0 To inCount - 1
        AddPhrase inLabels(i), inTexts(i), inNewlines(i)
    Next i

    ImportFrom = SaveToDisk()
    Exit Function

Fail:
    errMsg = Err.Description
    ImportFrom = False
End Function

Public Function ExportTo(ByVal path As String, ByRef errMsg As String) As Boolean
    On Error GoTo Fail

    EnsureLoaded
    WriteUtf8 path, JsonLite.SerializePhrases(mLabels, mTexts, mNewlines, mCount)
    ExportTo = True
    Exit Function

Fail:
    errMsg = Err.Description
    ExportTo = False
End Function

'--- Helpers -------------------------------------------------------------------

Private Sub Grow(ByVal needed As Long)
    Dim capacity As Long

    On Error Resume Next
    capacity = UBound(mLabels) + 1
    If Err.Number <> 0 Then capacity = 0
    Err.Clear
    On Error GoTo 0

    If needed <= capacity Then Exit Sub

    Do While capacity < needed
        If capacity = 0 Then capacity = 16 Else capacity = capacity * 2
    Loop

    ReDim Preserve mLabels(0 To capacity - 1)
    ReDim Preserve mTexts(0 To capacity - 1)
    ReDim Preserve mNewlines(0 To capacity - 1)
End Sub

Private Sub SeedDefaults()
    ' Two examples, one per script, so a new user can see immediately that both
    ' Latin and Persian text work. Written with ChrW so the source file stays
    ' pure ASCII and cannot be mangled by an editor saving in the wrong encoding.
    AddPhrase "Hello", "Hello"

    Dim salam As String
    salam = ChrW$(&H633) & ChrW$(&H644) & ChrW$(&H627) & ChrW$(&H645)
    AddPhrase salam, salam
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub

' Reads a UTF-8 file into a VBA string, tolerating a leading BOM.
Public Function ReadUtf8(ByVal path As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2                ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    ReadUtf8 = stm.ReadText(-1) ' adReadAll
    stm.Close

    If Len(ReadUtf8) > 0 Then
        If AscW(Left$(ReadUtf8, 1)) = -257 Then ReadUtf8 = Mid$(ReadUtf8, 2)
    End If
End Function

' Writes a VBA string as UTF-8 with no BOM.
'
' ADODB.Stream always emits a BOM in text mode, so the text is staged in a
' text stream, then re-read as binary from byte 4 onward and saved from a
' second, binary stream.
Public Sub WriteUtf8(ByVal path As String, ByVal content As String)
    Dim textStm As Object, binStm As Object

    Set textStm = CreateObject("ADODB.Stream")
    textStm.Type = 2            ' adTypeText
    textStm.Charset = "utf-8"
    textStm.Open
    textStm.WriteText content

    ' Type may only be switched while Position is 0, so reset, switch, then
    ' seek past the 3-byte UTF-8 BOM.
    textStm.Position = 0
    textStm.Type = 1            ' adTypeBinary
    textStm.Position = 3

    Set binStm = CreateObject("ADODB.Stream")
    binStm.Type = 1
    binStm.Open
    textStm.CopyTo binStm
    binStm.SaveToFile path, 2   ' adSaveCreateOverWrite

    binStm.Close
    textStm.Close
End Sub
