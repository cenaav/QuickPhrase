Attribute VB_Name = "QuickPhraseMain"
'===============================================================================
' QuickPhraseMain - the actions behind the ribbon.
'
' Insertion, the manager dialog, import/export and the about box live here.
' Ribbon plumbing (callbacks, invalidation) lives in QuickPhraseRibbon.
'===============================================================================
Option Explicit

Public Const QP_NAME As String = "QuickPhrase"
Public Const QP_VERSION As String = "1.0.0"
Public Const QP_URL As String = "https://github.com/cenaav/QuickPhrase"

'--- Insertion -----------------------------------------------------------------

' Inserts a phrase at the insertion point, replacing any current selection.
' Stored text uses vbLf for line breaks; Word wants vbCr inside a range.
Public Sub InsertPhrase(ByVal index As Long)
    Dim text As String

    If index < 0 Or index >= SnippetStore.PhraseCount Then Exit Sub
    text = SnippetStore.PhraseText(index)
    If Len(text) = 0 Then Exit Sub

    If Application.Documents.count = 0 Then
        MsgBox "Open a document first.", vbInformation, QP_NAME
        Exit Sub
    End If

    Select Case Application.Selection.Type
        Case wdNoSelection
            MsgBox "Place the cursor where the text should go.", vbInformation, QP_NAME
            Exit Sub
    End Select

    ' Group the insertion into a single undo step where Word supports it.
    Dim rec As Object
    If Val(Application.Version) >= 14 Then
        On Error Resume Next
        Set rec = Application.UndoRecord
        rec.StartCustomRecord "QuickPhrase: insert phrase"
        On Error GoTo 0
    End If

    On Error GoTo Fail
    Application.Selection.TypeText text:=NormalizeLineBreaks(text)

CleanUp:
    If Not rec Is Nothing Then
        On Error Resume Next
        rec.EndCustomRecord
        On Error GoTo 0
    End If
    Exit Sub

Fail:
    MsgBox "QuickPhrase could not insert the text:" & vbCrLf & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & _
           "This usually means the document is protected or read-only.", _
           vbExclamation, QP_NAME
    Resume CleanUp
End Sub

' Converts stored line breaks into the carriage returns Word expects.
Private Function NormalizeLineBreaks(ByVal s As String) As String
    s = Replace(s, vbCrLf, vbCr)
    s = Replace(s, vbLf, vbCr)
    NormalizeLineBreaks = s
End Function

'--- Manager -------------------------------------------------------------------

' Shows the manager dialog. The form saves to disk itself; this refreshes the
' ribbon afterwards so new or reordered phrases appear immediately.
Public Sub ShowManager()
    frmManager.Show vbModal
    QuickPhraseRibbon.RefreshRibbon
End Sub

'--- Import / export -----------------------------------------------------------

Public Sub ImportPhrases()
    Dim path As String
    path = PickOpenFile()
    If Len(path) = 0 Then Exit Sub

    Dim answer As VbMsgBoxResult
    answer = MsgBox("Replace your current phrases with the ones in this file?" & vbCrLf & vbCrLf & _
                    "Yes" & vbTab & "- replace everything" & vbCrLf & _
                    "No" & vbTab & "- add them to the end of the current list" & vbCrLf & _
                    "Cancel" & vbTab & "- do nothing", _
                    vbYesNoCancel + vbQuestion, QP_NAME)
    If answer = vbCancel Then Exit Sub

    Dim errMsg As String
    If SnippetStore.ImportFrom(path, (answer = vbYes), errMsg) Then
        QuickPhraseRibbon.RefreshRibbon
        MsgBox "Imported " & SnippetStore.PhraseCount & " phrase(s).", _
               vbInformation, QP_NAME
    Else
        MsgBox "Import failed:" & vbCrLf & vbCrLf & errMsg, vbExclamation, QP_NAME
    End If
End Sub

Public Sub ExportPhrases()
    Dim path As String
    path = PickSaveFile()
    If Len(path) = 0 Then Exit Sub

    If LCase$(Right$(path, 5)) <> ".json" Then path = path & ".json"

    Dim errMsg As String
    If SnippetStore.ExportTo(path, errMsg) Then
        MsgBox "Exported " & SnippetStore.PhraseCount & " phrase(s) to:" & vbCrLf & _
               path, vbInformation, QP_NAME
    Else
        MsgBox "Export failed:" & vbCrLf & vbCrLf & errMsg, vbExclamation, QP_NAME
    End If
End Sub

Private Function PickOpenFile() As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)

    With dlg
        .Title = QP_NAME & " - import phrases"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "QuickPhrase JSON", "*.json"
        .Filters.Add "All files", "*.*"
        .InitialFileName = SnippetStore.StoreFolder() & Application.PathSeparator
        If .Show = -1 Then PickOpenFile = .SelectedItems(1)
    End With
End Function

Private Function PickSaveFile() As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogSaveAs)

    With dlg
        .Title = QP_NAME & " - export phrases"
        .InitialFileName = Environ$("USERPROFILE") & Application.PathSeparator & _
                           "Documents" & Application.PathSeparator & _
                           "quickphrase-export.json"
        If .Show = -1 Then PickSaveFile = .SelectedItems(1)
    End With
End Function

'--- About ---------------------------------------------------------------------

Public Sub ShowAbout()
    MsgBox QP_NAME & " " & QP_VERSION & vbCrLf & vbCrLf & _
           "Insert your frequently used phrases into Word with one click." & vbCrLf & vbCrLf & _
           "Phrase file:" & vbCrLf & SnippetStore.StorePath() & vbCrLf & vbCrLf & _
           "Phrases loaded: " & SnippetStore.PhraseCount & vbCrLf & _
           "Word version: " & Application.Version & vbCrLf & vbCrLf & _
           QP_URL & vbCrLf & _
           "MIT licence - free for personal and commercial use.", _
           vbInformation, QP_NAME & " " & QP_VERSION
End Sub

'--- Startup -------------------------------------------------------------------

' Word runs AutoExec when the template loads from the STARTUP folder.
Public Sub AutoExec()
    On Error Resume Next
    SnippetStore.EnsureLoaded
End Sub
