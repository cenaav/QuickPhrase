'===============================================================================
' frmManager - the phrase editor.
'
' The form edits SnippetStore's in-memory list directly. "Save & Close" writes
' it to disk; "Cancel" throws the edits away by re-reading the file. Edits in
' the label/text boxes are committed to the selected phrase automatically
' whenever the selection changes or the form closes, so there is no Apply step.
'
' NOTE: controls are created by build/build.ps1, not by a checked-in .frx blob,
' so this file is the only place the dialog's behaviour is defined. Control
' names must stay in sync with the $controls table in that script.
'===============================================================================
Option Explicit

' Index currently shown in the edit boxes, or -1 when nothing is loaded.
Private mCurrentIndex As Long

' Set while the list box is being repopulated, so Change events fired by code
' are not mistaken for the user picking a different phrase.
Private mSuspendEvents As Boolean

'--- Form lifecycle ------------------------------------------------------------

Private Sub UserForm_Initialize()
    Me.Caption = "QuickPhrase - Manage Phrases"
    mCurrentIndex = -1

    SnippetStore.EnsureLoaded
    RefreshList 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Closing with the X button behaves like Cancel: discard, do not save.
    If CloseMode = vbFormControlMenu Then
        mSuspendEvents = True
        SnippetStore.LoadFromDisk
    End If
End Sub

'--- List ----------------------------------------------------------------------

' Rebuilds the list box and selects selectIndex (clamped to the valid range).
Private Sub RefreshList(ByVal selectIndex As Long)
    Dim i As Long
    Dim count As Long

    mSuspendEvents = True
    lstPhrases.Clear

    count = SnippetStore.PhraseCount
    For i = 0 To count - 1
        lstPhrases.AddItem (i + 1) & ".  " & SnippetStore.PhraseLabel(i)
    Next i

    If count = 0 Then
        mCurrentIndex = -1
        mSuspendEvents = False
        LoadEditors -1
        UpdateEnabledState
        Exit Sub
    End If

    If selectIndex < 0 Then selectIndex = 0
    If selectIndex > count - 1 Then selectIndex = count - 1

    lstPhrases.ListIndex = selectIndex
    mSuspendEvents = False

    LoadEditors selectIndex
    UpdateEnabledState
End Sub

Private Sub lstPhrases_Change()
    If mSuspendEvents Then Exit Sub

    ' Keep whatever the user typed for the phrase they are leaving.
    CommitEditors
    LoadEditors lstPhrases.ListIndex
    UpdateEnabledState
End Sub

'--- Editors -------------------------------------------------------------------

Private Sub LoadEditors(ByVal index As Long)
    mSuspendEvents = True

    If index < 0 Or index >= SnippetStore.PhraseCount Then
        txtLabel.text = ""
        txtText.text = ""
        mCurrentIndex = -1
    Else
        txtLabel.text = SnippetStore.PhraseLabel(index)
        ' MSForms text boxes want vbCrLf for line breaks; the store uses vbLf.
        txtText.text = Replace(SnippetStore.PhraseText(index), vbLf, vbCrLf)
        mCurrentIndex = index
    End If

    mSuspendEvents = False
End Sub

' Writes the edit boxes back into the store entry they were loaded from.
Private Sub CommitEditors()
    If mCurrentIndex < 0 Then Exit Sub
    If mCurrentIndex >= SnippetStore.PhraseCount Then Exit Sub

    Dim label As String, text As String
    label = Trim$(txtLabel.text)
    text = Replace(txtText.text, vbCrLf, vbLf)

    ' An empty label would produce a blank ribbon button.
    If Len(label) = 0 Then label = Left$(Replace(text, vbLf, " "), 24)
    If Len(label) = 0 Then label = "(untitled)"

    SnippetStore.UpdatePhrase mCurrentIndex, label, text
End Sub

'--- Buttons -------------------------------------------------------------------

Private Sub btnNew_Click()
    CommitEditors

    SnippetStore.AddPhrase "New phrase", ""
    RefreshList SnippetStore.PhraseCount - 1

    txtLabel.SetFocus
    txtLabel.SelStart = 0
    txtLabel.SelLength = Len(txtLabel.text)
End Sub

Private Sub btnDelete_Click()
    Dim index As Long
    index = lstPhrases.ListIndex
    If index < 0 Then Exit Sub

    If MsgBox("Delete """ & SnippetStore.PhraseLabel(index) & """?", _
              vbYesNo + vbQuestion, "QuickPhrase") <> vbYes Then Exit Sub

    ' Deliberately not committing first - the entry is about to disappear.
    mCurrentIndex = -1
    SnippetStore.DeletePhrase index
    RefreshList index
End Sub

Private Sub btnUp_Click()
    MoveSelection -1
End Sub

Private Sub btnDown_Click()
    MoveSelection 1
End Sub

Private Sub MoveSelection(ByVal delta As Long)
    Dim index As Long
    index = lstPhrases.ListIndex
    If index < 0 Then Exit Sub

    CommitEditors

    Dim newIndex As Long
    newIndex = SnippetStore.MovePhrase(index, delta)
    If newIndex = index Then Exit Sub

    mCurrentIndex = -1
    RefreshList newIndex
End Sub

Private Sub btnSave_Click()
    CommitEditors

    If Not SnippetStore.SaveToDisk() Then Exit Sub

    mSuspendEvents = True
    Unload Me
End Sub

Private Sub btnCancel_Click()
    mSuspendEvents = True
    SnippetStore.LoadFromDisk
    Unload Me
End Sub

'--- State ---------------------------------------------------------------------

Private Sub UpdateEnabledState()
    Dim index As Long
    Dim count As Long

    index = lstPhrases.ListIndex
    count = SnippetStore.PhraseCount

    btnDelete.Enabled = (index >= 0)
    btnUp.Enabled = (index > 0)
    btnDown.Enabled = (index >= 0 And index < count - 1)
    txtLabel.Enabled = (index >= 0)
    txtText.Enabled = (index >= 0)

    If count > QuickPhraseRibbon.FAV_COUNT Then
        lblHint.Caption = "The first " & QuickPhraseRibbon.FAV_COUNT & _
                          " phrases get their own button on the QuickPhrase tab. " & _
                          "The rest are in the All Phrases menu. Use Up/Down to reorder."
    Else
        lblHint.Caption = "Every phrase gets its own button on the QuickPhrase tab. " & _
                          "Use Up/Down to change the button order."
    End If
End Sub
