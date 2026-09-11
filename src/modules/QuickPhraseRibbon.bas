Attribute VB_Name = "QuickPhraseRibbon"
'===============================================================================
' QuickPhraseRibbon - callbacks named by src/customUI/customUI*.xml.
'
' The ribbon cannot be rebuilt at run time, so the Favorites group ships with a
' fixed pool of FAV_COUNT buttons. Each one asks this module for its label and
' visibility, which turns a static tab into a list the user can edit. Anything
' past that pool is still reachable through the "All Phrases" dynamic menu,
' which IS generated at run time and has no size limit.
'===============================================================================
Option Explicit

' Must match the number of qpBtnNN buttons in the customUI XML.
Public Const FAV_COUNT As Long = 12

Private mRibbon As IRibbonUI

'--- Lifecycle -----------------------------------------------------------------

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
    SnippetStore.EnsureLoaded
End Sub

' Asks Word to re-query every callback, so label and visibility changes show up
' without restarting. The IRibbonUI reference is lost if the VBA project resets
' (for example after an unhandled error while debugging); in that case the
' ribbon refreshes the next time Word rebuilds it, so failing quietly is right.
Public Sub RefreshRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

'--- Favorites buttons ---------------------------------------------------------

Public Sub GetFavLabel(control As IRibbonControl, ByRef returnedVal)
    Dim index As Long
    index = IndexFromButtonId(control.id)

    If index >= 0 And index < SnippetStore.PhraseCount Then
        returnedVal = SnippetStore.PhraseLabel(index)
    Else
        returnedVal = " "
    End If
End Sub

Public Sub GetFavVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = (IndexFromButtonId(control.id) < SnippetStore.PhraseCount)
End Sub

Public Sub GetFavScreentip(control As IRibbonControl, ByRef returnedVal)
    Dim index As Long
    index = IndexFromButtonId(control.id)

    If index >= 0 And index < SnippetStore.PhraseCount Then
        returnedVal = Left$(Replace(SnippetStore.PhraseText(index), vbLf, " "), 200)
    Else
        returnedVal = ""
    End If
End Sub

'--- All Phrases dynamic menu --------------------------------------------------

' Builds the full phrase list as ribbon XML. The namespace differs between
' Word 2007 and Word 2010+, and a mismatch makes the menu come up empty, so it
' is chosen from the running version.
Public Sub GetMenuContent(control As IRibbonControl, ByRef returnedVal)
    Dim ns As String
    Dim sb As String
    Dim i As Long
    Dim count As Long

    If Val(Application.Version) >= 14 Then
        ns = "http://schemas.microsoft.com/office/2009/07/customui"
    Else
        ns = "http://schemas.microsoft.com/office/2006/01/customui"
    End If

    count = SnippetStore.PhraseCount
    sb = "<menu xmlns=""" & ns & """ itemSize=""normal"">"

    If count = 0 Then
        sb = sb & "<button id=""qpDynEmpty"" label=""(no phrases yet - use Manage Phrases)"" " & _
                  "enabled=""false"" />"
    Else
        For i = 0 To count - 1
            sb = sb & "<button id=""qpDyn" & i & """" & _
                      " label=""" & XmlAttr(MenuLabel(i)) & """" & _
                      " tag=""" & i & """" & _
                      " onAction=""OnPhraseClick"" />"
        Next i
    End If

    sb = sb & "<menuSeparator id=""qpDynSep"" />"
    sb = sb & "<button id=""qpDynManage"" label=""Manage Phrases..."" " & _
              "imageMso=""ControlProperties"" onAction=""OnManage"" />"
    sb = sb & "</menu>"

    returnedVal = sb
End Sub

' Numbers the first FAV_COUNT entries so the user can see which phrases also
' have their own button on the tab.
Private Function MenuLabel(ByVal index As Long) As String
    Dim label As String
    label = SnippetStore.PhraseLabel(index)

    If index < FAV_COUNT Then
        MenuLabel = (index + 1) & ". " & label
    Else
        MenuLabel = label
    End If
End Function

'--- Shared actions ------------------------------------------------------------

' Serves both the fixed Favorites buttons and the dynamic menu entries. Menu
' items carry their index in tag; fixed buttons carry it in their id.
Public Sub OnPhraseClick(control As IRibbonControl)
    Dim index As Long

    If Len(control.tag) > 0 Then
        index = CLng(control.tag)
    Else
        index = IndexFromButtonId(control.id)
    End If

    QuickPhraseMain.InsertPhrase index
End Sub

Public Sub OnManage(control As IRibbonControl)
    QuickPhraseMain.ShowManager
End Sub

Public Sub OnImport(control As IRibbonControl)
    QuickPhraseMain.ImportPhrases
End Sub

Public Sub OnExport(control As IRibbonControl)
    QuickPhraseMain.ExportPhrases
End Sub

Public Sub OnReload(control As IRibbonControl)
    SnippetStore.LoadFromDisk
    RefreshRibbon
End Sub

Public Sub OnAbout(control As IRibbonControl)
    QuickPhraseMain.ShowAbout
End Sub

'--- Helpers -------------------------------------------------------------------

' "qpBtn07" -> 6. Returns -1 for anything unexpected.
Private Function IndexFromButtonId(ByVal controlId As String) As Long
    Dim suffix As String

    IndexFromButtonId = -1
    If Len(controlId) < 6 Then Exit Function
    If Left$(controlId, 5) <> "qpBtn" Then Exit Function

    suffix = Mid$(controlId, 6)
    If Not IsNumeric(suffix) Then Exit Function

    IndexFromButtonId = CLng(suffix) - 1
End Function

' Escapes a string for use inside an XML attribute.
Private Function XmlAttr(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    XmlAttr = s
End Function
