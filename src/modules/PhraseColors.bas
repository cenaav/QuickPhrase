Attribute VB_Name = "PhraseColors"
'===============================================================================
' PhraseColors - the colour tag shown beside a phrase on the ribbon.
'
' The ribbon has no way to set a background colour on button text. customUI
' exposes label, screentip and image, and nothing else; there is no styling API
' in any Word version. What a button can have is an icon, so a phrase's colour
' is drawn as a small solid swatch and returned from the getImage callback.
'
' Producing that swatch means building a bitmap at run time, because the colour
' is chosen by the user and cannot be shipped as a package image. GDI draws it,
' and OleCreatePictureIndirect wraps the handle as the IPictureDisp the ribbon
' expects.
'
' Colours are stored as "#RRGGBB" strings. An empty string means no colour,
' which is the default and leaves the button exactly as it was.
'===============================================================================
Option Explicit

Private Const SWATCH_SIZE As Long = 16

' Ribbon buttons are small, and a pale swatch on a pale ribbon would vanish
' without an outline.
Private Const BORDER_COLOR As Long = &H808080

#If VBA7 Then
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As LongPtr
    Private Declare PtrSafe Function FillRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
    Private Declare PtrSafe Function FrameRect Lib "user32" (ByVal hDC As LongPtr, lpRect As RECT, ByVal hBrush As LongPtr) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPictureDisp) As Long
    Private Declare PtrSafe Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
#Else
    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPictureDisp) As Long
    Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
#End If

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    #If VBA7 Then
        hImage As LongPtr
        hPal As LongPtr
    #Else
        hImage As Long
        hPal As Long
    #End If
End Type

Private Type CHOOSECOLOR
    #If VBA7 Then
        lStructSize As Long
        hwndOwner As LongPtr
        hInstance As LongPtr
        rgbResult As Long
        lpCustColors As LongPtr
        flags As Long
        lCustData As LongPtr
        lpfnHook As LongPtr
        lpTemplateName As LongPtr
    #Else
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        rgbResult As Long
        lpCustColors As Long
        flags As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As Long
    #End If
End Type

Private Const PICTYPE_BITMAP As Long = 1
Private Const CC_RGBINIT As Long = &H1
Private Const CC_FULLOPEN As Long = &H2
Private Const CC_ANYCOLOR As Long = &H100

' Drawing a bitmap on every ribbon refresh would be wasteful - getImage fires
' for every button each time the ribbon is invalidated - so swatches are made
' once per colour and reused.
Private mCache As Collection

'--- Palette -------------------------------------------------------------------

' Muted shades rather than saturated ones: these sit next to Word's own ribbon
' icons, and a block of pure red would shout.
Private Const PALETTE As String = _
    "None|;" & _
    "Yellow|#FFE699;" & _
    "Green|#C6E0B4;" & _
    "Blue|#BDD7EE;" & _
    "Red|#F8B7B7;" & _
    "Orange|#F8CBAD;" & _
    "Purple|#D9C2E9;" & _
    "Teal|#B7E1E4;" & _
    "Pink|#F4C7DC;" & _
    "Lime|#DCE775;" & _
    "Grey|#D9D9D9"

Public Function PaletteCount() As Long
    PaletteCount = UBound(Split(PALETTE, ";")) + 1
End Function

Public Function PaletteName(ByVal index As Long) As String
    Dim parts() As String
    parts = Split(PALETTE, ";")
    If index < 0 Or index > UBound(parts) Then Exit Function
    PaletteName = Split(parts(index), "|")(0)
End Function

Public Function PaletteHex(ByVal index As Long) As String
    Dim parts() As String
    parts = Split(PALETTE, ";")
    If index < 0 Or index > UBound(parts) Then Exit Function
    PaletteHex = Split(parts(index), "|")(1)
End Function

' Index of a stored colour in the palette, or -1 when it is a custom value.
Public Function PaletteIndexOf(ByVal hexColor As String) As Long
    Dim i As Long

    PaletteIndexOf = -1
    For i = 0 To PaletteCount - 1
        If StrComp(PaletteHex(i), Trim$(hexColor), vbTextCompare) = 0 Then
            PaletteIndexOf = i
            Exit Function
        End If
    Next i
End Function

'--- Conversion ----------------------------------------------------------------

' "#RRGGBB" to the BGR long that GDI and VBA both use. Returns -1 when the
' string is empty or malformed, which callers read as "no colour".
Public Function HexToColor(ByVal hexColor As String) As Long
    Dim s As String

    HexToColor = -1
    s = Trim$(hexColor)
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    If Len(s) <> 6 Then Exit Function

    On Error GoTo Fail
    HexToColor = CLng("&H" & Mid$(s, 5, 2)) * 65536 + _
                 CLng("&H" & Mid$(s, 3, 2)) * 256 + _
                 CLng("&H" & Mid$(s, 1, 2))
    Exit Function

Fail:
    HexToColor = -1
End Function

Public Function ColorToHex(ByVal color As Long) As String
    Dim r As Long, g As Long, b As Long

    If color < 0 Then Exit Function

    b = (color \ 65536) And &HFF
    g = (color \ 256) And &HFF
    r = color And &HFF

    ColorToHex = "#" & Right$("0" & Hex$(r), 2) & _
                       Right$("0" & Hex$(g), 2) & _
                       Right$("0" & Hex$(b), 2)
End Function

'--- Swatch --------------------------------------------------------------------

' The ribbon image for a phrase colour, or Nothing when the phrase has none.
Public Function SwatchFor(ByVal hexColor As String) As IPictureDisp
    Dim color As Long
    Dim key As String

    color = HexToColor(hexColor)
    If color < 0 Then Exit Function

    If mCache Is Nothing Then Set mCache = New Collection

    key = "c" & CStr(color)
    On Error Resume Next
    Set SwatchFor = mCache(key)
    On Error GoTo 0
    If Not SwatchFor Is Nothing Then Exit Function

    Set SwatchFor = CreateSwatch(color)
    If Not SwatchFor Is Nothing Then
        On Error Resume Next
        mCache.Add SwatchFor, key
        On Error GoTo 0
    End If
End Function

' Drops every cached bitmap. Called when colours change, so a reused colour is
' not served from a stale entry.
Public Sub ClearCache()
    Set mCache = Nothing
End Sub

Private Function CreateSwatch(ByVal color As Long) As IPictureDisp
    #If VBA7 Then
        Dim hScreenDC As LongPtr, hMemDC As LongPtr
        Dim hBitmap As LongPtr, hOldBitmap As LongPtr, hBrush As LongPtr
    #Else
        Dim hScreenDC As Long, hMemDC As Long
        Dim hBitmap As Long, hOldBitmap As Long, hBrush As Long
    #End If

    Dim box As RECT
    Dim desc As PICTDESC
    Dim iid As GUID
    Dim pic As IPictureDisp

    On Error GoTo CleanUp

    hScreenDC = GetDC(0)
    If hScreenDC = 0 Then Exit Function

    hMemDC = CreateCompatibleDC(hScreenDC)
    hBitmap = CreateCompatibleBitmap(hScreenDC, SWATCH_SIZE, SWATCH_SIZE)
    If hMemDC = 0 Or hBitmap = 0 Then GoTo CleanUp

    hOldBitmap = SelectObject(hMemDC, hBitmap)

    box.Left = 0
    box.Top = 0
    box.Right = SWATCH_SIZE
    box.Bottom = SWATCH_SIZE

    hBrush = CreateSolidBrush(color)
    FillRect hMemDC, box, hBrush
    DeleteObject hBrush

    hBrush = CreateSolidBrush(BORDER_COLOR)
    FrameRect hMemDC, box, hBrush
    DeleteObject hBrush

    SelectObject hMemDC, hOldBitmap

    ' IID_IPictureDisp {7BF80981-BF32-101A-8BBB-00AA00300CAB}
    With iid
        .Data1 = &H7BF80981
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B: .Data4(1) = &HBB: .Data4(2) = &H0: .Data4(3) = &HAA
        .Data4(4) = &H0:  .Data4(5) = &H30: .Data4(6) = &HC: .Data4(7) = &HAB
    End With

    With desc
        .cbSizeOfStruct = LenB(desc)
        .picType = PICTYPE_BITMAP
        .hImage = hBitmap
        .hPal = 0
    End With

    ' fPictureOwnsHandle = True, so the returned picture frees the bitmap when
    ' it is released. Deleting hBitmap here would blank the image.
    If OleCreatePictureIndirect(desc, iid, 1, pic) = 0 Then
        Set CreateSwatch = pic
        hBitmap = 0
    End If

CleanUp:
    If hBitmap <> 0 Then DeleteObject hBitmap
    If hMemDC <> 0 Then DeleteDC hMemDC
    If hScreenDC <> 0 Then ReleaseDC 0, hScreenDC
End Function

'--- Custom colour picker ------------------------------------------------------

' Opens the standard Windows colour dialog. Returns True and updates hexColor
' when the user picks one, False when they cancel or the dialog cannot open.
Public Function PickCustomColor(ByRef hexColor As String) As Boolean
    Dim cc As CHOOSECOLOR
    Dim custom(0 To 15) As Long
    Dim i As Long
    Dim current As Long

    On Error GoTo Fail

    ' The dialog requires all sixteen custom slots to be initialised.
    For i = 0 To 15
        custom(i) = &HFFFFFF
    Next i

    current = HexToColor(hexColor)
    If current < 0 Then current = &HFFFFFF

    With cc
        .lStructSize = LenB(cc)
        .hwndOwner = 0
        .rgbResult = current
        .lpCustColors = VarPtr(custom(0))
        .flags = CC_RGBINIT Or CC_FULLOPEN Or CC_ANYCOLOR
    End With

    If ChooseColorAPI(cc) = 0 Then Exit Function

    hexColor = ColorToHex(cc.rgbResult)
    PickCustomColor = True
    Exit Function

Fail:
    PickCustomColor = False
End Function
