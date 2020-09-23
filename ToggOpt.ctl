VERSION 5.00
Begin VB.UserControl ToggOpt 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3885
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   259
End
Attribute VB_Name = "ToggOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'XP code is not mine ,borrowed from someones button control on PSC
Private Type RECT
    rLeft As Long
    rTop As Long
    rRight As Long
    rBottom As Long
End Type
Private Const DT_VCENTER                As Long = &H4
Private Const DT_SINGLELINE             As Long = &H20
Private Const DT_FLAGS                  As Long = DT_VCENTER + DT_SINGLELINE
Private Const DT_CENTER                 As Long = &H1
Private Const mdef_Enabled              As Boolean = True

Private Type POINT
   x As Long
   Y As Long
End Type

Private Type RGBColor
    R As Single
    G As Single
    B As Single
End Type

Private Type typeColors
    cBorders(0 To 4)        As Long
    cTopLine1(0 To 4)       As Long
    cTopLine2(0 To 4)       As Long
    cBottomLine1(0 To 4)    As Long
    cBottomLine2(0 To 4)    As Long
    cCornerPixel1(0 To 4)   As Long
    cCornerPixel2(0 To 4)   As Long
    cCornerPixel3(0 To 4)   As Long
    cSideGradTop(1 To 3)    As Long
    cSideGradBottom(1 To 3) As Long
End Type

Public Enum eValue
   tOff = 0
   tOn = 1
End Enum

Public Enum eAlign
   Right = 0
   Left = 1
End Enum

Public Enum eStyle
   OnOff = 0
   YesNo = 1
   CheckX = 2
End Enum

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINT) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Public Event Click()
Public Event DblClick()

Private bSkipDrawing        As Boolean '------- Pauses drawing when needed
Private tColors             As typeColors '---- Enum declare for typeColors
Private bParentActive       As Boolean '------- Tracks when parent form has the Windows focus
Private bDisplayAsDefault   As Boolean '------- USed for ambient default property changes
Private lParentHwnd         As Long '---------- Stores the parents window handle

Private pHWND               As Long
Private pCAPTION            As String
Private pENABLED            As Boolean
Private pFORECOLOR          As OLE_COLOR
Private mValue              As Integer
Private mAlign              As Integer
Private mStyle              As Integer
Private mOffset             As Integer
Private WithEvents pFONT    As StdFont
Attribute pFONT.VB_VarHelpID = -1
Private mShadow              As Boolean
Private mEnabled              As Boolean

Private Sub DrawXP()
On Error Resume Next
Dim lw          As Long
Dim lh          As Long
Dim lHdc        As Long
Dim R           As RECT
Dim hRgn        As Long
Dim x As Integer

With UserControl
   lw = .ScaleWidth
   lh = .ScaleHeight
   .Cls
End With
lHdc = UserControl.hDC

With tColors
    LineApi 3, 0, lw - 3, 0, .cBorders(1) '------------------------ Draw border lines
    LineApi 0, 3, 0, lh - 3, .cBorders(1)
    LineApi 3, lh - 1, lw - 3, lh - 1, .cBorders(1)
    LineApi lw - 1, 3, lw - 1, lh - 3, .cBorders(1)
    
    SetRect R, 1, 3, lw - 1, lh - 2 '------------------------------- Draw side gradients
    If Enabled = True Then
       Call DrawGradient(R, .cSideGradTop(1), .cSideGradBottom(1))
    Else
       Call DrawGradient(R, &HE1E1E3, &HE1E1E3)
    End If
    SetRect R, 3, 3, lw - 3, lh - 3 '------------------------------- Draw background gradient (IDLE, HOT, FOCUS)
    If Enabled = True Then
    Call DrawGradient(R, 16514300, 15133676)
    LineApi 1, 1, lw, 1, .cTopLine1(1) '----------------------- Draw fade at the top
    LineApi 1, 2, lw, 2, .cTopLine2(1)
    LineApi 1, lh - 3, lw, lh - 3, .cBottomLine1(1) '---------- Draw fade at the bottom
    LineApi 2, lh - 2, lw - 1, lh - 2, .cBottomLine2(1)
    Else
    Call DrawGradient(R, &HE1E1E3, &HE1E1E3)
    LineApi 1, 1, lw, 1, &HE1E1E3    '----------------------- Draw fade at the top
    LineApi 1, 2, lw, 2, &HE1E1E3
    LineApi 1, lh - 3, lw, lh - 3, &HE1E1E3    '---------- Draw fade at the bottom
    LineApi 2, lh - 2, lw - 1, lh - 2, &HE1E1E3
    End If
  
    SetPixel lHdc, 0, 1, .cCornerPixel2(1) '----------------------- Top left Corner
    SetPixel lHdc, 0, 2, .cCornerPixel1(1)
    SetPixel lHdc, 1, 0, .cCornerPixel2(1)
    SetPixel lHdc, 1, 1, .cCornerPixel3(1)
    SetPixel lHdc, 2, 0, .cCornerPixel1(1)
    
    SetPixel lHdc, (lw - 1), 1, .cCornerPixel2(1) '---------------- Top right corner
    SetPixel lHdc, lw - 1, 2, .cCornerPixel1(1)
    SetPixel lHdc, lw - 2, 0, .cCornerPixel2(1)
    SetPixel lHdc, lw - 2, 1, .cCornerPixel3(1)
    SetPixel lHdc, lw - 3, 0, .cCornerPixel1(1)
    
    SetPixel lHdc, 0, lh - 2, .cCornerPixel2(1) '------------------ Bottom left corner
    SetPixel lHdc, 0, lh - 3, .cCornerPixel1(1)
    SetPixel lHdc, 1, lh - 1, .cCornerPixel2(1)
    SetPixel lHdc, 1, lh - 2, .cCornerPixel3(1)
    SetPixel lHdc, 2, lh - 1, .cCornerPixel1(1)
    
    SetPixel lHdc, lw - 1, lh - 2, .cCornerPixel2(1) '------------- Bottom right corner
    SetPixel lHdc, lw - 1, lh - 3, .cCornerPixel1(1)
    SetPixel lHdc, lw - 2, lh - 1, .cCornerPixel2(1)
    SetPixel lHdc, lw - 2, lh - 2, .cCornerPixel3(1)
    SetPixel lHdc, lw - 3, lh - 1, .cCornerPixel1(1)
    
    hRgn = CreateRoundRectRgn(0, 0, lw + 1, lh + 1, 3, 3) '------------- Clip extreme corner pixels
    Call SetWindowRgn(UserControl.hwnd, hRgn, True)
    DeleteObject hRgn
End With

bSkipDrawing = True '--------------------------------------------------- Draw caption
 If Align = 0 Then  ' ------Right Align
   SetRect R, 3, 3, lw - 3, lh - 3
   If Enabled = False Then
      UserControl.ForeColor = &H8000000C
   Else
      UserControl.ForeColor = ForeColor
   End If
   Call DrawText(lHdc, pCAPTION, -1, R, DT_FLAGS)

   LineApi lw - (UserControl.FontSize * 4) + (3 - OffSet), 4, lw - (UserControl.FontSize * 4) + (3 - OffSet), lh - 4, vbBlack '-----Draw vertical line
   Select Case Style
      Case 0
      If Value = 0 Then '------------------------------Draw Value "ON/OFF"
         If Shadow = True Then
            SetRect R, lw - (UserControl.FontSize * 4) + (7 - OffSet), 4, lw - 3, lh - 3
            If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
            UserControl.FontBold = Font.Bold
            Call DrawText(lHdc, "OFF", -1, R, DT_FLAGS)
         End If
         SetRect R, lw - (UserControl.FontSize * 4) + (6 - OffSet), 3, lw - 3, lh - 3
         If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbRed
            End If
         UserControl.FontBold = Font.Bold
         Call DrawText(lHdc, "OFF", -1, R, DT_FLAGS)
      Else
         If Shadow = True Then
            SetRect R, lw - (UserControl.FontSize * 4) + (9 - OffSet), 4, lw - 3, lh - 3
            If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
            UserControl.FontBold = Font.Bold
            Call DrawText(lHdc, "ON", -1, R, DT_FLAGS)
         End If
         SetRect R, lw - (UserControl.FontSize * 4) + (8 - OffSet), 3, lw - 3, lh - 3
         If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = &HC000&
            End If
         UserControl.FontBold = Font.Bold
         Call DrawText(lHdc, "ON", -1, R, DT_FLAGS)
      End If
      Case 1
      If Value = 0 Then '------------------------------Draw Value "Yes/No"
         If Shadow = True Then
            SetRect R, lw - (UserControl.FontSize * 4) + (10 - OffSet), 4, lw - 3, lh - 3
            If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
            UserControl.FontBold = Font.Bold
            Call DrawText(lHdc, "NO", -1, R, DT_FLAGS)
         End If
         SetRect R, lw - (UserControl.FontSize * 4) + (9 - OffSet), 3, lw - 3, lh - 3
         If Enabled = False Then
            UserControl.ForeColor = &H8000000C
         Else
            UserControl.ForeColor = vbRed
         End If
         UserControl.FontBold = Font.Bold
         Call DrawText(lHdc, "NO", -1, R, DT_FLAGS)
      Else
         If Shadow = True Then
            SetRect R, lw - (UserControl.FontSize * 4) + (7 - OffSet), 4, lw - 3, lh - 3
            If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
            UserControl.FontBold = Font.Bold
            Call DrawText(lHdc, "YES", -1, R, DT_FLAGS)
         End If
         SetRect R, lw - (UserControl.FontSize * 4) + (6 - OffSet), 3, lw - 3, lh - 3
         If Enabled = False Then
            UserControl.ForeColor = &H8000000C
         Else
            UserControl.ForeColor = &HC000&
         End If
         UserControl.FontBold = Font.Bold
         Call DrawText(lHdc, "YES", -1, R, DT_FLAGS)
      End If
      Case 2
         If Value = 0 Then '------------------------------Draw Value "Check/X"
         If Shadow = True Then
         For x = 0 To 12  '--------draw X
            If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
            UserControl.Line (lw - x - 6, x + 3)-(lw - x - 13, x + 3)
            UserControl.Line (lw - x - 6, 15 - x)-(lw - x - 13, 15 - x)
         Next x
         End If
         For x = 0 To 10
            If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbRed
            End If
            UserControl.Line (lw - x - 8, 4 + x)-(lw - x - 13, 4 + x)
            UserControl.Line (lw - x - 8, 14 - x)-(lw - x - 13, 14 - x)
         Next x
      Else
        If Shadow = True Then
        For x = 0 To 12 '-----draw check
           If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
           If x >= 7 Then UserControl.Line ((lw - 35) + x, x + 3)-((lw - 33) + (x + 5), x + 3)
           UserControl.Line ((lw - 20) + x, 15 - x)-((lw - 18) + x + 4, 15 - x)
        Next x
        End If
        For x = 0 To 10
           If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbGreen
            End If
           If x >= 7 Then UserControl.Line ((lw - 33) + x, 4 + x)-((lw - 32) + (x + 4), 4 + x)
           UserControl.Line ((lw - 18) + x, 14 - x)-((lw - 18) + x + 4, 14 - x)
        Next x
      End If
   End Select
 Else    '------------------Left Align
   SetRect R, (UserControl.FontSize * 4) + (3 + OffSet), 3, lw - 3, lh - 3
   If Enabled = False Then
      UserControl.ForeColor = &H8000000C
   Else
      UserControl.ForeColor = ForeColor
   End If
   Call DrawText(lHdc, pCAPTION, -1, R, DT_FLAGS)

   LineApi (UserControl.FontSize * 4) - (3 - OffSet), 4, (UserControl.FontSize * 4) - (3 - OffSet), lh - 4, vbBlack ' ---------------------------Draw vertical line
   Select Case Style
      Case 0
      If Value = 0 Then '------------------------------Draw Value "ON/OFF"
        If Shadow = True Then
           SetRect R, 4, 4, lw - 3, lh - 3
           If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
           UserControl.FontBold = Font.Bold
           Call DrawText(lHdc, "OFF", -1, R, DT_FLAGS)
        End If
        SetRect R, 3, 3, lw - 3, lh - 3
        If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbRed
            End If
        UserControl.FontBold = Font.Bold
        Call DrawText(lHdc, "OFF", -1, R, DT_FLAGS)
      Else
        If Shadow = True Then
           SetRect R, 7, 4, lw - 3, lh - 3
           If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
           UserControl.FontBold = Font.Bold
           Call DrawText(lHdc, "ON", -1, R, DT_FLAGS)
        End If
        SetRect R, 6, 3, lw - 3, lh - 3
        If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = &HC000&
            End If
        UserControl.FontBold = Font.Bold
        Call DrawText(lHdc, "ON", -1, R, DT_FLAGS)
      End If
      Case 1
      If Value = 0 Then '------------------------------Draw Value "Yes/No"
        If Shadow = True Then
           SetRect R, 7, 5, lw - 3, lh - 3
           If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
           UserControl.FontBold = Font.Bold
           Call DrawText(lHdc, "NO", -1, R, DT_FLAGS)
        End If
        SetRect R, 6, 3, lw - 3, lh - 3
        If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbRed
            End If
        UserControl.FontBold = Font.Bold
        Call DrawText(lHdc, "NO", -1, R, DT_FLAGS)
      Else
        If Shadow = True Then
           SetRect R, 4, 5, lw - 3, lh - 3
           If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
           UserControl.FontBold = Font.Bold
           Call DrawText(lHdc, "YES", -1, R, DT_FLAGS)
        End If
        SetRect R, 3, 3, lw - 3, lh - 3
        If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = &HC000&
            End If
        UserControl.FontBold = Font.Bold
        Call DrawText(lHdc, "YES", -1, R, DT_FLAGS)
      End If
      Case 2
      If Value = 0 Then '------------------------------Draw Value "Check/X"
         If Shadow = True Then
         For x = 0 To 12  '--------draw X
            If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
            UserControl.Line (x + 6, x + 4)-(x + 13, x + 4)
            UserControl.Line (x + 6, 16 - x)-(x + 13, 16 - x)
         Next x
         End If
         For x = 0 To 10
            If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbRed
            End If
            UserControl.Line (x + 8, 5 + x)-(x + 13, 5 + x)
            UserControl.Line (x + 8, 15 - x)-(x + 13, 15 - x)
         Next x
      Else
        If Shadow = True Then
        For x = 0 To 12 '-----draw check
           If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbBlack
            End If
           If x >= 7 Then UserControl.Line (x - 4, x + 4)-(x + 3, x + 4)
           UserControl.Line (x + 10, 16 - x)-(x + 16, 16 - x)
        Next x
        End If
        For x = 0 To 10
           If Enabled = False Then
               UserControl.ForeColor = &H8000000C
            Else
               UserControl.ForeColor = vbGreen
            End If
           If x >= 7 Then UserControl.Line (x - 2, 5 + x)-(x + 3, 5 + x)
           UserControl.Line (x + 12, 15 - x)-(x + 16, 15 - x)
        Next x
      End If
   End Select
 End If
bSkipDrawing = False
End Sub

Private Sub UserControl_Initialize()
bSkipDrawing = 1
Call FillColorScheme
Set pFONT = UserControl.Font
pHWND = UserControl.hwnd
mEnabled = mdef_Enabled
End Sub

Private Sub UserControl_InitProperties()
   ForeColor = &H0
   Caption = Extender.Name
   Align = 0
   Style = 0
   OffSet = 0
   Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lParentHwnd = UserControl.Parent.hwnd
    With PropBag
        Caption = .ReadProperty("Caption", Extender.Name)
        ForeColor = .ReadProperty("ForeColor", 0)
        Value = .ReadProperty("Value", 0)
        Align = .ReadProperty("Align", 0)
        Style = .ReadProperty("Style", 0)
        OffSet = .ReadProperty("OffSet", 0)
        Shadow = .ReadProperty("Shadow", True)
        Set Font = .ReadProperty("Font", pFONT)
        Enabled = .ReadProperty("Enabled", mdef_Enabled)
    End With
    bSkipDrawing = False: Call DrawXP
End Sub

Private Sub UserControl_Resize()
With UserControl
        If .Height < 100 Then bSkipDrawing = True: .Height = 100
        If .Width < 100 Then bSkipDrawing = True: .Width = 100
    End With
    If Not bSkipDrawing Then Call DrawXP
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Caption", pCAPTION, Extender.Name
        .WriteProperty "ForeColor", pFORECOLOR, 0
        .WriteProperty "Value", mValue, 0
        .WriteProperty "Align", mAlign, 0
        .WriteProperty "Style", mStyle, 0
        .WriteProperty "OffSet", mOffset, 0
        .WriteProperty "Shadow", mShadow, True
        .WriteProperty "Font", pFONT, "Verdana"
        .WriteProperty "Enabled", mEnabled, mdef_Enabled
    End With
End Sub

Private Sub UserControl_Click()
If Value = 0 Then
   Value = 1
   Else
   Value = 0
   End If
RaiseEvent Click
End Sub

Private Sub UserControl_Terminate()
   Set pFONT = Nothing
End Sub

Public Property Get Align() As eAlign
    Align = mAlign
End Property

Public Property Let Align(NewValue As eAlign)
    mAlign = NewValue
    PropertyChanged "Align"
    Call DrawXP
End Property

Public Property Get Caption() As String
    Caption = pCAPTION
End Property

Public Property Let Caption(ByVal NewValue As String)
    pCAPTION = NewValue
    Call DrawXP
    UserControl.PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = mEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    mEnabled = NewValue
    UserControl.Enabled = mEnabled
    Call DrawXP
    UserControl.PropertyChanged "Enabled"
End Property

Public Property Get Value() As eValue
    Value = mValue
End Property

Public Property Let Value(ByVal NewValue As eValue)
    mValue = NewValue
    DrawXP
    UserControl.PropertyChanged "Value"
End Property

Public Property Get Font() As StdFont
    Set Font = pFONT
End Property

Public Property Set Font(NewValue As StdFont)
    Set pFONT = NewValue
    Call pFONT_FontChanged("")
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = pFORECOLOR
End Property

Public Property Let ForeColor(NewValue As OLE_COLOR)
    pFORECOLOR = NewValue
    UserControl.ForeColor = pFORECOLOR
    Call DrawXP
    UserControl.PropertyChanged "ForeColor"
End Property

Public Property Get OffSet() As Integer
    OffSet = mOffset
End Property

Public Property Let OffSet(ByVal NewValue As Integer)
    mOffset = NewValue
    DrawXP
    UserControl.PropertyChanged "OffSet"
End Property

Public Property Get Shadow() As Boolean
    Shadow = mShadow
End Property

Public Property Let Shadow(ByVal NewValue As Boolean)
    mShadow = NewValue
    DrawXP
    UserControl.PropertyChanged "Shadow"
End Property

Public Property Get Style() As eStyle
    Style = mStyle
End Property

Public Property Let Style(ByVal NewValue As eStyle)
    mStyle = NewValue
    DrawXP
    UserControl.PropertyChanged "Style"
End Property

Private Sub DrawGradient(R As RECT, ByVal StartColor As Long, ByVal EndColor As Long)
Dim s       As RGBColor '--- Start RGB colors
Dim e       As RGBColor '--- End RBG colors
Dim I       As RGBColor '--- Increment RGB colors
Dim x       As Long
Dim lSteps  As Long
Dim lHdc    As Long
    lHdc = UserControl.hDC
    lSteps = R.rBottom - R.rTop
    s.R = (StartColor And &HFF)
    s.G = (StartColor \ &H100) And &HFF
    s.B = (StartColor And &HFF0000) / &H10000
    e.R = (EndColor And &HFF)
    e.G = (EndColor \ &H100) And &HFF
    e.B = (EndColor And &HFF0000) / &H10000
    With I
        .R = (s.R - e.R) / lSteps
        .G = (s.G - e.G) / lSteps
        .B = (s.B - e.B) / lSteps
        For x = 0 To lSteps
            Call LineApi(R.rLeft, (lSteps - x) + R.rTop, R.rRight, (lSteps - x) + R.rTop, RGB(e.R + (x * .R), e.G + (x * .G), e.B + (x * .B)))
        Next x
    End With
End Sub

Private Sub DrawFilled(tR As RECT, ByVal cBackColor As Long)
Dim hBrush As Long
    hBrush = CreateSolidBrush(cBackColor) '----------------- Fill with solid brush
    FillRect UserControl.hDC, tR, hBrush
    DeleteObject hBrush
End Sub

Private Sub LineApi(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
Dim pt      As POINT
Dim hPen    As Long
Dim hPenOld As Long
Dim lHdc    As Long
    lHdc = UserControl.hDC
    hPen = CreatePen(0, 1, Color)
    hPenOld = SelectObject(lHdc, hPen)
    MoveToEx lHdc, X1, Y1, pt
    LineTo lHdc, X2, Y2
    SelectObject lHdc, hPenOld
    DeleteObject hPen
End Sub

Private Sub FillColorScheme()
    With tColors
        .cBorders(1) = 7617536
        .cTopLine1(1) = 16777215
        .cTopLine2(1) = 16711422
        .cBottomLine1(1) = 14082018
        .cBottomLine2(1) = 12964054
        .cCornerPixel1(1) = 8672545
        .cCornerPixel2(1) = 11376251
        .cCornerPixel3(1) = 10845522
        .cSideGradTop(1) = 16514300
        .cSideGradBottom(1) = 15133676
    End With
End Sub

Private Sub pFONT_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = pFONT
    Call DrawXP
    UserControl.PropertyChanged "Font"
End Sub

