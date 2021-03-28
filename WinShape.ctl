VERSION 5.00
Begin VB.UserControl WinShape 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   2880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picAux 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   804
      Left            =   1656
      ScaleHeight     =   804
      ScaleWidth      =   588
      TabIndex        =   0
      Top             =   1404
      Visible         =   0   'False
      Width           =   588
   End
   Begin VB.Shape Shape1 
      Height          =   660
      Left            =   252
      Top             =   216
      Width           =   444
   End
   Begin VB.Shape DefShape 
      Height          =   336
      Left            =   396
      Top             =   1944
      Visible         =   0   'False
      Width           =   408
   End
End
Attribute VB_Name = "WinShape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type
 
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, ByRef graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef lpInput As GDIPlusStartupInput, Optional ByRef lpOutput As Any) As Long
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipFillEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal color As Long, ByVal Width As Single, ByVal unit As Long, pen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GdipFillPieI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Private Const UnitPixel = 2
Private Const SmoothingModeAntiAlias    As Long = &H4

Private Type OSVERSIONINFO 'for GetVersionEx API call
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type
 
Private Type Size
    cx As Long
    cy As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
 
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal hdcDst As Long, pptDst As Any, psize As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Any, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Const DIB_RGB_COLORS As Long = 0
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_LAYERED As Long = &H80000
Private Const AC_SRC_ALPHA As Long = &H1
Private Const AC_SRC_OVER As Long = &H0
Private Const ULW_ALPHA As Long = &H2
Private Const ULW_OPAQUE As Long = &H4

Private Const WM_PAINT As Long = &HF
Private Const WM_PRINT As Long = &H317
Private Const PRF_CHILDREN As Long = &H10&
Private Const PRF_CLIENT As Long = &H4&
Private Const PRF_OWNED As Long = &H20&

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Public Enum shBackStyleConstants
    shTransparent = 0
    shOpaque = 1
End Enum

Private mAntiAlias As Boolean
Private mOpacity As Single

Private mGdipToken As Long
Private mOldBitmap As Long
Private m32BitsBitmap As Long
Private mDC As Long

Private mTesting As Boolean


Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()
#If Testing Then
    mTesting = InIDE2
#End If
End Sub

Private Sub UserControl_InitProperties()
    mOpacity = 100
    mAntiAlias = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_Resize()
    Dim iWinStyle As Long
    
    picAux.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    If mDC Then
        iWinStyle = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
        If (iWinStyle And WS_EX_LAYERED) <> 0 Then
            RemoveLayeredTransparency
            SetWindowLong UserControl.hwnd, GWL_EXSTYLE, iWinStyle And Not WS_EX_LAYERED
        End If
        Destroy32BitsBitmap
    End If
    Me.Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Shape1.BackColor = PropBag.ReadProperty("BackColor", DefShape.BackColor)
    Shape1.BackStyle = PropBag.ReadProperty("BackStyle", DefShape.BackStyle)
    Shape1.BorderColor = PropBag.ReadProperty("BorderColor", DefShape.BorderColor)
    Shape1.Shape = PropBag.ReadProperty("Shape", DefShape.Shape)
    Shape1.FillColor = PropBag.ReadProperty("FillColor", DefShape.FillColor)
    Shape1.FillStyle = PropBag.ReadProperty("FillStyle", DefShape.FillStyle)
    Shape1.BorderStyle = PropBag.ReadProperty("BorderStyle", DefShape.BorderStyle)
    Shape1.BorderWidth = PropBag.ReadProperty("BorderWidth", DefShape.BorderWidth)
    mAntiAlias = PropBag.ReadProperty("AntiAlias", True)
    mOpacity = PropBag.ReadProperty("Opacity", 100)
End Sub

Private Sub UserControl_Terminate()
    If mGdipToken <> 0 Then
        Destroy32BitsBitmap
        TerminateGDI
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", Shape1.BackColor, DefShape.BackColor
    PropBag.WriteProperty "BackStyle", Shape1.BackStyle, DefShape.BackStyle
    PropBag.WriteProperty "BorderColor", Shape1.BorderColor, DefShape.BorderColor
    PropBag.WriteProperty "Shape", Shape1.Shape, DefShape.Shape
    PropBag.WriteProperty "FillColor", Shape1.FillColor, DefShape.FillColor
    PropBag.WriteProperty "FillStyle", Shape1.FillStyle, DefShape.FillStyle
    PropBag.WriteProperty "BorderStyle", Shape1.BorderStyle, DefShape.BorderStyle
    PropBag.WriteProperty "BorderWidth", Shape1.BorderWidth, DefShape.BorderWidth
    PropBag.WriteProperty "AntiAlias", mAntiAlias, True
    PropBag.WriteProperty "Opacity", mOpacity, 100
End Sub


Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = Shape1.BorderColor
End Property

Public Property Let BorderColor(ByVal nValue As OLE_COLOR)
    If nValue <> Shape1.BorderColor Then
        Shape1.BorderColor = nValue
        Me.Refresh
        PropertyChanged "BorderColor"
    End If
End Property


Public Property Get Shape() As ShapeConstants
Attribute Shape.VB_Description = "Returns/sets a value indicating the appearance of a control."
    Shape = Shape1.Shape
End Property

Public Property Let Shape(ByVal nValue As ShapeConstants)
    If nValue <> Shape1.Shape Then
        Shape1.Shape = nValue
        Me.Refresh
        PropertyChanged "Shape"
    End If
End Property


Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Shape1.BackColor
End Property

Public Property Let BackColor(ByVal nValue As OLE_COLOR)
    If nValue <> Shape1.BackColor Then
        Shape1.BackColor = nValue
        Me.Refresh
        PropertyChanged "BackColor"
    End If
End Property


Public Property Get BackStyle() As shBackStyleConstants
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = Shape1.BackStyle
End Property

Public Property Let BackStyle(ByVal nValue As shBackStyleConstants)
    If nValue <> Shape1.BackStyle Then
        Shape1.BackStyle = nValue
        Me.Refresh
        PropertyChanged "BackStyle"
    End If
End Property


Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = Shape1.FillColor
End Property

Public Property Let FillColor(ByVal nValue As OLE_COLOR)
    If nValue <> Shape1.FillColor Then
        Shape1.FillColor = nValue
        Me.Refresh
        PropertyChanged "FillColor"
    End If
End Property


Public Property Get FillStyle() As FillStyleConstants
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    FillStyle = Shape1.FillStyle
End Property

Public Property Let FillStyle(ByVal nValue As FillStyleConstants)
    If nValue <> Shape1.FillStyle Then
        Shape1.FillStyle = nValue
        Me.Refresh
        PropertyChanged "FillStyle"
    End If
End Property


Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = Shape1.BorderStyle
End Property

Public Property Let BorderStyle(ByVal nValue As BorderStyleConstants)
    If nValue <> Shape1.BorderStyle Then
        Shape1.BorderStyle = nValue
        Me.Refresh
        PropertyChanged "BorderStyle"
    End If
End Property


Public Property Get BorderWidth() As Long
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."
    BorderWidth = Shape1.BorderWidth
End Property

Public Property Let BorderWidth(ByVal nValue As Long)
    If nValue <> Shape1.BorderWidth Then
        Shape1.BorderWidth = nValue
        UserControl_Resize
        PropertyChanged "BorderWidth"
    End If
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
    
    
Public Property Get AntiAlias() As Boolean
    AntiAlias = mAntiAlias
End Property
    
Public Property Let AntiAlias(ByVal nValue As Boolean)
    If nValue <> mAntiAlias Then
        mAntiAlias = nValue
        If Not mAntiAlias Then UserControl.Cls
        Shape1.Visible = Not mAntiAlias
        UserControl.Refresh
        Me.Refresh
        PropertyChanged "AntiAlias"
    End If
End Property


Public Property Get Opacity() As Single
    Opacity = mOpacity
End Property
    
Public Property Let Opacity(ByVal nValue As Single)
    If nValue <> mOpacity Then
        mOpacity = nValue
        If (mOpacity < 0) Or (mOpacity > 100) Then mOpacity = 100
        Me.Refresh
        PropertyChanged "Opacity"
    End If
End Property

    
Public Sub Refresh()
    Dim iBorderX As Long
    Dim iBorderY As Long
    Dim iWinStyle As Long
    
    If mAntiAlias And IsWindows8OrMore And Not InIDE2 Then
        Set UserControl.MaskPicture = Nothing
        Shape1.Visible = False
        UserControl.BackStyle = shOpaque
        iWinStyle = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
        If (iWinStyle And WS_EX_LAYERED) = 0 Then
            SetWindowLong UserControl.hwnd, GWL_EXSTYLE, iWinStyle Or WS_EX_LAYERED
        End If
        Draw
        MakeLayeredTransparency
    Else
        iWinStyle = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
        If (iWinStyle And WS_EX_LAYERED) <> 0 Then
            'RemoveLayeredTransparency
            SetWindowLong UserControl.hwnd, GWL_EXSTYLE, iWinStyle And Not WS_EX_LAYERED
        End If
        Shape1.Visible = True
        UserControl.BackStyle = shTransparent
        UserControl.Refresh
        
        iBorderX = UserControl.ScaleX(Shape1.BorderWidth, vbPixels, UserControl.ScaleMode)
        iBorderY = UserControl.ScaleY(Shape1.BorderWidth, vbPixels, UserControl.ScaleMode)
        Shape1.Move iBorderX / 2, iBorderY / 2, UserControl.ScaleWidth - iBorderX, UserControl.ScaleHeight - iBorderY
        
        SendMessage UserControl.hwnd, WM_PAINT, picAux.hdc, 0
        SendMessage UserControl.hwnd, WM_PRINT, picAux.hdc, PRF_CHILDREN Or PRF_CLIENT Or PRF_OWNED
        
        Set UserControl.MaskPicture = picAux.Image
        picAux.Cls
    End If
End Sub

Private Sub Draw()
    Dim iShape As ShapeConstants
    Dim iDiameter As Long
    Dim iGraphics As Long
    Dim iFillColor As Long
    Dim iFilled As Boolean
    Dim iHeight As Long
    
    UserControl.Cls
    iShape = Shape1.Shape
    
    If mGdipToken = 0 Then InitGDI
    If mDC = 0 Then Create32BitsBitmap
    If GdipCreateFromHDC(IIf(mTesting, UserControl.hdc, mDC), iGraphics) = 0 Then
        
        Create32BitsBitmap
        
        If Shape1.FillStyle = vbFSSolid Then
            iFilled = True
            iFillColor = Shape1.FillColor
        ElseIf Shape1.BackStyle = shOpaque Then
            iFilled = True
            iFillColor = Shape1.BackColor
        End If
        
        If iShape = vbShapeOval Then
            If iFilled Then
                FillEllipse iGraphics, iFillColor, Shape1.BorderWidth / 2, Shape1.BorderWidth / 2, UserControl.ScaleWidth - Shape1.BorderWidth - 2, UserControl.ScaleHeight - Shape1.BorderWidth - 2
            End If
            If Shape1.BorderStyle <> vbTransparent Then
                DrawEllipse iGraphics, Shape1.BorderColor, Shape1.BorderWidth, Shape1.BorderWidth / 2, Shape1.BorderWidth / 2, UserControl.ScaleWidth - Shape1.BorderWidth - 1, UserControl.ScaleHeight - Shape1.BorderWidth - 1
            End If
        ElseIf iShape = vbShapeCircle Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iDiameter = UserControl.ScaleWidth - Shape1.BorderWidth
            Else
                iDiameter = UserControl.ScaleHeight - Shape1.BorderWidth
            End If
            If iFilled Then
                FillEllipse iGraphics, iFillColor, UserControl.ScaleWidth / 2 - iDiameter / 2, UserControl.ScaleHeight / 2 - iDiameter / 2, iDiameter - 2, iDiameter - 2
            End If
            If Shape1.BorderStyle <> vbTransparent Then
                DrawEllipse iGraphics, Shape1.BorderColor, Shape1.BorderWidth, UserControl.ScaleWidth / 2 - iDiameter / 2, UserControl.ScaleHeight / 2 - iDiameter / 2, iDiameter - 1, iDiameter - 1
            End If
        ElseIf iShape = vbShapeRectangle Then
            If iFilled Then
                FillRectangle iGraphics, iFillColor, Shape1.BorderWidth / 2, Shape1.BorderWidth / 2, UserControl.ScaleWidth - Shape1.BorderWidth - 2, UserControl.ScaleHeight - Shape1.BorderWidth - 2
            End If
            If Shape1.BorderStyle <> vbTransparent Then
                DrawRectangle iGraphics, Shape1.BorderColor, Shape1.BorderWidth, Shape1.BorderWidth / 2, Shape1.BorderWidth / 2, UserControl.ScaleWidth - Shape1.BorderWidth - 2, UserControl.ScaleHeight - Shape1.BorderWidth - 2
            End If
        ElseIf iShape = vbShapeSquare Then
            If UserControl.ScaleWidth < UserControl.ScaleHeight Then
                iHeight = UserControl.ScaleWidth
            Else
                iHeight = UserControl.ScaleHeight
            End If
            If iFilled Then
                FillRectangle iGraphics, iFillColor, UserControl.ScaleWidth / 2 - iHeight / 2 + Shape1.BorderWidth / 2, UserControl.ScaleHeight / 2 - iHeight / 2 + Shape1.BorderWidth / 2, iHeight - Shape1.BorderWidth - 2, iHeight - Shape1.BorderWidth - 2
            End If
            If Shape1.BorderStyle <> vbTransparent Then
                DrawRectangle iGraphics, Shape1.BorderColor, Shape1.BorderWidth, UserControl.ScaleWidth / 2 - iHeight / 2 + Shape1.BorderWidth / 2, UserControl.ScaleHeight / 2 - iHeight / 2 + Shape1.BorderWidth / 2, iHeight - Shape1.BorderWidth - 2, iHeight - Shape1.BorderWidth - 2
            End If
        ElseIf iShape = vbShapeRoundedRectangle Then
            If iFilled Then
                FillRoundRect iGraphics, iFillColor, Shape1.BorderWidth / 2 + 0.5, Shape1.BorderWidth / 2 + 0.5, UserControl.ScaleWidth - Shape1.BorderWidth - 1.5 - IIf(Shape1.BorderStyle = vbTransparent, 0.5, 0), UserControl.ScaleHeight - Shape1.BorderWidth - 1.5 - IIf(Shape1.BorderStyle = vbTransparent, 0.5, 0)
            End If
            If Shape1.BorderStyle <> vbTransparent Then
                DrawRoundRect iGraphics, Shape1.BorderColor, Shape1.BorderWidth, Shape1.BorderWidth / 2 + 0.5, Shape1.BorderWidth / 2 + 0.5, UserControl.ScaleWidth - Shape1.BorderWidth - 1.5, UserControl.ScaleHeight - Shape1.BorderWidth - 1.5
            End If
            
        ElseIf iShape = vbShapeRoundedSquare Then
            
        End If
        
        Call GdipDeleteGraphics(iGraphics)
    End If
End Sub

Private Sub FillRectangle(ByVal nGraphics As Long, ByVal nColor As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nAlpha As Single = 100)
    Dim hBrush As Long
    
    If GdipCreateSolidFill(ConvertColor(nColor, nAlpha), hBrush) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingModeAntiAlias)
        GdipFillRectangleI nGraphics, hBrush, x, y, nWidth, nHeight
        Call GdipDeleteBrush(hBrush)
    End If
    
End Sub

Private Sub DrawRectangle(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawnWidth As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nAlpha As Single = 100)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, nAlpha), nDrawnWidth, UnitPixel, hPen) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingModeAntiAlias)
        GdipDrawRectangleI nGraphics, hPen, x, y, nWidth, nHeight
        Call GdipDeletePen(hPen)
    End If
    
End Sub

Private Sub FillEllipse(ByVal nGraphics As Long, ByVal nColor As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nAlpha As Single = 100)
    Dim hBrush As Long
    
    If GdipCreateSolidFill(ConvertColor(nColor, nAlpha), hBrush) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingModeAntiAlias)
        GdipFillEllipseI nGraphics, hBrush, x, y, nWidth, nHeight
        Call GdipDeleteBrush(hBrush)
    End If
End Sub

Private Sub DrawEllipse(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawnWidth As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nAlpha As Single = 100)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, nAlpha), nDrawnWidth, UnitPixel, hPen) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingModeAntiAlias)
        GdipDrawEllipseI nGraphics, hPen, x, y, nWidth, nHeight
        Call GdipDeletePen(hPen)
    End If
    
End Sub

Private Sub FillRoundRect(ByVal nGraphics As Long, ByVal nColor As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nRoundSize As Long = 10, Optional ByVal nAlpha As Single = 100)
    Dim hBrush As Long
    
    If GdipCreateSolidFill(ConvertColor(nColor, nAlpha), hBrush) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingModeAntiAlias)
        
        GdipFillPieI nGraphics, hBrush, x, y, nRoundSize * 2, nRoundSize * 2, 180, 90 ''
        GdipFillRectangleI nGraphics, hBrush, x + nRoundSize - 1, y, nWidth - 2 * nRoundSize + 2, nRoundSize + 2 '
        GdipFillPieI nGraphics, hBrush, x + nWidth - nRoundSize * 2, y, nRoundSize * 2, nRoundSize * 2, 270, 90
        GdipFillRectangleI nGraphics, hBrush, x, y + nRoundSize - 1, nWidth, y + nHeight - nRoundSize * 2
        GdipFillPieI nGraphics, hBrush, x + nWidth - nRoundSize * 2, y + nHeight - nRoundSize * 2, nRoundSize * 2, nRoundSize * 2, 0, 90
        GdipFillPieI nGraphics, hBrush, x, y + nHeight - nRoundSize * 2, nRoundSize * 2, nRoundSize * 2, 90, 90
        GdipFillRectangleI nGraphics, hBrush, x + nRoundSize - 1, y + nHeight - nRoundSize - 2, nWidth - 2 * nRoundSize + 2, nRoundSize + 2
        
        Call GdipDeleteBrush(hBrush)
    End If
    
End Sub

Private Sub DrawRoundRect(ByVal nGraphics As Long, ByVal nColor As Long, ByVal nDrawnWidth As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, Optional ByVal nRoundSize As Long = 10, Optional ByVal nAlpha As Single = 100)
    Dim hPen As Long
    
    If GdipCreatePen1(ConvertColor(nColor, nAlpha), nDrawnWidth, UnitPixel, hPen) = 0 Then
        Call GdipSetSmoothingMode(nGraphics, SmoothingModeAntiAlias)
        
        GdipDrawArcI nGraphics, hPen, x, y, nRoundSize * 2, nRoundSize * 2, 180, 90
        GdipDrawLineI nGraphics, hPen, x + nRoundSize - 1, y, x + nWidth - nRoundSize + 2, y
        GdipDrawArcI nGraphics, hPen, x + nWidth - nRoundSize * 2, y, nRoundSize * 2, nRoundSize * 2, 270, 90
        GdipDrawLineI nGraphics, hPen, x + nWidth, y + nRoundSize - 1, x + nWidth, y + nHeight - nRoundSize + 2
        GdipDrawArcI nGraphics, hPen, x + nWidth - nRoundSize * 2, y + nHeight - nRoundSize * 2, nRoundSize * 2, nRoundSize * 2, 0, 90
        GdipDrawLineI nGraphics, hPen, x + nRoundSize - 1, y + nHeight, x + nWidth - nRoundSize + 2, y + nHeight
        GdipDrawArcI nGraphics, hPen, x, y + nHeight - nRoundSize * 2, nRoundSize * 2, nRoundSize * 2, 90, 90
        GdipDrawLineI nGraphics, hPen, x, y + nRoundSize - 1, x, y + nHeight - nRoundSize + 2
        
        Call GdipDeletePen(hPen)
    End If
    
End Sub

Private Function ConvertColor(nColor As Long, nOpacity As Single) As Long
    Dim BGRA(0 To 3) As Byte
    Dim iColor As Long
    
    TranslateColor nColor, 0&, iColor
    
    BGRA(3) = CByte((nOpacity / 100) * 255)
    BGRA(0) = ((iColor \ &H10000) And &HFF)
    BGRA(1) = ((iColor \ &H100) And &HFF)
    BGRA(2) = (iColor And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function

Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(mGdipToken, GdipStartupInput, ByVal 0)
End Sub

Private Sub Create32BitsBitmap()
    Dim tempBI As BITMAPINFO
    
    If mTesting Then Exit Sub
    
    Destroy32BitsBitmap
    With tempBI.bmiHeader
       .biSize = Len(tempBI.bmiHeader)
       .biBitCount = 32
       .biHeight = UserControl.ScaleHeight
       .biWidth = UserControl.ScaleWidth
       .biPlanes = 1
       .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
    End With
    mDC = CreateCompatibleDC(UserControl.hdc)
    m32BitsBitmap = CreateDIBSection(mDC, tempBI, DIB_RGB_COLORS, ByVal 0, 0, 0)
    mOldBitmap = SelectObject(mDC, m32BitsBitmap)
End Sub

Private Sub Destroy32BitsBitmap()
    If mDC <> 0 Then
        SelectObject mDC, mOldBitmap
        DeleteObject m32BitsBitmap
        DeleteDC mDC
        mDC = 0
        mOldBitmap = 0
    End If
End Sub

Private Sub TerminateGDI()
    Call GdiplusShutdown(mGdipToken)
    mGdipToken = 0
End Sub

Private Sub MakeLayeredTransparency()
    Dim BlendF As BLENDFUNCTION
    Dim winSize As Size
    Dim srcPoint As POINTAPI
    
    If mTesting Then Exit Sub
    
    srcPoint.x = 0
    srcPoint.y = 0
    winSize.cx = UserControl.ScaleWidth
    winSize.cy = UserControl.ScaleHeight
      
    With BlendF
       .AlphaFormat = AC_SRC_ALPHA
       .BlendFlags = 0
       .BlendOp = AC_SRC_OVER
       .SourceConstantAlpha = 255
    End With
      
    Call UpdateLayeredWindow(UserControl.hwnd, UserControl.hdc, ByVal 0&, winSize, mDC, srcPoint, 0, BlendF, ULW_ALPHA)
End Sub

Private Sub RemoveLayeredTransparency()
    Dim BlendF As BLENDFUNCTION
    Dim winSize As Size
    Dim srcPoint As POINTAPI

    If mTesting Then Exit Sub

    srcPoint.x = 0
    srcPoint.y = 0
    winSize.cx = UserControl.ScaleWidth
    winSize.cy = UserControl.ScaleHeight

    With BlendF
       .AlphaFormat = AC_SRC_ALPHA
       .BlendFlags = 0
       .BlendOp = AC_SRC_OVER
       .SourceConstantAlpha = 255
    End With

    Call UpdateLayeredWindow(UserControl.hwnd, UserControl.hdc, ByVal 0&, winSize, 0&, srcPoint, 0, BlendF, ULW_OPAQUE)
End Sub

Private Function IsWindows8OrMore() As Boolean
    Static sValue As Long

    If sValue = 0 Then
        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer
        
        osinfo.dwOSVersionInfoSize = 148
        osinfo.szCSDVersion = Space$(128)
        retvalue = GetVersionEx(osinfo)
        sValue = 1
        If osinfo.dwPlatformID = 2 Then ' if it is NT
            If osinfo.dwMajorVersion > 6 Then
                sValue = 2
            Else
                If osinfo.dwMajorVersion = 6 Then
                    If osinfo.dwMinorVersion >= 2 Then
                        sValue = 2
                    End If
                End If
            End If
        End If
    End If
    
    IsWindows8OrMore = (sValue = 2)
End Function

Private Function InIDE2() As Boolean
    Dim s As String
    Static sValue As Long
    
    If mTesting Then Exit Function
    
    If sValue = 0 Then
        s = Space$(255)
        Call GetModuleFileName(GetModuleHandle(vbNullString), s, Len(s))
        If (UCase$(Trim$(s)) Like "*VB6.EXE*") Then
            sValue = 2
        Else
            sValue = 1
        End If
    End If
    InIDE2 = sValue = 2
End Function
