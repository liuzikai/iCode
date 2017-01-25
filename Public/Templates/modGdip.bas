Attribute VB_Name = "GDIP"
Option Explicit

'vIstaswx GDI+ 声明模块
'vIstaswx GDI+ Declare Module

'vIstaswx 整理扩展
'Extended by vIstaswx

'===========================================
'最后修改：2011/2/8
'Latest edit: 2011/2/8
'
'2011-2-8
'1.增加Gdi+1.1的函数,结构体,枚举和常数的声明
'2.增加GdipSetImageAttributesCachedBackground
'  和GdipTestControl函数声明
'3.修改InitGdiPlus(To)的参数
'4.修正一些bug
'5.格式化了API函数和结构体使之更易读
'6.Enum ImageType -> Enum GdipImageType
'7.增加 NewPointF,NewPointL,NewPointsF,NewPointsL,NewColors 函数
'8.增加 Zero(Point/Rect)(F/L) 0变量
'
'2011-2-7
'1.修正GdipSetLinePresetBlend等4个函数参数声明的错误
'
'2010-6-5:
'1.保存图片过程优化
'2.InitGDIPlus(To) 错误时可选显示错误对话框及退出程序；
'  支持自定义错误对话框内容；增加返回值；增加已经初始化的判断
'3.TerminateGDIPlus(From) 增加已经关闭的判断
'4.删除RtlMoveMemory(CopyMemory)声明；修改CLSIDFromString声明为Private级
'===========================================

'http://vIstaswx..com
'QQ     : 490241327

#Const GdipVersion = 1#

'===================================================================================
'  常用内容
'===================================================================================

'=================================
'== Structures                  ==
'=================================

'=================================
'Point Structure
Public Type POINTL
    X As Long
    Y As Long
End Type

Public Type POINTF
    X As Single
    Y As Single
End Type

'=================================
'Rectange Structure
Public Type RECTL
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Public Type RECTF
    Left   As Single
    Top    As Single
    Right  As Single
    Bottom As Single
End Type

'=================================
'Size Structure
Public Type SIZEL
    cX As Long
    cY As Long
End Type

Public Type SIZEF
    cX As Single
    cY As Single
End Type

'=================================
'Bitmap Structure
Public Type RGBQUAD
    rgbBlue     As Byte
    rgbGreen    As Byte
    rgbRed      As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Public Type BitmapData
    Width       As Long
    Height      As Long
    stride      As Long
    PixelFormat As GpPixelFormat
    scan0       As Long
    Reserved    As Long
End Type

'=================================
'Color Structure
Public Type COLORBYTES
    BlueByte  As Byte
    GreenByte As Byte
    RedByte   As Byte
    AlphaByte As Byte
End Type

Public Type COLORLONG
    longval As Long
End Type

Public Type ColorMap
    oldColor As Long
    newColor As Long
End Type

Public Type ColorMatrix
    m(0 To 4, 0 To 4) As Single
End Type

'=================================
'Path
Public Type PathData
    Count   As Long
    pPoints As Long
    pTypes  As Long
End Type

'=================================
'EnCoder
Public Type Clsid
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type

Public Type EnCoderParameter
    guid           As Clsid
    NumberOfValues As Long

    type           As EnCoderParameterValueType
    Value          As Long
End Type

Public Type EnCoderParameters
    Count     As Long
    Parameter As EnCoderParameter
End Type

'=================================
'== Enums                       ==
'=================================

'=================================
'Pixel
Public Enum GpPixelFormat
    PixelFormat1bppIndexed = &H30101
    PixelFormat4bppIndexed = &H30402
    PixelFormat8bppIndexed = &H30803
    PixelFormat16bppGreyScale = &H101004
    PixelFormat16bppRGB555 = &H21005
    PixelFormat16bppRGB565 = &H21006
    PixelFormat16bppARGB1555 = &H61007
    PixelFormat24bppRGB = &H21808
    PixelFormat32bppRGB = &H22009
    PixelFormat32bppARGB = &H26200A
    PixelFormat32bppPARGB = &HE200B
    PixelFormat48bppRGB = &H10300C
    PixelFormat64bppARGB = &H34400D
    PixelFormat64bppPARGB = &H1C400E
End Enum

'=================================
'Unit
Public Enum GpUnit
    UnitWorld = 0
    UnitDisplay
    UnitPixel
    UnitPoint
    UnitInch
    UnitDocument
    UnitMillimeter
End Enum

'=================================
'Path
Public Enum PathPointType                                                       'GdipGetPathTypes,GdipCreatePath2,GdipCreatePath2I
    PathPointTypeStart = 0
    PathPointTypeLine = 1
    PathPointTypeBezier = 3
    PathPointTypePathTypeMask = &H7
    PathPointTypePathDashMode = &H10
    PathPointTypePathMarker = &H20
    PathPointTypeCloseSubpath = &H80
    PathPointTypeBezier3 = 3
End Enum

'=================================
'Font / String
Public Enum GenericFontFamily
    GenericFontFamilySerif = 0
    GenericFontFamilySansSerif
    GenericFontFamilyMonospace
End Enum

Public Enum FontStyle
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Public Enum StringAlignment
    StringAlignmentNear = 0
    StringAlignmentCenter = 1
    StringAlignmentFar = 2
End Enum

'=================================
'Fill / Wrap
Public Enum FillMode
    FillModeAlternate = 0
    FillModeWinding
End Enum

Public Enum WrapMode
    WrapModeTile = 0
    WrapModeTileFlipX
    WrapModeTileFlipY
    WrapModeTileFlipXY
    WrapModeClamp
End Enum

Public Enum LinearGradientMode
    LinearGradientModeHorizontal = 0
    LinearGradientModeVertical
    LinearGradientModeForwardDiagonal
    LinearGradientModeBackwardDiagonal
End Enum

'=================================
'Quality
Public Enum QualityMode
    QualityModeInvalid = -1
    QualityModeDefault = 0
    QualityModeLow = 1
    QualityModeHigh = 2
End Enum

Public Enum CompositingMode
    CompositingModeSourceOver = 0
    CompositingModeSourceCopy
End Enum

Public Enum CompositingQuality
    CompositingQualityInvalid = QualityModeInvalid
    CompositingQualityDefault = QualityModeDefault
    CompositingQualityHighSpeed = QualityModeLow
    CompositingQualityHighQuality = QualityModeHigh
    CompositingQualityGammaCorrected
    CompositingQualityAssumeLinear
End Enum

Public Enum SmoothingMode
    SmoothingModeInvalid = QualityModeInvalid
    SmoothingModeDefault = QualityModeDefault
    SmoothingModeHighSpeed = QualityModeLow
    SmoothingModeHighQuality = QualityModeHigh
    SmoothingModeNone
    SmoothingModeAntiAlias
#If GdipVersion >= 1.1 Then
    SmoothingModeAntiAlias8x4 = SmoothingModeAntiAlias
    SmoothingModeAntiAlias8x8
#End If
End Enum

Public Enum InterpolationMode
    InterpolationModeInvalid = QualityModeInvalid
    InterpolationModeDefault = QualityModeDefault
    InterpolationModeLowQuality = QualityModeLow
    InterpolationModeHighQuality = QualityModeHigh
    InterpolationModeBilinear
    InterpolationModeBicubic
    InterpolationModeNearestNeighbor
    InterpolationModeHighQualityBilinear
    InterpolationModeHighQualityBicubic
End Enum

Public Enum PixelOffsetMode
    PixelOffsetModeInvalid = QualityModeInvalid
    PixelOffsetModeDefault = QualityModeDefault
    PixelOffsetModeHighSpeed = QualityModeLow
    PixelOffsetModeHighQuality = QualityModeHigh
    PixelOffsetModeNone
    PixelOffsetModeHalf
End Enum

Public Enum TextRenderingHint
    TextRenderingHintSystemDefault = 0                                          ' Glyph with system default rendering hint
    TextRenderingHintSingleBitPerPixelGridFit                                   ' Glyph bitmap with hinting
    TextRenderingHintSingleBitPerPixel                                          ' Glyph bitmap without hinting
    TextRenderingHintAntiAliasGridFit                                           ' Glyph anti-alias bitmap with hinting
    TextRenderingHintAntiAlias                                                  ' Glyph anti-alias bitmap without hinting
    TextRenderingHintClearTypeGridFit                                           ' Glyph CT bitmap with hinting
End Enum

'=================================
'Color Matrix
Public Enum MatrixOrder
    MatrixOrderPrepend = 0
    MatrixOrderAppend = 1
End Enum

Public Enum ColorAdjustType
    ColorAdjustTypeDefault = 0
    ColorAdjustTypeBitmap
    ColorAdjustTypeBrush
    ColorAdjustTypePen
    ColorAdjustTypeText
    ColorAdjustTypeCount
    ColorAdjustTypeAny
End Enum

Public Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = 0
    ColorMatrixFlagsSkipGrays = 1
    ColorMatrixFlagsAltGray = 2
End Enum

Public Enum WarpMode
    WarpModePerspective = 0
    WarpModeBilinear
End Enum

Public Enum CombineMode
    CombineModeReplace = 0
    CombineModeIntersect
    CombineModeUnion
    CombineModeXor
    CombineModeExclude
    CombineModeComplement
End Enum

Public Enum ImageLockMode
    ImageLockModeRead = 1
    ImageLockModeWrite = 2
    ImageLockModeUserInputBuf = 4
End Enum

Public Declare Function GdipGetDC _
    Lib "GDIPlus" (ByVal graphics As Long, _
    hDC As Long) As GpStatus
Public Declare Function GdipReleaseDC _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal hDC As Long) As GpStatus

    '==================================================

Public Declare Function GdipCreateFromHDC _
    Lib "GDIPlus" (ByVal hDC As Long, _
    graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND _
    Lib "GDIPlus" (ByVal hWnd As Long, _
    graphics As Long) As GpStatus
Public Declare Function GdipGetImageGraphicsContext _
    Lib "GDIPlus" (ByVal Image As Long, _
    graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics _
    Lib "GDIPlus" (ByVal graphics As Long) As GpStatus

Public Declare Function GdipGraphicsClear _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal lColor As Long) As GpStatus

Public Declare Function GdipSetCompositingMode _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal CompositingMd As CompositingMode) As GpStatus
Public Declare Function GdipGetCompositingMode _
    Lib "GDIPlus" (ByVal graphics As Long, _
    CompositingMd As CompositingMode) As GpStatus
Public Declare Function GdipSetRenderingOrigin _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal X As Long, _
    ByVal Y As Long) As GpStatus
Public Declare Function GdipGetRenderingOrigin _
    Lib "GDIPlus" (ByVal graphics As Long, _
    X As Long, _
    Y As Long) As GpStatus
Public Declare Function GdipSetCompositingQuality _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal CompositingQlty As CompositingQuality) As GpStatus
Public Declare Function GdipGetCompositingQuality _
    Lib "GDIPlus" (ByVal graphics As Long, _
    CompositingQlty As CompositingQuality) As GpStatus
Public Declare Function GdipSetSmoothingMode _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal SmoothingMd As SmoothingMode) As GpStatus
Public Declare Function GdipGetSmoothingMode _
    Lib "GDIPlus" (ByVal graphics As Long, _
    SmoothingMd As SmoothingMode) As GpStatus
Public Declare Function GdipSetPixelOffsetMode _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal PixOffsetMode As PixelOffsetMode) As GpStatus
Public Declare Function GdipGetPixelOffsetMode _
    Lib "GDIPlus" (ByVal graphics As Long, _
    PixOffsetMode As PixelOffsetMode) As GpStatus
Public Declare Function GdipSetTextRenderingHint _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Mode As TextRenderingHint) As GpStatus
Public Declare Function GdipGetTextRenderingHint _
    Lib "GDIPlus" (ByVal graphics As Long, _
    Mode As TextRenderingHint) As GpStatus
Public Declare Function GdipSetTextContrast _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal contrast As Long) As GpStatus
Public Declare Function GdipGetTextContrast _
    Lib "GDIPlus" (ByVal graphics As Long, _
    contrast As Long) As GpStatus
Public Declare Function GdipSetInterpolationMode _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal interpolation As InterpolationMode) As GpStatus
Public Declare Function GdipGetInterpolationMode _
    Lib "GDIPlus" (ByVal graphics As Long, _
    interpolation As InterpolationMode) As GpStatus

Public Declare Function GdipSetWorldTransform _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetWorldTransform _
    Lib "GDIPlus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipMultiplyWorldTransform _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal matrix As Long, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateWorldTransform _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal dx As Single, _
    ByVal dy As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleWorldTransform _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal sx As Single, _
    ByVal sy As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateWorldTransform _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal angle As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipGetWorldTransform _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetPageTransform _
    Lib "GDIPlus" (ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetPageUnit _
    Lib "GDIPlus" (ByVal graphics As Long, _
    unit As GpUnit) As GpStatus
Public Declare Function GdipGetPageScale _
    Lib "GDIPlus" (ByVal graphics As Long, _
    sScale As Single) As GpStatus
Public Declare Function GdipSetPageUnit _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal unit As GpUnit) As GpStatus
Public Declare Function GdipSetPageScale _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal sScale As Single) As GpStatus
Public Declare Function GdipGetDpiX _
    Lib "GDIPlus" (ByVal graphics As Long, _
    dpi As Single) As GpStatus
Public Declare Function GdipGetDpiY _
    Lib "GDIPlus" (ByVal graphics As Long, _
    dpi As Single) As GpStatus
Public Declare Function GdipTransformPoints _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal destSpace As CoordinateSpace, _
    ByVal srcSpace As CoordinateSpace, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformPointsI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal destSpace As CoordinateSpace, _
    ByVal srcSpace As CoordinateSpace, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformPoints_ _
    Lib "GDIPlus" _
    Alias "GdipTransformPoints" _
    (ByVal graphics As Long, _
    ByVal destSpace As CoordinateSpace, _
    ByVal srcSpace As CoordinateSpace, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformPointsI_ _
    Lib "GDIPlus" _
    Alias "GdipTransformPointsI" _
    (ByVal graphics As Long, _
    ByVal destSpace As CoordinateSpace, _
    ByVal srcSpace As CoordinateSpace, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetNearestColor _
    Lib "GDIPlus" (ByVal graphics As Long, _
    argb As Long) As GpStatus
Public Declare Function GdipCreateHalftonePalette Lib "GDIPlus" () As Long

Public Declare Function GdipSetClipGraphics _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal srcgraphics As Long, _
    ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRect _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRectI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipPath _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Path As Long, _
    ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRegion _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal region As Long, _
    ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipHrgn _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal hRgn As Long, _
    ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipResetClip Lib "GDIPlus" (ByVal graphics As Long) As GpStatus

Public Declare Function GdipTranslateClip _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal dx As Single, _
    ByVal dy As Single) As GpStatus
Public Declare Function GdipTranslateClipI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal dx As Long, _
    ByVal dy As Long) As GpStatus
Public Declare Function GdipGetClip _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal region As Long) As GpStatus
Public Declare Function GdipGetClipBounds _
    Lib "GDIPlus" (ByVal graphics As Long, _
    RECT As RECTF) As GpStatus
Public Declare Function GdipGetClipBoundsI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    RECT As RECTL) As GpStatus

Public Declare Function GdipIsClipEmpty _
    Lib "GDIPlus" (ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipGetVisibleClipBounds _
    Lib "GDIPlus" (ByVal graphics As Long, _
    RECT As RECTF) As GpStatus
Public Declare Function GdipGetVisibleClipBoundsI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    RECT As RECTL) As GpStatus
Public Declare Function GdipIsVisibleClipEmpty _
    Lib "GDIPlus" (ByVal graphics As Long, _
    result As Long) As GpStatus

Public Declare Function GdipIsVisiblePoint _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    result As Long) As GpStatus
Public Declare Function GdipIsVisiblePointI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsVisibleRect _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    result As Long) As GpStatus
Public Declare Function GdipIsVisibleRectI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    result As Long) As GpStatus

Public Declare Function GdipSaveGraphics _
    Lib "GDIPlus" (ByVal graphics As Long, _
    state As Long) As GpStatus
Public Declare Function GdipRestoreGraphics _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal state As Long) As GpStatus
Public Declare Function GdipBeginContainer _
    Lib "GDIPlus" (ByVal graphics As Long, _
    dstRect As RECTF, _
    srcRect As RECTF, _
    ByVal unit As GpUnit, _
    state As Long) As GpStatus
Public Declare Function GdipBeginContainerI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    dstRect As RECTL, _
    srcRect As RECTL, _
    ByVal unit As GpUnit, _
    state As Long) As GpStatus
Public Declare Function GdipBeginContainer2 _
    Lib "GDIPlus" (ByVal graphics As Long, _
    state As Long) As GpStatus
Public Declare Function GdipEndContainer _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal state As Long) As GpStatus

    '==================================================

Public Declare Function GdipDrawLine _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X1 As Single, _
    ByVal Y1 As Single, _
    ByVal X2 As Single, _
    ByVal Y2 As Single) As GpStatus
Public Declare Function GdipDrawLineI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long) As GpStatus
Public Declare Function GdipDrawLines _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawLinesI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawLines_ _
    Lib "GDIPlus" _
    Alias "GdipDrawLines" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawLinesI_ _
    Lib "GDIPlus" _
    Alias "GdipDrawLinesI" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
    '==================================================

Public Declare Function GdipDrawArc _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    ByVal startAngle As Single, _
    ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawArcI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal startAngle As Single, _
    ByVal sweepAngle As Single) As GpStatus

    '==================================================

Public Declare Function GdipDrawBezier _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X1 As Single, _
    ByVal Y1 As Single, _
    ByVal X2 As Single, _
    ByVal Y2 As Single, _
    ByVal x3 As Single, _
    ByVal y3 As Single, _
    ByVal x4 As Single, _
    ByVal y4 As Single) As GpStatus
Public Declare Function GdipDrawBezierI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    ByVal x3 As Long, _
    ByVal y3 As Long, _
    ByVal x4 As Long, _
    ByVal y4 As Long) As GpStatus
Public Declare Function GdipDrawBeziers _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawBeziersI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawBeziers_ _
    Lib "GDIPlus" _
    Alias "GdipDrawBeziers" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawBeziersI_ _
    Lib "GDIPlus" _
    Alias "GdipDrawBeziersI" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
    '==================================================

Public Declare Function GdipDrawRectangle _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawRectangleI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawRectangles _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    rects As RECTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawRectanglesI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    rects As RECTL, _
    ByVal Count As Long) As GpStatus

Public Declare Function GdipFillRectangle _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single) As GpStatus
Public Declare Function GdipFillRectangleI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As GpStatus
Public Declare Function GdipFillRectangles _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    rects As RECTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipFillRectanglesI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    rects As RECTL, _
    ByVal Count As Long) As GpStatus

    '==================================================

Public Declare Function GdipDrawEllipse _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawEllipseI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As GpStatus

Public Declare Function GdipFillEllipse _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single) As GpStatus
Public Declare Function GdipFillEllipseI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As GpStatus

    '==================================================

Public Declare Function GdipDrawPie _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    ByVal startAngle As Single, _
    ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawPieI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal startAngle As Single, _
    ByVal sweepAngle As Single) As GpStatus

Public Declare Function GdipFillPie _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    ByVal startAngle As Single, _
    ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipFillPieI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal startAngle As Single, _
    ByVal sweepAngle As Single) As GpStatus

    '==================================================

Public Declare Function GdipDrawPolygon _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawPolygonI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus

Public Declare Function GdipFillPolygon _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As POINTF, _
    ByVal Count As Long, _
    ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygonI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As POINTL, _
    ByVal Count As Long, _
    ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygon2 _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipFillPolygon2I _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawPolygon_ _
    Lib "GDIPlus" _
    Alias "GdipDrawPolygon" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawPolygonI_ _
    Lib "GDIPlus" _
    Alias "GdipDrawPolygonI" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus

Public Declare Function GdipFillPolygon_ _
    Lib "GDIPlus" _
    Alias "GdipFillPolygon" _
    (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygonI_ _
    Lib "GDIPlus" _
    Alias "GdipFillPolygonI" _
    (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygon2_ _
    Lib "GDIPlus" _
    Alias "GdipFillPolygon2" _
    (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipFillPolygon2I_ _
    Lib "GDIPlus" _
    Alias "GdipFillPolygon2I" _
    (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus

    '==================================================

Public Declare Function GdipDrawPath _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    ByVal Path As Long) As GpStatus

Public Declare Function GdipFillPath _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    ByVal Path As Long) As GpStatus

    '==================================================

Public Declare Function GdipDrawCurve _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurveI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurve2 _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTF, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve2I _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTL, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3 _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTF, _
    ByVal Count As Long, _
    ByVal Offset As Long, _
    ByVal numberOfSegments As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3I _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTL, _
    ByVal Count As Long, _
    ByVal Offset As Long, _
    ByVal numberOfSegments As Long, _
    ByVal tension As Single) As GpStatus

Public Declare Function GdipDrawClosedCurve _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurveI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurve2 _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTF, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawClosedCurve2I _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As POINTL, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus

Public Declare Function GdipFillClosedCurve _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurveI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurve2 _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As POINTF, _
    ByVal Count As Long, _
    ByVal tension As Single, _
    ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillClosedCurve2I _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As POINTL, _
    ByVal Count As Long, _
    ByVal tension As Single, _
    ByVal FillMd As FillMode) As GpStatus

Public Declare Function GdipDrawCurve_ _
    Lib "GDIPlus" _
    Alias "GdipDrawCurve" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurveI_ _
    Lib "GDIPlus" _
    Alias "GdipDrawCurveI" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurve2_ _
    Lib "GDIPlus" _
    Alias "GdipDrawCurve2" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve2I_ _
    Lib "GDIPlus" _
    Alias "GdipDrawCurve2I" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3_ _
    Lib "GDIPlus" _
    Alias "GdipDrawCurve3" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal Offset As Long, _
    ByVal numberOfSegments As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3I_ _
    Lib "GDIPlus" _
    Alias "GdipDrawCurve3I" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal Offset As Long, _
    ByVal numberOfSegments As Long, _
    ByVal tension As Single) As GpStatus

Public Declare Function GdipDrawClosedCurve_ _
    Lib "GDIPlus" _
    Alias "GdipDrawClosedCurve" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurveI_ _
    Lib "GDIPlus" _
    Alias "GdipDrawClosedCurveI" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurve2_ _
    Lib "GDIPlus" _
    Alias "GdipDrawClosedCurve2" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawClosedCurve2I_ _
    Lib "GDIPlus" _
    Alias "GdipDrawClosedCurve2I" _
    (ByVal graphics As Long, _
    ByVal pen As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus

Public Declare Function GdipFillClosedCurve_ _
    Lib "GDIPlus" _
    Alias "GdipFillClosedCurve" _
    (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurveI_ _
    Lib "GDIPlus" _
    Alias "GdipFillClosedCurveI" _
    (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurve2_ _
    Lib "GDIPlus" _
    Alias "GdipFillClosedCurve2" _
    (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal tension As Single, _
    ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillClosedCurve2I_ _
    Lib "GDIPlus" _
    Alias "GdipFillClosedCurve2I" _
    (ByVal graphics As Long, _
    ByVal brush As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal tension As Single, _
    ByVal FillMd As FillMode) As GpStatus


    '==================================================

Public Declare Function GdipFillRegion _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal brush As Long, _
    ByVal region As Long) As GpStatus

    '==================================================

Public Declare Function GdipDrawImage _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    ByVal X As Single, _
    ByVal Y As Single) As GpStatus
Public Declare Function GdipDrawImageI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    ByVal X As Long, _
    ByVal Y As Long) As GpStatus

Public Declare Function GdipDrawImageRect _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawImageRectI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawImagePoints _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    dstpoints As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointsI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    dstpoints As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePoints_ _
    Lib "GDIPlus" _
    Alias "GdipDrawImagePoints" _
    (ByVal graphics As Long, _
    ByVal Image As Long, _
    dstpoints As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointsI_ _
    Lib "GDIPlus" _
    Alias "GdipDrawImagePointsI" _
    (ByVal graphics As Long, _
    ByVal Image As Long, _
    dstpoints As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointRect _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal srcx As Single, _
    ByVal srcy As Single, _
    ByVal srcwidth As Single, _
    ByVal srcheight As Single, _
    ByVal srcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImagePointRectI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal srcx As Long, _
    ByVal srcy As Long, _
    ByVal srcwidth As Long, _
    ByVal srcheight As Long, _
    ByVal srcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImagePointsRect _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    Points As POINTF, _
    ByVal Count As Long, _
    ByVal srcx As Single, _
    ByVal srcy As Single, _
    ByVal srcwidth As Single, _
    ByVal srcheight As Single, _
    ByVal srcUnit As GpUnit, _
    Optional ByVal imageAttributes As Long = 0, _
    Optional ByVal callback As Long = 0, _
    Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRectI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    Points As POINTL, _
    ByVal Count As Long, _
    ByVal srcx As Long, _
    ByVal srcy As Long, _
    ByVal srcwidth As Long, _
    ByVal srcheight As Long, _
    ByVal srcUnit As GpUnit, _
    Optional ByVal imageAttributes As Long = 0, _
    Optional ByVal callback As Long = 0, _
                              Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRect_ _
    Lib "GDIPlus" _
    Alias "GdipDrawImagePointsRect" _
    (ByVal graphics As Long, _
    ByVal Image As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal srcx As Single, _
    ByVal srcy As Single, _
    ByVal srcwidth As Single, _
    ByVal srcheight As Single, _
    ByVal srcUnit As GpUnit, _
    Optional ByVal imageAttributes As Long = 0, _
    Optional ByVal callback As Long = 0, _
                              Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRectI_ _
    Lib "GDIPlus" _
    Alias "GdipDrawImagePointsRectI" _
    (ByVal graphics As Long, _
    ByVal Image As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal srcx As Long, _
    ByVal srcy As Long, _
    ByVal srcwidth As Long, _
    ByVal srcheight As Long, _
    ByVal srcUnit As GpUnit, _
    Optional ByVal imageAttributes As Long = 0, _
    Optional ByVal callback As Long = 0, _
    Optional ByVal callbackData As Long = 0) As GpStatus

Public Declare Function GdipGetImageDecoders _
    Lib "GDIPlus" (ByVal numDecoders As Long, _
    ByVal size As Long, _
    decoders As Any) As GpStatus
Public Declare Function GdipGetImageEnCodersSize _
    Lib "GDIPlus" (numEnCoders As Long, _
    size As Long) As GpStatus
Public Declare Function GdipGetImageEnCoders _
    Lib "GDIPlus" (ByVal numEnCoders As Long, _
    ByVal size As Long, _
    enCoders As Any) As GpStatus
Public Declare Function GdipComment _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal sizeData As Long, _
    data As Any) As GpStatus

Public Declare Function GdipLoadImageFromFile _
    Lib "GDIPlus" (ByVal filename As Long, _
    Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromFileICM _
    Lib "GDIPlus" (ByVal filename As Long, _
    Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromStream _
    Lib "GDIPlus" (ByVal stream As Any, _
    Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromStreamICM _
    Lib "GDIPlus" (ByVal stream As Any, _
    Image As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As GpStatus
Public Declare Function GdipCloneImage _
    Lib "GDIPlus" (ByVal Image As Long, _
    cloneImage As Long) As GpStatus

Public Declare Function GdipSaveImageToFile _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal filename As Long, _
    clsidEnCoder As Clsid, _
    enCoderParams As Any) As GpStatus
Public Declare Function GdipSaveImageToStream _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal stream As Any, _
    clsidEnCoder As Clsid, _
    enCoderParams As Any) As GpStatus

Public Declare Function GdipSaveAdd _
    Lib "GDIPlus" (ByVal Image As Long, _
    enCoderParams As EnCoderParameters) As GpStatus
Public Declare Function GdipSaveAddImage _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal newImage As Long, _
    enCoderParams As EnCoderParameters) As GpStatus

Public Declare Function GdipGetImageBounds _
    Lib "GDIPlus" (ByVal Image As Long, _
    srcRect As RECTF, _
    srcUnit As GpUnit) As GpStatus
Public Declare Function GdipGetImageDimension _
    Lib "GDIPlus" (ByVal Image As Long, _
    Width As Single, _
    Height As Single) As GpStatus
Public Declare Function GdipGetImageType _
    Lib "GDIPlus" (ByVal Image As Long, _
    itype As Image_Type) As GpStatus
Public Declare Function GdipGetImageWidth _
    Lib "GDIPlus" (ByVal Image As Long, _
    Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight _
    Lib "GDIPlus" (ByVal Image As Long, _
    Height As Long) As GpStatus
Public Declare Function GdipGetImageHorizontalResolution _
    Lib "GDIPlus" (ByVal Image As Long, _
    resolution As Single) As GpStatus
Public Declare Function GdipGetImageVerticalResolution _
    Lib "GDIPlus" (ByVal Image As Long, _
    resolution As Single) As GpStatus
Public Declare Function GdipGetImageFlags _
    Lib "GDIPlus" (ByVal Image As Long, _
    flags As Long) As GpStatus
Public Declare Function GdipGetImageRawFormat _
    Lib "GDIPlus" (ByVal Image As Long, _
    format As Clsid) As GpStatus
Public Declare Function GdipGetImagePixelFormat _
    Lib "GDIPlus" (ByVal Image As Long, _
    PixelFormat As GpPixelFormat) As GpStatus
Public Declare Function GdipGetImageThumbnail _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal thumbWidth As Long, _
    ByVal thumbHeight As Long, _
    thumbImage As Long, _
    Optional ByVal callback As Long = 0, _
    Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipGetEnCoderParameterListSize _
    Lib "GDIPlus" (ByVal Image As Long, _
    clsidEnCoder As Clsid, _
    size As Long) As GpStatus
Public Declare Function GdipGetEnCoderParameterList _
    Lib "GDIPlus" (ByVal Image As Long, _
    clsidEnCoder As Clsid, _
    ByVal size As Long, _
    buffer As EnCoderParameters) As GpStatus

Public Declare Function GdipImageGetFrameDimensionsCount _
    Lib "GDIPlus" (ByVal Image As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipImageGetFrameDimensionsList _
    Lib "GDIPlus" (ByVal Image As Long, _
    dimensionIDs As Clsid, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipImageGetFrameCount _
    Lib "GDIPlus" (ByVal Image As Long, _
    dimensionID As Clsid, _
    Count As Long) As GpStatus
Public Declare Function GdipImageSelectActiveFrame _
    Lib "GDIPlus" (ByVal Image As Long, _
    dimensionID As Clsid, _
    ByVal frameIndex As Long) As GpStatus
Public Declare Function GdipImageRotateFlip _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal rfType As RotateFlipType) As GpStatus
Public Declare Function GdipGetImagePalette _
    Lib "GDIPlus" (ByVal Image As Long, _
    palette As ColorPalette, _
    ByVal size As Long) As GpStatus
Public Declare Function GdipSetImagePalette _
    Lib "GDIPlus" (ByVal Image As Long, _
    palette As ColorPalette) As GpStatus
Public Declare Function GdipGetImagePaletteSize _
    Lib "GDIPlus" (ByVal Image As Long, _
    size As Long) As GpStatus
Public Declare Function GdipGetPropertyCount _
    Lib "GDIPlus" (ByVal Image As Long, _
    numOfProperty As Long) As GpStatus
Public Declare Function GdipGetPropertyIdList _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal numOfProperty As Long, _
    list As Long) As GpStatus
Public Declare Function GdipGetPropertyItemSize _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal propId As Long, _
    size As Long) As GpStatus
Public Declare Function GdipGetPropertyItem _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal propId As Long, _
    ByVal propSize As Long, _
    buffer As PropertyItem) As GpStatus
Public Declare Function GdipGetPropertySize _
    Lib "GDIPlus" (ByVal Image As Long, _
    totalBufferSize As Long, _
    numProperties As Long) As GpStatus
Public Declare Function GdipGetAllPropertyItems _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal totalBufferSize As Long, _
    ByVal numProperties As Long, _
    allItems As PropertyItem) As GpStatus
Public Declare Function GdipRemovePropertyItem _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal propId As Long) As GpStatus
Public Declare Function GdipSetPropertyItem _
    Lib "GDIPlus" (ByVal Image As Long, _
    item As PropertyItem) As GpStatus
Public Declare Function GdipImageForceValidation _
    Lib "GDIPlus" (ByVal Image As Long) As GpStatus

    '==================================================

Public Declare Function GdipCreatePen1 _
    Lib "GDIPlus" (ByVal Color As Long, _
    ByVal Width As Single, _
    ByVal unit As GpUnit, _
    pen As Long) As GpStatus
Public Declare Function GdipCreatePen2 _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal Width As Single, _
    ByVal unit As GpUnit, _
    pen As Long) As GpStatus
Public Declare Function GdipClonePen _
    Lib "GDIPlus" (ByVal pen As Long, _
    clonepen As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "GDIPlus" (ByVal pen As Long) As GpStatus

Public Declare Function GdipSetPenWidth _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal Width As Single) As GpStatus
Public Declare Function GdipGetPenWidth _
    Lib "GDIPlus" (ByVal pen As Long, _
    Width As Single) As GpStatus
Public Declare Function GdipSetPenUnit _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal unit As GpUnit) As GpStatus
Public Declare Function GdipGetPenUnit _
    Lib "GDIPlus" (ByVal pen As Long, _
    unit As GpUnit) As GpStatus

Public Declare Function GdipSetPenLineCap _
    Lib "GDIPlus" _
    Alias "GdipSetPenLineCap197819" (ByVal pen As Long, _
    ByVal startCap As LineCap, _
    ByVal endCap As LineCap, _
    ByVal dcap As DashCap) As GpStatus
Public Declare Function GdipSetPenStartCap _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal startCap As LineCap) As GpStatus
Public Declare Function GdipSetPenEndCap _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal endCap As LineCap) As GpStatus
Public Declare Function GdipSetPenDashCap _
    Lib "GDIPlus" _
    Alias "GdipSetPenDashCap197819" (ByVal pen As Long, _
    ByVal dcap As DashCap) As GpStatus
Public Declare Function GdipGetPenStartCap _
    Lib "GDIPlus" (ByVal pen As Long, _
    startCap As LineCap) As GpStatus
Public Declare Function GdipGetPenEndCap _
    Lib "GDIPlus" (ByVal pen As Long, _
    endCap As LineCap) As GpStatus
Public Declare Function GdipGetPenDashCap _
    Lib "GDIPlus" _
    Alias "GdipGetPenDashCap197819" (ByVal pen As Long, _
    dcap As DashCap) As GpStatus
Public Declare Function GdipSetPenLineJoin _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal lnJoin As LineJoin) As GpStatus
Public Declare Function GdipGetPenLineJoin _
    Lib "GDIPlus" (ByVal pen As Long, _
    lnJoin As LineJoin) As GpStatus
Public Declare Function GdipSetPenCustomStartCap _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal customCap As Long) As GpStatus
Public Declare Function GdipGetPenCustomStartCap _
    Lib "GDIPlus" (ByVal pen As Long, _
    customCap As Long) As GpStatus
Public Declare Function GdipSetPenCustomEndCap _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal customCap As Long) As GpStatus
Public Declare Function GdipGetPenCustomEndCap _
    Lib "GDIPlus" (ByVal pen As Long, _
    customCap As Long) As GpStatus

Public Declare Function GdipSetPenMiterLimit _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal miterLimit As Single) As GpStatus
Public Declare Function GdipGetPenMiterLimit _
    Lib "GDIPlus" (ByVal pen As Long, _
    miterLimit As Single) As GpStatus
Public Declare Function GdipSetPenMode _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal penMode As PenAlignment) As GpStatus
Public Declare Function GdipGetPenMode _
    Lib "GDIPlus" (ByVal pen As Long, _
    penMode As PenAlignment) As GpStatus
Public Declare Function GdipSetPenTransform _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetPenTransform _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetPenTransform _
    Lib "GDIPlus" (ByVal pen As Long) As GpStatus
Public Declare Function GdipMultiplyPenTransform _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal matrix As Long, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslatePenTransform _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal dx As Single, _
    ByVal dy As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScalePenTransform _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal sx As Single, _
    ByVal sy As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotatePenTransform _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal angle As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipSetPenColor _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal argb As Long) As GpStatus
Public Declare Function GdipGetPenColor _
    Lib "GDIPlus" (ByVal pen As Long, _
    argb As Long) As GpStatus
Public Declare Function GdipSetPenBrushFill _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal brush As Long) As GpStatus
Public Declare Function GdipGetPenBrushFill _
    Lib "GDIPlus" (ByVal pen As Long, _
    brush As Long) As GpStatus
Public Declare Function GdipGetPenFillType _
    Lib "GDIPlus" (ByVal pen As Long, _
    ptype As PenType) As GpStatus
Public Declare Function GdipGetPenDashStyle _
    Lib "GDIPlus" (ByVal pen As Long, _
    dStyle As DashStyle) As GpStatus
Public Declare Function GdipSetPenDashStyle _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal dStyle As DashStyle) As GpStatus
Public Declare Function GdipGetPenDashOffset _
    Lib "GDIPlus" (ByVal pen As Long, _
    Offset As Single) As GpStatus
Public Declare Function GdipSetPenDashOffset _
    Lib "GDIPlus" (ByVal pen As Long, _
    ByVal Offset As Single) As GpStatus
Public Declare Function GdipGetPenDashCount _
    Lib "GDIPlus" (ByVal pen As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipSetPenDashArray _
    Lib "GDIPlus" (ByVal pen As Long, _
    dash As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenDashArray _
    Lib "GDIPlus" (ByVal pen As Long, _
    dash As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenCompoundCount _
    Lib "GDIPlus" (ByVal pen As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipSetPenCompoundArray _
    Lib "GDIPlus" (ByVal pen As Long, _
    dash As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenCompoundArray _
    Lib "GDIPlus" (ByVal pen As Long, _
    dash As Single, _
    ByVal Count As Long) As GpStatus

Public Declare Function GdipCreateCustomLineCap _
    Lib "GDIPlus" (ByVal fillPath As Long, _
    ByVal strokePath As Long, _
    ByVal baseCap As LineCap, _
    ByVal baseInset As Single, _
    customCap As Long) As GpStatus
Public Declare Function GdipDeleteCustomLineCap _
    Lib "GDIPlus" (ByVal customCap As Long) As GpStatus
Public Declare Function GdipCloneCustomLineCap _
    Lib "GDIPlus" (ByVal customCap As Long, _
    clonedCap As Long) As GpStatus
Public Declare Function GdipGetCustomLineCapType _
    Lib "GDIPlus" (ByVal customCap As Long, _
    capType As CustomLineCapType) As GpStatus
Public Declare Function GdipSetCustomLineCapStrokeCaps _
    Lib "GDIPlus" (ByVal customCap As Long, _
    ByVal startCap As LineCap, _
    ByVal endCap As LineCap) As GpStatus
Public Declare Function GdipGetCustomLineCapStrokeCaps _
    Lib "GDIPlus" (ByVal customCap As Long, _
    startCap As LineCap, _
    endCap As LineCap) As GpStatus
Public Declare Function GdipSetCustomLineCapStrokeJoin _
    Lib "GDIPlus" (ByVal customCap As Long, _
    ByVal lnJoin As LineJoin) As GpStatus
Public Declare Function GdipGetCustomLineCapStrokeJoin _
    Lib "GDIPlus" (ByVal customCap As Long, _
    lnJoin As LineJoin) As GpStatus
Public Declare Function GdipSetCustomLineCapBaseCap _
    Lib "GDIPlus" (ByVal customCap As Long, _
    ByVal baseCap As LineCap) As GpStatus
Public Declare Function GdipGetCustomLineCapBaseCap _
    Lib "GDIPlus" (ByVal customCap As Long, _
    baseCap As LineCap) As GpStatus
Public Declare Function GdipSetCustomLineCapBaseInset _
    Lib "GDIPlus" (ByVal customCap As Long, _
    ByVal inset As Single) As GpStatus
Public Declare Function GdipGetCustomLineCapBaseInset _
    Lib "GDIPlus" (ByVal customCap As Long, _
    inset As Single) As GpStatus
Public Declare Function GdipSetCustomLineCapWidthScale _
    Lib "GDIPlus" (ByVal customCap As Long, _
    ByVal widthScale As Single) As GpStatus
Public Declare Function GdipGetCustomLineCapWidthScale _
    Lib "GDIPlus" (ByVal customCap As Long, _
    widthScale As Single) As GpStatus

Public Declare Function GdipCreateAdjustableArrowCap _
    Lib "GDIPlus" (ByVal Height As Single, _
    ByVal Width As Single, _
    ByVal isFilled As Long, _
    cap As Long) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapHeight _
    Lib "GDIPlus" (ByVal cap As Long, _
    ByVal Height As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapHeight _
    Lib "GDIPlus" (ByVal cap As Long, _
    Height As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapWidth _
    Lib "GDIPlus" (ByVal cap As Long, _
    ByVal Width As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapWidth _
    Lib "GDIPlus" (ByVal cap As Long, _
    Width As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapMiddleInset _
    Lib "GDIPlus" (ByVal cap As Long, _
    ByVal middleInset As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapMiddleInset _
    Lib "GDIPlus" (ByVal cap As Long, _
    middleInset As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapFillState _
    Lib "GDIPlus" (ByVal cap As Long, _
    ByVal bFillState As Long) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapFillState _
    Lib "GDIPlus" (ByVal cap As Long, _
    bFillState As Long) As GpStatus

    '==================================================

Public Declare Function GdipCreateBitmapFromFile _
    Lib "GDIPlus" (ByVal filename As Long, _
    Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFileICM _
    Lib "GDIPlus" (ByVal filename As Long, _
    Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromStream _
    Lib "GDIPlus" (ByVal stream As Any, _
    Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromStreamICM _
    Lib "GDIPlus" (ByVal stream As Any, _
    Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromScan0 _
    Lib "GDIPlus" (ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal stride As Long, _
    ByVal PixelFormat As GpPixelFormat, _
    scan0 As Any, _
    Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGraphics _
    Lib "GDIPlus" (ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal graphics As Long, _
    Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGdiDib _
    Lib "GDIPlus" (gdiBitmapInfo As BITMAPINFO, _
    ByVal gdiBitmapData As Long, _
    Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP _
    Lib "GDIPlus" (ByVal hbm As Long, _
    ByVal hpal As Long, _
    Bitmap As Long) As GpStatus
Public Declare Function GdipCreateHBITMAPFromBitmap _
    Lib "GDIPlus" (ByVal Bitmap As Long, _
    hbmReturn As Long, _
    ByVal background As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHICON _
    Lib "GDIPlus" (ByVal hicon As Long, _
    Bitmap As Long) As GpStatus
Public Declare Function GdipCreateHICONFromBitmap _
    Lib "GDIPlus" (ByVal Bitmap As Long, _
    hbmReturn As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromResource _
    Lib "GDIPlus" (ByVal hInstance As Long, _
    ByVal lpBitmapName As Long, _
    Bitmap As Long) As GpStatus

Public Declare Function GdipCloneBitmapArea _
    Lib "GDIPlus" (ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    ByVal PixelFormat As GpPixelFormat, _
    ByVal srcBitmap As Long, _
    dstBitmap As Long) As GpStatus
Public Declare Function GdipCloneBitmapAreaI _
    Lib "GDIPlus" (ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal PixelFormat As GpPixelFormat, _
    ByVal srcBitmap As Long, _
    dstBitmap As Long) As GpStatus

Public Declare Function GdipBitmapLockBits _
    Lib "GDIPlus" (ByVal Bitmap As Long, _
    RECT As RECTL, _
    ByVal flags As ImageLockMode, _
    ByVal PixelFormat As GpPixelFormat, _
    lockedBitmapData As BitmapData) As GpStatus
Public Declare Function GdipBitmapUnlockBits _
    Lib "GDIPlus" (ByVal Bitmap As Long, _
    lockedBitmapData As BitmapData) As GpStatus

Public Declare Function GdipBitmapGetPixel _
    Lib "GDIPlus" (ByVal Bitmap As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    Color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel _
    Lib "GDIPlus" (ByVal Bitmap As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Color As Long) As GpStatus

Public Declare Function GdipBitmapSetResolution _
    Lib "GDIPlus" (ByVal Bitmap As Long, _
    ByVal xdpi As Single, _
    ByVal ydpi As Single) As GpStatus

Public Declare Function GdipCreateCachedBitmap _
    Lib "GDIPlus" (ByVal Bitmap As Long, _
    ByVal graphics As Long, _
    cachedBitmap As Long) As GpStatus
Public Declare Function GdipDeleteCachedBitmap _
    Lib "GDIPlus" (ByVal cachedBitmap As Long) As GpStatus
Public Declare Function GdipDrawCachedBitmap _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal cachedBitmap As Long, _
    ByVal X As Long, _
    ByVal Y As Long) As GpStatus

    '==================================================

Public Declare Function GdipCloneBrush _
    Lib "GDIPlus" (ByVal brush As Long, _
    cloneBrush As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "GDIPlus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipGetBrushType _
    Lib "GDIPlus" (ByVal brush As Long, _
    brshType As BrushType) As GpStatus
Public Declare Function GdipCreateHatchBrush _
    Lib "GDIPlus" (ByVal Style As HatchStyle, _
    ByVal forecolr As Long, _
    ByVal backcolr As Long, _
    brush As Long) As GpStatus
Public Declare Function GdipGetHatchStyle _
    Lib "GDIPlus" (ByVal brush As Long, _
    Style As HatchStyle) As GpStatus
Public Declare Function GdipGetHatchForegroundColor _
    Lib "GDIPlus" (ByVal brush As Long, _
    forecolr As Long) As GpStatus
Public Declare Function GdipGetHatchBackgroundColor _
    Lib "GDIPlus" (ByVal brush As Long, _
    backcolr As Long) As GpStatus
Public Declare Function GdipCreateSolidFill _
    Lib "GDIPlus" (ByVal argb As Long, _
    brush As Long) As GpStatus
Public Declare Function GdipSetSolidFillColor _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal argb As Long) As GpStatus
Public Declare Function GdipGetSolidFillColor _
    Lib "GDIPlus" (ByVal brush As Long, _
    argb As Long) As GpStatus
Public Declare Function GdipCreateLineBrush _
    Lib "GDIPlus" (Point1 As POINTF, _
    Point2 As POINTF, _
    ByVal color1 As Long, _
    ByVal color2 As Long, _
    ByVal WrapMd As WrapMode, _
    lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushI _
    Lib "GDIPlus" (Point1 As POINTL, _
    Point2 As POINTL, _
    ByVal color1 As Long, _
    ByVal color2 As Long, _
    ByVal WrapMd As WrapMode, _
    lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRect _
    Lib "GDIPlus" (RECT As RECTF, _
    ByVal color1 As Long, _
    ByVal color2 As Long, _
    ByVal Mode As LinearGradientMode, _
    ByVal WrapMd As WrapMode, _
    lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectI _
    Lib "GDIPlus" (RECT As RECTL, _
    ByVal color1 As Long, _
    ByVal color2 As Long, _
    ByVal Mode As LinearGradientMode, _
    ByVal WrapMd As WrapMode, _
    lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectWithAngle _
    Lib "GDIPlus" (RECT As RECTF, _
    ByVal color1 As Long, _
    ByVal color2 As Long, _
    ByVal angle As Single, _
    ByVal isAngleScalable As Long, _
    ByVal WrapMd As WrapMode, _
    lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectWithAngleI _
    Lib "GDIPlus" (RECT As RECTL, _
    ByVal color1 As Long, _
    ByVal color2 As Long, _
    ByVal angle As Single, _
    ByVal isAngleScalable As Long, _
    ByVal WrapMd As WrapMode, _
    lineGradient As Long) As GpStatus
Public Declare Function GdipSetLineColors _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal color1 As Long, _
    ByVal color2 As Long) As GpStatus
Public Declare Function GdipGetLineColors _
    Lib "GDIPlus" (ByVal brush As Long, _
    lColors As Long) As GpStatus
Public Declare Function GdipGetLineRect _
    Lib "GDIPlus" (ByVal brush As Long, _
    RECT As RECTF) As GpStatus
Public Declare Function GdipGetLineRectI _
    Lib "GDIPlus" (ByVal brush As Long, _
    RECT As RECTL) As GpStatus
Public Declare Function GdipSetLineGammaCorrection _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetLineGammaCorrection _
    Lib "GDIPlus" (ByVal brush As Long, _
    useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetLineBlendCount _
    Lib "GDIPlus" (ByVal brush As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipGetLineBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    blend As Long, _
    positions As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLineBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    blend As Long, _
    positions As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetLineBlend_ _
    Lib "GDIPlus" _
    Alias "GdipGetLineBlend" _
    (ByVal brush As Long, _
    blend As Any, _
    positions As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLineBlend_ _
    Lib "GDIPlus" _
    Alias "GdipSetLineBlend" _
    (ByVal brush As Long, _
    blend As Any, _
    positions As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlendCount _
    Lib "GDIPlus" (ByVal brush As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    blend As Long, _
    positions As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLinePresetBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    blend As Long, _
    positions As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlend_ _
    Lib "GDIPlus" _
    Alias "GdipGetLinePresetBlend" _
    (ByVal brush As Long, _
    blend As Any, _
    positions As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLinePresetBlend_ _
    Lib "GDIPlus" _
    Alias "GdipSetLinePresetBlend" _
    (ByVal brush As Long, _
    blend As Any, _
    positions As Any, _
    ByVal Count As Long) As GpStatus

Public Declare Function GdipSetLineSigmaBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal focus As Single, _
    ByVal theScale As Single) As GpStatus
Public Declare Function GdipSetLineLinearBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal focus As Single, _
    ByVal theScale As Single) As GpStatus
Public Declare Function GdipSetLineWrapMode _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetLineWrapMode _
    Lib "GDIPlus" (ByVal brush As Long, _
    WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetLineTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    matrix As Long) As GpStatus
Public Declare Function GdipSetLineTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetLineTransform _
    Lib "GDIPlus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipMultiplyLineTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal matrix As Long, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateLineTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal dx As Single, _
    ByVal dy As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleLineTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal sx As Single, _
    ByVal sy As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateLineTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal angle As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipCreateTexture _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal WrapMd As WrapMode, _
    texture As Long) As GpStatus
Public Declare Function GdipCreateTexture2 _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal WrapMd As WrapMode, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    texture As Long) As GpStatus
Public Declare Function GdipCreateTextureIA _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal imageAttributes As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    texture As Long) As GpStatus
Public Declare Function GdipCreateTexture2I _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal WrapMd As WrapMode, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    texture As Long) As GpStatus
Public Declare Function GdipCreateTextureIAI _
    Lib "GDIPlus" (ByVal Image As Long, _
    ByVal imageAttributes As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    texture As Long) As GpStatus
Public Declare Function GdipGetTextureTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipSetTextureTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetTextureTransform _
    Lib "GDIPlus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipTranslateTextureTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal dx As Single, _
    ByVal dy As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipMultiplyTextureTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal matrix As Long, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleTextureTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal sx As Single, _
    ByVal sy As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateTextureTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal angle As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipSetTextureWrapMode _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetTextureWrapMode _
    Lib "GDIPlus" (ByVal brush As Long, _
    WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetTextureImage _
    Lib "GDIPlus" (ByVal brush As Long, _
    Image As Long) As GpStatus
Public Declare Function GdipCreatePathGradient _
    Lib "GDIPlus" (Points As POINTF, _
    ByVal Count As Long, _
    ByVal WrapMd As WrapMode, _
    polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientI _
    Lib "GDIPlus" (Points As POINTL, _
    ByVal Count As Long, _
    ByVal WrapMd As WrapMode, _
    polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradient_ _
    Lib "GDIPlus" _
    Alias "GdipCreatePathGradient" _
    (Points As Any, _
    ByVal Count As Long, _
    ByVal WrapMd As WrapMode, _
    polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientI_ _
    Lib "GDIPlus" _
    Alias "GdipCreatePathGradientI" _
    (Points As Any, _
    ByVal Count As Long, _
    ByVal WrapMd As WrapMode, _
    polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientFromPath _
    Lib "GDIPlus" (ByVal Path As Long, _
    polyGradient As Long) As GpStatus
Public Declare Function GdipGetPathGradientCenterColor _
    Lib "GDIPlus" (ByVal brush As Long, _
    lColors As Long) As GpStatus
Public Declare Function GdipSetPathGradientCenterColor _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal lColors As Long) As GpStatus
Public Declare Function GdipGetPathGradientSurroundColorsWithCount _
    Lib "GDIPlus" (ByVal brush As Long, _
    argb As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientSurroundColorsWithCount _
    Lib "GDIPlus" (ByVal brush As Long, _
    argb As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPath _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal Path As Long) As GpStatus
Public Declare Function GdipSetPathGradientPath _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal Path As Long) As GpStatus
Public Declare Function GdipGetPathGradientCenterPoint _
    Lib "GDIPlus" (ByVal brush As Long, _
    Points As POINTF) As GpStatus
Public Declare Function GdipGetPathGradientCenterPointI _
    Lib "GDIPlus" (ByVal brush As Long, _
    Points As POINTL) As GpStatus
Public Declare Function GdipSetPathGradientCenterPoint _
    Lib "GDIPlus" (ByVal brush As Long, _
    Points As POINTF) As GpStatus
Public Declare Function GdipSetPathGradientCenterPointI _
    Lib "GDIPlus" (ByVal brush As Long, _
    Points As POINTL) As GpStatus
Public Declare Function GdipGetPathGradientRect _
    Lib "GDIPlus" (ByVal brush As Long, _
    RECT As RECTF) As GpStatus
Public Declare Function GdipGetPathGradientRectI _
    Lib "GDIPlus" (ByVal brush As Long, _
    RECT As RECTL) As GpStatus
Public Declare Function GdipGetPathGradientPointCount _
    Lib "GDIPlus" (ByVal brush As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientSurroundColorCount _
    Lib "GDIPlus" (ByVal brush As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientGammaCorrection _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetPathGradientGammaCorrection _
    Lib "GDIPlus" (ByVal brush As Long, _
    useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlendCount _
    Lib "GDIPlus" (ByVal brush As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    blend As Long, _
    positions As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    blend As Long, _
    positions As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlend_ _
    Lib "GDIPlus" _
    Alias "GdipGetPathGradientBlend" _
    (ByVal brush As Long, _
    blend As Any, _
    positions As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientBlend_ _
    Lib "GDIPlus" _
    Alias "GdipSetPathGradientBlend" _
    (ByVal brush As Long, _
    blend As Any, _
    positions As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlendCount _
    Lib "GDIPlus" (ByVal brush As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    blend As Long, _
    positions As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientPresetBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    blend As Long, _
    positions As Single, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlend_ _
    Lib "GDIPlus" _
    Alias "GdipGetPathGradientPresetBlend" _
    (ByVal brush As Long, _
    blend As Any, _
    positions As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientPresetBlend_ _
    Lib "GDIPlus" _
    Alias "GdipSetPathGradientPresetBlend" _
    (ByVal brush As Long, _
    blend As Any, _
    positions As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientSigmaBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal focus As Single, _
    ByVal sScale As Single) As GpStatus
Public Declare Function GdipSetPathGradientLinearBlend _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal focus As Single, _
    ByVal sScale As Single) As GpStatus
Public Declare Function GdipGetPathGradientWrapMode _
    Lib "GDIPlus" (ByVal brush As Long, _
    WrapMd As WrapMode) As GpStatus
Public Declare Function GdipSetPathGradientWrapMode _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetPathGradientTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipSetPathGradientTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetPathGradientTransform _
    Lib "GDIPlus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipMultiplyPathGradientTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal matrix As Long, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslatePathGradientTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal dx As Single, _
    ByVal dy As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScalePathGradientTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal sx As Single, _
    ByVal sy As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotatePathGradientTransform _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal angle As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipGetPathGradientFocusScales _
    Lib "GDIPlus" (ByVal brush As Long, _
    xScale As Single, _
    yScale As Single) As GpStatus
Public Declare Function GdipSetPathGradientFocusScales _
    Lib "GDIPlus" (ByVal brush As Long, _
    ByVal xScale As Single, _
    ByVal yScale As Single) As GpStatus
Public Declare Function GdipCreatePath _
    Lib "GDIPlus" (ByVal brushmode As FillMode, _
    Path As Long) As GpStatus
Public Declare Function GdipCreatePath2 _
    Lib "GDIPlus" (Points As POINTF, _
    types As Any, _
    ByVal Count As Long, _
    brushmode As FillMode, _
    Path As Long) As GpStatus
Public Declare Function GdipCreatePath2I _
    Lib "GDIPlus" (Points As POINTL, _
    types As Any, _
    ByVal Count As Long, _
    brushmode As FillMode, _
    Path As Long) As GpStatus
Public Declare Function GdipCreatePath2_ _
    Lib "GDIPlus" _
    Alias "GdipCreatePath2" _
    (Points As Any, _
    types As Any, _
    ByVal Count As Long, _
    brushmode As FillMode, _
    Path As Long) As GpStatus
Public Declare Function GdipCreatePath2I_ _
    Lib "GDIPlus" _
    Alias "GdipCreatePath2I" _
    (Points As Any, _
    types As Any, _
    ByVal Count As Long, _
    brushmode As FillMode, _
    Path As Long) As GpStatus
Public Declare Function GdipClonePath _
    Lib "GDIPlus" (ByVal Path As Long, _
    clonePath As Long) As GpStatus
Public Declare Function GdipDeletePath Lib "GDIPlus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipResetPath Lib "GDIPlus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipGetPointCount _
    Lib "GDIPlus" (ByVal Path As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipGetPathTypes _
    Lib "GDIPlus" (ByVal Path As Long, _
    types As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPoints _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPointsI _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPoints_ _
    Lib "GDIPlus" _
    Alias "GdipGetPathPoints" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPointsI_ _
    Lib "GDIPlus" _
    Alias "GdipGetPathPointsI" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathFillMode _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal brushmode As FillMode) As GpStatus
Public Declare Function GdipSetPathFillMode _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal brushmode As FillMode) As GpStatus
Public Declare Function GdipGetPathData _
    Lib "GDIPlus" (ByVal Path As Long, _
    pData As PathData) As GpStatus
Public Declare Function GdipStartPathFigure Lib "GDIPlus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipClosePathFigure Lib "GDIPlus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipClosePathFigures _
    Lib "GDIPlus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipSetPathMarker Lib "GDIPlus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipClearPathMarkers _
    Lib "GDIPlus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipReversePath Lib "GDIPlus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipGetPathLastPoint _
    Lib "GDIPlus" (ByVal Path As Long, _
    lastPoint As POINTF) As GpStatus
Public Declare Function GdipAddPathLine _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X1 As Single, _
    ByVal Y1 As Single, _
    ByVal X2 As Single, _
    ByVal Y2 As Single) As GpStatus
Public Declare Function GdipAddPathLine2 _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathLine2_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathLine2" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathArc _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    ByVal startAngle As Single, _
    ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathBezier _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X1 As Single, _
    ByVal Y1 As Single, _
    ByVal X2 As Single, _
    ByVal Y2 As Single, _
    ByVal x3 As Single, _
    ByVal y3 As Single, _
    ByVal x4 As Single, _
    ByVal y4 As Single) As GpStatus
Public Declare Function GdipAddPathBeziers _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2 _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTF, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathCurve3 _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTF, _
    ByVal Count As Long, _
    ByVal Offset As Long, _
    ByVal numberOfSegments As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurve _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2 _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTF, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathBeziers_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathBeziers" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathCurve" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathCurve2" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathCurve3_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathCurve3" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal Offset As Long, _
    ByVal numberOfSegments As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurve_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathClosedCurve" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathClosedCurve2" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathRectangle _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single) As GpStatus
Public Declare Function GdipAddPathRectangles _
    Lib "GDIPlus" (ByVal Path As Long, _
    RECT As RECTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathEllipse _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single) As GpStatus
Public Declare Function GdipAddPathPie _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    ByVal startAngle As Single, _
    ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathPolygon _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathPolygon_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathPolygon" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathPath _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal addingPath As Long, _
    ByVal bConnect As Long) As GpStatus
Public Declare Function GdipAddPathString _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal Str As Long, _
    ByVal Length As Long, _
    ByVal family As Long, _
    ByVal Style As FontStyle, _
    ByVal emSize As Single, _
    layoutRect As RECTF, _
    ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipAddPathStringI _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal Str As Long, _
    ByVal Length As Long, _
    ByVal family As Long, _
    ByVal Style As FontStyle, _
    ByVal emSize As Single, _
    layoutRect As RECTL, _
    ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipAddPathLineI _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long) As GpStatus
Public Declare Function GdipAddPathLine2I _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathLine2I_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathLine2I" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathArcI _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal startAngle As Single, _
    ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathBezierI _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X1 As Long, _
    ByVal Y1 As Long, _
    ByVal X2 As Long, _
    ByVal Y2 As Long, _
    ByVal x3 As Long, _
    ByVal y3 As Long, _
    ByVal x4 As Long, _
    ByVal y4 As Long) As GpStatus
Public Declare Function GdipAddPathBeziersI _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurveI _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2I _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTL, _
    ByVal Count As Long, _
    ByVal tension As Long) As GpStatus
Public Declare Function GdipAddPathCurve3I _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTL, _
    ByVal Count As Long, _
    ByVal Offset As Long, _
    ByVal numberOfSegments As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurveI _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2I _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTL, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathBeziersI_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathBeziersI" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurveI_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathCurveI" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2I_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathCurve2I" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal tension As Long) As GpStatus
Public Declare Function GdipAddPathCurve3I_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathCurve3I" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal Offset As Long, _
    ByVal numberOfSegments As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurveI_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathClosedCurveI" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2I_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathClosedCurve2I" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathRectangleI _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathRectanglesI _
    Lib "GDIPlus" (ByVal Path As Long, _
    rects As RECTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathEllipseI _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathPieI _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal startAngle As Single, _
    ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathPolygonI _
    Lib "GDIPlus" (ByVal Path As Long, _
    Points As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathPolygonI_ _
    Lib "GDIPlus" _
    Alias "GdipAddPathPolygonI" _
    (ByVal Path As Long, _
    Points As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipFlattenPath _
    Lib "GDIPlus" (ByVal Path As Long, _
    Optional ByVal matrix As Long = 0, _
    Optional ByVal flatness As Single = 0.25) As GpStatus
Public Declare Function GdipWindingModeOutline _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal matrix As Long, _
    ByVal flatness As Single) As GpStatus
Public Declare Function GdipWidenPath _
    Lib "GDIPlus" (ByVal NativePath As Long, _
    ByVal pen As Long, _
    ByVal matrix As Long, _
    ByVal flatness As Single) As GpStatus
Public Declare Function GdipWarpPath _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal matrix As Long, _
    Points As POINTF, _
    ByVal Count As Long, _
    ByVal srcx As Single, _
    ByVal srcy As Single, _
    ByVal srcwidth As Single, _
    ByVal srcheight As Single, _
    ByVal WarpMd As WarpMode, _
    ByVal flatness As Single) As GpStatus
Public Declare Function GdipWarpPath_ _
    Lib "GDIPlus" _
    Alias "GdipWarpPath" _
    (ByVal Path As Long, _
    ByVal matrix As Long, _
    Points As Any, _
    ByVal Count As Long, _
    ByVal srcx As Single, _
    ByVal srcy As Single, _
    ByVal srcwidth As Single, _
    ByVal srcheight As Single, _
    ByVal WarpMd As WarpMode, _
    ByVal flatness As Single) As GpStatus
Public Declare Function GdipTransformPath _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetPathWorldBounds _
    Lib "GDIPlus" (ByVal Path As Long, _
    bounds As RECTF, _
    ByVal matrix As Long, _
    ByVal pen As Long) As GpStatus
Public Declare Function GdipGetPathWorldBoundsI _
    Lib "GDIPlus" (ByVal Path As Long, _
    bounds As RECTL, _
    ByVal matrix As Long, _
    ByVal pen As Long) As GpStatus
Public Declare Function GdipIsVisiblePathPoint _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsVisiblePathPointI _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsOutlineVisiblePathPoint _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal pen As Long, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsOutlineVisiblePathPointI _
    Lib "GDIPlus" (ByVal Path As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal pen As Long, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipCreatePathIter _
    Lib "GDIPlus" (iterator As Long, _
    ByVal Path As Long) As GpStatus
Public Declare Function GdipDeletePathIter _
    Lib "GDIPlus" (ByVal iterator As Long) As GpStatus
Public Declare Function GdipPathIterNextSubpath _
    Lib "GDIPlus" (ByVal iterator As Long, _
    resultCount As Long, _
    startIndex As Long, _
    endIndex As Long, _
    isClosed As Long) As GpStatus
Public Declare Function GdipPathIterNextSubpathPath _
    Lib "GDIPlus" (ByVal iterator As Long, _
    resultCount As Long, _
    ByVal Path As Long, _
    isClosed As Long) As GpStatus
Public Declare Function GdipPathIterNextPathType _
    Lib "GDIPlus" (ByVal iterator As Long, _
    resultCount As Long, _
    pathType As Any, _
    startIndex As Long, _
    endIndex As Long) As GpStatus
Public Declare Function GdipPathIterNextMarker _
    Lib "GDIPlus" (ByVal iterator As Long, _
    resultCount As Long, _
    startIndex As Long, _
    endIndex As Long) As GpStatus
Public Declare Function GdipPathIterNextMarkerPath _
    Lib "GDIPlus" (ByVal iterator As Long, _
    resultCount As Long, _
    ByVal Path As Long) As GpStatus
Public Declare Function GdipPathIterGetCount _
    Lib "GDIPlus" (ByVal iterator As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipPathIterGetSubpathCount _
    Lib "GDIPlus" (ByVal iterator As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipPathIterIsValid _
    Lib "GDIPlus" (ByVal iterator As Long, _
    valid As Long) As GpStatus
Public Declare Function GdipPathIterHasCurve _
    Lib "GDIPlus" (ByVal iterator As Long, _
    hasCurve As Long) As GpStatus
Public Declare Function GdipPathIterRewind _
    Lib "GDIPlus" (ByVal iterator As Long) As GpStatus
Public Declare Function GdipPathIterEnumerate _
    Lib "GDIPlus" (ByVal iterator As Long, _
    resultCount As Long, _
    Points As POINTF, _
    types As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipPathIterCopyData _
    Lib "GDIPlus" (ByVal iterator As Long, _
    resultCount As Long, _
    Points As POINTF, _
    types As Any, _
    ByVal startIndex As Long, _
    ByVal endIndex As Long) As GpStatus
Public Declare Function GdipPathIterEnumerate_ _
    Lib "GDIPlus" _
    Alias "GdipPathIterEnumerate" _
    (ByVal iterator As Long, _
    resultCount As Long, _
    Points As Any, _
    types As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipPathIterCopyData_ _
    Lib "GDIPlus" _
    Alias "GdipPathIterCopyData" _
    (ByVal iterator As Long, _
    resultCount As Long, _
    Points As Any, _
    types As Any, _
    ByVal startIndex As Long, _
    ByVal endIndex As Long) As GpStatus
Public Declare Function GdipCreateMatrix Lib "GDIPlus" (matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix2 _
    Lib "GDIPlus" (ByVal m11 As Single, _
    ByVal m12 As Single, _
    ByVal m21 As Single, _
    ByVal m22 As Single, _
    ByVal dx As Single, _
    ByVal dy As Single, _
    matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix3 _
    Lib "GDIPlus" (RECT As RECTF, _
    dstplg As POINTF, _
    matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix3I _
    Lib "GDIPlus" (RECT As RECTL, _
    dstplg As POINTL, _
    matrix As Long) As GpStatus
Public Declare Function GdipCloneMatrix _
    Lib "GDIPlus" (ByVal matrix As Long, _
    cloneMatrix As Long) As GpStatus
Public Declare Function GdipDeleteMatrix Lib "GDIPlus" (ByVal matrix As Long) As GpStatus
Public Declare Function GdipSetMatrixElements _
    Lib "GDIPlus" (ByVal matrix As Long, _
    ByVal m11 As Single, _
    ByVal m12 As Single, _
    ByVal m21 As Single, _
    ByVal m22 As Single, _
    ByVal dx As Single, _
    ByVal dy As Single) As GpStatus
Public Declare Function GdipMultiplyMatrix _
    Lib "GDIPlus" (ByVal matrix As Long, _
    ByVal matrix2 As Long, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateMatrix _
    Lib "GDIPlus" (ByVal matrix As Long, _
    ByVal offsetX As Single, _
    ByVal offsetY As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleMatrix _
    Lib "GDIPlus" (ByVal matrix As Long, _
    ByVal ScaleX As Single, _
    ByVal ScaleY As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateMatrix _
    Lib "GDIPlus" (ByVal matrix As Long, _
    ByVal angle As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipShearMatrix _
    Lib "GDIPlus" (ByVal matrix As Long, _
    ByVal shearX As Single, _
    ByVal shearY As Single, _
    ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipInvertMatrix Lib "GDIPlus" (ByVal matrix As Long) As GpStatus
Public Declare Function GdipTransformMatrixPoints _
    Lib "GDIPlus" (ByVal matrix As Long, _
    pts As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformMatrixPointsI _
    Lib "GDIPlus" (ByVal matrix As Long, _
    pts As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPoints _
    Lib "GDIPlus" (ByVal matrix As Long, _
    pts As POINTF, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPointsI _
    Lib "GDIPlus" (ByVal matrix As Long, _
    pts As POINTL, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformMatrixPoints_ _
    Lib "GDIPlus" _
    Alias "GdipTransformMatrixPoints" _
    (ByVal matrix As Long, _
    pts As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformMatrixPointsI_ _
    Lib "GDIPlus" _
    Alias "GdipTransformMatrixPointsI" _
    (ByVal matrix As Long, _
    pts As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPoints_ _
    Lib "GDIPlus" _
    Alias "GdipVectorTransformMatrixPoints" _
    (ByVal matrix As Long, _
    pts As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPointsI_ _
    Lib "GDIPlus" _
    Alias "GdipVectorTransformMatrixPointsI" _
    (ByVal matrix As Long, _
    pts As Any, _
    ByVal Count As Long) As GpStatus
Public Declare Function GdipGetMatrixElements _
    Lib "GDIPlus" (ByVal matrix As Long, _
    matrixOut As Single) As GpStatus
Public Declare Function GdipIsMatrixInvertible _
    Lib "GDIPlus" (ByVal matrix As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsMatrixIdentity _
    Lib "GDIPlus" (ByVal matrix As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsMatrixEqual _
    Lib "GDIPlus" (ByVal matrix As Long, _
    ByVal matrix2 As Long, _
    result As Long) As GpStatus
Public Declare Function GdipCreateRegion Lib "GDIPlus" (region As Long) As GpStatus
Public Declare Function GdipCreateRegionRect _
    Lib "GDIPlus" (RECT As RECTF, _
    region As Long) As GpStatus
Public Declare Function GdipCreateRegionRectI _
    Lib "GDIPlus" (RECT As RECTL, _
    region As Long) As GpStatus
Public Declare Function GdipCreateRegionPath _
    Lib "GDIPlus" (ByVal Path As Long, _
    region As Long) As GpStatus
Public Declare Function GdipCreateRegionRgnData _
    Lib "GDIPlus" (regionData As Any, _
    ByVal size As Long, _
    region As Long) As GpStatus
Public Declare Function GdipCreateRegionHrgn _
    Lib "GDIPlus" (ByVal hRgn As Long, _
    region As Long) As GpStatus
Public Declare Function GdipCloneRegion _
    Lib "GDIPlus" (ByVal region As Long, _
    cloneRegion As Long) As GpStatus
Public Declare Function GdipDeleteRegion Lib "GDIPlus" (ByVal region As Long) As GpStatus
Public Declare Function GdipSetInfinite Lib "GDIPlus" (ByVal region As Long) As GpStatus
Public Declare Function GdipSetEmpty Lib "GDIPlus" (ByVal region As Long) As GpStatus
Public Declare Function GdipCombineRegionRect _
    Lib "GDIPlus" (ByVal region As Long, _
    RECT As RECTF, _
    ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionRectI _
    Lib "GDIPlus" (ByVal region As Long, _
    RECT As RECTL, _
    ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionPath _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal Path As Long, _
    ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionRegion _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal region2 As Long, _
    ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipTranslateRegion _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal dx As Single, _
    ByVal dy As Single) As GpStatus
Public Declare Function GdipTranslateRegionI _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal dx As Long, _
    ByVal dy As Long) As GpStatus
Public Declare Function GdipTransformRegion _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetRegionBounds _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal graphics As Long, _
    RECT As RECTF) As GpStatus
Public Declare Function GdipGetRegionBoundsI _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal graphics As Long, _
    RECT As RECTL) As GpStatus
Public Declare Function GdipGetRegionHRgn _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal graphics As Long, _
    hRgn As Long) As GpStatus
Public Declare Function GdipIsEmptyRegion _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsInfiniteRegion _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsEqualRegion _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal region2 As Long, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipGetRegionDataSize _
    Lib "GDIPlus" (ByVal region As Long, _
    bufferSize As Long) As GpStatus
Public Declare Function GdipGetRegionData _
    Lib "GDIPlus" (ByVal region As Long, _
    buffer As Any, _
    ByVal bufferSize As Long, _
    sizeFilled As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionPoint _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionPointI _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionRect _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal X As Single, _
    ByVal Y As Single, _
    ByVal Width As Single, _
    ByVal Height As Single, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionRectI _
    Lib "GDIPlus" (ByVal region As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    ByVal graphics As Long, _
    result As Long) As GpStatus
Public Declare Function GdipGetRegionScansCount _
    Lib "GDIPlus" (ByVal region As Long, _
    Ucount As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetRegionScans _
    Lib "GDIPlus" (ByVal region As Long, _
    rects As RECTF, _
    Count As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetRegionScansI _
    Lib "GDIPlus" (ByVal region As Long, _
    rects As RECTL, _
    Count As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipCreateImageAttributes _
    Lib "GDIPlus" (imageattr As Long) As GpStatus
Public Declare Function GdipCloneImageAttributes _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    cloneImageattr As Long) As GpStatus
Public Declare Function GdipDisposeImageAttributes _
    Lib "GDIPlus" (ByVal imageattr As Long) As GpStatus
Public Declare Function GdipSetImageAttributesToIdentity _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipResetImageAttributes _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipSetImageAttributesColorMatrix _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal ClrAdjType As ColorAdjustType, _
    ByVal enableFlag As Long, _
    colourMatrix As Any, _
    grayMatrix As Any, _
    ByVal flags As ColorMatrixFlags) As GpStatus
Public Declare Function GdipSetImageAttributesThreshold _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal ClrAdjType As ColorAdjustType, _
    ByVal enableFlag As Long, _
    ByVal threshold As Single) As GpStatus
Public Declare Function GdipSetImageAttributesGamma _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal ClrAdjType As ColorAdjustType, _
    ByVal enableFlag As Long, _
    ByVal gamma As Single) As GpStatus
Public Declare Function GdipSetImageAttributesNoOp _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal ClrAdjType As ColorAdjustType, _
    ByVal enableFlag As Long) As GpStatus
Public Declare Function GdipSetImageAttributesColorKeys _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal ClrAdjType As ColorAdjustType, _
    ByVal enableFlag As Long, _
    ByVal colorLow As Long, _
    ByVal colorHigh As Long) As GpStatus
Public Declare Function GdipSetImageAttributesOutputChannel _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal ClrAdjstType As ColorAdjustType, _
    ByVal enableFlag As Long, _
    ByVal channelFlags As ColorChannelFlags) As GpStatus
Public Declare Function GdipSetImageAttributesOutputChannelColorProfile _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal ClrAdjType As ColorAdjustType, _
    ByVal enableFlag As Long, _
    ByVal colorProfileFilename As Long) As GpStatus
Public Declare Function GdipSetImageAttributesRemapTable _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal ClrAdjType As ColorAdjustType, _
    ByVal enableFlag As Long, _
    ByVal mapSize As Long, _
    map As Any) As GpStatus
Public Declare Function GdipSetImageAttributesWrapMode _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal wrap As WrapMode, _
    ByVal argb As Long, _
    ByVal bClamp As Long) As GpStatus
Public Declare Function GdipSetImageAttributesICMMode _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal bOn As Long) As GpStatus
Public Declare Function GdipGetImageAttributesAdjustedPalette _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    colorPal As ColorPalette, _
    ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipCreateFontFamilyFromName _
    Lib "GDIPlus" (ByVal Name As Long, _
    ByVal fontCollection As Long, _
    fontFamily As Long) As GpStatus
Public Declare Function GdipDeleteFontFamily _
    Lib "GDIPlus" (ByVal fontFamily As Long) As GpStatus
Public Declare Function GdipCloneFontFamily _
    Lib "GDIPlus" (ByVal fontFamily As Long, _
    clonedFontFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySansSerif _
    Lib "GDIPlus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySerif _
    Lib "GDIPlus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilyMonospace _
    Lib "GDIPlus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetFamilyName _
    Lib "GDIPlus" (ByVal family As Long, _
    ByVal Name As Long, _
    ByVal language As Integer) As GpStatus
Public Declare Function GdipIsStyleAvailable _
    Lib "GDIPlus" (ByVal family As Long, _
    ByVal Style As Long, _
    IsStyleAvailable As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerable _
    Lib "GDIPlus" (ByVal fontCollection As Long, _
    ByVal graphics As Long, _
    numFound As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerate _
    Lib "GDIPlus" (ByVal fontCollection As Long, _
    ByVal numSought As Long, _
    gpFamilies As Long, _
    ByVal numFound As Long, _
    ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetEmHeight _
    Lib "GDIPlus" (ByVal family As Long, _
    ByVal Style As FontStyle, _
    EmHeight As Integer) As GpStatus
Public Declare Function GdipGetCellAscent _
    Lib "GDIPlus" (ByVal family As Long, _
    ByVal Style As FontStyle, _
    CellAscent As Integer) As GpStatus
Public Declare Function GdipGetCellDescent _
    Lib "GDIPlus" (ByVal family As Long, _
    ByVal Style As FontStyle, _
    CellDescent As Integer) As GpStatus
Public Declare Function GdipGetLineSpacing _
    Lib "GDIPlus" (ByVal family As Long, _
    ByVal Style As FontStyle, _
    LineSpacing As Integer) As GpStatus
Public Declare Function GdipCreateFontFromDC _
    Lib "GDIPlus" (ByVal hDC As Long, _
    createdfont As Long) As GpStatus
Public Declare Function GdipCreateFontFromLogfontA _
    Lib "GDIPlus" (ByVal hDC As Long, _
    logfont As LOGFONTA, _
    createdfont As Long) As GpStatus
Public Declare Function GdipCreateFontFromLogfontW _
    Lib "GDIPlus" (ByVal hDC As Long, _
    logfont As LOGFONTW, _
    createdfont As Long) As GpStatus
Public Declare Function GdipCreateFont _
    Lib "GDIPlus" (ByVal fontFamily As Long, _
    ByVal emSize As Single, _
    ByVal Style As FontStyle, _
    ByVal unit As GpUnit, _
    createdfont As Long) As GpStatus
Public Declare Function GdipCloneFont _
    Lib "GDIPlus" (ByVal curFont As Long, _
    cloneFont As Long) As GpStatus
Public Declare Function GdipDeleteFont Lib "GDIPlus" (ByVal curFont As Long) As GpStatus
Public Declare Function GdipGetFamily _
    Lib "GDIPlus" (ByVal curFont As Long, _
    family As Long) As GpStatus
Public Declare Function GdipGetFontStyle _
    Lib "GDIPlus" (ByVal curFont As Long, _
    Style As FontStyle) As GpStatus
Public Declare Function GdipGetFontSize _
    Lib "GDIPlus" (ByVal curFont As Long, _
    size As Single) As GpStatus
Public Declare Function GdipGetFontUnit _
    Lib "GDIPlus" (ByVal curFont As Long, _
    unit As GpUnit) As GpStatus
Public Declare Function GdipGetFontHeight _
    Lib "GDIPlus" (ByVal curFont As Long, _
    ByVal graphics As Long, _
    Height As Single) As GpStatus
Public Declare Function GdipGetFontHeightGivenDPI _
    Lib "GDIPlus" (ByVal curFont As Long, _
    ByVal dpi As Single, _
    Height As Single) As GpStatus
Public Declare Function GdipGetLogFontA _
    Lib "GDIPlus" (ByVal curFont As Long, _
    ByVal graphics As Long, _
    logfont As LOGFONTA) As GpStatus
Public Declare Function GdipGetLogFontW _
    Lib "GDIPlus" (ByVal curFont As Long, _
    ByVal graphics As Long, _
    logfont As LOGFONTW) As GpStatus
Public Declare Function GdipNewInstalledFontCollection _
    Lib "GDIPlus" (fontCollection As Long) As GpStatus
Public Declare Function GdipNewPrivateFontCollection _
    Lib "GDIPlus" (fontCollection As Long) As GpStatus
Public Declare Function GdipDeletePrivateFontCollection _
    Lib "GDIPlus" (fontCollection As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyCount _
    Lib "GDIPlus" (ByVal fontCollection As Long, _
    numFound As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyList _
    Lib "GDIPlus" (ByVal fontCollection As Long, _
    ByVal numSought As Long, _
    gpFamilies As Long, _
    numFound As Long) As GpStatus
Public Declare Function GdipPrivateAddFontFile _
    Lib "GDIPlus" (ByVal fontCollection As Long, _
    ByVal filename As Long) As GpStatus
Public Declare Function GdipPrivateAddMemoryFont _
    Lib "GDIPlus" (ByVal fontCollection As Long, _
    ByVal memory As Long, _
    ByVal Length As Long) As GpStatus
Public Declare Function GdipDrawString _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Str As Long, _
    ByVal Length As Long, _
    ByVal thefont As Long, _
    layoutRect As RECTF, _
    ByVal StringFormat As Long, _
    ByVal brush As Long) As GpStatus
Public Declare Function GdipMeasureString _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Str As Long, _
    ByVal Length As Long, _
    ByVal thefont As Long, _
    layoutRect As RECTF, _
    ByVal StringFormat As Long, _
    boundingBox As RECTF, _
    codepointsFitted As Long, _
    linesFilled As Long) As GpStatus
Public Declare Function GdipMeasureCharacterRanges _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Str As Long, _
    ByVal Length As Long, _
    ByVal thefont As Long, _
    layoutRect As RECTF, _
    ByVal StringFormat As Long, _
    ByVal regionCount As Long, _
    regions As Long) As GpStatus
Public Declare Function GdipDrawDriverString _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Str As Long, _
    ByVal Length As Long, _
    ByVal thefont As Long, _
    ByVal brush As Long, _
    positions As POINTF, _
    ByVal flags As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipMeasureDriverString _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Str As Long, _
    ByVal Length As Long, _
    ByVal thefont As Long, _
    positions As POINTF, _
    ByVal flags As Long, _
    ByVal matrix As Long, _
    boundingBox As RECTF) As GpStatus
Public Declare Function GdipDrawDriverString_ _
    Lib "GDIPlus" _
    Alias "GdipDrawDriverString" _
    (ByVal graphics As Long, _
    ByVal Str As Long, _
    ByVal Length As Long, _
    ByVal thefont As Long, _
    ByVal brush As Long, _
    positions As Any, _
    ByVal flags As Long, _
    ByVal matrix As Long) As GpStatus
Public Declare Function GdipMeasureDriverString_ _
    Lib "GDIPlus" _
    Alias "GdipMeasureDriverString" _
    (ByVal graphics As Long, _
    ByVal Str As Long, _
    ByVal Length As Long, _
    ByVal thefont As Long, _
    positions As Any, _
    ByVal flags As Long, _
    ByVal matrix As Long, _
    boundingBox As RECTF) As GpStatus
Public Declare Function GdipCreateStringFormat _
    Lib "GDIPlus" (ByVal formatAttributes As Long, _
    ByVal language As Integer, _
    StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericDefault _
    Lib "GDIPlus" (StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericTypographic _
    Lib "GDIPlus" (StringFormat As Long) As GpStatus
Public Declare Function GdipDeleteStringFormat _
    Lib "GDIPlus" (ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipCloneStringFormat _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    newFormat As Long) As GpStatus
Public Declare Function GdipSetStringFormatFlags _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    ByVal flags As Long) As GpStatus
Public Declare Function GdipGetStringFormatFlags _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    flags As Long) As GpStatus
Public Declare Function GdipSetStringFormatAlign _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatAlign _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    Align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatLineAlign _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatLineAlign _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    Align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatTrimming _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    ByVal trimming As StringTrimming) As GpStatus
Public Declare Function GdipGetStringFormatTrimming _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    trimming As Long) As GpStatus
Public Declare Function GdipSetStringFormatHotkeyPrefix _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    ByVal hkPrefix As HotkeyPrefix) As GpStatus
Public Declare Function GdipGetStringFormatHotkeyPrefix _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    hkPrefix As HotkeyPrefix) As GpStatus
Public Declare Function GdipSetStringFormatTabStops _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    ByVal firstTabOffset As Single, _
    ByVal Count As Long, _
    tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStops _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    ByVal Count As Long, _
    firstTabOffset As Single, _
    tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStopCount _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipSetStringFormatDigitSubstitution _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    ByVal language As Integer, _
    ByVal substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatDigitSubstitution _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    language As Integer, _
    substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatMeasurableCharacterRangeCount _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    Count As Long) As GpStatus
Public Declare Function GdipSetStringFormatMeasurableCharacterRanges _
    Lib "GDIPlus" (ByVal StringFormat As Long, _
    ByVal rangeCount As Long, _
    ranges As CharacterRange) As GpStatus

    '===================================================================================
    '  GdiPlus 1.1 新内容
    '===================================================================================
#If GdipVersion >= 1.1 Then

Public Const BlurEffectGuid                   As String = "{633C80A4-1843-482B-9EF2-BE2834C5FDD4}"
Public Const BrightnessContrastEffectGuid     As String = "{D3A1DBE1-8EC4-4C17-9F4C-EA97AD1C343D}"
Public Const ColorBalanceEffectGuid           As String = "{537E597D-251E-48DA-9664-29CA496B70F8}"
Public Const ColorCurveEffectGuid             As String = "{DD6A0022-58E4-4A67-9D9B-D48EB881A53D}"
Public Const ColorLookupTableEffectGuid       As String = "{A7CE72A9-0F7F-40D7-B3CC-D0C02D5C3212}"
Public Const ColorMatrixEffectGuid            As String = "{718F2615-7933-40E3-A511-5F68FE14DD74}"
Public Const HueSaturationLightnessEffectGuid As String = "{8B2DD6C3-EB07-4D87-A5F0-7108E26A9C5F}"
Public Const LevelsEffectGuid                 As String = "{99C354EC-2A31-4F3A-8C34-17A803B33A25}"
Public Const RedEyeCorrectionEffectGuid       As String = "{74D29D05-69A4-4266-9549-3CC52836B632}"
Public Const SharpenEffectGuid                As String = "{63CBF3EE-C526-402C-8F71-62C540BF5142}"
Public Const TintEffectGuid                   As String = "{1077AF00-2848-4441-9489-44AD4C2D7A2C}"
    
Public Enum GdipEffectType
    Blur
    BrightnessContrast
    ColorBalance
    ColorCurve
    ColorLookupTable
    ColorMatrix
    HueSaturationLightness
    Levels
    RedEyeCorrection
    Sharpen
    Tint
End Enum
    
Public Enum HistogramFormat
    HistogramFormatARGB
    HistogramFormatPARGB
    HistogramFormatRGB
    HistogramFormatGray
    HistogramFormatB
    HistogramFormatG
    HistogramFormatR
    HistogramFormatA
End Enum

Public Enum CurveAdjustments
    AdjustExposure
    AdjustDensity
    AdjustContrast
    AdjustHighlight
    AdjustShadow
    AdjustMidtone
    AdjustWhiteSaturation
    AdjustBlackSaturation
End Enum

Public Enum CurveChannel
    CurveChannelAll
    CurveChannelRed
    CurveChannelGreen
    CurveChannelBlue
End Enum

Public Enum PaletteType
    PaletteTypeCustom = 0
    ' Optimal palette generated using a median-cut algorithm.
    PaletteTypeOptimal = 1
    ' Black and white palette.
    PaletteTypeFixedBW = 2
    ' Symmetric halftone palettes.
    ' Each of these halftone palettes will be a superset of the system palette.
    ' E.g. Halftone8 will have it's 8-color on-off primaries and the 16 system
    ' colors added. With duplicates removed, that leaves 16 colors.
    PaletteTypeFixedHalftone8 = 3                                               ' 8-color, on-off primaries
    PaletteTypeFixedHalftone27 = 4                                              ' 3 intensity levels of each color
    PaletteTypeFixedHalftone64 = 5                                              ' 4 intensity levels of each color
    PaletteTypeFixedHalftone125 = 6                                             ' 5 intensity levels of each color
    PaletteTypeFixedHalftone216 = 7                                             ' 6 intensity levels of each color
    ' Assymetric halftone palettes.
    ' These are somewhat less useful than the symmetric ones, but are
    ' included for completeness. These do not include all of the system
    ' colors.
    PaletteTypeFixedHalftone252 = 8                                             ' 6-red, 7-green, 6-blue intensities
    PaletteTypeFixedHalftone256 = 9                                             ' 8-red, 8-green, 4-blue intensities
End Enum

Public Enum DitherType
    DitherTypeNone = 0
    ' Solid color - picks the nearest matching color with no attempt to
    ' halftone or dither. May be used on an arbitrary palette.
    DitherTypeSolid = 1
    ' Ordered dithers and spiral dithers must be used with a fixed palette.
    ' NOTE: DitherOrdered4x4 is unique in that it may apply to 16bpp
    ' conversions also.
    DitherTypeOrdered4x4 = 2
    DitherTypeOrdered8x8 = 3
    DitherTypeOrdered16x16 = 4
    DitherTypeSpiral4x4 = 5
    DitherTypeSpiral8x8 = 6
    DitherTypeDualSpiral4x4 = 7
    DitherTypeDualSpiral8x8 = 8
    ' Error diffusion. May be used with any palette.
    DitherTypeErrorDiffusion = 9
    DitherTypeMax = 10
End Enum

Public Enum ItemDataPosition
    ItemDataPositionAfterHeader = 0
    ItemDataPositionAfterPalette = 1
    ItemDataPositionAfterBits = 2
End Enum

'struct __declspec(novtable) GdiplusAbort
'{
'    virtual HRESULT __stdcall Abort(void) = 0;
'};
Public Type GdiplusAbort
    AbortCallback As Long
End Type

Public Type ImageItemData
    size     As Long
    Position As Long
    pDesc    As Long
    DescSize As Long
    pData    As Long
    dataSize As Long
    Cookie   As Long
End Type

Public Type SharpenParams
    radius As Single
    amount As Single
End Type

Public Type BlurParams
    radius     As Single
    expandEdge As Long
End Type

Public Type BrightnessContrastParams
    brightnessLevel As Long
    contrastLevel   As Long
End Type

Public Type RedEyeCorrectionParams
    numberOfAreas As Long
    areas         As RECTL
End Type

Public Type HueSaturationLightnessParams
    hueLevel        As Long
    saturationLevel As Long
    lightnessLevel  As Long
End Type

Public Type TintParams
    hue    As Long
    amount As Long
End Type

Public Type LevelsParams
    highlight As Long
    midtone   As Long
    shadow    As Long
End Type

Public Type ColorBalanceParams
    cyanRed      As Long
    magentaGreen As Long
    yellowBlue   As Long
End Type

Public Type ColorLUTParams
    lutB(0 To 255) As Byte
    lutG(0 To 255) As Byte
    lutR(0 To 255) As Byte
    lutA(0 To 255) As Byte
End Type

Public Type ColorCurveParams
    adjustment  As CurveAdjustments
    channel     As CurveChannel
    adjustValue As Long
End Type

Public Declare Function GdipCreateEffect _
    Lib "GDIPlus" (ByVal Guid41 As Long, _
    ByVal Guid42 As Long, _
    ByVal Guid43 As Long, _
    ByVal Guid44 As Long, _
    Effect As Long) As GpStatus
                                  
Public Declare Function GdipDeleteEffect _
    Lib "GDIPlus" (ByVal Effect As Long) As GpStatus
Public Declare Function GdipGetEffectParameterSize _
    Lib "GDIPlus" (ByVal Effect As Long, _
    size As Long) As GpStatus
Public Declare Function GdipSetEffectParameters _
    Lib "GDIPlus" (ByVal Effect As Long, _
    Params As Any, _
    ByVal size As Long) As GpStatus
Public Declare Function GdipGetEffectParameters _
    Lib "GDIPlus" (ByVal Effect As Long, _
    size As Long, _
    Params As Any) As GpStatus

Public Declare Function GdipImageSetAbort _
    Lib "GDIPlus" (ByVal Image As Long, _
    IAbort As GdiplusAbort) As GpStatus
Public Declare Function GdipGraphicsSetAbort _
    Lib "GDIPlus" (ByVal graphics As Long, _
    IAbort As GdiplusAbort) As GpStatus
Public Declare Function GdipBitmapConvertFormat _
    Lib "GDIPlus" (ByVal InputBitmap As Long, _
    ByVal format As GpPixelFormat, _
    ByVal DitherType As DitherType, _
    ByVal PaletteType As PaletteType, _
    palette As ColorPalette, _
    ByVal alphaThresholdPercent As Single) As GpStatus
Public Declare Function GdipInitializePalette _
    Lib "GDIPlus" (palette As ColorPalette, _
    ByVal PaletteType As PaletteType, _
    ByVal optimalColors As Long, _
    ByVal useTransparentColor As Long, _
    Optional ByVal Bitmap As Long) As GpStatus
Public Declare Function GdipBitmapApplyEffect _
    Lib "GDIPlus" (ByVal Bitmap As Long, _
    ByVal Effect As Long, _
    roi As RECTL, _
    ByVal useAuxData As Long, _
    auxData As Any, _
    auxDataSize As Long) As GpStatus
Public Declare Function GdipBitmapCreateApplyEffect _
    Lib "GDIPlus" (inputBitmaps As Any, _
    ByVal numInputs As Long, _
    ByVal Effect As Long, _
    roi As RECTL, _
    outputRect As RECTL, _
    outputBitmap As Long, _
    ByVal useAuxData As Long, _
    auxData As Any, _
    auxDataSize As Long) As GpStatus
Public Declare Function GdipBitmapGetHistogram _
    Lib "GDIPlus" (ByVal Bitmap As Long, _
    ByVal format As HistogramFormat, _
    ByVal NumberOfEntries As Long, _
    channel0 As Any, _
    channel1 As Any, _
    channel2 As Any, _
    channel3 As Any) As GpStatus
Public Declare Function GdipBitmapGetHistogramSize _
    Lib "GDIPlus" (ByVal format As HistogramFormat, _
    NumberOfEntries As Long) As GpStatus

Public Declare Function GdipFindFirstImageItem _
    Lib "GDIPlus" (ByVal Image As Long, _
    item As ImageItemData) As GpStatus
Public Declare Function GdipFindNextImageItem _
    Lib "GDIPlus" (ByVal Image As Long, _
    item As ImageItemData) As GpStatus
Public Declare Function GdipGetImageItemData _
    Lib "GDIPlus" (ByVal Image As Long, _
    item As ImageItemData) As GpStatus

Public Declare Function GdipDrawImageFX _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal Image As Long, _
    Source As RECTF, _
    ByVal xForm As Long, _
    ByVal Effect As Long, _
    ByVal imageAttributes As Long, _
    ByVal srcUnit As GpUnit) As GpStatus

#End If

'===================================================================================
'  不怎么常用的东西
'===================================================================================

'=================================
'== Structures                  ==
'=================================

'=================================
'Log Font Structure
Public Type LOGFONTA
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32)   As Byte
End Type

Public Type LOGFONTW
    lfHeight         As Long
    lfWidth          As Long
    lfEscapement     As Long
    lfOrientation    As Long
    lfWeight         As Long
    lfItalic         As Byte
    lfUnderline      As Byte
    lfStrikeOut      As Byte
    lfCharSet        As Byte
    lfOutPrecision   As Byte
    lfClipPrecision  As Byte
    lfQuality        As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32)   As Byte
End Type

'=================================
'Image
Public Type ImageCodecInfo
    ClassID           As Clsid
    FormatID          As Clsid
    CodecName         As Long
    DllName           As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType          As Long
    flags             As ImageCodecFlags
    Version           As Long
    SigCount          As Long
    SigSize           As Long
    SigPattern        As Long
    SigMask           As Long
End Type

'=================================
'Colors
Public Type ColorPalette
    flags             As PaletteFlags
    Count             As Long
    Entries(0 To 255) As Long
End Type

'=================================
'Meta File
Public Type PWMFRect16
    Left   As Integer
    Top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Public Type WmfPlaceableFileHeader
    Key         As Long                                                         ' GDIP_WMF_PLACEABLEKEY
    Hmf         As Integer                                                      ' Metafile HANDLE number (always 0)
    boundingBox As PWMFRect16                                                   ' Coordinates in metafile units
    Inch        As Integer                                                      ' Number of metafile units per inch
    Reserved    As Long                                                         ' Reserved (always 0)
    Checksum    As Integer                                                      ' Checksum value for previous 10 WORDs
End Type

Public Type ENHMETAHEADER3
    itype          As Long                                                      ' Record type EMR_HEADER
    nSize          As Long                                                      ' Record size in bytes.  This may be greater
    ' than the sizeof(ENHMETAHEADER).
    rclBounds      As RECTL                                                     ' Inclusive-inclusive bounds in device units
    rclFrame       As RECTL                                                     ' Inclusive-inclusive Picture Frame .01mm unit
    dSignature     As Long                                                      ' Signature.  Must be ENHMETA_SIGNATURE.
    nVersion       As Long                                                      ' Version number
    nBytes         As Long                                                      ' Size of the metafile in bytes
    nRecords       As Long                                                      ' Number of records in the metafile
    nHandles       As Integer                                                   ' Number of handles in the handle table
    ' Handle index zero is reserved.
    sReserved      As Integer                                                   ' Reserved.  Must be zero.
    nDescription   As Long                                                      ' Number of chars in the unicode desc string
    ' This is 0 if there is no description string
    offDescription As Long                                                      ' Offset to the metafile description record.
    ' This is 0 if there is no description string
    nPalEntries    As Long                                                      ' Number of entries in the metafile palette.
    szlDevice      As SIZEL                                                     ' Size of the reference device in pels
    szlMillimeters As SIZEL                                                     ' Size of the reference device in millimeters
End Type

Public Type METAHEADER
    mtType         As Integer
    mtHeaderSize   As Integer
    mtVersion      As Integer
    mtSize         As Long
    mtNoObjects    As Integer
    mtMaxRecord    As Long
    mtNoParameters As Integer
End Type

Public Type MetafileHeader
    mType             As MetafileType
    size              As Long                                                   ' Size of the metafile (in bytes)
    Version           As Long                                                   ' EMF+, EMF, or WMF version
    EmfPlusFlags      As Long
    DpiX              As Single
    DpiY              As Single
    X                 As Long                                                   ' Bounds in device units
    Y                 As Long
    Width             As Long
    Height            As Long

    EmfHeader         As ENHMETAHEADER3                                         ' NOTE: You'll have to use CopyMemory to view the METAHEADER type
    EmfPlusHeaderSize As Long                                                   ' size of the EMF+ header in file
    LogicalDpiX       As Long                                                   ' Logical Dpi of reference Hdc
    LogicalDpiY       As Long                                                   ' usually valid only for EMF+
End Type

'=================================
'Other
Public Type PropertyItem
    propId As Long                                                              ' ID of this property
    Length As Long                                                              ' Length of the property value, in bytes

    type   As Integer                                                           ' Type of the value, as one of TAG_TYPE_XXX
    ' defined above
    Value  As Long                                                              ' property value
End Type

Public Type CharacterRange
    First  As Long
    Length As Long
End Type

'=================================
'== Enums                       ==
'=================================

'=================================
'Image
Public Enum GpImageSaveFormat
    GpSaveBMP = 0
    GpSaveJPEG = 1
    GpSaveGIF = 2
    GpSavePNG = 3
    GpSaveTIFF = 4
End Enum

Public Enum GpImageFormatIdentifiers
    GpImageFormatUndefined = 0
    GpImageFormatMemoryBMP = 1
    GpImageFormatBMP = 2
    GpImageFormatEMF = 3
    GpImageFormatWMF = 4
    GpImageFormatJPEG = 5
    GpImageFormatPNG = 6
    GpImageFormatGIF = 7
    GpImageFormatTIFF = 8
    GpImageFormatEXIF = 9
    GpImageFormatIcon = 10
End Enum

Public Enum Image_Type
    ImageTypeUnknown = 0
    ImageTypeBitmap = 1
    ImageTypeMetafile = 2
End Enum

Public Enum Image_Property_Types
    PropertyTagTypeByte = 1
    PropertyTagTypeASCII = 2
    PropertyTagTypeShort = 3
    PropertyTagTypeLong = 4
    PropertyTagTypeRational = 5
    PropertyTagTypeUndefined = 7
    PropertyTagTypeSLONG = 9
    PropertyTagTypeSRational = 10
End Enum

Public Enum ImageCodecFlags
    ImageCodecFlagsEnCoder = &H1
    ImageCodecFlagsDecoder = &H2
    ImageCodecFlagsSupportBitmap = &H4
    ImageCodecFlagsSupportVector = &H8
    ImageCodecFlagsSeekableEnCode = &H10
    ImageCodecFlagsBlockingDecode = &H20
    
    ImageCodecFlagsBuiltin = &H10000
    ImageCodecFlagsSystem = &H20000
    ImageCodecFlagsUser = &H40000
End Enum

Public Enum Image_Property_ID_Tags
    PropertyTagExifIFD = &H8769
    PropertyTagGpsIFD = &H8825

    PropertyTagNewSubfileType = &HFE
    PropertyTagSubfileType = &HFF
    PropertyTagImageWidth = &H100
    PropertyTagImageHeight = &H101
    PropertyTagBitsPerSample = &H102
    PropertyTagCompression = &H103
    PropertyTagPhotometricInterp = &H106
    PropertyTagThreshHolding = &H107
    PropertyTagCellWidth = &H108
    PropertyTagCellHeight = &H109
    PropertyTagFillOrder = &H10A
    PropertyTagDocumentName = &H10D
    PropertyTagImageDescription = &H10E
    PropertyTagEquipMake = &H10F
    PropertyTagEquipModel = &H110
    PropertyTagStripOffsets = &H111
    PropertyTagOrientation = &H112
    PropertyTagSamplesPerPixel = &H115
    PropertyTagRowsPerStrip = &H116
    PropertyTagStripBytesCount = &H117
    PropertyTagMinSampleValue = &H118
    PropertyTagMaxSampleValue = &H119
    PropertyTagXResolution = &H11A                                              ' Image resolution in width direction
    PropertyTagYResolution = &H11B                                              ' Image resolution in height direction
    PropertyTagPlanarConfig = &H11C                                             ' Image data arrangement
    PropertyTagPageName = &H11D
    PropertyTagXPosition = &H11E
    PropertyTagYPosition = &H11F
    PropertyTagFreeOffset = &H120
    PropertyTagFreeByteCounts = &H121
    PropertyTagGrayResponseUnit = &H122
    PropertyTagGrayResponseCurve = &H123
    PropertyTagT4Option = &H124
    PropertyTagT6Option = &H125
    PropertyTagResolutionUnit = &H128                                           ' Unit of X and Y resolution
    PropertyTagPageNumber = &H129
    PropertyTagTransferFuncition = &H12D
    PropertyTagSoftwareUsed = &H131
    PropertyTagDateTime = &H132
    PropertyTagArtist = &H13B
    PropertyTagHostComputer = &H13C
    PropertyTagPredictor = &H13D
    PropertyTagWhitePoint = &H13E
    PropertyTagPrimaryChromaticities = &H13F
    PropertyTagColorMap = &H140
    PropertyTagHalftoneHints = &H141
    PropertyTagTileWidth = &H142
    PropertyTagTileLength = &H143
    PropertyTagTileOffset = &H144
    PropertyTagTileByteCounts = &H145
    PropertyTagInkSet = &H14C
    PropertyTagInkNames = &H14D
    PropertyTagNumberOfInks = &H14E
    PropertyTagDotRange = &H150
    PropertyTagTargetPrinter = &H151
    PropertyTagExtraSamples = &H152
    PropertyTagSampleFormat = &H153
    PropertyTagSMinSampleValue = &H154
    PropertyTagSMaxSampleValue = &H155
    PropertyTagTransferRange = &H156

    PropertyTagJPEGProc = &H200
    PropertyTagJPEGInterFormat = &H201
    PropertyTagJPEGInterLength = &H202
    PropertyTagJPEGRestartInterval = &H203
    PropertyTagJPEGLosslessPredictors = &H205
    PropertyTagJPEGPointTransforms = &H206
    PropertyTagJPEGQTables = &H207
    PropertyTagJPEGDCTables = &H208
    PropertyTagJPEGACTables = &H209

    PropertyTagYCbCrCoefficients = &H211
    PropertyTagYCbCrSubsampling = &H212
    PropertyTagYCbCrPositioning = &H213
    PropertyTagREFBlackWhite = &H214

    PropertyTagICCProfile = &H8773                                              ' This TAG is defined by ICC
    ' for embedded ICC in TIFF
    PropertyTagGamma = &H301
    PropertyTagICCProfileDescriptor = &H302
    PropertyTagSRGBRenderingIntent = &H303

    PropertyTagImageTitle = &H320
    PropertyTagCopyright = &H8298

    PropertyTagResolutionXUnit = &H5001
    PropertyTagResolutionYUnit = &H5002
    PropertyTagResolutionXLengthUnit = &H5003
    PropertyTagResolutionYLengthUnit = &H5004
    PropertyTagPrintFlags = &H5005
    PropertyTagPrintFlagsVersion = &H5006
    PropertyTagPrintFlagsCrop = &H5007
    PropertyTagPrintFlagsBleedWidth = &H5008
    PropertyTagPrintFlagsBleedWidthScale = &H5009
    PropertyTagHalftoneLPI = &H500A
    PropertyTagHalftoneLPIUnit = &H500B
    PropertyTagHalftoneDegree = &H500C
    PropertyTagHalftoneShape = &H500D
    PropertyTagHalftoneMisc = &H500E
    PropertyTagHalftoneScreen = &H500F
    PropertyTagJPEGQuality = &H5010
    PropertyTagGridSize = &H5011
    PropertyTagThumbnailFormat = &H5012                                         ' 1 = JPEG, 0 = RAW RGB
    PropertyTagThumbnailWidth = &H5013
    PropertyTagThumbnailHeight = &H5014
    PropertyTagThumbnailColorDepth = &H5015
    PropertyTagThumbnailPlanes = &H5016
    PropertyTagThumbnailRawBytes = &H5017
    PropertyTagThumbnailSize = &H5018
    PropertyTagThumbnailCompressedSize = &H5019
    PropertyTagColorTransferFunction = &H501A
    PropertyTagThumbnailData = &H501B
    PropertyTagThumbnailImageWidth = &H5020                                     ' Thumbnail width
    PropertyTagThumbnailImageHeight = &H5021                                    ' Thumbnail height
    PropertyTagThumbnailBitsPerSample = &H5022                                  ' Number of bits per
    ' component
    PropertyTagThumbnailCompression = &H5023                                    ' Compression Scheme
    PropertyTagThumbnailPhotometricInterp = &H5024                              ' Pixel composition
    PropertyTagThumbnailImageDescription = &H5025                               ' Image Tile
    PropertyTagThumbnailEquipMake = &H5026                                      ' Manufacturer of Image
    ' Input equipment
    PropertyTagThumbnailEquipModel = &H5027                                     ' Model of Image input
    ' equipment
    PropertyTagThumbnailStripOffsets = &H5028                                   ' Image data location
    PropertyTagThumbnailOrientation = &H5029                                    ' Orientation of image
    PropertyTagThumbnailSamplesPerPixel = &H502A                                ' Number of components
    PropertyTagThumbnailRowsPerStrip = &H502B                                   ' Number of rows per strip
    PropertyTagThumbnailStripBytesCount = &H502C                                ' Bytes per compressed
    ' strip
    PropertyTagThumbnailResolutionX = &H502D                                    ' Resolution in width
    ' direction
    PropertyTagThumbnailResolutionY = &H502E                                    ' Resolution in height
    ' direction
    PropertyTagThumbnailPlanarConfig = &H502F                                   ' Image data arrangement
    PropertyTagThumbnailResolutionUnit = &H5030                                 ' Unit of X and Y
    ' Resolution
    PropertyTagThumbnailTransferFunction = &H5031                               ' Transfer function
    PropertyTagThumbnailSoftwareUsed = &H5032                                   ' Software used
    PropertyTagThumbnailDateTime = &H5033                                       ' File change date and
    ' time
    PropertyTagThumbnailArtist = &H5034                                         ' Person who created the
    ' image
    PropertyTagThumbnailWhitePoint = &H5035                                     ' White point chromaticity
    PropertyTagThumbnailPrimaryChromaticities = &H5036
    ' Chromaticities of
    ' primaries
    PropertyTagThumbnailYCbCrCoefficients = &H5037                              ' Color space transforma-
    ' tion coefficients
    PropertyTagThumbnailYCbCrSubsampling = &H5038                               ' Subsampling ratio of Y
    ' to C
    PropertyTagThumbnailYCbCrPositioning = &H5039                               ' Y and C position
    PropertyTagThumbnailRefBlackWhite = &H503A                                  ' Pair of black and white
    ' reference values
    PropertyTagThumbnailCopyRight = &H503B                                      ' CopyRight holder

    PropertyTagLuminanceTable = &H5090
    PropertyTagChrominanceTable = &H5091

    PropertyTagFrameDelay = &H5100
    PropertyTagLoopCount = &H5101
#If GdipVersion >= 1.1 Then
    PropertyTagGlobalPalette = &H5102
    PropertyTagIndexBackground = &H5103
    PropertyTagIndexTransparent = &H5104
#End If

    PropertyTagPixelUnit = &H5110                                               ' Unit specifier for pixel/unit
    PropertyTagPixelPerUnitX = &H5111                                           ' Pixels per unit in X
    PropertyTagPixelPerUnitY = &H5112                                           ' Pixels per unit in Y
    PropertyTagPaletteHistogram = &H5113                                        ' Palette histogram

    PropertyTagExifExposureTime = &H829A
    PropertyTagExifFNumber = &H829D

    PropertyTagExifExposureProg = &H8822
    PropertyTagExifSpectralSense = &H8824
    PropertyTagExifISOSpeed = &H8827
    PropertyTagExifOECF = &H8828

    PropertyTagExifVer = &H9000
    PropertyTagExifDTOrig = &H9003                                              ' Date & time of original
    PropertyTagExifDTDigitized = &H9004                                         ' Date & time of digital data generation

    PropertyTagExifCompConfig = &H9101
    PropertyTagExifCompBPP = &H9102

    PropertyTagExifShutterSpeed = &H9201
    PropertyTagExifAperture = &H9202
    PropertyTagExifBrightness = &H9203
    PropertyTagExifExposureBias = &H9204
    PropertyTagExifMaxAperture = &H9205
    PropertyTagExifSubjectDist = &H9206
    PropertyTagExifMeteringMode = &H9207
    PropertyTagExifLightSource = &H9208
    PropertyTagExifFlash = &H9209
    PropertyTagExifFocalLength = &H920A
    PropertyTagExifMakerNote = &H927C
    PropertyTagExifUserComment = &H9286
    PropertyTagExifDTSubsec = &H9290                                            ' Date & Time subseconds
    PropertyTagExifDTOrigSS = &H9291                                            ' Date & Time original subseconds
    PropertyTagExifDTDigSS = &H9292                                             ' Date & TIme digitized subseconds

    PropertyTagExifFPXVer = &HA000
    PropertyTagExifColorSpace = &HA001
    PropertyTagExifPixXDim = &HA002
    PropertyTagExifPixYDim = &HA003
    PropertyTagExifRelatedWav = &HA004                                          ' related sound file
    PropertyTagExifInterop = &HA005
    PropertyTagExifFlashEnergy = &HA20B
    PropertyTagExifSpatialFR = &HA20C                                           ' Spatial Frequency Response
    PropertyTagExifFocalXRes = &HA20E                                           ' Focal Plane X Resolution
    PropertyTagExifFocalYRes = &HA20F                                           ' Focal Plane Y Resolution
    PropertyTagExifFocalResUnit = &HA210                                        ' Focal Plane Resolution Unit
    PropertyTagExifSubjectLoc = &HA214
    PropertyTagExifExposureIndex = &HA215
    PropertyTagExifSensingMethod = &HA217
    PropertyTagExifFileSource = &HA300
    PropertyTagExifSceneType = &HA301
    PropertyTagExifCfaPattern = &HA302

    PropertyTagGpsVer = &H0
    PropertyTagGpsLatitudeRef = &H1
    PropertyTagGpsLatitude = &H2
    PropertyTagGpsLongitudeRef = &H3
    PropertyTagGpsLongitude = &H4
    PropertyTagGpsAltitudeRef = &H5
    PropertyTagGpsAltitude = &H6
    PropertyTagGpsGpsTime = &H7
    PropertyTagGpsGpsSatellites = &H8
    PropertyTagGpsGpsStatus = &H9
    PropertyTagGpsGpsMeasureMode = &HA
    PropertyTagGpsGpsDop = &HB                                                  ' Measurement precision
    PropertyTagGpsSpeedRef = &HC
    PropertyTagGpsSpeed = &HD
    PropertyTagGpsTrackRef = &HE
    PropertyTagGpsTrack = &HF
    PropertyTagGpsImgDirRef = &H10
    PropertyTagGpsImgDir = &H11
    PropertyTagGpsMapDatum = &H12
    PropertyTagGpsDestLatRef = &H13
    PropertyTagGpsDestLat = &H14
    PropertyTagGpsDestLongRef = &H15
    PropertyTagGpsDestLong = &H16
    PropertyTagGpsDestBearRef = &H17
    PropertyTagGpsDestBear = &H18
    PropertyTagGpsDestDistRef = &H19
    PropertyTagGpsDestDist = &H1A
End Enum

'=================================
'Palette
Public Enum PaletteFlags
    PaletteFlagsHasAlpha = &H1
    PaletteFlagsGrayScale = &H2
    PaletteFlagsHalftone = &H4
End Enum

'=================================
'Rotate
Public Enum RotateFlipType
    RotateNoneFlipNone = 0
    Rotate90FlipNone = 1
    Rotate180FlipNone = 2
    Rotate270FlipNone = 3

    RotateNoneFlipX = 4
    Rotate90FlipX = 5
    Rotate180FlipX = 6
    Rotate270FlipX = 7

    RotateNoneFlipY = Rotate180FlipX
    Rotate90FlipY = Rotate270FlipX
    Rotate180FlipY = RotateNoneFlipX
    Rotate270FlipY = Rotate90FlipX

    RotateNoneFlipXY = Rotate180FlipNone
    Rotate90FlipXY = Rotate270FlipNone
    Rotate180FlipXY = RotateNoneFlipNone
    Rotate270FlipXY = Rotate90FlipNone
End Enum

'=================================
'Colors
Public Enum colors
    AliceBlue = &HFFF0F8FF
    AntiqueWhite = &HFFFAEBD7
    Aqua = &HFF00FFFF
    Aquamarine = &HFF7FFFD4
    Azure = &HFFF0FFFF
    Beige = &HFFF5F5DC
    Bisque = &HFFFFE4C4
    Black = &HFF000000
    BlanchedAlmond = &HFFFFEBCD
    Blue = &HFF0000FF
    BlueViolet = &HFF8A2BE2
    Brown = &HFFA52A2A
    BurlyWood = &HFFDEB887
    CadetBlue = &HFF5F9EA0
    Chartreuse = &HFF7FFF00
    Chocolate = &HFFD2691E
    Coral = &HFFFF7F50
    CornflowerBlue = &HFF6495ED
    Cornsilk = &HFFFFF8DC
    Crimson = &HFFDC143C
    Cyan = &HFF00FFFF
    DarkBlue = &HFF00008B
    DarkCyan = &HFF008B8B
    DarkGoldenrod = &HFFB8860B
    DarkGray = &HFFA9A9A9
    DarkGreen = &HFF006400
    DarkKhaki = &HFFBDB76B
    DarkMagenta = &HFF8B008B
    DarkOliveGreen = &HFF556B2F
    DarkOrange = &HFFFF8C00
    DarkOrchid = &HFF9932CC
    DarkRed = &HFF8B0000
    DarkSalmon = &HFFE9967A
    DarkSeaGreen = &HFF8FBC8B
    DarkSlateBlue = &HFF483D8B
    DarkSlateGray = &HFF2F4F4F
    DarkTurquoise = &HFF00CED1
    DarkViolet = &HFF9400D3
    DeepPink = &HFFFF1493
    DeepSkyBlue = &HFF00BFFF
    DimGray = &HFF696969
    DodgerBlue = &HFF1E90FF
    Firebrick = &HFFB22222
    FloralWhite = &HFFFFFAF0
    ForestGreen = &HFF228B22
    Fuchsia = &HFFFF00FF
    Gainsboro = &HFFDCDCDC
    GhostWhite = &HFFF8F8FF
    Gold = &HFFFFD700
    Goldenrod = &HFFDAA520
    Gray = &HFF808080
    Green = &HFF008000
    GreenYellow = &HFFADFF2F
    Honeydew = &HFFF0FFF0
    HotPink = &HFFFF69B4
    IndianRed = &HFFCD5C5C
    Indigo = &HFF4B0082
    Ivory = &HFFFFFFF0
    Khaki = &HFFF0E68C
    Lavender = &HFFE6E6FA
    LavenderBlush = &HFFFFF0F5
    LawnGreen = &HFF7CFC00
    LemonChiffon = &HFFFFFACD
    LightBlue = &HFFADD8E6
    LightCoral = &HFFF08080
    LightCyan = &HFFE0FFFF
    LightGoldenrodYellow = &HFFFAFAD2
    LightGray = &HFFD3D3D3
    LightGreen = &HFF90EE90
    LightPink = &HFFFFB6C1
    LightSalmon = &HFFFFA07A
    LightSeaGreen = &HFF20B2AA
    LightSkyBlue = &HFF87CEFA
    LightSlateGray = &HFF778899
    LightSteelBlue = &HFFB0C4DE
    LightYellow = &HFFFFFFE0
    Lime = &HFF00FF00
    LimeGreen = &HFF32CD32
    Linen = &HFFFAF0E6
    Magenta = &HFFFF00FF
    Maroon = &HFF800000
    MediumAquamarine = &HFF66CDAA
    MediumBlue = &HFF0000CD
    MediumOrchid = &HFFBA55D3
    MediumPurple = &HFF9370DB
    MediumSeaGreen = &HFF3CB371
    MediumSlateBlue = &HFF7B68EE
    MediumSpringGreen = &HFF00FA9A
    MediumTurquoise = &HFF48D1CC
    MediumVioletRed = &HFFC71585
    MidnightBlue = &HFF191970
    MintCream = &HFFF5FFFA
    MistyRose = &HFFFFE4E1
    Moccasin = &HFFFFE4B5
    NavajoWhite = &HFFFFDEAD
    Navy = &HFF000080
    OldLace = &HFFFDF5E6
    Olive = &HFF808000
    OliveDrab = &HFF6B8E23
    Orange = &HFFFFA500
    OrangeRed = &HFFFF4500
    Orchid = &HFFDA70D6
    PaleGoldenrod = &HFFEEE8AA
    PaleGreen = &HFF98FB98
    PaleTurquoise = &HFFAFEEEE
    PaleVioletRed = &HFFDB7093
    PapayaWhip = &HFFFFEFD5
    PeachPuff = &HFFFFDAB9
    Peru = &HFFCD853F
    Pink = &HFFFFC0CB
    Plum = &HFFDDA0DD
    PowderBlue = &HFFB0E0E6
    Purple = &HFF800080
    Red = &HFFFF0000
    RosyBrown = &HFFBC8F8F
    RoyalBlue = &HFF4169E1
    SaddleBrown = &HFF8B4513
    Salmon = &HFFFA8072
    SandyBrown = &HFFF4A460
    SeaGreen = &HFF2E8B57
    SeaShell = &HFFFFF5EE
    Sienna = &HFFA0522D
    Silver = &HFFC0C0C0
    SkyBlue = &HFF87CEEB
    SlateBlue = &HFF6A5ACD
    SlateGray = &HFF708090
    Snow = &HFFFFFAFA
    SpringGreen = &HFF00FF7F
    SteelBlue = &HFF4682B4
    Tan = &HFFD2B48C
    Teal = &HFF008080
    Thistle = &HFFD8BFD8
    Tomato = &HFFFF6347
    Transparent = &HFFFFFF
    Turquoise = &HFF40E0D0
    Violet = &HFFEE82EE
    Wheat = &HFFF5DEB3
    White = &HFFFFFFFF
    WhiteSmoke = &HFFF5F5F5
    Yellow = &HFFFFFF00
    YellowGreen = &HFF9ACD32
End Enum

Public Enum ColorMode
    ColorModeARGB32 = 0
    ColorModeARGB64 = 1
End Enum

Public Enum ColorChannelFlags
    ColorChannelFlagsC = 0
    ColorChannelFlagsM
    ColorChannelFlagsY
    ColorChannelFlagsK
    ColorChannelFlagsLast
End Enum

Public Enum ColorShiftComponents
    AlphaShift = 24
    RedShift = 16
    GreenShift = 8
    BlueShift = 0
End Enum

Public Enum ColorMaskComponents
    AlphaMask = &HFF000000
    RedMask = &HFF0000
    GreenMask = &HFF00
    BlueMask = &HFF
End Enum

'=================================
'String
Public Enum StringFormatFlags
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000

    StringFormatFlagsNoClip = &H4000
End Enum

Public Enum StringTrimming
    StringTrimmingNone = 0
    StringTrimmingCharacter = 1
    StringTrimmingWord = 2
    StringTrimmingEllipsisCharacter = 3
    StringTrimmingEllipsisWord = 4
    StringTrimmingEllipsisPath = 5
End Enum

Public Enum StringDigitSubstitute
    StringDigitSubstituteUser = 0
    StringDigitSubstituteNone = 1
    StringDigitSubstituteNational = 2
    StringDigitSubstituteTraditional = 3
End Enum

'=================================
'Pen / Brush
Public Enum HatchStyle
    HatchStyleHorizontal                                                        ' 0
    HatchStyleVertical                                                          ' 1
    HatchStyleForwardDiagonal                                                   ' 2
    HatchStyleBackwardDiagonal                                                  ' 3
    HatchStyleCross                                                             ' 4
    HatchStyleDiagonalCross                                                     ' 5
    HatchStyle05Percent                                                         ' 6
    HatchStyle10Percent                                                         ' 7
    HatchStyle20Percent                                                         ' 8
    HatchStyle25Percent                                                         ' 9
    HatchStyle30Percent                                                         ' 10
    HatchStyle40Percent                                                         ' 11
    HatchStyle50Percent                                                         ' 12
    HatchStyle60Percent                                                         ' 13
    HatchStyle70Percent                                                         ' 14
    HatchStyle75Percent                                                         ' 15
    HatchStyle80Percent                                                         ' 16
    HatchStyle90Percent                                                         ' 17
    HatchStyleLightDownwardDiagonal                                             ' 18
    HatchStyleLightUpwardDiagonal                                               ' 19
    HatchStyleDarkDownwardDiagonal                                              ' 20
    HatchStyleDarkUpwardDiagonal                                                ' 21
    HatchStyleWideDownwardDiagonal                                              ' 22
    HatchStyleWideUpwardDiagonal                                                ' 23
    HatchStyleLightVertical                                                     ' 24
    HatchStyleLightHorizontal                                                   ' 25
    HatchStyleNarrowVertical                                                    ' 26
    HatchStyleNarrowHorizontal                                                  ' 27
    HatchStyleDarkVertical                                                      ' 28
    HatchStyleDarkHorizontal                                                    ' 29
    HatchStyleDashedDownwardDiagonal                                            ' 30
    HatchStyleDashedUpwardDiagonal                                              ' 31
    HatchStyleDashedHorizontal                                                  ' 32
    HatchStyleDashedVertical                                                    ' 33
    HatchStyleSmallConfetti                                                     ' 34
    HatchStyleLargeConfetti                                                     ' 35
    HatchStyleZigZag                                                            ' 36
    HatchStyleWave                                                              ' 37
    HatchStyleDiagonalBrick                                                     ' 38
    HatchStyleHorizontalBrick                                                   ' 39
    HatchStyleWeave                                                             ' 40
    HatchStylePlaid                                                             ' 41
    HatchStyleDivot                                                             ' 42
    HatchStyleDottedGrid                                                        ' 43
    HatchStyleDottedDiamond                                                     ' 44
    HatchStyleShingle                                                           ' 45
    HatchStyleTrellis                                                           ' 46
    HatchStyleSphere                                                            ' 47
    HatchStyleSmallGrid                                                         ' 48
    HatchStyleSmallCheckerBoard                                                 ' 49
    HatchStyleLargeCheckerBoard                                                 ' 50
    HatchStyleOutlinedDiamond                                                   ' 51
    HatchStyleSolidDiamond                                                      ' 52

    HatchStyleTotal
    HatchStyleLargeGrid = HatchStyleCross                                       ' 4

    HatchStyleMin = HatchStyleHorizontal
    HatchStyleMax = HatchStyleTotal - 1
End Enum

Public Enum PenAlignment
    PenAlignmentCenter = 0
    PenAlignmentInset = 1
End Enum

Public Enum BrushType
    BrushTypeSolidColor = 0
    BrushTypeHatchFill = 1
    BrushTypeTextureFill = 2
    BrushTypePathGradient = 3
    BrushTypeLinearGradient = 4
End Enum

Public Enum DashStyle
    DashStyleSolid
    DashStyleDash
    DashStyleDot
    DashStyleDashDot
    DashStyleDashDotDot
    DashStyleCustom
End Enum

Public Enum DashCap
    DashCapFlat = 0
    DashCapRound = 2
    DashCapTriangle = 3
End Enum

Public Enum LineCap
    LineCapFlat = 0
    LineCapSquare = 1
    LineCapRound = 2
    LineCapTriangle = 3

    LineCapNoAnchor = &H10                                                      ' corresponds to flat cap
    LineCapSquareAnchor = &H11                                                  ' corresponds to square cap
    LineCapRoundAnchor = &H12                                                   ' corresponds to round cap
    LineCapDiamondAnchor = &H13                                                 ' corresponds to triangle cap
    LineCapArrowAnchor = &H14                                                   ' no correspondence

    LineCapCustom = &HFF                                                        ' custom cap

    LineCapAnchorMask = &HF0                                                    ' mask to check for anchor or not.
End Enum

Public Enum CustomLineCapType
    CustomLineCapTypeDefault = 0
    CustomLineCapTypeAdjustableArrow = 1
End Enum

Public Enum LineJoin
    LineJoinMiter = 0
    LineJoinBevel = 1
    LineJoinRound = 2
    LineJoinMiterClipped = 3
End Enum

Public Enum PenType
    PenTypeSolidColor = BrushTypeSolidColor
    PenTypeHatchFill = BrushTypeHatchFill
    PenTypeTextureFill = BrushTypeTextureFill
    PenTypePathGradient = BrushTypePathGradient
    PenTypeLinearGradient = BrushTypeLinearGradient
    PenTypeUnknown = -1
End Enum

'=================================
'Meta File
Public Enum MetafileType
    MetafileTypeInvalid                                                         ' Invalid metafile
    MetafileTypeWmf                                                             ' Standard WMF
    MetafileTypeWmfPlaceable                                                    ' Placeable WMF
    MetafileTypeEmf                                                             ' EMF (not EMF+)
    MetafileTypeEmfPlusOnly                                                     ' EMF+ without dual down-level records
    MetafileTypeEmfPlusDual                                                     ' EMF+ with dual down-level records
End Enum

Public Enum emfType
    EmfTypeEmfOnly = MetafileTypeEmf                                            ' no EMF+  only EMF
    EmfTypeEmfPlusOnly = MetafileTypeEmfPlusOnly                                ' no EMF  only EMF+
    EmfTypeEmfPlusDual = MetafileTypeEmfPlusDual                                ' both EMF+ and EMF
End Enum

Public Enum ObjectType
    ObjectTypeInvalid
    ObjectTypeBrush
    ObjectTypePen
    ObjectTypePath
    ObjectTypeRegion
    ObjectTypeImage
    ObjectTypeFont
    ObjectTypeStringFormat
    ObjectTypeImageAttributes
    ObjectTypeCustomLineCap
#If GdipVersion >= 1.1 Then
    ObjectTypeGraphics
    ObjectTypeMax = ObjectTypeGraphics
#Else
    ObjectTypeMax = ObjectTypeCustomLineCap
#End If
    ObjectTypeMin = ObjectTypeBrush
End Enum

Public Enum MetafileFrameUnit
    MetafileFrameUnitPixel = UnitPixel
    MetafileFrameUnitPoint = UnitPoint
    MetafileFrameUnitInch = UnitInch
    MetafileFrameUnitDocument = UnitDocument
    MetafileFrameUnitMillimeter = UnitMillimeter
    MetafileFrameUnitGdi                                                        ' GDI compatible .01 MM units
End Enum

' Coordinate space identifiers
Public Enum CoordinateSpace
    CoordinateSpaceWorld                                                        ' 0
    CoordinateSpacePage                                                         ' 1
    CoordinateSpaceDevice                                                       ' 2
End Enum

Public Enum EmfPlusRecordType
    WmfRecordTypeSetBkColor = &H10201
    WmfRecordTypeSetBkMode = &H10102
    WmfRecordTypeSetMapMode = &H10103
    WmfRecordTypeSetROP2 = &H10104
    WmfRecordTypeSetRelAbs = &H10105
    WmfRecordTypeSetPolyFillMode = &H10106
    WmfRecordTypeSetStretchBltMode = &H10107
    WmfRecordTypeSetTextCharExtra = &H10108
    WmfRecordTypeSetTextColor = &H10209
    WmfRecordTypeSetTextJustification = &H1020A
    WmfRecordTypeSetWindowOrg = &H1020B
    WmfRecordTypeSetWindowExt = &H1020C
    WmfRecordTypeSetViewportOrg = &H1020D
    WmfRecordTypeSetViewportExt = &H1020E
    WmfRecordTypeOffsetWindowOrg = &H1020F
    WmfRecordTypeScaleWindowExt = &H10410
    WmfRecordTypeOffsetViewportOrg = &H10211
    WmfRecordTypeScaleViewportExt = &H10412
    WmfRecordTypeLineTo = &H10213
    WmfRecordTypeMoveTo = &H10214
    WmfRecordTypeExcludeClipRect = &H10415
    WmfRecordTypeIntersectClipRect = &H10416
    WmfRecordTypeArc = &H10817
    WmfRecordTypeEllipse = &H10418
    WmfRecordTypeFloodFill = &H10419
    WmfRecordTypePie = &H1081A
    WmfRecordTypeRectangle = &H1041B
    WmfRecordTypeRoundRect = &H1061C
    WmfRecordTypePatBlt = &H1061D
    WmfRecordTypeSaveDC = &H1001E
    WmfRecordTypeSetPixel = &H1041F
    WmfRecordTypeOffsetClipRgn = &H10220
    WmfRecordTypeTextOut = &H10521
    WmfRecordTypeBitBlt = &H10922
    WmfRecordTypeStretchBlt = &H10B23
    WmfRecordTypePolygon = &H10324
    WmfRecordTypePolyline = &H10325
    WmfRecordTypeEscape = &H10626
    WmfRecordTypeRestoreDC = &H10127
    WmfRecordTypeFillRegion = &H10228
    WmfRecordTypeFrameRegion = &H10429
    WmfRecordTypeInvertRegion = &H1012A
    WmfRecordTypePaintRegion = &H1012B
    WmfRecordTypeSelectClipRegion = &H1012C
    WmfRecordTypeSelectObject = &H1012D
    WmfRecordTypeSetTextAlign = &H1012E
    WmfRecordTypeDrawText = &H1062F
    WmfRecordTypeChord = &H10830
    WmfRecordTypeSetMapperFlags = &H10231
    WmfRecordTypeExtTextOut = &H10A32
    WmfRecordTypeSetDIBToDev = &H10D33
    WmfRecordTypeSelectPalette = &H10234
    WmfRecordTypeRealizePalette = &H10035
    WmfRecordTypeAnimatePalette = &H10436
    WmfRecordTypeSetPalEntries = &H10037
    WmfRecordTypePolyPolygon = &H10538
    WmfRecordTypeResizePalette = &H10139
    WmfRecordTypeDIBBitBlt = &H10940
    WmfRecordTypeDIBStretchBlt = &H10B41
    WmfRecordTypeDIBCreatePatternBrush = &H10142
    WmfRecordTypeStretchDIB = &H10F43
    WmfRecordTypeExtFloodFill = &H10548
    WmfRecordTypeSetLayout = &H10149
    WmfRecordTypeResetDC = &H1014C
    WmfRecordTypeStartDoc = &H1014D
    WmfRecordTypeStartPage = &H1004F
    WmfRecordTypeEndPage = &H10050
    WmfRecordTypeAbortDoc = &H10052
    WmfRecordTypeEndDoc = &H1005E
    WmfRecordTypeDeleteObject = &H101F0
    WmfRecordTypeCreatePalette = &H100F7
    WmfRecordTypeCreateBrush = &H100F8
    WmfRecordTypeCreatePatternBrush = &H101F9
    WmfRecordTypeCreatePenIndirect = &H102FA
    WmfRecordTypeCreateFontIndirect = &H102FB
    WmfRecordTypeCreateBrushIndirect = &H102FC
    WmfRecordTypeCreateBitmapIndirect = &H102FD
    WmfRecordTypeCreateBitmap = &H106FE
    WmfRecordTypeCreateRegion = &H106FF
    EmfRecordTypeHeader = 1
    EmfRecordTypePolyBezier = 2
    EmfRecordTypePolygon = 3
    EmfRecordTypePolyline = 4
    EmfRecordTypePolyBezierTo = 5
    EmfRecordTypePolyLineTo = 6
    EmfRecordTypePolyPolyline = 7
    EmfRecordTypePolyPolygon = 8
    EmfRecordTypeSetWindowExtEx = 9
    EmfRecordTypeSetWindowOrgEx = 10
    EmfRecordTypeSetViewportExtEx = 11
    EmfRecordTypeSetViewportOrgEx = 12
    EmfRecordTypeSetBrushOrgEx = 13
    EmfRecordTypeEOF = 14
    EmfRecordTypeSetPixelV = 15
    EmfRecordTypeSetMapperFlags = 16
    EmfRecordTypeSetMapMode = 17
    EmfRecordTypeSetBkMode = 18
    EmfRecordTypeSetPolyFillMode = 19
    EmfRecordTypeSetROP2 = 20
    EmfRecordTypeSetStretchBltMode = 21
    EmfRecordTypeSetTextAlign = 22
    EmfRecordTypeSetColorAdjustment = 23
    EmfRecordTypeSetTextColor = 24
    EmfRecordTypeSetBkColor = 25
    EmfRecordTypeOffsetClipRgn = 26
    EmfRecordTypeMoveToEx = 27
    EmfRecordTypeSetMetaRgn = 28
    EmfRecordTypeExcludeClipRect = 29
    EmfRecordTypeIntersectClipRect = 30
    EmfRecordTypeScaleViewportExtEx = 31
    EmfRecordTypeScaleWindowExtEx = 32
    EmfRecordTypeSaveDC = 33
    EmfRecordTypeRestoreDC = 34
    EmfRecordTypeSetWorldTransform = 35
    EmfRecordTypeModifyWorldTransform = 36
    EmfRecordTypeSelectObject = 37
    EmfRecordTypeCreatePen = 38
    EmfRecordTypeCreateBrushIndirect = 39
    EmfRecordTypeDeleteObject = 40
    EmfRecordTypeAngleArc = 41
    EmfRecordTypeEllipse = 42
    EmfRecordTypeRectangle = 43
    EmfRecordTypeRoundRect = 44
    EmfRecordTypeArc = 45
    EmfRecordTypeChord = 46
    EmfRecordTypePie = 47
    EmfRecordTypeSelectPalette = 48
    EmfRecordTypeCreatePalette = 49
    EmfRecordTypeSetPaletteEntries = 50
    EmfRecordTypeResizePalette = 51
    EmfRecordTypeRealizePalette = 52
    EmfRecordTypeExtFloodFill = 53
    EmfRecordTypeLineTo = 54
    EmfRecordTypeArcTo = 55
    EmfRecordTypePolyDraw = 56
    EmfRecordTypeSetArcDirection = 57
    EmfRecordTypeSetMiterLimit = 58
    EmfRecordTypeBeginPath = 59
    EmfRecordTypeEndPath = 60
    EmfRecordTypeCloseFigure = 61
    EmfRecordTypeFillPath = 62
    EmfRecordTypeStrokeAndFillPath = 63
    EmfRecordTypeStrokePath = 64
    EmfRecordTypeFlattenPath = 65
    EmfRecordTypeWidenPath = 66
    EmfRecordTypeSelectClipPath = 67
    EmfRecordTypeAbortPath = 68
    EmfRecordTypeReserved_069 = 69
    EmfRecordTypeGdiComment = 70
    EmfRecordTypeFillRgn = 71
    EmfRecordTypeFrameRgn = 72
    EmfRecordTypeInvertRgn = 73
    EmfRecordTypePaintRgn = 74
    EmfRecordTypeExtSelectClipRgn = 75
    EmfRecordTypeBitBlt = 76
    EmfRecordTypeStretchBlt = 77
    EmfRecordTypeMaskBlt = 78
    EmfRecordTypePlgBlt = 79
    EmfRecordTypeSetDIBitsToDevice = 80
    EmfRecordTypeStretchDIBits = 81
    EmfRecordTypeExtCreateFontIndirect = 82
    EmfRecordTypeExtTextOutA = 83
    EmfRecordTypeExtTextOutW = 84
    EmfRecordTypePolyBezier16 = 85
    EmfRecordTypePolygon16 = 86
    EmfRecordTypePolyline16 = 87
    EmfRecordTypePolyBezierTo16 = 88
    EmfRecordTypePolylineTo16 = 89
    EmfRecordTypePolyPolyline16 = 90
    EmfRecordTypePolyPolygon16 = 91
    EmfRecordTypePolyDraw16 = 92
    EmfRecordTypeCreateMonoBrush = 93
    EmfRecordTypeCreateDIBPatternBrushPt = 94
    EmfRecordTypeExtCreatePen = 95
    EmfRecordTypePolyTextOutA = 96
    EmfRecordTypePolyTextOutW = 97
    EmfRecordTypeSetICMMode = 98
    EmfRecordTypeCreateColorSpace = 99
    EmfRecordTypeSetColorSpace = 100
    EmfRecordTypeDeleteColorSpace = 101
    EmfRecordTypeGLSRecord = 102
    EmfRecordTypeGLSBoundedRecord = 103
    EmfRecordTypePixelFormat = 104
    EmfRecordTypeDrawEscape = 105
    EmfRecordTypeExtEscape = 106
    EmfRecordTypeStartDoc = 107
    EmfRecordTypeSmallTextOut = 108
    EmfRecordTypeForceUFIMapping = 109
    EmfRecordTypeNamedEscape = 110
    EmfRecordTypeColorCorrectPalette = 111
    EmfRecordTypeSetICMProfileA = 112
    EmfRecordTypeSetICMProfileW = 113
    EmfRecordTypeAlphaBlend = 114
    EmfRecordTypeSetLayout = 115
    EmfRecordTypeTransparentBlt = 116
    EmfRecordTypeReserved_117 = 117
    EmfRecordTypeGradientFill = 118
    EmfRecordTypeSetLinkedUFIs = 119
    EmfRecordTypeSetTextJustification = 120
    EmfRecordTypeColorMatchToTargetW = 121
    EmfRecordTypeCreateColorSpaceW = 122
    EmfRecordTypeMax = 122
    EmfRecordTypeMin = 1
    
    EmfPlusRecordTypeInvalid = 16384                                            '//GDIP_EMFPLUS_RECORD_BASE
    EmfPlusRecordTypeHeader = 16385
    EmfPlusRecordTypeEndOfFile = 16386
    EmfPlusRecordTypeComment = 16387
    EmfPlusRecordTypeGetDC = 16388
    EmfPlusRecordTypeMultiFormatStart = 16389
    EmfPlusRecordTypeMultiFormatSection = 16390
    EmfPlusRecordTypeMultiFormatEnd = 16391
    
    EmfPlusRecordTypeObject = 16392
    
    EmfPlusRecordTypeClear = 16393
    EmfPlusRecordTypeFillRects = 16394
    EmfPlusRecordTypeDrawRects = 16395
    EmfPlusRecordTypeFillPolygon = 16396
    EmfPlusRecordTypeDrawLines = 16397
    EmfPlusRecordTypeFillEllipse = 16398
    EmfPlusRecordTypeDrawEllipse = 16399
    EmfPlusRecordTypeFillPie = 16400
    EmfPlusRecordTypeDrawPie = 16401
    EmfPlusRecordTypeDrawArc = 16402
    EmfPlusRecordTypeFillRegion = 16403
    EmfPlusRecordTypeFillPath = 16404
    EmfPlusRecordTypeDrawPath = 16405
    EmfPlusRecordTypeFillClosedCurve = 16406
    EmfPlusRecordTypeDrawClosedCurve = 16407
    EmfPlusRecordTypeDrawCurve = 16408
    EmfPlusRecordTypeDrawBeziers = 16409
    EmfPlusRecordTypeDrawImage = 16410
    EmfPlusRecordTypeDrawImagePoints = 16411
    EmfPlusRecordTypeDrawString = 16412
    
    EmfPlusRecordTypeSetRenderingOrigin = 16413
    EmfPlusRecordTypeSetAntiAliasMode = 16414
    EmfPlusRecordTypeSetTextRenderingHint = 16415
    EmfPlusRecordTypeSetTextContrast = 16416
    EmfPlusRecordTypeSetInterpolationMode = 16417
    EmfPlusRecordTypeSetPixelOffsetMode = 16418
    EmfPlusRecordTypeSetCompositingMode = 16419
    EmfPlusRecordTypeSetCompositingQuality = 16420
    EmfPlusRecordTypeSave = 16421
    EmfPlusRecordTypeRestore = 16422
    EmfPlusRecordTypeBeginContainer = 16423
    EmfPlusRecordTypeBeginContainerNoParams = 16424
    EmfPlusRecordTypeEndContainer = 16425
    EmfPlusRecordTypeSetWorldTransform = 16426
    EmfPlusRecordTypeResetWorldTransform = 16427
    EmfPlusRecordTypeMultiplyWorldTransform = 16428
    EmfPlusRecordTypeTranslateWorldTransform = 16429
    EmfPlusRecordTypeScaleWorldTransform = 16430
    EmfPlusRecordTypeRotateWorldTransform = 16431
    EmfPlusRecordTypeSetPageTransform = 16432
    EmfPlusRecordTypeResetClip = 16433
    EmfPlusRecordTypeSetClipRect = 16434
    EmfPlusRecordTypeSetClipPath = 16435
    EmfPlusRecordTypeSetClipRegion = 16436
    EmfPlusRecordTypeOffsetClip = 16437
    EmfPlusRecordTypeDrawDriverString = 16438
#If GdipVersion >= 1.1 Then
    EmfPlusRecordTypeStrokeFillPath = 16439
    EmfPlusRecordTypeSerializableObject = 16440
    EmfPlusRecordTypeSetTSGraphics = 16441
    EmfPlusRecordTypeSetTSClip = 16442
    EmfPlusRecordTotal = 16443
#Else
    EmfPlusRecordTotal = 16439
#End If
    EmfPlusRecordTypeMax = EmfPlusRecordTotal - 1
    EmfPlusRecordTypeMin = EmfPlusRecordTypeHeader
End Enum

'=================================
'Other
Public Enum HotkeyPrefix
    HotkeyPrefixNone = 0
    HotkeyPrefixShow = 1
    HotkeyPrefixHide = 2
End Enum

Public Enum FlushIntention
    FlushIntentionFlush = 0                                                     ' Flush all batched rendering operations
    FlushIntentionSync = 1                                                      ' Flush all batched rendering operations
End Enum

Public Enum EnCoderParameterValueType
    EnCoderParameterValueTypeByte = 1                                           ' 8-bit unsigned int
    EnCoderParameterValueTypeASCII = 2                                          ' 8-bit byte containing one 7-bit ASCII
    ' code. NULL terminated.
    EnCoderParameterValueTypeShort = 3                                          ' 16-bit unsigned int
    EnCoderParameterValueTypeLong = 4                                           ' 32-bit unsigned int
    EnCoderParameterValueTypeRational = 5                                       ' Two Longs. The first Long is the
    ' numerator the second Long expresses the
    ' denomintor.
    EnCoderParameterValueTypeLongRange = 6                                      ' Two longs which specify a range of
    ' integer values. The first Long specifies
    ' the lower end and the second one
    ' specifies the higher end. All values
    ' are inclusive at both ends
    EnCoderParameterValueTypeUndefined = 7                                      ' 8-bit byte that can take any value
    ' depending on field definition
    EnCoderParameterValueTypeRationalRange = 8                                  ' Two Rationals. The first Rational
    ' specifies the lower end and the second
    ' specifies the higher end. All values
    ' are inclusive at both ends
#If GdipVersion >= 1.1 Then
    EnCoderParameterValueTypePointer = 9                                        ' a pointer to a parameter defined data.
#End If
End Enum

Public Enum EnCoderValue
    EnCoderValueColorTypeCMYK = 0
    EnCoderValueColorTypeYCCK
    EnCoderValueCompressionLZW
    EnCoderValueCompressionCCITT3
    EnCoderValueCompressionCCITT4
    EnCoderValueCompressionRle
    EnCoderValueCompressionNone
    EnCoderValueScanMethodInterlaced
    EnCoderValueScanMethodNonInterlaced
    EnCoderValueVersionGif87
    EnCoderValueVersionGif89
    EnCoderValueRenderProgressive
    EnCoderValueRenderNonProgressive
    EnCoderValueTransformRotate90
    EnCoderValueTransformRotate180
    EnCoderValueTransformRotate270
    EnCoderValueTransformFlipHorizontal
    EnCoderValueTransformFlipVertical
    EnCoderValueMultiFrame
    EnCoderValueLastFrame
    EnCoderValueFlush
    EnCoderValueFrameDimensionTime
    EnCoderValueFrameDimensionResolution
    EnCoderValueFrameDimensionPage
#If GdipVersion >= 1.1 Then
    EnCoderValueColorTypeGray
    EnCoderValueColorTypeRGB
#End If
End Enum

#If GdipVersion >= 1.1 Then
Public Enum ConvertToEmfPlusFlags
    ConvertToEmfPlusFlagsDefault = 0
    ConvertToEmfPlusFlagsRopUsed = 1
    ConvertToEmfPlusFlagsText = 2
    ConvertToEmfPlusFlagsInvalidRecord = 4
End Enum
#End If

Public Enum DebugEventLevel
    DebugEventLevelFatal = 0
    DebugEventLevelWarning
End Enum

Public Enum GpTestControlEnum
    TestControlForceBilinear = 0
    TestControlNoICM = 1
    TestControlGetBuildNumber = 2
End Enum

Public Declare Function GdipCreateFromHDC2 _
    Lib "GDIPlus" (ByVal hDC As Long, _
    ByVal hDevice As Long, _
    graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWNDICM _
    Lib "GDIPlus" (ByVal hWnd As Long, _
    graphics As Long) As GpStatus

Public Declare Function GdipEnumerateMetafileDestPoint _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destPoint As POINTF, _
    lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPointI _
    Lib "GDIPlus" (graphics As Long, _
    ByVal metafile As Long, _
    destPoint As POINTL, _
    ByVal lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestRect _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destRect As RECTF, _
    lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestRectI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destRect As RECTL, _
    lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPoints _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destPoint As POINTF, _
    ByVal Count As Long, _
    lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPointsI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destPoint As POINTL, _
    ByVal Count As Long, _
    lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoint _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destPoint As POINTF, _
    srcRect As RECTF, _
    ByVal srcUnit As GpUnit, _
    ByVal lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destPoint As POINTL, _
    srcRect As RECTL, _
    ByVal srcUnit As GpUnit, _
    ByVal lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestRect _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destRect As RECTF, _
    srcRect As RECTF, _
    ByVal srcUnit As GpUnit, _
    ByVal lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestRectI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destRect As RECTL, _
    srcRect As RECTL, _
    ByVal srcUnit As GpUnit, _
    ByVal lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoints _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destPoints As POINTF, _
    ByVal Count As Long, _
    srcRect As RECTF, _
    ByVal srcUnit As GpUnit, _
    ByVal lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointsI _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destPoints As POINTL, _
    ByVal Count As Long, _
    srcRect As RECTL, _
    ByVal srcUnit As GpUnit, _
    ByVal lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoints_ _
    Lib "GDIPlus" _
    Alias "GdipEnumerateMetafileSrcRectDestPoints" _
    (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destPoints As Any, _
    ByVal Count As Long, _
    srcRect As RECTF, _
    ByVal srcUnit As GpUnit, _
    ByVal lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointsI_ _
    Lib "GDIPlus" _
    Alias "GdipEnumerateMetafileSrcRectDestPointsI" _
    (ByVal graphics As Long, _
    ByVal metafile As Long, _
    destPoints As Any, _
    ByVal Count As Long, _
    srcRect As RECTL, _
    ByVal srcUnit As GpUnit, _
    ByVal lpEnumerateMetafileProc As Long, _
    ByVal callbackData As Long, _
    ByVal imageAttributes As Long) As GpStatus

Public Declare Function GdipPlayMetafileRecord _
    Lib "GDIPlus" (ByVal metafile As Long, _
    ByVal recordType As EmfPlusRecordType, _
    ByVal flags As Long, _
    ByVal dataSize As Long, _
    byteData As Any) As GpStatus

Public Declare Function GdipGetMetafileHeaderFromWmf _
    Lib "GDIPlus" (ByVal hWmf As Long, _
    WmfPlaceableFileHdr As WmfPlaceableFileHeader, _
    header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromEmf _
    Lib "GDIPlus" (ByVal hEmf As Long, _
    header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromFile _
    Lib "GDIPlus" (ByVal filename As Long, _
    header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromStream _
    Lib "GDIPlus" (ByVal stream As Any, _
    header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromMetafile _
    Lib "GDIPlus" (ByVal metafile As Long, _
    header As MetafileHeader) As GpStatus
Public Declare Function GdipGetHemfFromMetafile _
    Lib "GDIPlus" (ByVal metafile As Long, _
    hEmf As Long) As GpStatus
Public Declare Function GdipCreateStreamOnFile Lib "GDIPlus" (ByVal filename As Long, ByVal access As Long, stream As Any) As GpStatus
Public Declare Function GdipCreateMetafileFromWmf _
    Lib "GDIPlus" (ByVal hWmf As Long, _
    ByVal bDeleteWmf As Long, _
    WmfPlaceableFileHdr As WmfPlaceableFileHeader, _
    ByVal metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromEmf _
    Lib "GDIPlus" (ByVal hEmf As Long, _
    ByVal bDeleteEmf As Long, _
    metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromFile _
    Lib "GDIPlus" (ByVal file As Long, _
    metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromWmfFile _
    Lib "GDIPlus" (ByVal file As Long, _
    WmfPlaceableFileHdr As WmfPlaceableFileHeader, _
    metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromStream _
    Lib "GDIPlus" (ByVal stream As Any, _
    metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafile _
    Lib "GDIPlus" (ByVal referenceHdc As Long, _
    etype As emfType, _
    frameRect As RECTF, _
    ByVal frameUnit As MetafileFrameUnit, _
    ByVal description As Long, _
    metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileI _
    Lib "GDIPlus" (ByVal referenceHdc As Long, _
    etype As emfType, _
    frameRect As RECTL, _
    ByVal frameUnit As MetafileFrameUnit, _
    ByVal description As Long, _
    metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileFileName _
    Lib "GDIPlus" (ByVal filename As Long, _
    ByVal referenceHdc As Long, _
    etype As emfType, _
    frameRect As RECTF, _
    ByVal frameUnit As MetafileFrameUnit, _
    ByVal description As Long, _
    metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileFileNameI _
    Lib "GDIPlus" (ByVal filename As Long, _
    ByVal referenceHdc As Long, _
    etype As emfType, _
    frameRect As RECTL, _
    ByVal frameUnit As MetafileFrameUnit, _
    ByVal description As Long, _
    metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileStream _
    Lib "GDIPlus" (ByVal stream As Any, _
    ByVal referenceHdc As Long, _
    etype As emfType, _
    frameRect As RECTF, _
    ByVal frameUnit As MetafileFrameUnit, _
    ByVal description As Long, _
    metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileStreamI _
    Lib "GDIPlus" (ByVal stream As Any, _
    ByVal referenceHdc As Long, _
    etype As emfType, _
    frameRect As RECTL, _
    ByVal frameUnit As MetafileFrameUnit, _
    ByVal description As Long, _
    metafile As Long) As GpStatus
Public Declare Function GdipSetMetafileDownLevelRasterizationLimit _
    Lib "GDIPlus" (ByVal metafile As Long, _
    ByVal metafileRasterizationLimitDpi As Long) As GpStatus
Public Declare Function GdipGetMetafileDownLevelRasterizationLimit _
    Lib "GDIPlus" (ByVal metafile As Long, _
    metafileRasterizationLimitDpi As Long) As GpStatus
Public Declare Function GdipGetImageDecodersSize _
    Lib "GDIPlus" (numDecoders As Long, _
    size As Long) As GpStatus
Public Declare Function GdipSetImageAttributesCachedBackground _
    Lib "GDIPlus" (ByVal imageattr As Long, _
    ByVal enableFlag As Long) As GpStatus
Public Declare Function GdipTestControl _
    Lib "GDIPlus" (ByVal control As GpTestControlEnum, _
    param As Any) As GpStatus
#If GdipVersion >= 1.1 Then

Public Declare Function GdipConvertToEmfPlus _
    Lib "GDIPlus" (ByVal refGraphics As Long, _
    conversionFailureFlag As Long, _
    ByVal emfType As emfType, _
    ByVal description As Long, _
    ByVal out_metafile As Long) As GpStatus
Public Declare Function GdipConvertToEmfPlusToFile _
    Lib "GDIPlus" (ByVal refGraphics As Long, _
    ByVal metafile As Long, _
    conversionFailureFlag As Long, _
    ByVal filename As Long, _
    ByVal emfType As emfType, _
    ByVal description As Long, _
    out_metafile As Long) As GpStatus
Public Declare Function GdipConvertToEmfPlusToStream _
    Lib "GDIPlus" (ByVal refGraphics As Long, _
    ByVal metafile As Long, _
    conversionFailureFlag As Long, _
    stream As Any, _
    ByVal emfType As emfType, _
    ByVal description As Long, _
    out_metafile As Long) As GpStatus
#End If

Public Declare Function GdipFlush _
    Lib "GDIPlus" (ByVal graphics As Long, _
    ByVal intention As FlushIntention) As GpStatus
Public Declare Function GdipAlloc Lib "GDIPlus" (ByVal size As Long) As Long
Public Declare Sub GdipFree Lib "GDIPlus" (ByVal Ptr As Long)

'===================================================================================
'  公共部分 / 其他部分
'===================================================================================

Public Declare Function GdiplusStartup _
    Lib "GDIPlus" (token As Long, _
    inputbuf As GdiplusStartupInput, _
    Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As GpStatus

Public Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Public Enum GpStatus
    Ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
#If GdipVersion >= 1.1 Then
    ProfileNotFound = 21
#End If
End Enum

Private Declare Function CLSIDFromString _
    Lib "ole32.dll" (ByVal lpszProgID As Long, _
    pCLSID As Clsid) As Long

Public Enum GdipImageType
    BMP
    EMF
    WMF
    JPG
    PNG
    GIF
    TIF
    ICO
End Enum

Public Const ImageEnCoderSuffix       As String = "-1A04-11D3-9A73-0000F81EF32E}"
Public Const ImageEnCoderBMP          As String = "{557CF400" & ImageEnCoderSuffix
Public Const ImageEnCoderJPG          As String = "{557CF401" & ImageEnCoderSuffix
Public Const ImageEnCoderGIF          As String = "{557CF402" & ImageEnCoderSuffix
Public Const ImageEnCoderEMF          As String = "{557CF403" & ImageEnCoderSuffix
Public Const ImageEnCoderWMF          As String = "{557CF404" & ImageEnCoderSuffix
Public Const ImageEnCoderTIF          As String = "{557CF405" & ImageEnCoderSuffix
Public Const ImageEnCoderPNG          As String = "{557CF406" & ImageEnCoderSuffix
Public Const ImageEnCoderICO          As String = "{557CF407" & ImageEnCoderSuffix
Public Const EnCoderCompression       As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Public Const EnCoderColorDepth        As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Public Const EnCoderScanMethod        As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Public Const EnCoderVersion           As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Public Const EnCoderRenderMethod      As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Public Const EnCoderQuality           As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Public Const EnCoderTransformation    As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Public Const EnCoderLuminanceTable    As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Public Const EnCoderChrominanceTable  As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Public Const EnCoderSaveFlag          As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
#If GdipVersion >= 1.1 Then
Public Const EnCoderColorSpace        As String = "{AE7A62A0-EE2C-49D8-9D07-1BA8A927596E}"
Public Const EnCoderImageItems        As String = "{63875E13-1F1D-45AB-9195-A29B6066A650}"
Public Const EnCoderSaveAsCMYK        As String = "{A219BBC9-0A9D-4005-A3EE-3A421B8BB06C}"
#End If

Public ZeroPointF As POINTF, ZeroPointL As POINTL
Public ZeroRectF As RECTF, ZeroRectL As RECTL

Dim mToken As Long

Dim Pens() As Long, PenCount As Long
Dim Brushes() As Long, BrushCount As Long
Dim StrFormats() As Long, StrFormatCount As Long
Dim Matrixes() As Long, MatrixCount As Long

Public Function DeleteObjects()
    Dim i As Long
    
    For i = 1 To PenCount: GdipDeletePen Pens(i): Next
        For i = 1 To BrushCount: GdipDeleteBrush Brushes(i): Next
            For i = 1 To StrFormatCount: GdipDeleteStringFormat StrFormats(i): Next
                For i = 1 To MatrixCount: GdipDeleteMatrix Matrixes(i): Next
                    PenCount = 0
                    BrushCount = 0
                    StrFormatCount = 0
                    MatrixCount = 0
End Function

Public Function NewPen(ByVal Color As Long, ByVal Width As Single) As Long
    PenCount = PenCount + 1
    ReDim Preserve Pens(PenCount)
    
    GdipCreatePen1 Color, Width, UnitPixel, Pens(PenCount)
    NewPen = Pens(PenCount)
End Function

Public Function NewBrush(ByVal Color As Long) As Long
    BrushCount = BrushCount + 1
    ReDim Preserve Brushes(BrushCount)
    
    GdipCreateSolidFill Color, Brushes(BrushCount)
    NewBrush = Brushes(BrushCount)
End Function

Public Function NewStringFormat(ByVal Align As StringAlignment) As Long
    StrFormatCount = StrFormatCount + 1
    ReDim Preserve StrFormats(StrFormatCount)
    
    GdipCreateStringFormat 0, 0, StrFormats(StrFormatCount)
    GdipSetStringFormatAlign StrFormats(StrFormatCount), Align
    NewStringFormat = StrFormats(StrFormatCount)
End Function

Public Function NewMatrix(ByVal m11 As Single, _
    ByVal m12 As Single, _
    ByVal m21 As Single, _
    ByVal m22 As Single, _
    ByVal dx As Single, _
    ByVal dy As Single) As Long
    
    MatrixCount = MatrixCount + 1
    ReDim Preserve Matrixes(MatrixCount)
    
    GdipCreateMatrix Matrixes(MatrixCount)
    GdipSetMatrixElements Matrixes(MatrixCount), m11, m12, m21, m22, dx, dy
    NewMatrix = Matrixes(MatrixCount)
End Function

Public Function NewRectF(ByVal Left As Single, _
    ByVal Top As Single, _
    ByVal Width As Single, _
    ByVal Height As Single) As RECTF
    
    With NewRectF
        .Left = Left
        .Top = Top
        .Right = Width
        .Bottom = Height
    End With
End Function

Public Function NewRectL(ByVal Left As Single, _
    ByVal Top As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As RECTL
    
    With NewRectL
        .Left = Left
        .Top = Top
        .Right = Width
        .Bottom = Height
    End With
End Function

Public Function NewPointF(ByVal X As Single, _
    ByVal Y As Single) As POINTF
    
    NewPointF.X = X
    NewPointF.Y = Y
End Function

Public Function NewPointL(ByVal X As Single, _
    ByVal Y As Single) As POINTL
    
    NewPointL.X = X
    NewPointL.Y = Y
End Function

Public Function NewPointsFPtr(ParamArray ptXY()) As Long
    If (UBound(ptXY) And 1) = 0 Then GoTo ErrHandle
    
    Dim ret() As POINTF, i As Long
    ReDim ret(0 To UBound(ptXY) \ 2)
    
    For i = 0 To UBound(ptXY) Step 2
        ret(i \ 2).X = ptXY(i)
        ret(i \ 2).Y = ptXY(i + 1)
    Next
    
    NewPointsFPtr = VarPtr(ret(0))
    
    Exit Function
ErrHandle:
    NewPointsFPtr = 0
End Function

Public Function NewPointsLPtr(ParamArray ptXY()) As Long
    If (UBound(ptXY) And 1) = 0 Then GoTo ErrHandle
    
    Dim ret() As POINTL, i As Long
    ReDim ret(0 To UBound(ptXY) \ 2)
    
    For i = 0 To UBound(ptXY) Step 2
        ret(i \ 2).X = ptXY(i)
        ret(i \ 2).Y = ptXY(i + 1)
    Next
    
    NewPointsLPtr = VarPtr(ret(0))
    
    Exit Function
ErrHandle:
    NewPointsLPtr = 0
End Function

Public Function NewColors(ParamArray colors()) As Long()
    Dim ret() As Long, i As Long
    
    ReDim ret(UBound(colors))
    For i = 0 To UBound(colors)
        ret(i) = colors(i)
    Next
    
    NewColors = ret
End Function

Public Function InitGDIPlus(Optional OnErrorMsgbox, _
    Optional ByVal OnErrorEnd As Boolean = True) As GpStatus
    
    If mToken <> 0 Then
        Debug.Print "InitGDIPlus> GdiPlus已被初始化"
        Exit Function
    End If
    
    Dim uInput As GdiplusStartupInput
    Dim ret    As GpStatus
    
    uInput.GdiplusVersion = 1
    ret = GdiplusStartup(mToken, uInput)
    
    If ret <> Ok Then
        If Not IsMissing(OnErrorMsgbox) Then MsgBox OnErrorMsgbox
        'If OnErrorEnd Then End
    End If
    
    InitGDIPlus = ret
End Function

Public Sub TerminateGDIPlus()
    If mToken = 0 Then
        Debug.Print "TerminateGDIPlus> GdiPlus已被结束"
        Exit Sub
    End If
    
    DeleteObjects
    GdiplusShutdown mToken
    
    mToken = 0
End Sub

Public Function InitGDIPlusTo(ByRef token As Long, _
    Optional OnErrorMsgbox, _
    Optional ByVal OnErrorEnd As Boolean = True) As GpStatus
    
    If token <> 0 Then
        Debug.Print "InitGDIPlusTo> GdiPlus已被初始化"
        Exit Function
    End If
    
    Dim uInput As GdiplusStartupInput
    Dim ret As GpStatus
    
    uInput.GdiplusVersion = 1
    ret = GdiplusStartup(token, uInput)
    
    If ret <> Ok Then
        If Not IsMissing(OnErrorMsgbox) Then MsgBox OnErrorMsgbox
        'If OnErrorEnd Then End
    End If
    
    InitGDIPlusTo = ret
End Function

Public Sub TerminateGDIPlusFrom(ByVal token As Long)
    If token = 0 Then
        Debug.Print "TerminateGDIPlusFrom> GdiPlus已被结束"
        Exit Sub
    End If
    
    DeleteObjects
    GdiplusShutdown token
    
    token = 0
End Sub

#If GdipVersion >= 1.1 Then

Public Sub GdipCreateEffect2(ByVal EffectType As GdipEffectType, Effect As Long)
    Select Case EffectType
        Case GdipEffectType.Blur:                   GdipCreateEffect &H633C80A4, &H482B1843, &H28BEF29E, &HD4FDC534, Effect 'CLSIDFromString StrPtr(BlurEffectGuid), GetEffectClsid
        Case GdipEffectType.BrightnessContrast:     GdipCreateEffect &HD3A1DBE1, &H4C178EC4, &H97EA4C9F, &H3D341CAD, Effect 'CLSIDFromString StrPtr(BrightnessContrastEffectGuid), GetEffectClsid
        Case GdipEffectType.ColorBalance:           GdipCreateEffect &H537E597D, &H48DA251E, &HCA296496, &HF8706B49, Effect 'CLSIDFromString StrPtr(ColorBalanceEffectGuid), GetEffectClsid
        Case GdipEffectType.ColorCurve:             GdipCreateEffect &HDD6A0022, &H4A6758E4, &H8ED49B9D, &H3DA581B8, Effect 'CLSIDFromString StrPtr(ColorCurveEffectGuid), GetEffectClsid
        Case GdipEffectType.ColorLookupTable:       GdipCreateEffect &HA7CE72A9, &H40D70F7F, &HC0D0CCB3, &H12325C2D, Effect 'CLSIDFromString StrPtr(ColorLookupTableEffectGuid), GetEffectClsid
        Case GdipEffectType.ColorMatrix:            GdipCreateEffect &H718F2615, &H40E37933, &H685F11A5, &H74DD14FE, Effect 'CLSIDFromString StrPtr(ColorMatrixEffectGuid), GetEffectClsid
        Case GdipEffectType.HueSaturationLightness: GdipCreateEffect &H8B2DD6C3, &H4D87EB07, &H871F0A5, &H5F9C6AE2, Effect 'CLSIDFromString StrPtr(HueSaturationLightnessEffectGuid), GetEffectClsid
        Case GdipEffectType.Levels:                 GdipCreateEffect &H99C354EC, &H4F3A2A31, &HA817348C, &H253AB303, Effect 'CLSIDFromString StrPtr(LevelsEffectGuid), GetEffectClsid
        Case GdipEffectType.RedEyeCorrection:       GdipCreateEffect &H74D29D05, &H426669A4, &HC53C4995, &H32B63628, Effect 'CLSIDFromString StrPtr(RedEyeCorrectionEffectGuid), GetEffectClsid
        Case GdipEffectType.Sharpen:                GdipCreateEffect &H63CBF3EE, &H402CC526, &HC562718F, &H4251BF40, Effect 'CLSIDFromString StrPtr(SharpenEffectGuid), GetEffectClsid
        Case GdipEffectType.Tint:                   GdipCreateEffect &H1077AF00, &H44412848, &HAD448994, &H2C7A2D4C, Effect 'CLSIDFromString StrPtr(TintEffectGuid), GetEffectClsid
    End Select
End Sub

Public Function GetAddress(ByVal lngAddr As Long) As Long
    GetAddress = lngAddr
End Function

#End If

Public Function GetImageEnCoderClsid(ByVal ImageType As GdipImageType) As Clsid
    Select Case ImageType
    Case GdipImageType.PNG: CLSIDFromString StrPtr(ImageEnCoderPNG), GetImageEnCoderClsid
    Case GdipImageType.JPG: CLSIDFromString StrPtr(ImageEnCoderJPG), GetImageEnCoderClsid
    Case GdipImageType.GIF: CLSIDFromString StrPtr(ImageEnCoderGIF), GetImageEnCoderClsid
    Case GdipImageType.BMP: CLSIDFromString StrPtr(ImageEnCoderBMP), GetImageEnCoderClsid
    Case GdipImageType.ICO: CLSIDFromString StrPtr(ImageEnCoderICO), GetImageEnCoderClsid
    Case GdipImageType.EMF: CLSIDFromString StrPtr(ImageEnCoderEMF), GetImageEnCoderClsid
    Case GdipImageType.WMF: CLSIDFromString StrPtr(ImageEnCoderWMF), GetImageEnCoderClsid
    Case GdipImageType.TIF: CLSIDFromString StrPtr(ImageEnCoderTIF), GetImageEnCoderClsid
    End Select
End Function

Public Function SaveImageToPNG(ByVal Image As Long, ByVal Path As String) As GpStatus
    SaveImageToPNG = GdipSaveImageToFile(Image, StrPtr(Path), GetImageEnCoderClsid( _
    PNG), ByVal 0)
End Function

Public Function SaveImageToJPG(ByVal Image As Long, _
    ByVal Path As String, _
    ByVal Quality As Long) As GpStatus
    
    Dim Params As EnCoderParameters
    
    Params.Count = 1
    CLSIDFromString StrPtr(EnCoderQuality), Params.Parameter.guid
    Params.Parameter.NumberOfValues = 1
    Params.Parameter.type = 4
    Params.Parameter.Value = VarPtr(Quality)
    
    SaveImageToJPG = GdipSaveImageToFile(Image, StrPtr(Path), GetImageEnCoderClsid( _
    JPG), Params)
End Function

Public Function SaveImageToGIF(ByVal Image As Long, ByVal Path As String) As GpStatus
    SaveImageToGIF = GdipSaveImageToFile(Image, StrPtr(Path), GetImageEnCoderClsid( _
    GIF), ByVal 0)
End Function

Public Function SaveImageToBMP(ByVal Image As Long, ByVal Path As String) As GpStatus
    SaveImageToBMP = GdipSaveImageToFile(Image, StrPtr(Path), GetImageEnCoderClsid( _
    BMP), ByVal 0)
End Function

Public Function CreateBitmap(ByRef Bitmap As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    Optional ByVal PixelFormat As GpPixelFormat = PixelFormat32bppARGB) As GpStatus
    
    GdipCreateBitmapFromScan0 Width, Height, 0, PixelFormat, ByVal 0, Bitmap
End Function

Public Function CreateBitmapWithGraphics(ByRef Bitmap As Long, _
    ByRef graphics As Long, _
    ByVal Width As Long, _
    ByVal Height As Long, _
    Optional ByVal PixelFormat As GpPixelFormat = PixelFormat32bppARGB) As GpStatus
    
    GdipCreateBitmapFromScan0 Width, Height, 0, PixelFormat, ByVal 0, Bitmap
    GdipGetImageGraphicsContext Bitmap, graphics
End Function







