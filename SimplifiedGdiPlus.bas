Attribute VB_Name = "SimplifiedGdiPlus"
'SimplifiedGdiPlus
'根据modGdip模块（原作者：vIstaSwx）制作
'部分注释来自于微软msdn手册（learn.microsoft.com），可能有含义偏差
'注释中包含“[?]”为意义暂不明确
'By 马云爱逛京东

'1.0.0 (2024-07-31)
'创建了此模块。

'1.0.1（2024-08-28）
'修复：DrawPath过程实现抗锯齿；
'改动：所有过程/函数的ReleaseHandles参数缺省为False；
'改动：将CreateBasicShapePath更名为NewBasicShapePath；
'新增：NewPolygonPath。

Option Explicit

'##################  枚  举  ##################
'像素格式
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

'图像度量单位
Public Enum GpUnit
    UnitWorld                                                                   '世界
    UnitDisplay                                                                 '显示器
    UnitPixel                                                                   '像素
    UnitPoint                                                                   '点
    UnitInch                                                                    '英寸
    UnitDocument                                                                '文档
    UnitMillimeter                                                              '毫米
End Enum

'路径点类型
Public Enum PathPointType
    PathPointTypeStart = 0                                                      '起始点
    PathPointTypeLine = 1                                                       '直线
    PathPointTypeBezier = 3                                                     '贝塞尔曲线
    PathPointTypePathTypeMask = &H7                                             '路径类型遮罩
    PathPointTypePathDashMode = &H10                                            '路径虚线模式[?]
    PathPointTypePathMarker = &H20                                              '路径标记
    PathPointTypeCloseSubpath = &H80                                            '关闭子路径
    PathPointTypeBezier3 = PathPointTypeBezier
End Enum

'通用字体族（不常用）
Public Enum GenericFontFamily
    GenericFontFamilySerif
    GenericFontFamilySansSerif
    GenericFontFamilyMonospace
End Enum

'字型（字体样式），可以组合使用，最大不超过15
Public Enum FontStyle
    FontStyleRegular = 0                                                        '常规
    FontStyleBold = 1                                                           '粗体
    FontStyleItalic = 2                                                         '斜体
    FontStyleBoldItalic = FontStyleBold + FontStyleItalic
    FontStyleUnderline = 4                                                      '下划线
    FontStyleStrikeout = 8                                                      '删除线
End Enum

'对齐
Public Enum StringAlignment
    StringAlignmentNear = 0                                                     '靠近
    StringAlignmentCenter = 1                                                   '居中
    StringAlignmentFar = 2                                                      '远离
End Enum

'填充模式
Public Enum FillMode
    FillModeAlternate                                                           '交替填充（缺省）
    FillModeWinding                                                             '围绕填充
    'FillModeAlternate：从封闭区域中的一个点向无穷远处水平画一条射线，当该射线穿越奇数条边框线时，填充封闭区域；
    '当该射线穿越偶数条边框线时，不填充封闭区域
    'FillModeWinding：从封闭区域中的一个点向无穷远处水平画一条射线，当该射线穿越奇数条边框线时，填充封闭区域；当
    '该射线穿越偶数条边框线时，则要根据边框线的方向来判断：如果穿过的边框线在不同方向的边框线数目相等，则不填充
    '封闭区域；如不相等，则填充封闭区域
End Enum

'平铺模式（注意不是扭曲模式WarpMode）
Public Enum WrapMode
    WrapModeTile                                                                '不翻转的平铺
    WrapModeTileFlipX                                                           '在一行中从一个磁贴移动到下一个磁贴时水平翻转磁贴
    WrapModeTileFlipY                                                           '在列中从一个磁贴移动到下一个磁贴时垂直翻转磁贴
    WrapModeTileFlipXY                                                          '在沿行移动时水平翻转磁贴，在沿列移动时垂直翻转磁贴
    WrapModeClamp                                                               '不进行平铺
End Enum

'线性渐变模式
Public Enum LinearGradientMode
    LinearGradientModeHorizontal                                                '水平渐变 ——
    LinearGradientModeVertical                                                  '垂直渐变 ｜
    LinearGradientModeForwardDiagonal                                           '左上到右下渐变 ／
    LinearGradientModeBackwardDiagonal                                          '右上到左下渐变 ＼
End Enum

'质量模式
Public Enum QualityMode
    QualityModeInvalid = -1                                                     '无效
    QualityModeDefault = 0                                                      '默认
    QualityModeLow = 1                                                          '低质量
    QualityModeHigh = 2                                                         '高质量
End Enum

'颜色合成模式
Public Enum CompositingMode
    CompositingModeSourceOver                                                   '混合模式：呈现颜色将与背景色混合，混合比例由呈现颜色的α分量决定
    CompositingModeSourceCopy                                                   '覆盖模式：呈现颜色将直接覆盖背景色
End Enum

'颜色合成质量
Public Enum CompositingQuality
    CompositingQualityInvalid = QualityModeInvalid
    CompositingQualityDefault = QualityModeDefault
    CompositingQualityHighSpeed = QualityModeLow
    CompositingQualityHighQuality = QualityModeHigh
    CompositingQualityGammaCorrected                                            '使用γ校正
    CompositingQualityAssumeLinear                                              '假定模型为线性
End Enum

'平滑模式
Public Enum SmoothingMode
    SmoothingModeInvalid = QualityModeInvalid
    SmoothingModeDefault = QualityModeDefault
    SmoothingModeHighSpeed = QualityModeLow
    SmoothingModeHighQuality = QualityModeHigh
    SmoothingModeNone                                                           '不使用
    SmoothingModeAntiAlias                                                      '抗锯齿（使用 8 × 4 框筛选器）
End Enum

'图像插值模式
Public Enum InterpolationMode
    InterpolationModeInvalid = QualityModeInvalid
    InterpolationModeDefault = QualityModeDefault
    InterpolationModeLowQuality = QualityModeLow
    InterpolationModeHighQuality = QualityModeHigh
    InterpolationModeBilinear                                                   '双线性插值
    InterpolationModeBicubic                                                    '双三次插值
    InterpolationModeNearestNeighbor                                            '最临近插值
    InterpolationModeHighQualityBilinear                                        '高质量的双线性插值
    InterpolationModeHighQualityBicubic                                         '高质量的双三次插值
End Enum

'像素偏移模式
Public Enum PixelOffsetMode
    PixelOffsetModeInvalid = QualityModeInvalid
    PixelOffsetModeDefault = QualityModeDefault
    PixelOffsetModeHighSpeed = QualityModeLow
    PixelOffsetModeHighQuality = QualityModeHigh
    PixelOffsetModeNone                                                         '像素中心具有整数坐标（不采用偏移）
    PixelOffsetModeHalf                                                         '像素中心的坐标介于整数值之间（半数偏移）
    '假定图像左上角的像素为（0,0）：
    'PixelOffsetModeNone：像素覆盖 x 和 y 方向 –0.5 至 0.5 之间的区域，即像素中心位于（0，0）。
    'PixelOffsetModeHalf：像素覆盖 x 和 y 方向 0 至 1 之间的区域，即像素中心位于（0.5，0.5）。
End Enum

'文本渲染提示
Public Enum TextRenderingHint
    TextRenderingHintSystemDefault = 0                                          '使用当前所选系统字体平滑模式绘制字符
    TextRenderingHintSingleBitPerPixelGridFit                                   '使用字符字形位图和提示绘制字符
    TextRenderingHintSingleBitPerPixel                                          '使用字符字形位图绘制字符，但不显示提示
    TextRenderingHintAntiAliasGridFit                                           '使用字符抗锯齿字形位图和提示绘制字符
    TextRenderingHintAntiAlias                                                  '使用字符抗锯齿字形位图绘制字符，但不显示提示
    TextRenderingHintClearTypeGridFit                                           '使用字符字形 ClearType 位图和提示绘制字符
End Enum

'颜色矩阵顺序
Public Enum MatrixOrder
    MatrixOrderPrepend = 0                                                      '新矩阵位于现有矩阵位左侧
    MatrixOrderAppend = 1                                                       '新矩阵位于现有矩阵位右侧
End Enum

'颜色调整类型
Public Enum ColorAdjustType
    ColorAdjustTypeDefault                                                      '默认
    ColorAdjustTypeBitmap                                                       '位图
    ColorAdjustTypeBrush                                                        '画刷
    ColorAdjustTypePen                                                          '画笔
    ColorAdjustTypeText                                                         '文本
    ColorAdjustTypeCount                                                        '数量
    ColorAdjustTypeAny                                                          '（预留）
End Enum

'颜色矩阵标志
Public Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = 0                                                 '的所有颜色值都由同一颜色调整矩阵调整
    ColorMatrixFlagsSkipGrays = 1                                               '调整颜色，但不调整灰色底纹*
    ColorMatrixFlagsAltGray = 2                                                 '颜色由一个矩阵调整，灰色底纹由另一个矩阵调整
    '注：灰色底纹是指其红色、绿色和蓝色分量的值都相同的任何颜色。
End Enum

'扭曲模式（注意不是平铺模式WrapMode）
Public Enum WarpMode
    WarpModePerspective                                                         '透视扭曲
    WarpModeBilinear                                                            '双线性扭曲
End Enum

'合并模式
Public Enum CombineMode
    CombineModeReplace                                                          '现有区域替换为新区域
    CombineModeIntersect                                                        '现有区域替换为该区域和新区域的交集
    CombineModeUnion                                                            '现有区域替换为该区域和新区域的并集
    CombineModeXor                                                              '现有区域替换为该区域和新区域的异或
    CombineModeExclude                                                          '现有区域替换为位于新区域之外的该区域部分
    CombineModeComplement                                                       '现有区域替换为该区域外部的新区域部分
End Enum

'图像锁定模式
Public Enum ImageLockMode
    ImageLockModeRead = &H1                                                     '锁定图像的一部分以便读取
    ImageLockModeWrite = &H2                                                    '锁定图像的一部分以便写入
    ImageLockModeUserInputBuf = &H4                                             '由用户读取或写入图像所使用的缓冲区
End Enum

'形状类型
Public Enum ShapeType
    ShapeTypeRectangle = 0                                                      '矩形
    ShapeTypeEllipse = 1                                                        '椭圆（包含圆形）
    ShapeTypeRoundedRectangle = 2                                               '圆角矩形
End Enum

'形状绘制模式
Public Enum ShapeDrawingMode
    ShapeDrawingModeEdge = 0                                                    '描边
    ShapeDrawingModeFill = 1                                                    '填充
    ShapeDrawingModeEdgeAndFill = 2                                             '描边和填充
End Enum

'图像保存格式
Public Enum GpImageSaveFormat
    GpSaveBMP = 0
    GpSaveJPEG = 1
    GpSaveGIF = 2
    GpSavePNG = 3
    GpSaveTIFF = 4
End Enum

'图像格式标识符
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

'图像类型
Public Enum ImageType
    ImageTypeUnknown = 0                                                        '未知
    ImageTypeBitmap = 1                                                         '位图
    ImageTypeMetafile = 2                                                       '图元文件
End Enum

'图像属性类型（不常用）
Public Enum ImagePropertyType
    ImagePropertyTypeByte = 1
    ImagePropertyTypeASCII = 2
    ImagePropertyTypeShort = 3
    ImagePropertyTypeLong = 4
    ImagePropertyTypeRational = 5                                               '有理数的[?]
    ImagePropertyTypeUndefined = 7                                              '未定义的
    ImagePropertyTypeSLONG = 9
    ImagePropertyTypeSRational = 10
End Enum

'图像编/解码器标志
Public Enum ImageCodecFlags
    ImageCodecFlagsEncoder = &H1                                                '编码器
    ImageCodecFlagsDecoder = &H2                                                '解码器
    ImageCodecFlagsSupportBitmap = &H4                                          '支持位图
    ImageCodecFlagsSupportVector = &H8                                          '支持向量[?]
    ImageCodecFlagsSeekableEncode = &H10                                        '可定位编码器
    ImageCodecFlagsBlockingDecode = &H20                                        '锁定的解码器[?]
    ImageCodecFlagsBuiltin = &H10000                                            '内建式[?]
    ImageCodecFlagsSystem = &H20000                                             '由系统
    ImageCodecFlagsUser = &H40000                                               '由用户
End Enum

'调色板标志
Public Enum PaletteFlags
    PaletteFlagsHasAlpha = &H1                                                  '具备α分量（透明度）
    PaletteFlagsGrayScale = &H2                                                 '仅包含灰度
    PaletteFlagsHalftone = &H4                                                  'Windows半色调调色板
End Enum

'旋转/翻转类型
Public Enum RotateFlipType
    RotateNoneFlipNone = 0                                                      '无
    Rotate90FlipNone = 1                                                        '旋转90°
    Rotate180FlipNone = 2                                                       '旋转180°
    Rotate270FlipNone = 3                                                       '旋转270°
    RotateNoneFlipX = 4                                                         '水平翻转
    Rotate90FlipX = 5                                                           '先旋转90°，然后再水平翻转
    Rotate180FlipX = 6                                                          '先旋转180°，然后再水平翻转
    Rotate270FlipX = 7                                                          '先旋转270°，然后再水平翻转
    RotateNoneFlipY = Rotate180FlipX
    Rotate90FlipY = Rotate270FlipX
    Rotate180FlipY = RotateNoneFlipX
    Rotate270FlipY = Rotate90FlipX
    RotateNoneFlipXY = Rotate180FlipNone
    Rotate90FlipXY = Rotate270FlipNone
    Rotate180FlipXY = RotateNoneFlipNone
    Rotate270FlipXY = Rotate90FlipNone
End Enum

'颜色深度模式
Public Enum ColorMode
    ColorModeARGB32 = 0
    ColorModeARGB64 = 1
End Enum

'CMYK模式通道标志（不常用）
Public Enum ColorChannelFlags
    ColorChannelFlagsC = 0                                                      '青色
    ColorChannelFlagsM                                                          '洋红色
    ColorChannelFlagsY                                                          '黄色
    ColorChannelFlagsK                                                          '黑色
    ColorChannelFlagsLast                                                       '上一个[?]
End Enum

'ARGB分量
Public Enum ColorShiftComponents
    AlphaShift = 24                                                             'α分量
    RedShift = 16                                                               '红分量
    GreenShift = 8                                                              '绿分量
    BlueShift = 0                                                               '蓝分量
End Enum

'ARGB颜色掩码
Public Enum ColorMaskComponents
    AlphaMask = &HFF000000
    RedMask = &HFF0000
    GreenMask = &HFF00
    BlueMask = &HFF
End Enum

'字重（值越大，字看起来越粗）
Public Enum FontWeight
    FW_DONTCARE = 0&                                                            '使用默认的字重
    FW_THIN = 100&                                                              '最细（-3）
    FW_EXTRALIGHT = 200&                                                        '特别细（-2）
    FW_ULTRALIGHT = FW_EXTRALIGHT                                               '特别细（-2）
    FW_LIGHT = 300&                                                             '略细（-1）
    FW_NORMAL = 400&                                                            '正常粗细（0）
    FW_REGULAR = FW_NORMAL                                                      '正常粗细（0）
    FW_MEDIUM = 500&                                                            '略粗（+1）
    FW_SEMIBOLD = 600&                                                          '中等粗（+2）
    FW_DEMIBOLD = FW_SEMIBOLD                                                   '
    FW_BOLD = 700&                                                              '粗（+3）
    FW_EXTRABOLD = 800&                                                         '特别粗（+4）
    FW_ULTRABOLD = FW_EXTRABOLD                                                 '
    FW_HEAVY = 900&                                                             '最粗（+5）
    FW_BLACK = FW_HEAVY                                                         '
End Enum

'字符集类型
Public Enum CharSetType
    ANSI_CHARSET = 0
    DEFAULT_CHARSET = 1                                                         '根据当前系统区域设置（默认）
    SYMBOL_CHARSET = 2
    SHIFTJIS_CHARSET = 128
    HANGEUL_CHARSET = 129
    HANGUL_CHARSET = 129
    GB2312_CHARSET = 134                                                        '简体中文（国标2312）
    CHINESEBIG5_CHARSET = 136                                                   '繁體中文（大五碼）
    GREEK_CHARSET = 161
    TURKISH_CHARSET = 162
    HEBREW_CHARSET = 177
    ARABIC_CHARSET = 178
    BALTIC_CHARSET = 186
    RUSSIAN_CHARSET = 204
    THAI_CHARSET = 222
    EASTEUROPE_CHARSET = 238
    OEM_CHARSET = 255                                                           '依赖于操作系统的字符集
    JOHAB_CHARSET = 130
    VIETNAMESE_CHARSET = 163
    MAC_CHARSET = 77
End Enum

'字体输出精度类型
Public Enum OutPrecisionType
    OUT_DEFAULT_PRECIS = 0                                                      '默认字体映射器行为
    OUT_STRING_PRECIS = 1                                                       '字体映射器不使用此值，但在枚举光栅字体时会返回此值
    OUT_CHARACTER_PRECIS = 2                                                    '未使用
    OUT_STROKE_PRECIS = 3                                                       '字体映射器不使用此值，但在枚举TrueType、其他基于轮廓的字体和矢量字体时返回此值
    OUT_TT_PRECIS = 4                                                           '当系统包含多个同名字体时，指示字体映射器选择 TrueType 字体
    OUT_DEVICE_PRECIS = 5                                                       '当系统包含多个同名字体时，指示字体映射器选择设备字体
    OUT_RASTER_PRECIS = 6                                                       '当系统包含多个同名字体时，指示字体映射器选择光栅字体
    OUT_TT_ONLY_PRECIS = 7                                                      '指示字体映射器仅从TrueType字体中进行选择（如果系统中没有安装TrueType字体，字体映射器将返回到默认行为）
    OUT_OUTLINE_PRECIS = 8                                                      '指示字体映射器从TrueType和其他基于大纲的字体中进行选择
End Enum

'字体剪裁精度类型
Public Enum ClipPrecisionType
    CLIP_DEFAULT_PRECIS = 0                                                     '指定默认剪辑行为
    CLIP_CHARACTER_PRECIS = 1                                                   '未使用
    CLIP_STROKE_PRECIS = 2                                                      '字体映射器不使用，但在枚举光栅、矢量或TrueType字体时返回*
    CLIP_MASK = 15                                                              '未使用
    CLIP_LH_ANGLES = 16                                                         '所有字体的旋转取决于坐标系的方向是左手还是右手*
    CLIP_TT_ALWAYS = 32                                                         '未使用
    CLIP_EMBEDDED = 128                                                         '必须指定此标志才能使用嵌入的只读字体
    '注1：CLIP_STROKE_PRECIS - 为了兼容，枚举字体时始终返回此值。
    '注2：CLIP_LH_ANGLES - 如果未使用，设备字体始终逆时针旋转，但其他字体的旋转取决于坐标系的方向。
End Enum

'字体质量类型
Public Enum FontQualityType
    DEFAULT_QUALITY = 0                                                         '使用默认的字体质量呈现
    DRAFT_QUALITY = 1                                                           '草稿质量（逻辑字体属性的精确匹配优先于字体质量）
    PROOF_QUALITY = 2                                                           '样张质量（字体质量优先于逻辑字体属性的精确匹配）
    NONANTIALIASED_QUALITY = 3                                                  '字体始终为非抗锯齿*
    ANTIALIASED_QUALITY = 4                                                     '如果字体支持该字体，并且字体大小不是太小或太大，则字体始终为抗锯齿
    '注：如果ANTIALIASED_QUALITY和NONANTIALIASED_QUALITY均未选中，则仅当用户在控制面板中选择平滑屏幕字体时，字体才会抗锯齿
End Enum

'字符串格式标志
Public Enum StringFormatFlags
    StringFormatFlagsNoUse = &H0                                                '不使用
    StringFormatFlagsDirectionRightToLeft = &H1                                 '从右到左的顺序
    StringFormatFlagsDirectionVertical = &H2                                    '垂直绘制单个文本行
    StringFormatFlagsNoFitBlackBox = &H4                                        '允许部分字符悬停在字符串的布局矩形上
    StringFormatFlagsDisplayFormatControl = &H20                                '使用代表性字符显示Unicode格式控制字符
    StringFormatFlagsNoFontFallback = &H400                                     '替换字符串中无效的字符（默认的缺失字符为“囟”字去掉上面一撇）
    StringFormatFlagsMeasureTrailingSpaces = &H800                              '行末空格包含在字符串度量中
    StringFormatFlagsNoWrap = &H1000                                            '禁用文本换行
    StringFormatFlagsLineLimit = &H2000                                         '在布局矩形中限制布局整行
    StringFormatFlagsNoClip = &H4000                                            '允许显示悬在布局矩形上方的字符和布局矩形外延伸的文本[?]
    StringFormatFlagsBypassGDI = &H80000000                                     '绕过GDI绘制[?]
End Enum

'字符串裁剪
Public Enum StringTrimming
    StringTrimmingNone = 0                                                      '不裁剪
    StringTrimmingCharacter = 1                                                 '在布局矩形内最后一个字符的边界处断开字符串（默认）
    StringTrimmingWord = 2                                                      '在布局矩形内最后一个单词的边界处断开字符串
    StringTrimmingEllipsisCharacter = 3                                         '在布局矩形内最后一个字符的边界处断开字符串，并在字符后面插入“...”
    StringTrimmingEllipsisWord = 4                                              '在布局矩形内最后一个单词的边界处断开，并在字符后面插入“...”
    StringTrimmingEllipsisPath = 5                                              '在布局矩形内最后一个路径的边界处断开，并在字符后面插入“...”[?]
End Enum

'字符串数字替换（不常用）
Public Enum StringDigitSubstitute
    StringDigitSubstituteUser = 0                                               '用户
    StringDigitSubstituteNone = 1                                               '禁用
    StringDigitSubstituteNational = 2                                           '按照国家（地区）替换
    StringDigitSubstituteTraditional = 3                                        '按照本机设定替换
End Enum

'画刷阴影样式（不常用）
Public Enum HatchStyle
    HatchStyleHorizontal                                                        '水平线
    HatchStyleVertical                                                          '垂直线
    HatchStyleForwardDiagonal                                                   '＼（抗锯齿）
    HatchStyleBackwardDiagonal                                                  '／（抗锯齿）
    HatchStyleCross                                                             '水平线和垂直线交叉
    HatchStyleDiagonalCross                                                     '斜线交叉（抗锯齿）
    HatchStyle05Percent                                                         '阴影比例为5%，以下枚举项类似
    HatchStyle10Percent                                                         '
    HatchStyle20Percent                                                         '
    HatchStyle25Percent                                                         '
    HatchStyle30Percent                                                         '
    HatchStyle40Percent                                                         '
    HatchStyle50Percent                                                         '
    HatchStyle60Percent                                                         '
    HatchStyle70Percent                                                         '
    HatchStyle75Percent                                                         '
    HatchStyle80Percent                                                         '
    HatchStyle90Percent                                                         '
    HatchStyleLightDownwardDiagonal                                             '＼，比HatchStyleForwardDiagonal更密（不抗锯齿）
    HatchStyleLightUpwardDiagonal                                               '／，比HatchStyleBackwardDiagonal更密（不抗锯齿）
    HatchStyleDarkDownwardDiagonal                                              '粗细为HatchStyleLightDownwardDiagonal的二倍，其余相同
    HatchStyleDarkUpwardDiagonal                                                '粗细为HatchStyleLightUpwardDiagonal的二倍，其余相同
    HatchStyleWideDownwardDiagonal                                              '
    HatchStyleWideUpwardDiagonal                                                '
    HatchStyleLightVertical                                                     '
    HatchStyleLightHorizontal                                                   '
    HatchStyleNarrowVertical                                                    '
    HatchStyleNarrowHorizontal                                                  '
    HatchStyleDarkVertical                                                      '
    HatchStyleDarkHorizontal                                                    '
    HatchStyleDashedDownwardDiagonal                                            '由＼组成的水平线
    HatchStyleDashedUpwardDiagonal                                              '由／组成的水平线
    HatchStyleDashedHorizontal                                                  '水平虚线
    HatchStyleDashedVertical                                                    '垂直虚线
    HatchStyleSmallConfetti                                                     '小斑点·
    HatchStyleLargeConfetti                                                     '大斑点●
    HatchStyleZigZag                                                            '锯齿形水平线
    HatchStyleWave                                                              '波浪形水平线
    HatchStyleDiagonalBrick                                                     '／砖块
    HatchStyleHorizontalBrick                                                   '水平砖块
    HatchStyleWeave                                                             '竹篾编织
    HatchStylePlaid                                                             '苏格兰格子花呢
    HatchStyleDivot                                                             '草皮层
    HatchStyleDottedGrid                                                        '水平交叉虚线
    HatchStyleDottedDiamond                                                     '斜线交叉虚线
    HatchStyleShingle                                                           '瓦片
    HatchStyleTrellis                                                           '格架
    HatchStyleSphere                                                            '球体棋盘
    HatchStyleSmallGrid                                                         '小网格
    HatchStyleSmallCheckerBoard                                                 '小棋盘外观
    HatchStyleLargeCheckerBoard                                                 '大棋盘外观
    HatchStyleOutlinedDiamond                                                   '斜线交叉（不抗锯齿）
    HatchStyleSolidDiamond                                                      '斜向大棋盘外观
    HatchStyleTotal                                                             '无阴影（允许画笔透明）
    HatchStyleLargeGrid = HatchStyleCross
    HatchStyleMin = HatchStyleHorizontal
    HatchStyleMax = HatchStyleTotal - 1
End Enum

'画笔对齐
Public Enum PenAlignment
    PenAlignmentCenter = 0                                                      '画笔在绘制的线条的中心对齐
    PenAlignmentInset = 1                                                       '在绘制多边形时，画笔在多边形边缘的内部对齐
End Enum

'画刷类型
Public Enum BrushType
    BrushTypeSolidColor = 0                                                     '实色
    BrushTypeHatchFill = 1                                                      '阴影画笔（参见HatchStyle枚举）
    BrushTypeTextureFill = 2                                                    '纹理画笔
    BrushTypePathGradient = 3                                                   '沿路径渐变
    BrushTypeLinearGradient = 4                                                 '线性渐变
End Enum

'虚线样式
Public Enum DashStyle
    DashStyleSolid                                                              '——————
    DashStyleDash                                                               '- - - -
    DashStyleDot                                                                '·······
    DashStyleDashDot                                                            '-·-·-·-·
    DashStyleDashDotDot                                                         '-··-··-··
    DashStyleCustom                                                             '用户自定义
End Enum

'虚线端点形状
Public Enum DashCap
    DashCapFlat = 0                                                             '平头
    DashCapRound = 2                                                            '●
    DashCapTriangle = 3                                                         '▲ *
    '注：如果画笔对齐设置为PenAlignmentInset，则不能使用DashCapTriangle。
End Enum

'直线端点形状
Public Enum LineCap
    LineCapFlat = 0                                                             '平头
    LineCapSquare = 1                                                           '■
    LineCapRound = 2                                                            '●
    LineCapTriangle = 3                                                         '▲
    LineCapNoAnchor = &H10                                                      '无锚点
    LineCapSquareAnchor = &H11                                                  '■锚点
    LineCapRoundAnchor = &H12                                                   '●锚点
    LineCapDiamondAnchor = &H13                                                 '◆锚点
    LineCapArrowAnchor = &H14                                                   '→锚点
    LineCapCustom = &HFF                                                        '自定义（参见CustomLineCapType）
    LineCapAnchorMask = &HF0                                                    '遮罩锚点[?]
End Enum

'自定义直线端点形状类型
Public Enum CustomLineCapType
    CustomLineCapTypeDefault = 0                                                '默认
    CustomLineCapTypeAdjustableArrow = 1                                        '自适应箭头
End Enum

'接线交点样式
Public Enum LineJoin
    LineJoinMiter = 0                                                           '斜方连接[?]
    LineJoinBevel = 1                                                           '斜面连接
    LineJoinRound = 2                                                           '圆角连接
    LineJoinMiterClipped = 3                                                    '斜方连接并剪出多余部分[?]
End Enum

'画笔类型，参见BrushType
Public Enum PenType
    PenTypeSolidColor = BrushTypeSolidColor
    PenTypeHatchFill = BrushTypeHatchFill
    PenTypeTextureFill = BrushTypeTextureFill
    PenTypePathGradient = BrushTypePathGradient
    PenTypeLinearGradient = BrushTypeLinearGradient
    PenTypeUnknown = -1                                                         '未知
End Enum

'图元文件类型（不常用）
Public Enum MetafileType
    MetafileTypeInvalid                                                         'Gdip中不能识别的图元文件格式
    MetafileTypeWmf                                                             'WMF文件（只包含GDI记录）
    MetafileTypeWmfPlaceable                                                    'WMF文件（文件前面有一个可放置的图元文件标头）[?]
    MetafileTypeEmf                                                             'EMF文件（只包含GDI记录）
    MetafileTypeEmfPlusOnly                                                     'EMF文件（只包含GDI+记录）
    MetafileTypeEmfPlusDual                                                     'EMF文件（包含GDI+记录和GDI记录，使用GDI绘制将导致质量下降）
    'WMF：Windows图元文件（Windows Metafile）是由简单的线条和封闭线条（图形）组成的矢量图。
    'EMF：增强型图元文件（Enhanced Metafile）是对WMF的改进和扩展，支持更多的颜色和更复杂的图像表示。
End Enum

'EMF文件类型（不常用）
Public Enum EmfType
    EmfTypeEmfOnly = MetafileTypeEmf
    EmfTypeEmfPlusOnly = MetafileTypeEmfPlusOnly
    EmfTypeEmfPlusDual = MetafileTypeEmfPlusDual
End Enum

'对象类型
Public Enum ObjectType
    ObjectTypeInvalid                                                           '无效的（保留）
    ObjectTypeBrush                                                             '画刷
    ObjectTypePen                                                               '画笔
    ObjectTypePath                                                              '路径
    ObjectTypeRegion                                                            '区域
    ObjectTypeImage                                                             '图像
    ObjectTypeFont                                                              '字体
    ObjectTypeStringFormat                                                      '字符串格式
    ObjectTypeImageAttributes                                                   '图像属性
    ObjectTypeCustomLineCap                                                     '自定义直线端点
    ObjectTypeGraphics                                                          '图形
    ObjectTypeMax = ObjectTypeGraphics                                          '
    ObjectTypeMin = ObjectTypeBrush                                             '
End Enum

'图元文件框架矩形度量单位
Public Enum MetafileFrameUnit
    MetafileFrameUnitPixel = UnitPixel
    MetafileFrameUnitPoint = UnitPoint
    MetafileFrameUnitInch = UnitInch
    MetafileFrameUnitDocument = UnitDocument                                    '文档定义（通常为 1/300 英寸）
    MetafileFrameUnitMillimeter = UnitMillimeter
    MetafileFrameUnitGdi                                                        '1/100 毫米（与GDI兼容）
End Enum

'坐标空间（不常用）
Public Enum CoordinateSpace
    CoordinateSpaceWorld                                                        '世界定义[?]
    CoordinateSpacePage                                                         '页面定义
    CoordinateSpaceDevice                                                       '设备定义
End Enum

'热键前缀[?]
Public Enum HotkeyPrefix
    HotkeyPrefixNone = 0                                                        '无前缀
    HotkeyPrefixShow = 1                                                        '显示前缀
    HotkeyPrefixHide = 2                                                        '隐藏前缀
End Enum

'刷新缓冲区（不常用）
Public Enum FlushIntention
    FlushIntentionFlush = 0                                                     '刷新所有批处理呈现操作（可能在呈现操作完成之前返回）
    FlushIntentionSync = 1                                                      '刷新所有批处理呈现操作（在呈现操作完成后才会返回）
End Enum

'编码器参数值类型（参见ImagePropertyType）
Public Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1
    EncoderParameterValueTypeASCII = 2
    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
    EncoderParameterValueTypeRational = 5
    EncoderParameterValueTypeLongRange = 6                                      '长范围
    EncoderParameterValueTypeUndefined = 7                                      '未定义
    EncoderParameterValueTypeRationalRange = 8                                  '实数范围[?]
End Enum

'编码器值[?]
Public Enum EncoderValue
    EncoderValueColorTypeCMYK                                                   'CMYK显色模式（GDIP 1.0 中无效）
    EncoderValueColorTypeYCCK                                                   'YCCK显色模式[?]（GDIP 1.0 中无效）
    EncoderValueCompressionLZW                                                  '使用LZW算法*压缩Tiff格式图像
    EncoderValueCompressionCCITT3                                               '使用CCTII3算法*压缩Tiff格式图像
    EncoderValueCompressionCCITT4                                               '使用CCTII4算法压缩Tiff格式图像
    EncoderValueCompressionRle                                                  '使用RLE算法*压缩Tiff格式图像
    EncoderValueCompressionNone                                                 '不压缩Tiff格式图像
    EncoderValueScanMethodInterlaced                                            '交错扫描方法[?]（GDIP 1.0 中无效）
    EncoderValueScanMethodNonInterlaced                                         '非交错扫描方法[?]（GDIP 1.0 中无效）
    EncoderValueVersionGif87                                                    '[?]
    EncoderValueVersionGif89                                                    '[?]
    EncoderValueRenderProgressive                                               '逐行方式渲染[?]（GDIP 1.0 中无效）
    EncoderValueRenderNonProgressive                                            '非逐行方式渲染[?]（GDIP 1.0 中无效）
    EncoderValueTransformRotate90                                               '将Jpeg图像*顺时针旋转90°（无损失）
    EncoderValueTransformRotate180                                              '将Jpeg图像顺时针旋转180°（无损失）
    EncoderValueTransformRotate270                                              '将Jpeg图像顺时针旋转270°（无损失）
    EncoderValueTransformFlipHorizontal                                         '将Jpeg图像水平翻转（无损失）
    EncoderValueTransformFlipVertical                                           '将Jpeg图像垂直翻转（无损失）
    EncoderValueMultiFrame                                                      '图像采用多帧编码
    EncoderValueLastFrame                                                       '多帧编码图像的最后一帧
    EncoderValueFlush                                                           '关闭编码器对象
    EncoderValueFrameDimensionTime                                              '以时间定义的帧维度[?]（GDIP 1.0 中无效）
    EncoderValueFrameDimensionResolution                                        '以分辨率定义的帧维度[?]（GDIP 1.0 中无效）
    EncoderValueFrameDimensionPage                                              '以页面定义的帧维度[?]（GDIP 1.0 中无效）
    EncoderValueColorTypeGray                                                   '灰度颜色
    EncoderValueColorTypeRGB                                                    'RGB颜色
    '注1：LZW算法，即“串表压缩算法”（Lempel-Ziv-Welch Encoding），该算法通过建立一个字符串表，用较短的代码来表示较长
    '的字符串来实现压缩。
    '注2：暂未找到CCTII3算法和CCTII4算法介绍。
    '注3：RLE算法，即“行程长度压缩算法”（Run Length Encoding），该算法以一个表示块数长度的属性字节加上一个数据块，来
    '代表原来连续的若干块数据，从而达到节省存储空间的目的。
    '注4：Jpeg，即“联合图像专家组”（Joint Photographic Experts Group）。JPEG图像格式属于有损压缩格式，是最常用的图
    '像文件格式之一。
End Enum

'位图压缩模式
Public Enum BitmapCompressionMode
    BL_RGB = &H0&
    BI_RLE8 = &H1&
    BI_RLE4 = &H2&
    BI_BITFIELDS = &H3&
    BI_JPEG = &H4&
    BI_PNG = &H5&
End Enum

'Debug事件级别（不常用）
Public Enum DebugEventLevel
    DebugEventLevelFatal                                                        '关键错误[?]
    DebugEventLevelWarning                                                      '警告
End Enum

'调用Gdip函数返回的状态
Public Enum GpStatus
    Ok = 0                                                                      '成功调用Gdip函数
    GenericError = 1                                                            '调用Gdip函数时出现了一般性的错误
    InvalidParameter = 2                                                        '调用Gdip函数时输入的参数无效
    OutOfMemory = 3                                                             '调用Gdip函数时内存不足
    ObjectBusy = 4                                                              '调用Gdip函数时目标对象忙碌无响应
    InsufficientBuffer = 5                                                      '调用Gdip函数时缓冲区大小不足
    NotImplemented = 6                                                          '调用Gdip函数时尚未实现操作
    Win32Error = 7                                                              '引发Win32错误
    WrongState = 8                                                              '状态错误
    Aborted = 9                                                                 '调用Gdip函数时操作被终止
    FileNotFound = 10                                                           '找不到文件
    ValueOverflow = 11                                                          '调用Gdip函数时参数值溢出
    AccessDenied = 12                                                           '调用Gdip函数时访问被拒绝
    UnknownImageFormat = 13                                                     '未知的图像格式
    FontFamilyNotFound = 14                                                     '找不到字体族
    FontStyleNotFound = 15                                                      '找不到字体类型
    NotTrueTypeFont = 16                                                        '不是TrueType格式字体
    UnsupportedGdiplusVersion = 17                                              '使用的是不支持的GDIP版本
    GdiplusNotInitialized = 18                                                  'GDIP未初始化
    PropertyNotFound = 19                                                       '找不到对应属性
    PropertyNotSupported = 20                                                   '不支持的对应属性
End Enum

'图像文件后缀
Public Enum ImageFileSuffix
    Bmp
    JPG
    GIF
    EMF
    WMF
    TIF
    PNG
    ICO
End Enum

'Gdip通用对象
Public Enum GdiplusCommonObject
    GdiplusPen = &H0
    GdiplusBrush = &H1
    GdiplusStringFormat = &H2
    GdiplusMatrix = &H3
    GdiplusFont = &H4
    GdiplusFontFamily = &H5
    GdiplusGraphics = &H6
    GdiplusPath = &H7
    GdiplusRegion = &H8
    GdiplusPathIter = &H9
    GdiplusImage = &HA
    GdiplusCachedBitmap = &HB
    GdiplusDeviceContext = &HC
End Enum

'EMF+文件记录类型
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
    EmfPlusRecordTypeInvalid = 16384
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
    EmfPlusRecordTotal = 16439
    EmfPlusRecordTypeMax = 16438
    EmfPlusRecordTypeMin = 16385
End Enum

'计算方式
Public Enum CalculatingMethod
    RoundDown                                                                   '向下取整
    RoundNear                                                                   '临近取整
    RoundUp                                                                     '向上取整
End Enum

'################  结  构  体  ################
'点
Public Type PointL
    X As Long
    Y As Long
End Type

Public Type PointF
    X As Single
    Y As Single
End Type

'矩形
Public Type RectL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RectF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

'尺寸
Public Type SizeL
    Width As Long                                                               '宽
    Height As Long                                                              '高
End Type

Public Type SizeF
    Width As Single
    Height As Single
End Type

'RGB颜色
Public Type RGBQuad
    rgbBlue As Byte                                                             '蓝色分量
    rgbGreen As Byte                                                            '绿色分量
    rgbRed As Byte                                                              '红色分量
    rgbReserved As Byte                                                         '保留值，必须为0
End Type

'ARGB颜色
Public Type ARGBColor
    Alpha As Byte                                                               'α分量（透明度*）
    Red As Byte
    Green As Byte
    Blue As Byte
    '注：透明度值从0（完全透明）到255（完全不透明）。
End Type

'位图信息头部
Public Type BitmapInfoHeader
    biSize As Long                                                              '位图信息头部的大小（以字节计算）
    biWidth As Long                                                             '位图宽度（以像素计算，下同）
    biHeight As Long                                                            '位图高度
    biPlanes As Integer                                                         '目标设备面数，必须为1
    biBitCount As Integer                                                       '记录每个像素所需要的位（Bit）数
    biCompression As BitmapCompressionMode                                      '图片采用的压缩方式，默认为不压缩（BL_RGB）
    biSizeImage As Long                                                         '图像的大小*（以字节计算）
    biXPelsPerMeter As Long                                                     '位图的目标设备的水平分辨率（以像素/米计算）
    biYPelsPerMeter As Long                                                     '位图的目标设备的垂直分辨率（以像素/米计算）
    biClrUsed As Long                                                           '调色板中颜色数量，默认为0
    biClrImportant As Long                                                      '重要颜色的数量，默认为0*
    '注1：对于未压缩的位图，biSizeImage的值默认为0。
    '注2：如果biClrImportant的值为0，则表示所有颜色都很重要。
End Type

'位图信息
Public Type BitmapInfo
    bmiHeader As BitmapInfoHeader
    bmiColors As RGBQuad
End Type

'位图数据
Public Type BitmapData
    Width As Long
    Height As Long
    Stride As Long                                                              '位图对象的跨距宽度
    PixelFormat As Long                                                         '像素格式
    Scan0 As Long                                                               '位图中第一个像素数据的地址
    Reserved As Long                                                            '保留值，必须为0
End Type

'颜色矩阵
Public Type ColorMatrix
    Matrix(0 To 4, 0 To 4) As Double
End Type

'路径数据
Public Type PathData
    Count As Long                                                               '数量
    Points As Long                                                              '指向PointL数组的指针[?]
    Types As Long                                                               '指向Byte数组的指针[?]
End Type

'编/解码器类别标识符
Public Type ClsID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'单个编码器参数
Public Type EncoderParameter
    Guid As ClsID
    NumberOfValues As Long                                                      '值的数量[?]
    ValueType As EncoderParameterValueType
    value As EncoderValue
End Type

'编码器参数
Public Type EncoderParameters
    Count As Long                                                               '数量
    Parameter As EncoderParameter
End Type

'字体类型
Public Type FontType
    Name As String
    Size As Single
    Weight As FontWeight
    Style As FontStyle
End Type

'逻辑字体结构（Ascii）[?]
Public Type LogFontA
    lfHeight As Long                                                            '字符高度*
    lfWidth As Long                                                             '字符的平均宽度*
    lfEscapement As Long                                                        '转义向量与设备的x轴之间的角度（以1/10°计算）
    lfOrientation As Long                                                       '每个字符的基线和设备x轴之间的角度（以1/10°计算）
    lfWeight As FontWeight                                                      '字重（字的粗细）
    lfItalic As Byte                                                            '斜体
    lfUnderline As Byte                                                         '下划线
    lfStrikeOut As Byte                                                         '删除线
    lfCharSet As CharSetType                                                    '字符集
    lfOutPrecision As OutPrecisionType                                          '输出精度*
    lfClipPrecision As ClipPrecisionType                                        '剪裁精度*
    lfQuality As FontQualityType                                                '输出质量
    lfPitchAndFamily As Byte                                                    '字体的间距和系列
    lfFaceName(31) As Byte                                                      '以NULL结尾的字符串指定字体的字体名称
    '注1： 字体映射器按以下方式解释lfHeight中指定的值：
    '--lfHeight的值----描述----------------------------------------------------------------------------------
    '  >0              字体映射器将此值转换为设备单位，并将其与可用字体的单元格高度匹配。
    '  =0              字体映射器在搜索匹配项时使用默认高度值。
    '  <0              字体映射器将此值转换为设备单位，并将其绝对值与可用字体的字符高度匹配。
    '注2：如果lfWidth为零，则设备的纵横比将与可用字体的数字化纵横比进行匹配，以查找由差值的绝对值确定的最接近匹配项。
    '注3：lfOutPrecision用于定义输出与所请求字体的高度、宽度、字符方向、转义、间距和字体类型的匹配程度。
    '注4：lfClipPrecision用于定义如何剪裁部分超出剪裁区域的字符。
End Type

'逻辑字体结构（WideChar）[?]
Public Type LogFontW
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As FontWeight
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As CharSetType
    lfOutPrecision As OutPrecisionType
    lfClipPrecision As ClipPrecisionType
    lfQuality As FontQualityType
    lfPitchAndFamily As Byte
    lfFaceName(31) As Byte
End Type

'图像编/解码器信息
Public Type ImageCodecInfo
    ClassID As ClsID
    FormatID As ClsID
    CodecName As Long                                                           '编/解码器名称
    DllName As Long
    FormatDescription As Long                                                   '格式描述
    FilenameExtension As Long
    MimeType As Long
    Flags As ImageCodecFlags
    Version As Long
    SigCount As Long
    SigSize As Long
    SigPattern As Long
    SigMask As Long
End Type

'调色板
Public Type ColorPalette
    Flags As PaletteFlags
    Count As Long
    Entries(0 To 255) As Long
End Type

'WMF文件标头描述矩形
Public Type PwmfRect16
    Left As Integer
    Top As Integer
    Width As Integer
    Height As Integer
End Type

'WMF文件可放置的图元文件标头
Public Type WmfPlaceableFileHeader
    Key As Long
    Hmf As Integer
    BoundingBox As PwmfRect16
    Inch As Integer
    Reserved As Long
    Checksum As Integer
End Type

'ENH元数据的头部数据[?]
Public Type ENHMETAHEADER3
    itype As Long
    nSize As Long
    rclBounds As RectL
    rclFrame As RectL
    dSignature As Long
    nVersion As Long
    nBytes As Long
    nRecords As Long
    nHandles As Integer
    sReserved As Integer
    nDescription As Long
    offDescription As Long
    nPalEntries As Long
    szlDevice As SizeL
    szlMillimeters As SizeL
End Type

'元数据头部
Public Type METAHEADER
    mtType As Integer
    mtHeaderSize As Integer
    mtVersion As Integer
    mtSize As Long
    mtNoObjects As Integer
    mtMaxRecord As Long
    mtNoParameters As Integer
End Type

'图元文件头部[?]
Public Type MetafileHeader
    mType As MetafileType
    Size As Long
    Version As Long
    EmfPlusFlags As Long
    DpiX As Single
    DpiY As Single
    X As Long
    Y As Long
    Width As Long
    Height As Long
    EmfHeader As ENHMETAHEADER3
    EmfPlusHeaderSize As Long
    LogicalDpiX As Long
    LogicalDpiY As Long
End Type

'属性项目
Public Type PropertyItem
    iPropId As Long
    iLength As Long
    itype As Integer
    iValue As Long
End Type

'字符范围
Public Type CharacterRange
    First As Long
    Length As Long
End Type

'GDIP初始化输入
Public Type GdiplusStartupInput
    GdiplusVersion As Long                                                      '版本
    DebugEventCallback As Long                                                  'Debug事件回调
    SuppressBackgroundThread As Long                                            '抑制后台线程[?]
    SuppressExternalCodecs As Long                                              '抑制外部编/解码器[?]
End Type

'GDIP对象
Public Type GdiplusObject
    GdiplusObjectName As String
    GdiplusObjectType As GdiplusCommonObject
    GdiplusObjectHandle As Long
End Type

'##############  函  数  声  明  ##############
'1、设备场景上下文
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal Graphics As Long, hDC As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal Graphics As Long, ByVal hDC As Long) As GpStatus
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDCAPI Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

'2、画布
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, Graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hWnd As Long, Graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHDC2 Lib "gdiplus" Alias "GdipCreateFromHdc2" (ByVal hDC As Long, ByVal hDevice As Long, Graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWNDICM Lib "gdiplus" Alias "GdipCreateFromHWndICM" (ByVal hWnd As Long, Graphics As Long) As GpStatus
Public Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, Graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Public Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal Graphics As Long, ByVal lColor As Long) As GpStatus

'3、混合模式、渲染、平滑模式、像素偏移、文本渲染提示、文本差异和差值模式
Public Declare Function GdipSetCompositingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal CompositingMd As CompositingMode) As GpStatus
Public Declare Function GdipGetCompositingMode Lib "gdiplus" (ByVal Graphics As Long, CompositingMd As CompositingMode) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetRenderingOrigin Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Long, ByVal Y As Long) As GpStatus
Public Declare Function GdipGetRenderingOrigin Lib "gdiplus" (ByVal Graphics As Long, X As Long, Y As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetCompositingQuality Lib "gdiplus" (ByVal Graphics As Long, ByVal CompositingQlty As CompositingQuality) As GpStatus
Public Declare Function GdipGetCompositingQuality Lib "gdiplus" (ByVal Graphics As Long, CompositingQlty As CompositingQuality) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, ByVal SmoothingMd As SmoothingMode) As GpStatus
Public Declare Function GdipGetSmoothingMode Lib "gdiplus" (ByVal Graphics As Long, SmoothingMd As SmoothingMode) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, ByVal PixOffsetMode As PixelOffsetMode) As GpStatus
Public Declare Function GdipGetPixelOffsetMode Lib "gdiplus" (ByVal Graphics As Long, PixOffsetMode As PixelOffsetMode) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetTextRenderingHint Lib "gdiplus" (ByVal Graphics As Long, ByVal Mode As TextRenderingHint) As GpStatus
Public Declare Function GdipGetTextRenderingHint Lib "gdiplus" (ByVal Graphics As Long, Mode As TextRenderingHint) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetTextContrast Lib "gdiplus" (ByVal Graphics As Long, ByVal contrast As Long) As GpStatus
Public Declare Function GdipGetTextContrast Lib "gdiplus" (ByVal Graphics As Long, contrast As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, ByVal Interpolation As InterpolationMode) As GpStatus
Public Declare Function GdipGetInterpolationMode Lib "gdiplus" (ByVal Graphics As Long, Interpolation As InterpolationMode) As GpStatus

'4、世界变换、页面设置
Public Declare Function GdipSetWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipGetWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetWorldTransform Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Public Declare Function GdipMultiplyWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal Matrix As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal sX As Single, ByVal sY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipResetPageTransform Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Public Declare Function GdipGetPageUnit Lib "gdiplus" (ByVal Graphics As Long, Unit As GpUnit) As GpStatus
Public Declare Function GdipGetPageScale Lib "gdiplus" (ByVal Graphics As Long, mScale As Single) As GpStatus
Public Declare Function GdipSetPageUnit Lib "gdiplus" (ByVal Graphics As Long, ByVal Unit As GpUnit) As GpStatus
Public Declare Function GdipSetPageScale Lib "gdiplus" (ByVal Graphics As Long, ByVal mScale As Single) As GpStatus

'5、获取分辨率、点集变换
Public Declare Function GdipGetDpiX Lib "gdiplus" (ByVal Graphics As Long, DPI As Single) As GpStatus
Public Declare Function GdipGetDpiY Lib "gdiplus" (ByVal Graphics As Long, DPI As Single) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipTransformPointsI Lib "gdiplus" (ByVal Graphics As Long, ByVal DestSpace As CoordinateSpace, ByVal SrcSpace As CoordinateSpace, Points As PointL, ByVal Count As Long) As GpStatus

'6、获得最接近颜色[?]、创建半色调调色板
Public Declare Function GdipGetNearestColor Lib "gdiplus" (ByVal Graphics As Long, Argb As Long) As GpStatus
Public Declare Function GdipCreateHalftonePalette Lib "gdiplus" () As Long

'7、剪裁相关
Public Declare Function GdipSetClipGraphics Lib "gdiplus" (ByVal Graphics As Long, ByVal SrcGraphics As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipPath Lib "gdiplus" (ByVal Graphics As Long, ByVal Path As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipRegion Lib "gdiplus" (ByVal Graphics As Long, ByVal Region As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipSetClipHrgn Lib "gdiplus" (ByVal Graphics As Long, ByVal hRgn As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipResetClip Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipTranslateClipI Lib "gdiplus" (ByVal Graphics As Long, ByVal dX As Long, ByVal dY As Long) As GpStatus
Public Declare Function GdipGetClip Lib "gdiplus" (ByVal Graphics As Long, ByVal Region As Long) As GpStatus
Public Declare Function GdipGetClipBoundsI Lib "gdiplus" (ByVal Graphics As Long, Rect As RectL) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipIsClipEmpty Lib "gdiplus" (ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipGetVisibleClipBoundsI Lib "gdiplus" (ByVal Graphics As Long, Rect As RectL) As GpStatus
Public Declare Function GdipIsVisibleClipEmpty Lib "gdiplus" (ByVal Graphics As Long, Result As Long) As GpStatus

'8、可视化矩形和可视化点判断
Public Declare Function GdipIsVisibleRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Result As Long) As GpStatus
Public Declare Function GdipIsVisiblePointI Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Long, ByVal Y As Long, Result As Long) As GpStatus

'9、存储和复原画布
Public Declare Function GdipSaveGraphics Lib "gdiplus" (ByVal Graphics As Long, State As Long) As GpStatus
Public Declare Function GdipRestoreGraphics Lib "gdiplus" (ByVal Graphics As Long, ByVal State As Long) As GpStatus

'10、容器操作
Public Declare Function GdipBeginContainerI Lib "gdiplus" (ByVal Graphics As Long, dstRect As RectL, srcRect As RectL, ByVal Unit As GpUnit, State As Long) As GpStatus
Public Declare Function GdipBeginContainer2 Lib "gdiplus" (ByVal Graphics As Long, State As Long) As GpStatus
Public Declare Function GdipEndContainer Lib "gdiplus" (ByVal Graphics As Long, ByVal State As Long) As GpStatus

'11、绘制线段、弧线、贝塞尔曲线、矩形、椭圆、扇形、多边形、自由曲线、封闭自由曲线、路径
Public Declare Function GdipDrawLineI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As GpStatus
Public Declare Function GdipDrawLinesI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As PointL, ByVal Count As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipDrawArcI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal StartAngle As Single, ByVal SweepAngle As Single) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipDrawBezierI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As GpStatus
Public Declare Function GdipDrawBeziersI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As PointL, ByVal Count As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipDrawRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawRectanglesI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Rects As RectL, ByVal Count As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipDrawPieI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal StartAngle As Single, ByVal SweepAngle As Single) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As PointL, ByVal Count As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipDrawCurveI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurve2I Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As PointL, ByVal Count As Long, ByVal Tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3I Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As PointL, ByVal Count As Long, ByVal Offset As Long, ByVal NumberOfSegments As Long, ByVal Tension As Single) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipDrawClosedCurveI Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurve2I Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, Points As PointL, ByVal Count As Long, ByVal Tension As Single) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipDrawPath Lib "gdiplus" (ByVal Graphics As Long, ByVal Pen As Long, ByVal Path As Long) As GpStatus

'12、填充矩形、椭圆、扇形、多边形、封闭自由曲线、路径、区域
Public Declare Function GdipFillRectangleI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipFillRectanglesI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Rects As RectL, ByVal Count As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipFillPieI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal StartAngle As Single, ByVal SweepAngle As Single) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As PointL, ByVal Count As Long, ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygon2I Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As PointL, ByVal Count As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipFillClosedCurveI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As PointL, ByVal Count As Long, ByVal Tension As Single, ByVal FillMd As FillMode) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipFillPath Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal Path As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipFillRegion Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal Region As Long) As GpStatus

'13、图像绘制
Public Declare Function GdipDrawImageI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long) As GpStatus
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawImagePointsI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, DstPoints As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal SrcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal SrcUnit As GpUnit, Optional ByVal ImageAttributes As Long = 0, Optional ByVal CallBack As Long = 0, Optional ByVal CallBackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, Points As PointL, ByVal Count As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal SrcUnit As GpUnit, Optional ByVal ImageAttributes As Long = 0, Optional ByVal CallBack As Long = 0, Optional ByVal CallBackData As Long = 0) As GpStatus

'14、图像编/解码器参数
Public Declare Function GdipGetImageDecoders Lib "gdiplus" (ByVal NumDecoders As Long, ByVal Size As Long, Decoders As Any) As GpStatus
Public Declare Function GdipGetImageDecodersSize Lib "gdiplus" (NumDecoders As Long, Size As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipGetImageEncodersSize Lib "gdiplus" (NumEncoders As Long, Size As Long) As GpStatus
Public Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal NumEncoders As Long, ByVal Size As Long, encoders As Any) As GpStatus
Public Declare Function GdipGetEncoderParameterListSize Lib "gdiplus" (ByVal Image As Long, ClsIDEncoder As ClsID, Size As Long) As GpStatus
Public Declare Function GdipGetEncoderParameterList Lib "gdiplus" (ByVal Image As Long, ClsIDEncoder As ClsID, ByVal Size As Long, Buffer As EncoderParameters) As GpStatus

'15、图像加载、释放、复制、保存
Public Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromFileICM Lib "gdiplus" (ByVal FileName As Long, Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromStreamICM Lib "gdiplus" (ByVal Stream As Any, Image As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipCloneImage Lib "gdiplus" (ByVal Image As Long, CloneImage As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal Image As Long, ByVal FileName As Long, ClsIDEncoder As ClsID, EncoderParams As Any) As GpStatus
Public Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As Any, ClsIDEncoder As ClsID, EncoderParams As Any) As GpStatus
Public Declare Function GdipSaveAdd Lib "gdiplus" (ByVal Image As Long, EncoderParams As EncoderParameters) As GpStatus
Public Declare Function GdipSaveAddImage Lib "gdiplus" (ByVal Image As Long, ByVal NewImage As Long, EncoderParams As EncoderParameters) As GpStatus

'16、图像参数信息及相关操作
Public Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, Width As Single, Height As Single) As GpStatus
Public Declare Function GdipGetImageType Lib "gdiplus" (ByVal Image As Long, itype As ImageType) As GpStatus
Public Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus
Public Declare Function GdipGetImageHorizontalResolution Lib "gdiplus" (ByVal Image As Long, Resolution As Single) As GpStatus
Public Declare Function GdipGetImageVerticalResolution Lib "gdiplus" (ByVal Image As Long, Resolution As Single) As GpStatus
Public Declare Function GdipGetImageFlags Lib "gdiplus" (ByVal Image As Long, Flags As Long) As GpStatus
Public Declare Function GdipGetImageRawFormat Lib "gdiplus" (ByVal Image As Long, Format As ClsID) As GpStatus
Public Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal Image As Long, PixelFormat As Long) As GpStatus
Public Declare Function GdipGetImageThumbnail Lib "gdiplus" (ByVal Image As Long, ByVal ThumbWidth As Long, ByVal ThumbHeight As Long, ThumbImage As Long, Optional ByVal CallBack As Long = 0, Optional ByVal CallBackData As Long = 0) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipImageGetFrameDimensionsCount Lib "gdiplus" (ByVal Image As Long, Count As Long) As GpStatus
Public Declare Function GdipImageGetFrameDimensionsList Lib "gdiplus" (ByVal Image As Long, DimensionIDs As ClsID, ByVal Count As Long) As GpStatus
Public Declare Function GdipImageGetFrameCount Lib "gdiplus" (ByVal Image As Long, dimensionID As ClsID, Count As Long) As GpStatus
Public Declare Function GdipImageSelectActiveFrame Lib "gdiplus" (ByVal Image As Long, dimensionID As ClsID, ByVal FrameIndex As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As RotateFlipType) As GpStatus
Public Declare Function GdipImageForceValidation Lib "gdiplus" (ByVal Image As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipGetImagePalette Lib "gdiplus" (ByVal Image As Long, Palette As ColorPalette, ByVal Size As Long) As GpStatus
Public Declare Function GdipSetImagePalette Lib "gdiplus" (ByVal Image As Long, Palette As ColorPalette) As GpStatus
Public Declare Function GdipGetImagePaletteSize Lib "gdiplus" (ByVal Image As Long, Size As Long) As GpStatus

'17、属性操作
Public Declare Function GdipGetPropertyCount Lib "gdiplus" (ByVal Image As Long, NumOfProperty As Long) As GpStatus
Public Declare Function GdipGetPropertyIdList Lib "gdiplus" (ByVal Image As Long, ByVal NumOfProperty As Long, List As Long) As GpStatus
Public Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal Image As Long, ByVal PropId As Long, Size As Long) As GpStatus
Public Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal PropId As Long, ByVal PropSize As Long, Buffer As PropertyItem) As GpStatus
Public Declare Function GdipGetPropertySize Lib "gdiplus" (ByVal Image As Long, TotalBufferSize As Long, NumProperties As Long) As GpStatus
Public Declare Function GdipGetAllPropertyItems Lib "gdiplus" (ByVal Image As Long, ByVal TotalBufferSize As Long, ByVal NumProperties As Long, AllItems As PropertyItem) As GpStatus
Public Declare Function GdipRemovePropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal PropId As Long) As GpStatus
Public Declare Function GdipSetPropertyItem Lib "gdiplus" (ByVal Image As Long, Item As PropertyItem) As GpStatus

'18、画笔相关
Public Declare Function GdipCreatePen1 Lib "gdiplus" (ByVal Color As Long, ByVal Width As Single, ByVal Unit As GpUnit, Pen As Long) As GpStatus
Public Declare Function GdipCreatePen2 Lib "gdiplus" (ByVal Brush As Long, ByVal Width As Single, ByVal Unit As GpUnit, Pen As Long) As GpStatus
Public Declare Function GdipClonePen Lib "gdiplus" (ByVal Pen As Long, ClonePen As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal Pen As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetPenWidth Lib "gdiplus" (ByVal Pen As Long, ByVal Width As Single) As GpStatus
Public Declare Function GdipGetPenWidth Lib "gdiplus" (ByVal Pen As Long, Width As Single) As GpStatus
Public Declare Function GdipSetPenUnit Lib "gdiplus" (ByVal Pen As Long, ByVal Unit As GpUnit) As GpStatus
Public Declare Function GdipGetPenUnit Lib "gdiplus" (ByVal Pen As Long, Unit As GpUnit) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetPenLineCap Lib "gdiplus" Alias "GdipSetPenLineCap197819" (ByVal Pen As Long, ByVal StartCap As LineCap, ByVal EndCap As LineCap, ByVal DshCap As DashCap) As GpStatus
Public Declare Function GdipSetPenStartCap Lib "gdiplus" (ByVal Pen As Long, ByVal StartCap As LineCap) As GpStatus
Public Declare Function GdipSetPenEndCap Lib "gdiplus" (ByVal Pen As Long, ByVal EndCap As LineCap) As GpStatus
Public Declare Function GdipSetPenDashCap Lib "gdiplus" Alias "GdipSetPenDashCap197819" (ByVal Pen As Long, ByVal dcap As DashCap) As GpStatus
Public Declare Function GdipGetPenStartCap Lib "gdiplus" (ByVal Pen As Long, StartCap As LineCap) As GpStatus
Public Declare Function GdipGetPenEndCap Lib "gdiplus" (ByVal Pen As Long, EndCap As LineCap) As GpStatus
Public Declare Function GdipGetPenDashCap Lib "gdiplus" Alias "GdipGetPenDashCap197819" (ByVal Pen As Long, dcap As DashCap) As GpStatus
Public Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal Pen As Long, ByVal lnJoin As LineJoin) As GpStatus
Public Declare Function GdipGetPenLineJoin Lib "gdiplus" (ByVal Pen As Long, lnJoin As LineJoin) As GpStatus
Public Declare Function GdipSetPenCustomStartCap Lib "gdiplus" (ByVal Pen As Long, ByVal CustomCap As Long) As GpStatus
Public Declare Function GdipGetPenCustomStartCap Lib "gdiplus" (ByVal Pen As Long, CustomCap As Long) As GpStatus
Public Declare Function GdipSetPenCustomEndCap Lib "gdiplus" (ByVal Pen As Long, ByVal CustomCap As Long) As GpStatus
Public Declare Function GdipGetPenCustomEndCap Lib "gdiplus" (ByVal Pen As Long, CustomCap As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetPenMiterLimit Lib "gdiplus" (ByVal Pen As Long, ByVal MiterLimit As Single) As GpStatus
Public Declare Function GdipGetPenMiterLimit Lib "gdiplus" (ByVal Pen As Long, MiterLimit As Single) As GpStatus
Public Declare Function GdipSetPenMode Lib "gdiplus" (ByVal Pen As Long, ByVal PenMode As PenAlignment) As GpStatus
Public Declare Function GdipGetPenMode Lib "gdiplus" (ByVal Pen As Long, PenMode As PenAlignment) As GpStatus
Public Declare Function GdipSetPenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipGetPenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetPenTransform Lib "gdiplus" (ByVal Pen As Long) As GpStatus
Public Declare Function GdipMultiplyPenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal Matrix As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslatePenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScalePenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal sX As Single, ByVal sY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotatePenTransform Lib "gdiplus" (ByVal Pen As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipSetPenColor Lib "gdiplus" (ByVal Pen As Long, ByVal Argb As Long) As GpStatus
Public Declare Function GdipGetPenColor Lib "gdiplus" (ByVal Pen As Long, Argb As Long) As GpStatus
Public Declare Function GdipSetPenBrushFill Lib "gdiplus" (ByVal Pen As Long, ByVal Brush As Long) As GpStatus
Public Declare Function GdipGetPenBrushFill Lib "gdiplus" (ByVal Pen As Long, Brush As Long) As GpStatus
Public Declare Function GdipGetPenFillType Lib "gdiplus" (ByVal Pen As Long, PType As PenType) As GpStatus
Public Declare Function GdipGetPenDashStyle Lib "gdiplus" (ByVal Pen As Long, DshStyle As DashStyle) As GpStatus
Public Declare Function GdipSetPenDashStyle Lib "gdiplus" (ByVal Pen As Long, ByVal DshStyle As DashStyle) As GpStatus
Public Declare Function GdipGetPenDashOffset Lib "gdiplus" (ByVal Pen As Long, Offset As Single) As GpStatus
Public Declare Function GdipSetPenDashOffset Lib "gdiplus" (ByVal Pen As Long, ByVal Offset As Single) As GpStatus
Public Declare Function GdipGetPenDashCount Lib "gdiplus" (ByVal Pen As Long, Count As Long) As GpStatus
Public Declare Function GdipSetPenDashArray Lib "gdiplus" (ByVal Pen As Long, Dash As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenDashArray Lib "gdiplus" (ByVal Pen As Long, Dash As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenCompoundCount Lib "gdiplus" (ByVal Pen As Long, Count As Long) As GpStatus
Public Declare Function GdipSetPenCompoundArray Lib "gdiplus" (ByVal Pen As Long, Dash As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenCompoundArray Lib "gdiplus" (ByVal Pen As Long, Dash As Single, ByVal Count As Long) As GpStatus

'19、直线端点、箭头端点
Public Declare Function GdipCreateCustomLineCap Lib "gdiplus" (ByVal FillPath As Long, ByVal StrokePath As Long, ByVal baseCap As LineCap, ByVal baseInset As Single, CustomCap As Long) As GpStatus
Public Declare Function GdipDeleteCustomLineCap Lib "gdiplus" (ByVal CustomCap As Long) As GpStatus
Public Declare Function GdipCloneCustomLineCap Lib "gdiplus" (ByVal CustomCap As Long, clonedCap As Long) As GpStatus
Public Declare Function GdipGetCustomLineCapType Lib "gdiplus" (ByVal CustomCap As Long, capType As CustomLineCapType) As GpStatus
Public Declare Function GdipSetCustomLineCapStrokeCaps Lib "gdiplus" (ByVal CustomCap As Long, ByVal StartCap As LineCap, ByVal EndCap As LineCap) As GpStatus
Public Declare Function GdipGetCustomLineCapStrokeCaps Lib "gdiplus" (ByVal CustomCap As Long, StartCap As LineCap, EndCap As LineCap) As GpStatus
Public Declare Function GdipSetCustomLineCapStrokeJoin Lib "gdiplus" (ByVal CustomCap As Long, ByVal lnJoin As LineJoin) As GpStatus
Public Declare Function GdipGetCustomLineCapStrokeJoin Lib "gdiplus" (ByVal CustomCap As Long, lnJoin As LineJoin) As GpStatus
Public Declare Function GdipSetCustomLineCapBaseCap Lib "gdiplus" (ByVal CustomCap As Long, ByVal baseCap As LineCap) As GpStatus
Public Declare Function GdipGetCustomLineCapBaseCap Lib "gdiplus" (ByVal CustomCap As Long, baseCap As LineCap) As GpStatus
Public Declare Function GdipSetCustomLineCapBaseInset Lib "gdiplus" (ByVal CustomCap As Long, ByVal Inset As Single) As GpStatus
Public Declare Function GdipGetCustomLineCapBaseInset Lib "gdiplus" (ByVal CustomCap As Long, Inset As Single) As GpStatus
Public Declare Function GdipSetCustomLineCapWidthScale Lib "gdiplus" (ByVal CustomCap As Long, ByVal WidthScale As Single) As GpStatus
Public Declare Function GdipGetCustomLineCapWidthScale Lib "gdiplus" (ByVal CustomCap As Long, WidthScale As Single) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipCreateAdjustableArrowCap Lib "gdiplus" (ByVal Height As Single, ByVal Width As Single, ByVal isFilled As Long, Cap As Long) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapHeight Lib "gdiplus" (ByVal Cap As Long, ByVal Height As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapHeight Lib "gdiplus" (ByVal Cap As Long, Height As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapWidth Lib "gdiplus" (ByVal Cap As Long, ByVal Width As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapWidth Lib "gdiplus" (ByVal Cap As Long, Width As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapMiddleInset Lib "gdiplus" (ByVal Cap As Long, ByVal MiddleInset As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapMiddleInset Lib "gdiplus" (ByVal Cap As Long, MiddleInset As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapFillState Lib "gdiplus" (ByVal Cap As Long, ByVal bFillState As Long) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapFillState Lib "gdiplus" (ByVal Cap As Long, bFillState As Long) As GpStatus

'20、位图相关
Public Declare Function GdipCreateBitmapFromFile Lib "gdiplus" (ByVal FileName As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFileICM Lib "gdiplus" (ByVal FileName As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromStream Lib "gdiplus" (ByVal Stream As Any, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromStreamICM Lib "gdiplus" (ByVal Stream As Any, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Stride As Long, ByVal PixelFormat As Long, Scan0 As Any, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal Graphics As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGdiDib Lib "gdiplus" (gdiBitmapInfo As BitmapInfo, ByVal gdiBitmapData As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hBm As Long, ByVal hPal As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" (ByVal Bitmap As Long, hBmReturn As Long, ByVal Background As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHICON Lib "gdiplus" (ByVal hIcon As Long, Bitmap As Long) As GpStatus
Public Declare Function GdipCreateHICONFromBitmap Lib "gdiplus" (ByVal Bitmap As Long, hBmReturn As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromResource Lib "gdiplus" (ByVal hInstance As Long, ByVal lpBitmapName As Long, Bitmap As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal Bitmap As Long, Rect As RectL, ByVal Flags As ImageLockMode, ByVal PixelFormat As Long, LockedBitmapData As BitmapData) As GpStatus
Public Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal Bitmap As Long, LockedBitmapData As BitmapData) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal Bitmap As Long, ByVal X As Long, ByVal Y As Long, Color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel Lib "gdiplus" (ByVal Bitmap As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipBitmapSetResolution Lib "gdiplus" (ByVal Bitmap As Long, ByVal xDpi As Single, ByVal yDpi As Single) As GpStatus
Public Declare Function GdipCloneBitmapAreaI Lib "gdiplus" (ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal PixelFormat As Long, ByVal SrcBitmap As Long, DstBitmap As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipCreateCachedBitmap Lib "gdiplus" (ByVal Bitmap As Long, ByVal Graphics As Long, CachedBitmap As Long) As GpStatus
Public Declare Function GdipDeleteCachedBitmap Lib "gdiplus" (ByVal CachedBitmap As Long) As GpStatus
Public Declare Function GdipDrawCachedBitmap Lib "gdiplus" (ByVal Graphics As Long, ByVal CachedBitmap As Long, ByVal X As Long, ByVal Y As Long) As GpStatus

'21、画刷、画刷阴影、填充
Public Declare Function GdipCloneBrush Lib "gdiplus" (ByVal Brush As Long, CloneBrush As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As GpStatus
Public Declare Function GdipGetBrushType Lib "gdiplus" (ByVal Brush As Long, BrshType As BrushType) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipCreateLineBrushI Lib "gdiplus" (Point1 As PointL, Point2 As PointL, ByVal Color1 As Long, ByVal Color2 As Long, ByVal WrapMd As WrapMode, LineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectI Lib "gdiplus" (Rect As RectL, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Mode As LinearGradientMode, ByVal WrapMd As WrapMode, LineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "gdiplus" (Rect As RectL, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Angle As Single, ByVal IsAngleScalable As Long, ByVal WrapMd As WrapMode, LineGradient As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipCreateHatchBrush Lib "gdiplus" (ByVal Style As HatchStyle, ByVal mForeColor As Long, ByVal mBackColor As Long, Brush As Long) As GpStatus
Public Declare Function GdipGetHatchStyle Lib "gdiplus" (ByVal Brush As Long, Style As HatchStyle) As GpStatus
Public Declare Function GdipGetHatchForegroundColor Lib "gdiplus" (ByVal Brush As Long, mForeColor As Long) As GpStatus
Public Declare Function GdipGetHatchBackgroundColor Lib "gdiplus" (ByVal Brush As Long, mBackColor As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal Argb As Long, Brush As Long) As GpStatus
Public Declare Function GdipSetSolidFillColor Lib "gdiplus" (ByVal Brush As Long, ByVal Argb As Long) As GpStatus
Public Declare Function GdipGetSolidFillColor Lib "gdiplus" (ByVal Brush As Long, Argb As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipSetLineColors Lib "gdiplus" (ByVal Brush As Long, ByVal Color1 As Long, ByVal Color2 As Long) As GpStatus
Public Declare Function GdipGetLineColors Lib "gdiplus" (ByVal Brush As Long, lColors As Long) As GpStatus
Public Declare Function GdipGetLineRectI Lib "gdiplus" (ByVal Brush As Long, Rect As RectL) As GpStatus
Public Declare Function GdipSetLineGammaCorrection Lib "gdiplus" (ByVal Brush As Long, ByVal UseGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetLineGammaCorrection Lib "gdiplus" (ByVal Brush As Long, UseGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetLineBlendCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipGetLineBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Any, Positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLineBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Any, Positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlendCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Any, Positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLinePresetBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Any, Positions As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLineSigmaBlend Lib "gdiplus" (ByVal Brush As Long, ByVal Focus As Single, ByVal TheScale As Single) As GpStatus
Public Declare Function GdipSetLineLinearBlend Lib "gdiplus" (ByVal Brush As Long, ByVal Focus As Single, ByVal TheScale As Single) As GpStatus
Public Declare Function GdipSetLineWrapMode Lib "gdiplus" (ByVal Brush As Long, ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetLineWrapMode Lib "gdiplus" (ByVal Brush As Long, WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetLineTransform Lib "gdiplus" (ByVal Brush As Long, Matrix As Long) As GpStatus
Public Declare Function GdipSetLineTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetLineTransform Lib "gdiplus" (ByVal Brush As Long) As GpStatus
Public Declare Function GdipMultiplyLineTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateLineTransform Lib "gdiplus" (ByVal Brush As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleLineTransform Lib "gdiplus" (ByVal Brush As Long, ByVal sX As Single, ByVal sY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateLineTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus

'22、贴图相关
Public Declare Function GdipCreateTexture Lib "gdiplus" (ByVal Image As Long, ByVal WrapMd As WrapMode, Texture As Long) As GpStatus
Public Declare Function GdipCreateTexture2 Lib "gdiplus" (ByVal Image As Long, ByVal WrapMd As WrapMode, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, Texture As Long) As GpStatus
Public Declare Function GdipCreateTextureIA Lib "gdiplus" (ByVal Image As Long, ByVal ImageAttributes As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, Texture As Long) As GpStatus
Public Declare Function GdipCreateTexture2I Lib "gdiplus" (ByVal Image As Long, ByVal WrapMd As WrapMode, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Texture As Long) As GpStatus
Public Declare Function GdipCreateTextureIAI Lib "gdiplus" (ByVal Image As Long, ByVal ImageAttributes As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Texture As Long) As GpStatus
Public Declare Function GdipGetTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipSetTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetTextureTransform Lib "gdiplus" (ByVal Brush As Long) As GpStatus
Public Declare Function GdipTranslateTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipMultiplyTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal sX As Single, ByVal sY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateTextureTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipSetTextureWrapMode Lib "gdiplus" (ByVal Brush As Long, ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetTextureWrapMode Lib "gdiplus" (ByVal Brush As Long, WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetTextureImage Lib "gdiplus" (ByVal Brush As Long, Image As Long) As GpStatus

'23、路径操作
Public Declare Function GdipCreatePathGradientI Lib "gdiplus" (Points As PointL, ByVal Count As Long, ByVal WrapMd As WrapMode, PolyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientFromPath Lib "gdiplus" (ByVal Path As Long, PolyGradient As Long) As GpStatus
Public Declare Function GdipGetPathGradientCenterColor Lib "gdiplus" (ByVal Brush As Long, lColors As Long) As GpStatus
Public Declare Function GdipSetPathGradientCenterColor Lib "gdiplus" (ByVal Brush As Long, ByVal lColors As Long) As GpStatus
Public Declare Function GdipGetPathGradientSurroundColorsWithCount Lib "gdiplus" (ByVal Brush As Long, Argb As Long, Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "gdiplus" (ByVal Brush As Long, Argb As Long, Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPath Lib "gdiplus" (ByVal Brush As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipSetPathGradientPath Lib "gdiplus" (ByVal Brush As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipGetPathGradientCenterPointI Lib "gdiplus" (ByVal Brush As Long, Points As PointL) As GpStatus
Public Declare Function GdipSetPathGradientCenterPointI Lib "gdiplus" (ByVal Brush As Long, Points As PointL) As GpStatus
Public Declare Function GdipGetPathGradientRectI Lib "gdiplus" (ByVal Brush As Long, Rect As RectL) As GpStatus
Public Declare Function GdipGetPathGradientPointCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientSurroundColorCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientGammaCorrection Lib "gdiplus" (ByVal Brush As Long, ByVal UseGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetPathGradientGammaCorrection Lib "gdiplus" (ByVal Brush As Long, UseGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlendCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Single, Positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Single, Positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlendCount Lib "gdiplus" (ByVal Brush As Long, Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Long, Positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientPresetBlend Lib "gdiplus" (ByVal Brush As Long, Blend As Long, Positions As Single, ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientSigmaBlend Lib "gdiplus" (ByVal Brush As Long, ByVal Focus As Single, ByVal sScale As Single) As GpStatus
Public Declare Function GdipSetPathGradientLinearBlend Lib "gdiplus" (ByVal Brush As Long, ByVal Focus As Single, ByVal sScale As Single) As GpStatus
Public Declare Function GdipGetPathGradientWrapMode Lib "gdiplus" (ByVal Brush As Long, WrapMd As WrapMode) As GpStatus
Public Declare Function GdipSetPathGradientWrapMode Lib "gdiplus" (ByVal Brush As Long, ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetPathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipSetPathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipResetPathGradientTransform Lib "gdiplus" (ByVal Brush As Long) As GpStatus
Public Declare Function GdipMultiplyPathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Matrix As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslatePathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScalePathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal sX As Single, ByVal sY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotatePathGradientTransform Lib "gdiplus" (ByVal Brush As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipGetPathGradientFocusScales Lib "gdiplus" (ByVal Brush As Long, xScale As Single, yScale As Single) As GpStatus
Public Declare Function GdipSetPathGradientFocusScales Lib "gdiplus" (ByVal Brush As Long, ByVal xScale As Single, ByVal yScale As Single) As GpStatus
Public Declare Function GdipCreatePath Lib "gdiplus" (ByVal BrushMode As FillMode, Path As Long) As GpStatus
Public Declare Function GdipCreatePath2I Lib "gdiplus" (Points As PointL, Types As Any, ByVal Count As Long, BrushMode As FillMode, Path As Long) As GpStatus
Public Declare Function GdipClonePath Lib "gdiplus" (ByVal Path As Long, ClonePath As Long) As GpStatus
Public Declare Function GdipDeletePath Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipResetPath Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipGetPointCount Lib "gdiplus" (ByVal Path As Long, Count As Long) As GpStatus
Public Declare Function GdipGetPathTypes Lib "gdiplus" (ByVal Path As Long, Types As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPointsI Lib "gdiplus" (ByVal Path As Long, Points As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathFillMode Lib "gdiplus" (ByVal Path As Long, ByVal BrushMode As FillMode) As GpStatus
Public Declare Function GdipSetPathFillMode Lib "gdiplus" (ByVal Path As Long, ByVal BrushMode As FillMode) As GpStatus
Public Declare Function GdipGetPathData Lib "gdiplus" (ByVal Path As Long, pdata As PathData) As GpStatus
Public Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipClosePathFigures Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipSetPathMarker Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipClearPathMarkers Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipReversePath Lib "gdiplus" (ByVal Path As Long) As GpStatus
Public Declare Function GdipGetPathLastPoint Lib "gdiplus" (ByVal Path As Long, lastPoint As PointL) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipAddPathLine Lib "gdiplus" (ByVal Path As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single) As GpStatus
Public Declare Function GdipAddPathLine2 Lib "gdiplus" (ByVal Path As Long, Points As PointF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathArc Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal StartAngle As Single, ByVal SweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathBezier Lib "gdiplus" (ByVal Path As Long, ByVal X1 As Single, ByVal Y1 As Single, ByVal X2 As Single, ByVal Y2 As Single, ByVal X3 As Single, ByVal Y3 As Single, ByVal X4 As Single, ByVal Y4 As Single) As GpStatus
Public Declare Function GdipAddPathBeziers Lib "gdiplus" (ByVal Path As Long, Points As PointF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve Lib "gdiplus" (ByVal Path As Long, Points As PointF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2 Lib "gdiplus" (ByVal Path As Long, Points As PointF, ByVal Count As Long, ByVal Tension As Single) As GpStatus
Public Declare Function GdipAddPathCurve3 Lib "gdiplus" (ByVal Path As Long, Points As PointF, ByVal Count As Long, ByVal Offset As Long, ByVal NumberOfSegments As Long, ByVal Tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurve Lib "gdiplus" (ByVal Path As Long, Points As PointF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2 Lib "gdiplus" (ByVal Path As Long, Points As PointF, ByVal Count As Long, ByVal Tension As Single) As GpStatus
Public Declare Function GdipAddPathRectangle Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipAddPathRectangles Lib "gdiplus" (ByVal Path As Long, Rect As RectF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Public Declare Function GdipAddPathPie Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal StartAngle As Single, ByVal SweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathPolygon Lib "gdiplus" (ByVal Path As Long, Points As PointF, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathPath Lib "gdiplus" (ByVal Path As Long, ByVal addingPath As Long, ByVal bConnect As Long) As GpStatus
Public Declare Function GdipAddPathStringI Lib "gdiplus" (ByVal Path As Long, ByVal mText As Long, ByVal Length As Long, ByVal Family As Long, ByVal Style As Long, ByVal emSize As Single, LayoutRect As RectL, ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipAddPathLineI Lib "gdiplus" (ByVal Path As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As GpStatus
Public Declare Function GdipAddPathLine2I Lib "gdiplus" (ByVal Path As Long, Points As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathArcI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal StartAngle As Single, ByVal SweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathBezierI Lib "gdiplus" (ByVal Path As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As GpStatus
Public Declare Function GdipAddPathBeziersI Lib "gdiplus" (ByVal Path As Long, Points As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurveI Lib "gdiplus" (ByVal Path As Long, Points As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2I Lib "gdiplus" (ByVal Path As Long, Points As PointL, ByVal Count As Long, ByVal Tension As Long) As GpStatus
Public Declare Function GdipAddPathCurve3I Lib "gdiplus" (ByVal Path As Long, Points As PointL, ByVal Count As Long, ByVal Offset As Long, ByVal NumberOfSegments As Long, ByVal Tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurveI Lib "gdiplus" (ByVal Path As Long, Points As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2I Lib "gdiplus" (ByVal Path As Long, Points As PointL, ByVal Count As Long, ByVal Tension As Single) As GpStatus
Public Declare Function GdipAddPathRectangleI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathRectanglesI Lib "gdiplus" (ByVal Path As Long, Rects As RectL, ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathEllipseI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathPieI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal StartAngle As Single, ByVal SweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathPolygonI Lib "gdiplus" (ByVal Path As Long, Points As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipFlattenPath Lib "gdiplus" (ByVal Path As Long, Optional ByVal Matrix As Long = 0, Optional ByVal Flatness As Single = 0.25) As GpStatus
Public Declare Function GdipWindingModeOutline Lib "gdiplus" (ByVal Path As Long, ByVal Matrix As Long, ByVal Flatness As Single) As GpStatus
Public Declare Function GdipWidenPath Lib "gdiplus" (ByVal NativePath As Long, ByVal Pen As Long, ByVal Matrix As Long, ByVal Flatness As Single) As GpStatus
Public Declare Function GdipWarpPath Lib "gdiplus" (ByVal Path As Long, ByVal Matrix As Long, Points As PointF, ByVal Count As Long, ByVal SrcX As Single, ByVal SrcY As Single, ByVal SrcWidth As Single, ByVal SrcHeight As Single, ByVal WarpMd As WarpMode, ByVal Flatness As Single) As GpStatus
Public Declare Function GdipTransformPath Lib "gdiplus" (ByVal Path As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipGetPathWorldBounds Lib "gdiplus" (ByVal Path As Long, Bounds As RectF, ByVal Matrix As Long, ByVal Pen As Long) As GpStatus
Public Declare Function GdipGetPathWorldBoundsI Lib "gdiplus" (ByVal Path As Long, Bounds As RectL, ByVal Matrix As Long, ByVal Pen As Long) As GpStatus
Public Declare Function GdipIsVisiblePathPoint Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsVisiblePathPointI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsOutlineVisiblePathPoint Lib "gdiplus" (ByVal Path As Long, ByVal X As Single, ByVal Y As Single, ByVal Pen As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsOutlineVisiblePathPointI Lib "gdiplus" (ByVal Path As Long, ByVal X As Long, ByVal Y As Long, ByVal Pen As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipCreatePathIter Lib "gdiplus" (Iterator As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipDeletePathIter Lib "gdiplus" (ByVal Iterator As Long) As GpStatus
Public Declare Function GdipPathIterNextSubpath Lib "gdiplus" (ByVal Iterator As Long, ResultCount As Long, StartIndex As Long, EndIndex As Long, IsClosed As Long) As GpStatus
Public Declare Function GdipPathIterNextSubpathPath Lib "gdiplus" (ByVal Iterator As Long, ResultCount As Long, ByVal Path As Long, IsClosed As Long) As GpStatus
Public Declare Function GdipPathIterNextPathType Lib "gdiplus" (ByVal Iterator As Long, ResultCount As Long, PathType As Any, StartIndex As Long, EndIndex As Long) As GpStatus
Public Declare Function GdipPathIterNextMarker Lib "gdiplus" (ByVal Iterator As Long, ResultCount As Long, StartIndex As Long, EndIndex As Long) As GpStatus
Public Declare Function GdipPathIterNextMarkerPath Lib "gdiplus" (ByVal Iterator As Long, ResultCount As Long, ByVal Path As Long) As GpStatus
Public Declare Function GdipPathIterGetCount Lib "gdiplus" (ByVal Iterator As Long, Count As Long) As GpStatus
Public Declare Function GdipPathIterGetSubpathCount Lib "gdiplus" (ByVal Iterator As Long, Count As Long) As GpStatus
Public Declare Function GdipPathIterIsValid Lib "gdiplus" (ByVal Iterator As Long, Valid As Long) As GpStatus
Public Declare Function GdipPathIterHasCurve Lib "gdiplus" (ByVal Iterator As Long, HasCurve As Long) As GpStatus
Public Declare Function GdipPathIterRewind Lib "gdiplus" (ByVal Iterator As Long) As GpStatus
Public Declare Function GdipPathIterEnumerate Lib "gdiplus" (ByVal Iterator As Long, ResultCount As Long, Points As PointF, Types As Any, ByVal Count As Long) As GpStatus
Public Declare Function GdipPathIterCopyData Lib "gdiplus" (ByVal Iterator As Long, ResultCount As Long, Points As PointF, Types As Any, ByVal StartIndex As Long, ByVal EndIndex As Long) As GpStatus

'24、矩阵
Public Declare Function GdipCreateMatrix Lib "gdiplus" (Matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix2 Lib "gdiplus" (ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dX As Single, ByVal dY As Single, Matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix3I Lib "gdiplus" (Rect As RectL, dstplg As PointL, Matrix As Long) As GpStatus
Public Declare Function GdipCloneMatrix Lib "gdiplus" (ByVal Matrix As Long, cloneMatrix As Long) As GpStatus
Public Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal Matrix As Long) As GpStatus
Public Declare Function GdipSetMatrixElements Lib "gdiplus" (ByVal Matrix As Long, ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dX As Single, ByVal dY As Single) As GpStatus
Public Declare Function GdipMultiplyMatrix Lib "gdiplus" (ByVal Matrix As Long, ByVal Matrix2 As Long, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal Matrix As Long, ByVal OffsetX As Single, ByVal OffsetY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleMatrix Lib "gdiplus" (ByVal Matrix As Long, ByVal ScaleX As Single, ByVal ScaleY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal Matrix As Long, ByVal Angle As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipShearMatrix Lib "gdiplus" (ByVal Matrix As Long, ByVal ShearX As Single, ByVal ShearY As Single, ByVal Order As MatrixOrder) As GpStatus
Public Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal Matrix As Long) As GpStatus
Public Declare Function GdipTransformMatrixPointsI Lib "gdiplus" (ByVal Matrix As Long, pts As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPointsI Lib "gdiplus" (ByVal Matrix As Long, pts As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipGetMatrixElements Lib "gdiplus" (ByVal Matrix As Long, MatrixOut As Single) As GpStatus
Public Declare Function GdipIsMatrixInvertible Lib "gdiplus" (ByVal Matrix As Long, Result As Long) As GpStatus
Public Declare Function GdipIsMatrixIdentity Lib "gdiplus" (ByVal Matrix As Long, Result As Long) As GpStatus
Public Declare Function GdipIsMatrixEqual Lib "gdiplus" (ByVal Matrix As Long, ByVal Matrix2 As Long, Result As Long) As GpStatus

'25、区域
Public Declare Function GdipCreateRegion Lib "gdiplus" (Region As Long) As GpStatus
Public Declare Function GdipCreateRegionRectI Lib "gdiplus" (Rect As RectL, Region As Long) As GpStatus
Public Declare Function GdipCreateRegionPath Lib "gdiplus" (ByVal Path As Long, Region As Long) As GpStatus
Public Declare Function GdipCreateRegionRgnData Lib "gdiplus" (regionData As Any, ByVal Size As Long, Region As Long) As GpStatus
Public Declare Function GdipCreateRegionHrgn Lib "gdiplus" (ByVal hRgn As Long, Region As Long) As GpStatus
Public Declare Function GdipCloneRegion Lib "gdiplus" (ByVal Region As Long, CloneRegion As Long) As GpStatus
Public Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal Region As Long) As GpStatus
Public Declare Function GdipSetInfinite Lib "gdiplus" (ByVal Region As Long) As GpStatus
Public Declare Function GdipSetEmpty Lib "gdiplus" (ByVal Region As Long) As GpStatus
Public Declare Function GdipCombineRegionRectI Lib "gdiplus" (ByVal Region As Long, Rect As RectL, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionPath Lib "gdiplus" (ByVal Region As Long, ByVal Path As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionRegion Lib "gdiplus" (ByVal Region As Long, ByVal Region2 As Long, ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipTranslateRegionI Lib "gdiplus" (ByVal Region As Long, ByVal dX As Long, ByVal dY As Long) As GpStatus
Public Declare Function GdipTransformRegion Lib "gdiplus" (ByVal Region As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipGetRegionBoundsI Lib "gdiplus" (ByVal Region As Long, ByVal Graphics As Long, Rect As RectL) As GpStatus
Public Declare Function GdipGetRegionHRgn Lib "gdiplus" (ByVal Region As Long, ByVal Graphics As Long, hRgn As Long) As GpStatus
Public Declare Function GdipIsEmptyRegion Lib "gdiplus" (ByVal Region As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsInfiniteRegion Lib "gdiplus" (ByVal Region As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsEqualRegion Lib "gdiplus" (ByVal Region As Long, ByVal Region2 As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipGetRegionDataSize Lib "gdiplus" (ByVal Region As Long, BufferSize As Long) As GpStatus
Public Declare Function GdipGetRegionData Lib "gdiplus" (ByVal Region As Long, Buffer As Any, ByVal BufferSize As Long, SizeFilled As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionPointI Lib "gdiplus" (ByVal Region As Long, ByVal X As Long, ByVal Y As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionRectI Lib "gdiplus" (ByVal Region As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Graphics As Long, Result As Long) As GpStatus
Public Declare Function GdipGetRegionScansCount Lib "gdiplus" (ByVal Region As Long, Ucount As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipGetRegionScansI Lib "gdiplus" (ByVal Region As Long, Rects As RectL, Count As Long, ByVal Matrix As Long) As GpStatus

'26、图像属性
Public Declare Function GdipCreateImageAttributes Lib "gdiplus" (ImageAttr As Long) As GpStatus
Public Declare Function GdipCloneImageAttributes Lib "gdiplus" (ByVal ImageAttr As Long, CloneImageattr As Long) As GpStatus
Public Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal ImageAttr As Long) As GpStatus
Public Declare Function GdipSetImageAttributesToIdentity Lib "gdiplus" (ByVal ImageAttr As Long, ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipResetImageAttributes Lib "gdiplus" (ByVal ImageAttr As Long, ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal ImageAttr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal EnableFlag As Long, ColourMatrix As Any, GrayMatrix As Any, ByVal Flags As ColorMatrixFlags) As GpStatus
Public Declare Function GdipSetImageAttributesThreshold Lib "gdiplus" (ByVal ImageAttr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal EnableFlag As Long, ByVal Threshold As Single) As GpStatus
Public Declare Function GdipSetImageAttributesGamma Lib "gdiplus" (ByVal ImageAttr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal EnableFlag As Long, ByVal Gamma As Single) As GpStatus
Public Declare Function GdipSetImageAttributesNoOp Lib "gdiplus" (ByVal ImageAttr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal EnableFlag As Long) As GpStatus
Public Declare Function GdipSetImageAttributesColorKeys Lib "gdiplus" (ByVal ImageAttr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal EnableFlag As Long, ByVal ColorLow As Long, ByVal ColorHigh As Long) As GpStatus
Public Declare Function GdipSetImageAttributesOutputChannel Lib "gdiplus" (ByVal ImageAttr As Long, ByVal ClrAdjstType As ColorAdjustType, ByVal EnableFlag As Long, ByVal ChannelFlags As ColorChannelFlags) As GpStatus
Public Declare Function GdipSetImageAttributesOutputChannelColorProfile Lib "gdiplus" (ByVal ImageAttr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal EnableFlag As Long, ByVal ColorProfileFilename As Long) As GpStatus
Public Declare Function GdipSetImageAttributesRemapTable Lib "gdiplus" (ByVal ImageAttr As Long, ByVal ClrAdjType As ColorAdjustType, ByVal EnableFlag As Long, ByVal MapSize As Long, Map As Any) As GpStatus
Public Declare Function GdipSetImageAttributesWrapMode Lib "gdiplus" (ByVal ImageAttr As Long, ByVal Wrap As WrapMode, ByVal Argb As Long, ByVal bClamp As Long) As GpStatus
Public Declare Function GdipSetImageAttributesICMMode Lib "gdiplus" (ByVal ImageAttr As Long, ByVal bOn As Long) As GpStatus
Public Declare Function GdipGetImageAttributesAdjustedPalette Lib "gdiplus" (ByVal ImageAttr As Long, ColorPal As ColorPalette, ByVal ClrAdjType As ColorAdjustType) As GpStatus

'27、字符集、逻辑字体
Public Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal FontCollection As Long, FontFamily As Long) As GpStatus
Public Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal FontFamily As Long) As GpStatus
Public Declare Function GdipCloneFontFamily Lib "gdiplus" (ByVal FontFamily As Long, clonedFontFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySansSerif Lib "gdiplus" (NativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySerif Lib "gdiplus" (NativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilyMonospace Lib "gdiplus" (NativeFamily As Long) As GpStatus
Public Declare Function GdipGetFamilyName Lib "gdiplus" (ByVal Family As Long, ByVal Name As Long, ByVal language As Integer) As GpStatus
Public Declare Function GdipIsStyleAvailable Lib "gdiplus" (ByVal Family As Long, ByVal Style As Long, IsStyleAvailable As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerable Lib "gdiplus" (ByVal FontCollection As Long, ByVal Graphics As Long, NumFound As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerate Lib "gdiplus" (ByVal FontCollection As Long, ByVal NumSought As Long, gpFamilies As Long, ByVal NumFound As Long, ByVal Graphics As Long) As GpStatus
Public Declare Function GdipGetEmHeight Lib "gdiplus" (ByVal Family As Long, ByVal Style As Long, EmHeight As Integer) As GpStatus
Public Declare Function GdipGetCellAscent Lib "gdiplus" (ByVal Family As Long, ByVal Style As Long, CellAscent As Integer) As GpStatus
Public Declare Function GdipGetCellDescent Lib "gdiplus" (ByVal Family As Long, ByVal Style As Long, CellDescent As Integer) As GpStatus
Public Declare Function GdipGetLineSpacing Lib "gdiplus" (ByVal Family As Long, ByVal Style As Long, LineSpacing As Integer) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipCreateFontFromDC Lib "gdiplus" (ByVal hDC As Long, CreatedFont As Long) As GpStatus
Public Declare Function GdipCreateFontFromLogfontA Lib "gdiplus" (ByVal hDC As Long, LogFont As LogFontA, CreatedFont As Long) As GpStatus
Public Declare Function GdipCreateFontFromLogfontW Lib "gdiplus" (ByVal hDC As Long, LogFont As LogFontW, CreatedFont As Long) As GpStatus
Public Declare Function GdipCreateFont Lib "gdiplus" (ByVal FontFamily As Long, ByVal emSize As Single, ByVal Style As Long, ByVal Unit As GpUnit, CreatedFont As Long) As GpStatus
Public Declare Function GdipCloneFont Lib "gdiplus" (ByVal curFont As Long, CloneFont As Long) As GpStatus
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As GpStatus
Public Declare Function GdipGetFamily Lib "gdiplus" (ByVal curFont As Long, Family As Long) As GpStatus
Public Declare Function GdipGetFontStyle Lib "gdiplus" (ByVal curFont As Long, Style As Long) As GpStatus
Public Declare Function GdipGetFontSize Lib "gdiplus" (ByVal curFont As Long, Size As Single) As GpStatus
Public Declare Function GdipGetFontUnit Lib "gdiplus" (ByVal curFont As Long, Unit As GpUnit) As GpStatus
Public Declare Function GdipGetFontHeight Lib "gdiplus" (ByVal curFont As Long, ByVal Graphics As Long, Height As Single) As GpStatus
Public Declare Function GdipGetFontHeightGivenDPI Lib "gdiplus" (ByVal curFont As Long, ByVal DPI As Single, Height As Single) As GpStatus
Public Declare Function GdipGetLogFontA Lib "gdiplus" (ByVal curFont As Long, ByVal Graphics As Long, LogFont As LogFontA) As GpStatus
Public Declare Function GdipGetLogFontW Lib "gdiplus" (ByVal curFont As Long, ByVal Graphics As Long, LogFont As LogFontW) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipNewInstalledFontCollection Lib "gdiplus" (FontCollection As Long) As GpStatus
Public Declare Function GdipNewPrivateFontCollection Lib "gdiplus" (FontCollection As Long) As GpStatus
Public Declare Function GdipDeletePrivateFontCollection Lib "gdiplus" (FontCollection As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyCount Lib "gdiplus" (ByVal FontCollection As Long, NumFound As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyList Lib "gdiplus" (ByVal FontCollection As Long, ByVal NumSought As Long, gpFamilies As Long, NumFound As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipPrivateAddFontFile Lib "gdiplus" (ByVal FontCollection As Long, ByVal FileName As Long) As GpStatus
Public Declare Function GdipPrivateAddMemoryFont Lib "gdiplus" (ByVal FontCollection As Long, ByVal Memory As Long, ByVal Length As Long) As GpStatus

'28、字符串、字符串格式
Public Declare Function GdipDrawString Lib "gdiplus" (ByVal Graphics As Long, ByVal mText As Long, ByVal Length As Long, ByVal TheFont As Long, LayoutRect As RectL, ByVal StringFormat As Long, ByVal Brush As Long) As GpStatus
Public Declare Function GdipMeasureString Lib "gdiplus" (ByVal Graphics As Long, ByVal mText As Long, ByVal Length As Long, ByVal TheFont As Long, LayoutRect As RectF, ByVal StringFormat As Long, BoundingBox As RectF, CodePointsFitted As Long, LinesFilled As Long) As GpStatus
Public Declare Function GdipMeasureCharacterRanges Lib "gdiplus" (ByVal Graphics As Long, ByVal mText As Long, ByVal Length As Long, ByVal TheFont As Long, LayoutRect As RectL, ByVal StringFormat As Long, ByVal regionCount As Long, Regions As Long) As GpStatus
Public Declare Function GdipDrawDriverString Lib "gdiplus" (ByVal Graphics As Long, ByVal mText As Long, ByVal Length As Long, ByVal TheFont As Long, ByVal Brush As Long, Positions As PointL, ByVal Flags As Long, ByVal Matrix As Long) As GpStatus
Public Declare Function GdipMeasureDriverString Lib "gdiplus" (ByVal Graphics As Long, ByVal mText As Long, ByVal Length As Long, ByVal TheFont As Long, Positions As PointL, ByVal Flags As Long, ByVal Matrix As Long, BoundingBox As RectL) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericDefault Lib "gdiplus" (StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericTypographic Lib "gdiplus" (StringFormat As Long) As GpStatus
Public Declare Function GdipDeleteStringFormat Lib "gdiplus" (ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipCloneStringFormat Lib "gdiplus" (ByVal StringFormat As Long, newFormat As Long) As GpStatus
Public Declare Function GdipSetStringFormatFlags Lib "gdiplus" (ByVal StringFormat As Long, ByVal Flags As Long) As GpStatus
Public Declare Function GdipGetStringFormatFlags Lib "gdiplus" (ByVal StringFormat As Long, Flags As Long) As GpStatus
Public Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, Align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatLineAlign Lib "gdiplus" (ByVal StringFormat As Long, Align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatTrimming Lib "gdiplus" (ByVal StringFormat As Long, ByVal trimming As StringTrimming) As GpStatus
Public Declare Function GdipGetStringFormatTrimming Lib "gdiplus" (ByVal StringFormat As Long, trimming As Long) As GpStatus
Public Declare Function GdipSetStringFormatHotkeyPrefix Lib "gdiplus" (ByVal StringFormat As Long, ByVal hkPrefix As HotkeyPrefix) As GpStatus
Public Declare Function GdipGetStringFormatHotkeyPrefix Lib "gdiplus" (ByVal StringFormat As Long, hkPrefix As HotkeyPrefix) As GpStatus
Public Declare Function GdipSetStringFormatTabStops Lib "gdiplus" (ByVal StringFormat As Long, ByVal firstTabOffset As Single, ByVal Count As Long, tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStops Lib "gdiplus" (ByVal StringFormat As Long, ByVal Count As Long, firstTabOffset As Single, tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStopCount Lib "gdiplus" (ByVal StringFormat As Long, Count As Long) As GpStatus
Public Declare Function GdipSetStringFormatDigitSubstitution Lib "gdiplus" (ByVal StringFormat As Long, ByVal language As Integer, ByVal substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatDigitSubstitution Lib "gdiplus" (ByVal StringFormat As Long, language As Integer, substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatMeasurableCharacterRangeCount Lib "gdiplus" (ByVal StringFormat As Long, Count As Long) As GpStatus
Public Declare Function GdipSetStringFormatMeasurableCharacterRanges Lib "gdiplus" (ByVal StringFormat As Long, ByVal rangeCount As Long, ranges As CharacterRange) As GpStatus

'29、图元文件
Public Declare Function GdipEnumerateMetafileDestPointI Lib "gdiplus" (Graphics As Long, ByVal MetaFile As Long, destPoint As PointL, ByVal lpEnumerateMetafileProc As Long, ByVal CallBackData As Long, ByVal ImageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal MetaFile As Long, destRect As RectL, lpEnumerateMetafileProc As Long, ByVal CallBackData As Long, ImageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPointsI Lib "gdiplus" (ByVal Graphics As Long, ByVal MetaFile As Long, destPoint As PointL, ByVal Count As Long, lpEnumerateMetafileProc As Long, ByVal CallBackData As Long, ImageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointI Lib "gdiplus" (ByVal Graphics As Long, ByVal MetaFile As Long, destPoint As PointL, srcRect As RectL, ByVal SrcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallBackData As Long, ByVal ImageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal MetaFile As Long, destRect As RectL, srcRect As RectL, ByVal SrcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallBackData As Long, ByVal ImageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointsI Lib "gdiplus" (ByVal Graphics As Long, ByVal MetaFile As Long, DestPoints As PointL, ByVal Count As Long, srcRect As RectL, ByVal SrcUnit As GpUnit, ByVal lpEnumerateMetafileProc As Long, ByVal CallBackData As Long, ByVal ImageAttributes As Long) As GpStatus
Public Declare Function GdipPlayMetafileRecord Lib "gdiplus" (ByVal MetaFile As Long, ByVal recordType As EmfPlusRecordType, ByVal Flags As Long, ByVal dataSize As Long, byteData As Any) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipGetMetafileHeaderFromWmf Lib "gdiplus" (ByVal hWmf As Long, WmfPlaceableFileHdr As WmfPlaceableFileHeader, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromEmf Lib "gdiplus" (ByVal hEmf As Long, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromFile Lib "gdiplus" (ByVal FileName As Long, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromStream Lib "gdiplus" (ByVal Stream As Any, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromMetafile Lib "gdiplus" (ByVal MetaFile As Long, header As MetafileHeader) As GpStatus
Public Declare Function GdipGetHemfFromMetafile Lib "gdiplus" (ByVal MetaFile As Long, hEmf As Long) As GpStatus
Public Declare Function GdipCreateStreamOnFile Lib "gdiplus" (ByVal FileName As Long, ByVal access As Long, Stream As Any) As GpStatus
Public Declare Function GdipCreateMetafileFromWmf Lib "gdiplus" (ByVal hWmf As Long, ByVal bDeleteWmf As Long, WmfPlaceableFileHdr As WmfPlaceableFileHeader, ByVal MetaFile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromEmf Lib "gdiplus" (ByVal hEmf As Long, ByVal bDeleteEmf As Long, MetaFile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromFile Lib "gdiplus" (ByVal File As Long, MetaFile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromWmfFile Lib "gdiplus" (ByVal File As Long, WmfPlaceableFileHdr As WmfPlaceableFileHeader, MetaFile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromStream Lib "gdiplus" (ByVal Stream As Any, MetaFile As Long) As GpStatus
Public Declare Function GdipRecordMetafileI Lib "gdiplus" (ByVal referenceHdc As Long, etype As EmfType, frameRect As RectL, ByVal frameUnit As MetafileFrameUnit, ByVal description As Long, MetaFile As Long) As GpStatus
Public Declare Function GdipRecordMetafileFileNameI Lib "gdiplus" (ByVal FileName As Long, ByVal referenceHdc As Long, etype As EmfType, frameRect As RectL, ByVal frameUnit As MetafileFrameUnit, ByVal description As Long, MetaFile As Long) As GpStatus
Public Declare Function GdipRecordMetafileStreamI Lib "gdiplus" (ByVal Stream As Any, ByVal referenceHdc As Long, etype As EmfType, frameRect As RectL, ByVal frameUnit As MetafileFrameUnit, ByVal description As Long, MetaFile As Long) As GpStatus
Public Declare Function GdipSetMetafileDownLevelRasterizationLimit Lib "gdiplus" (ByVal MetaFile As Long, ByVal metafileRasterizationLimitDpi As Long) As GpStatus
Public Declare Function GdipGetMetafileDownLevelRasterizationLimit Lib "gdiplus" (ByVal MetaFile As Long, metafileRasterizationLimitDpi As Long) As GpStatus

'30、其他内容
Public Declare Function GdipComment Lib "gdiplus" (ByVal Graphics As Long, ByVal sizeData As Long, Data As Any) As GpStatus
Public Declare Function GdipFlush Lib "gdiplus" (ByVal Graphics As Long, ByVal Intention As FlushIntention) As GpStatus
Public Declare Function GdipAlloc Lib "gdiplus" (ByVal Size As Long) As Long
Public Declare Sub GdipFree Lib "gdiplus" (ByVal Ptr As Long)
Public Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As GpStatus
Private Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As Long, pCLSID As ClsID) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'##################  常  量  ##################
'1、图像解码器
Public Const ImageEncoderSuffix       As String = "-1A04-11D3-9A73-0000F81EF32E}"
Public Const ImageEncoderPrefix       As String = "{557CF40"
Public Const EncoderCompression       As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Public Const EncoderColorDepth        As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Public Const EncoderScanMethod        As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Public Const EncoderVersion           As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Public Const EncoderRenderMethod      As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Public Const EncoderQuality           As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Public Const EncoderTransformation    As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Public Const EncoderLuminanceTable    As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Public Const EncoderChrominanceTable  As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Public Const EncoderSaveFlag          As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
'2、DC相关
Private Const HORZRES As Long = 8&
Private Const VERTRES As Long = 10&
'##################  变  量  ##################

Private mToken As Long                                                          '口令（作为GDIP是否被初始化的依据）
Private Objects() As GdiplusObject, ObjCount As Long                            'Gdip通用对象

'##################  过  程  /  函  数  ##################

'取二者中的较大值、较小值
Private Function Max(ByVal A As Long, ByVal B As Long) As Long
    Max = IIf(A > B, A, B)
End Function

Private Function Min(ByVal A As Long, ByVal B As Long) As Long
    Min = IIf(A < B, A, B)
End Function

'根据文件后缀获得对应编码器类别标识符
Public Function GetImageEncoderClsid(ByVal FileSuffix As ImageFileSuffix) As ClsID
    CLSIDFromString StrPtr(ImageEncoderPrefix & CInt(FileSuffix) & ImageEncoderSuffix), GetImageEncoderClsid
End Function

'初始化GDIPlus
Public Sub InitGDIPlus(Optional ByVal ShowLog As Boolean = False)
    Dim uInput As GdiplusStartupInput, Retn As GpStatus
    If mToken <> 0 Then
        If ShowLog Then Debug.Print "GdiPlus已初始化。"
        Exit Sub
    End If
    uInput.GdiplusVersion = 1
    Retn = GdiplusStartup(mToken, uInput)
    If Retn <> Ok Then
        If ShowLog Then Debug.Print "GdiPlus未能成功初始化。错误原因：" & _
        Choose(Retn, "调用Gdip函数时出现了一般性的错误", _
        "调用Gdip函数时输入的参数无效", _
        "调用Gdip函数时内存不足", _
        "调用Gdip函数时目标对象忙碌无响应", _
        "调用Gdip函数时缓冲区大小不足", _
        "调用Gdip函数时尚未实现操作", _
        "引发Win32错误", _
        "状态错误", _
        "调用Gdip函数时操作被终止", _
        "找不到文件", _
        "调用Gdip函数时参数值溢出", _
        "调用Gdip函数时访问被拒绝", _
        "未知的图像格式", _
        "找不到字体族", _
        "找不到字体类型", _
        "不是TrueType格式字体", _
        "使用的是不支持的GDIP版本", _
        "GDIP未初始化", _
        "找不到对应属性", _
        "不支持的对应属性")
    Else
        If ShowLog Then Debug.Print "GdiPlus初始化完成。"
    End If
End Sub

'关闭GDIPlus
Public Sub CloseGDIPlus(Optional ByVal ShowLog As Boolean = False)
    If mToken = 0 Then Exit Sub
    DeleteAllGdipCommonObjects
    GdiplusShutdown mToken
    mToken = 0
    If ShowLog Then Debug.Print "GdiPlus已关闭。"
End Sub

'新建对象
Public Sub NewGdipCommonObject(ByVal ObjName As String, ByVal ObjType As GdiplusCommonObject, ByVal ObjHandle As Long)
    Dim T As Long
    If ObjCount > 0 Then                                                        '已有GDIP通用对象，需要进行重名检测，以确保名称唯一性
        For T = 0 To ObjCount - 1
            If Objects(T).GdiplusObjectName = ObjName Then
                MsgBox "调用 NewGdipCommonObject 时错误：已存在名称为“" & ObjName & "”的Gdiplus通用对象。", vbCritical, "错误"
                Exit Sub
            End If
        Next T
    End If
    ObjCount = ObjCount + 1
    ReDim Preserve Objects(ObjCount - 1) As GdiplusObject
    With Objects(ObjCount - 1)
        .GdiplusObjectHandle = ObjHandle
        .GdiplusObjectName = ObjName
        .GdiplusObjectType = ObjType
    End With
End Sub

'NewGdipCommonObject过程的简化调用
Public Sub AddGCO(ByVal ObjName As String, ByVal ObjType As GdiplusCommonObject, ByVal ObjHandle As Long)
    NewGdipCommonObject ObjName, ObjType, ObjHandle
End Sub

'删除所有对象
Public Sub DeleteAllGdipCommonObjects()
    InitGDIPlus
    If ObjCount = 0 Then Exit Sub
    Dim T As Long
    For T = 0 To UBound(Objects)
        RemoveGdipCommonObjectByIndex T
    Next T
    Erase Objects
    ObjCount = 0
End Sub

'根据名字移除对象
Public Sub RemoveGdipCommonObject(ByVal ObjectName As String)
    InitGDIPlus
    If ObjCount = 0 Then Exit Sub
    Dim T As Long
    For T = 0 To UBound(Objects)
        If Objects(T).GdiplusObjectName = ObjectName Then
            RemoveGdipCommonObjectByIndex T
            Exit Sub
        End If
    Next T
End Sub

'RemoveGdipCommonObject过程的简化调用
Public Sub DelGCO(ByVal ObjName As String)
    RemoveGdipCommonObject ObjName
End Sub

'根据索引移除对象（内部使用）
Private Sub RemoveGdipCommonObjectByIndex(ByVal Index As Long)
    InitGDIPlus
    If Index < 0 Or Index >= ObjCount Then Exit Sub
    Dim T As Long
    Select Case Objects(Index).GdiplusObjectType
    Case GdiplusCommonObject.GdiplusBrush                                       '画刷
        GdipDeleteBrush Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusMatrix                                      '矩阵
        GdipDeleteMatrix Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusPen                                         '画笔
        GdipDeletePen Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusStringFormat                                '字符串格式
        GdipDeleteStringFormat Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusFont                                        '字体
        GdipDeleteFont Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusFontFamily                                  '字符集
        GdipDeleteFontFamily Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusGraphics                                    '画布
        GdipDeleteGraphics Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusPath                                        '路径
        GdipDeletePath Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusRegion                                      '区域
        GdipDeleteRegion Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusPathIter                                    '路径迭代器
        GdipDeletePathIter Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusCachedBitmap                                '缓存位图
        GdipDeleteCachedBitmap Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusImage                                       '图像
        GdipDisposeImage Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusDeviceContext                               '设备场景上下文
        DeleteDC Objects(Index).GdiplusObjectHandle
    End Select
    If Index = 0 And ObjCount = 1 Then
        Erase Objects
        ObjCount = 0
    Else
        If Index <> ObjCount - 1 Then
            For T = Index To ObjCount - 2
                Objects(T) = Objects(T + 1)
            Next T
        End If
        ObjCount = ObjCount - 1
        ReDim Preserve Objects(ObjCount - 1) As GdiplusObject
    End If
End Sub

'根据名字获得对象句柄（返回0表示获取句柄无效）
Public Function GetGdipCommonObjectHandle(ByVal ObjectName As String) As Long
    Dim T As Long
    If ObjCount = 0 Then Exit Function
    For T = 0 To UBound(Objects)
        If Objects(T).GdiplusObjectName = ObjectName Then
            GetGdipCommonObjectHandle = Objects(T).GdiplusObjectHandle
            Exit Function
        End If
    Next T
End Function

'GetGdipCommonObjectHandle函数的简化调用
Public Function GetGCO(ByVal ObjectName As String) As Long
    GetGCO = GetGdipCommonObjectHandle(ObjectName)
End Function

'新建点
Public Function NewPoint(ByVal X As Long, ByVal Y As Long) As PointL
    With NewPoint
        .X = X
        .Y = Y
    End With
End Function

Public Function NewPointFloat(ByVal X As Single, ByVal Y As Single) As PointF
    With NewPointFloat
        .X = X
        .Y = Y
    End With
End Function

'两种点变换
Public Function PointF2PointL(ByRef mPoint As PointF) As PointL
    With PointF2PointL
        .X = CLng(mPoint.X)
        .Y = CLng(mPoint.Y)
    End With
End Function

Public Function PointL2PointF(ByRef mPoint As PointL) As PointF
    With PointL2PointF
        .X = CSng(mPoint.X)
        .Y = CSng(mPoint.Y)
    End With
End Function

'新建长整数型点的数组
'PointsText的格式为：X1, Y1, X2, Y2 …
Public Function NewArrayOfPointL(ByVal PointsText As String, Optional ByVal Delimiter As String = ",") As PointL()
    If Delimiter = "" Then
        MsgBox "无效的分隔符。分隔符为一个字符。", vbCritical, "错误"
        Exit Function
    ElseIf PointsText = "" Then
        MsgBox "PointsText不能为空。", vbCritical, "错误"
        Exit Function
    End If
    Dim Retn() As PointL, T As Long, s() As String
    s = Split(PointsText, Delimiter)
    If UBound(s) Mod 2 = 0 Then
        MsgBox "参数 PointsText 所描述的点的X坐标和Y坐标数据数量不一致。", vbCritical, "错误"
        Exit Function
    End If
    ReDim Retn((UBound(s) - 1) / 2) As PointL
    For T = 0 To UBound(Retn)
        Retn(T) = NewPoint(Val(s(T * 2)), Val(s(T * 2 + 1)))
    Next T
    NewArrayOfPointL = Retn
End Function

'新建矩形
Public Function NewRect(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As RectL
    If Width <= 0 Or Height <= 0 Then
        MsgBox "无效的属性值。", vbCritical, "错误"
        Exit Function
    End If
    With NewRect
        .Left = Left
        .Top = Top
        .Right = Width + Left
        .Bottom = Height + Top
    End With
End Function

Public Function NewRectFloat(ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single) As RectF
    If Width <= 0 Or Height <= 0 Then
        MsgBox "无效的属性值。", vbCritical, "错误"
        Exit Function
    End If
    With NewRectFloat
        .Left = Left
        .Top = Top
        .Right = Width + Left
        .Bottom = Height + Top
    End With
End Function

'两种矩形变换
Public Function RectF2RectL(ByRef mRect As RectF) As RectL
    With RectF2RectL
        .Left = CLng(mRect.Left)
        .Top = CLng(mRect.Top)
        .Bottom = CLng(mRect.Bottom)
        .Right = CLng(mRect.Right)
    End With
End Function

Public Function RectL2RectF(ByRef mRect As RectL) As RectF
    With RectL2RectF
        .Left = CSng(mRect.Left)
        .Top = CSng(mRect.Top)
        .Bottom = CSng(mRect.Bottom)
        .Right = CSng(mRect.Right)
    End With
End Function

'新建尺寸
Public Function NewSize(ByVal Width As Long, ByVal Height As Long) As SizeL
    If Width <= 0 Or Height <= 0 Then
        MsgBox "无效的属性值。", vbCritical, "错误"
        Exit Function
    End If
    With NewSize
        .Width = Width
        .Height = Height
    End With
End Function

Public Function NewSizeFloat(ByVal Width As Single, ByVal Height As Single) As SizeF
    If Width <= 0 Or Height <= 0 Then
        MsgBox "无效的属性值。", vbCritical, "错误"
        Exit Function
    End If
    With NewSizeFloat
        .Width = Width
        .Height = Height
    End With
End Function

'两种尺寸变换
Public Function SizeF2SizeL(ByRef mSize As SizeF) As SizeL
    With SizeF2SizeL
        .Width = CLng(mSize.Width)
        .Height = CLng(mSize.Height)
    End With
End Function

Public Function SizeL2SizeF(ByRef mSize As SizeL) As SizeF
    With SizeL2SizeF
        .Width = CSng(mSize.Width)
        .Height = CSng(mSize.Height)
    End With
End Function

'新建ARGB颜色
Public Function NewARGBColor(ByVal Alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As ARGBColor
    With NewARGBColor
        .Alpha = Alpha
        .Red = Red
        .Green = Green
        .Blue = Blue
    End With
End Function

'将ARGB颜色转换为长整数形式
Public Function ARGBColor2Long(ByRef mARGBColor As ARGBColor) As Long
    Dim Retn As String
    Retn = Right("00" & Hex(mARGBColor.Alpha), 2) & Right("00" & Hex(mARGBColor.Red), 2) & Right("00" & Hex(mARGBColor.Green), 2) & Right("00" & Hex(mARGBColor.Blue), 2)
    ARGBColor2Long = CLng(Val("&H" & Retn))
End Function

'将vb6自带的OLE_COLOR变为ARGB颜色
Public Function OleColor2ARGBColor(ByVal OleColor As Long, Optional ByVal Alpha As Byte = 255) As ARGBColor
    Dim R As Byte, G As Byte, B As Byte, C As String
    C = Right("000000" & Hex(OleColor), 6)
    R = CByte(Val("&h" & Right(C, 2)))
    G = CByte(Val("&h" & Mid(C, 3, 2)))
    B = CByte(Val("&h" & Left(C, 2)))
    OleColor2ARGBColor = NewARGBColor(Alpha, R, G, B)
End Function

'新建画笔
Public Function NewPen(ByVal Color As Long, ByVal Width As Single) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreatePen1 Color, Width, GpUnit.UnitPixel, Retn
    NewPen = Retn
End Function

'新建纯色画刷
Public Function NewSolidBrush(ByVal Color As Long) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreateSolidFill Color, Retn
    NewSolidBrush = Retn
End Function

'新建线性渐变画刷
Public Function NewGradientLineBrush(ByRef Point1 As PointL, ByRef Point2 As PointL, ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal WrapType As WrapMode = WrapModeTile) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreateLineBrushI Point1, Point2, Color1, Color2, WrapType, Retn
    NewGradientLineBrush = Retn
End Function

'新建纹理画刷
Public Function NewHatchBrush(ByVal BrushHatchStyle As HatchStyle, ByVal ForeColor As Long, ByVal BackColor As Long) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreateHatchBrush BrushHatchStyle, ForeColor, BackColor, Retn
    NewHatchBrush = Retn
End Function

'新建贴图画刷
Public Function NewChartletBrush(ByVal hImage As Long, ByRef DatumPoint As PointL, Optional ByVal mWrapMode As WrapMode = WrapModeTile, Optional ByVal ReleaseHandles As Boolean = False)
    InitGDIPlus
    Dim Retn As Long
    GdipCreateTexture hImage, mWrapMode, Retn
    GdipResetTextureTransform Retn
    GdipTranslateTextureTransform Retn, DatumPoint.X, DatumPoint.Y, MatrixOrderAppend
    NewChartletBrush = Retn
    If ReleaseHandles Then GdipDisposeImage hImage
End Function

'新建字符串格式
Public Function NewStringFormat(Optional ByVal Alignment As StringAlignment = StringAlignmentCenter, _
    Optional ByVal FormatFlags As StringFormatFlags = StringFormatFlagsNoUse, _
    Optional ByVal trimming As StringTrimming = StringTrimmingNone, _
    Optional ByVal HkeyPrefix As HotkeyPrefix = HotkeyPrefixNone) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreateStringFormat 0, 0, Retn
    GdipSetStringFormatAlign Retn, Alignment
    GdipSetStringFormatFlags Retn, FormatFlags
    GdipSetStringFormatTrimming Retn, StringTrimmingNone
    GdipGetStringFormatHotkeyPrefix Retn, HkeyPrefix
    NewStringFormat = Retn
End Function

'新建矩阵（用于图像的变换）
Public Function NewMatrix(ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dX As Single, ByVal dY As Single) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreateMatrix Retn
    GdipSetMatrixElements Retn, m11, m12, m21, m22, dX, dY
    NewMatrix = Retn
End Function

'新建平移矩阵
Public Function NewTranslationalMatrix(ByVal X As Long, ByVal Y As Long) As Long
    NewTranslationalMatrix = NewMatrix(1, 0, 0, 1, X, Y)
End Function

'新建缩放矩阵
Public Function NewScalingMatrix(ByVal ScalingRatio As Single) As Long
    NewScalingMatrix = NewMatrix(ScalingRatio, 0, 0, ScalingRatio, 0, 0)
End Function

'新建旋转矩阵（角度）
Public Function NewRotatingMatrix(ByVal Degree As Single) As Long
    Const PI As Single = 3.14159265358979
    NewRotatingMatrix = NewMatrix(Cos(Degree / 180 * PI), Sin(Degree / 180 * PI), -Sin(Degree / 180 * PI), Cos(Degree / 180 * PI), 0, 0)
End Function

'新建字体
Public Function NewFont(ByVal hFontFamily As Long, ByVal cFontSize As Single, Optional ByVal cFontStyle As FontStyle = FontStyleRegular) As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreateFont hFontFamily, cFontSize, cFontStyle, UnitPixel, Retn
    NewFont = Retn
End Function

'新建字体族
Public Function NewFontFamily(ByVal cFontName As String, Optional ByVal cFontCollection As Long = 0) As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreateFontFamilyFromName StrPtr(cFontName), cFontCollection, Retn
    NewFontFamily = Retn
End Function

'新建画布
Public Function NewGraphics(ByVal hDC As Long) As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreateFromHDC hDC, Retn
    NewGraphics = Retn
End Function

'新建内存DC，返回这个DC的句柄
Public Function CreateMemoryDC(ByRef mSize As SizeL) As Long
    Dim tDC As Long, rDC As Long, rBmp As Long
    tDC = CreateDCAPI("DISPLAY", "", "", ByVal 0&)
    If tDC <> 0 Then
        rDC = CreateCompatibleDC(tDC)
        If rDC <> 0 Then
            rBmp = CreateCompatibleBitmap(tDC, mSize.Width, mSize.Height)
            If rBmp <> 0 Then
                DeleteObject SelectObject(rDC, rBmp)
                CreateMemoryDC = rDC
                DeleteObject rBmp
            Else
                DeleteDC rDC
            End If
        End If
        DeleteDC tDC
    End If
End Function

'新建路径
Public Function NewPath(Optional ByVal BrushMode As FillMode = FillModeAlternate) As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreatePath BrushMode, Retn
    NewPath = Retn
End Function

'新建基本形状的路径，其中FilletRadius当且仅当mShape参数为ShapeTypeRoundedRectangle时生效
Public Function NewBasicShapePath(ByRef mRect As RectL, ByVal mShape As ShapeType, Optional ByVal FilletRadius As Long = -1) As Long
    InitGDIPlus
    Dim hPath As Long, RoundSize As Long, tPath(7) As Long, T As Long
    Dim mLeft As Long, mTop As Long, mWidth As Long, mHeight As Long
    hPath = NewPath(FillMode.FillModeWinding)
    Select Case mShape
    Case ShapeTypeRectangle
        GdipAddPathRectangleI hPath, mRect.Left, mRect.Top, mRect.Right - mRect.Left, mRect.Bottom - mRect.Top
    Case ShapeTypeEllipse
        GdipAddPathEllipseI hPath, mRect.Left, mRect.Top, mRect.Right - mRect.Left, mRect.Bottom - mRect.Top
    Case ShapeTypeRoundedRectangle
        If FilletRadius <= 0 And FilletRadius <> -1 Then
            MsgBox "参数错误：FilletRadius的值必须大于0。", vbCritical, "错误"
            Exit Function
        ElseIf FilletRadius > Min(mRect.Bottom - mRect.Top, mRect.Right - mRect.Left) / 2 Or FilletRadius = -1 Then
            RoundSize = Min(mRect.Bottom - mRect.Top, mRect.Right - mRect.Left) / 2
        Else
            RoundSize = FilletRadius
        End If
        For T = 0 To 7
            tPath(T) = NewPath(FillMode.FillModeWinding)
        Next T
        mLeft = mRect.Left
        mTop = mRect.Top
        mWidth = mRect.Right - mRect.Left
        mHeight = mRect.Bottom - mRect.Top
        GdipAddPathArcI tPath(0), mLeft, mTop, RoundSize * 2, RoundSize * 2, 180, 90
        GdipAddPathArcI tPath(2), mLeft + mWidth - 2 * RoundSize, mTop, RoundSize * 2, RoundSize * 2, 270, 90
        GdipAddPathArcI tPath(4), mLeft + mWidth - 2 * RoundSize, mTop + mHeight - 2 * RoundSize, RoundSize * 2, RoundSize * 2, 0, 90
        GdipAddPathArcI tPath(6), mLeft, mTop + mHeight - 2 * RoundSize, RoundSize * 2, RoundSize * 2, 90, 90
        GdipAddPathLineI tPath(1), mLeft + RoundSize, mTop, mLeft + mWidth - RoundSize, mTop
        GdipAddPathLineI tPath(3), mLeft + mWidth, mTop + RoundSize, mLeft + mWidth, mTop + mHeight - RoundSize
        GdipAddPathLineI tPath(5), mLeft + mWidth - RoundSize, mTop + mHeight, mLeft + RoundSize, mTop + mHeight
        GdipAddPathLineI tPath(7), mLeft, mTop + RoundSize, mLeft, mTop + mHeight - RoundSize
        For T = 0 To 7
            GdipAddPathPath hPath, tPath(T), 1
            GdipDeletePath tPath(T)
        Next T
    End Select
    NewBasicShapePath = hPath
End Function

'新建多边形路径
Public Function NewPolygonPath(ByRef mPoint() As PointL) As Long
    If UBound(mPoint) - LBound(mPoint) < 2 Then
        MsgBox "参数错误：点集数组 mPoint 包含的数据不足，至少需要3个点才能构成一个封闭多边形。", vbCritical, "错误"
        Exit Function
    End If
    InitGDIPlus
    Dim hPath As Long, tPath() As Long, T As Long
    ReDim tPath(UBound(mPoint) - LBound(mPoint)) As Long
    hPath = NewPath(FillMode.FillModeWinding)
    For T = 0 To UBound(tPath) - 1
        tPath(T) = NewPath(FillMode.FillModeWinding)
        GdipAddPathLineI tPath(T), mPoint(T + LBound(mPoint)).X, mPoint(T + LBound(mPoint)).Y, mPoint(T + LBound(mPoint) + 1).X, mPoint(T + LBound(mPoint) + 1).Y
        GdipAddPathPath hPath, tPath(T), 1
    Next T
    tPath(UBound(tPath)) = NewPath(FillMode.FillModeWinding)
    GdipAddPathLineI tPath(UBound(tPath)), mPoint(UBound(mPoint)).X, mPoint(UBound(mPoint)).Y, mPoint(0).X, mPoint(0).Y
    GdipAddPathPath hPath, tPath(UBound(tPath)), 1
    For T = 0 To UBound(tPath)
        GdipDeletePath tPath(T)
    Next T
    NewPolygonPath = hPath
End Function

'新建路径迭代器
Public Function NewPathIterator(ByVal hPath As Long) As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreatePathIter Retn, hPath
    NewPathIterator = Retn
End Function

'新建区域
Public Function NewRegion() As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreateRegion Retn
    NewRegion = Retn
End Function

'根据路径创建对应的区域
Public Function NewRegionFromPath(ByVal hPath As Long, Optional ByVal mCombineMode As CombineMode = CombineModeReplace, Optional ByVal ReleaseHandles As Boolean = False) As Long
    Dim Retn As Long
    InitGDIPlus
    Retn = NewRegion
    GdipCombineRegionPath Retn, hPath, mCombineMode
    If ReleaseHandles Then GdipDeletePath hPath
    NewRegionFromPath = Retn
End Function

'根据矩形创建对应的区域
Public Function NewRegionFromRect(ByRef mRect As RectL) As Long
    Dim Retn As Long
    InitGDIPlus
    Retn = NewRegion
    GdipCreateRegionRectI mRect, Retn
    NewRegionFromRect = Retn
End Function

'新建字体类型
Public Function NewFontType(ByVal FontName As String, ByVal FontSize As Single, Optional ByVal Style As FontStyle = FontStyleRegular, Optional ByVal Weight As FontWeight = FW_NORMAL) As FontType
    With NewFontType
        .Name = FontName
        .Size = FontSize
        .Weight = Weight
        .Style = Style
    End With
End Function

'将StdFont转换为FontType
Public Function StdFont2FontType(ByRef sFont As StdFont, Optional ByVal FontSizeCalculatingMethod As CalculatingMethod = RoundDown) As FontType
    With StdFont2FontType
        .Name = sFont.Name
        .Size = CSng(Choose(FontSizeCalculatingMethod + 1, Int(sFont.Size * 4 / 3), Round(sFont.Size * 4 / 3), Abs(Int(0 - (sFont.Size * 4 / 3)))))
        .Weight = sFont.Weight
        .Style = IIf(sFont.Bold, 1, 0) + IIf(sFont.Italic, 2, 0) + IIf(sFont.Underline, 4, 0) + IIf(sFont.Strikethrough, 8, 0)
    End With
End Function

'从StdPicture中获得图像句柄hImage
Public Function GetImageFromStdPicture(ByRef StandardPicture As StdPicture) As Long
    InitGDIPlus
    Dim Retn As Long
    On Error GoTo ErrHandler
    GdipCreateBitmapFromHBITMAP StandardPicture.Handle, StandardPicture.hPal, Retn
    GetImageFromStdPicture = Retn
    Exit Function
ErrHandler:
    GdipCreateBitmapFromHBITMAP StandardPicture.Handle, 0, Retn
    GetImageFromStdPicture = Retn
End Function

'加载图像，返回图像句柄hImage
Public Function LoadImage(ByVal FilePath As String) As Long
    If Dir(FilePath) = "" Then Exit Function
    InitGDIPlus
    Dim Retn As Long
    GdipLoadImageFromFile StrPtr(FilePath), Retn
    LoadImage = Retn
End Function

'从文件夹中批量导入图像，并存储在Gdip通用对象数组中（以文件主名作为名称）
'后缀名（Suffix参数）不填写表示加载所有Gdip支持的图像，否则只包含这类图像（如Suffix参数为.png或者png时表示只导入这个文件夹下的所有png图片文件）
Public Sub LoadImagesFromFolder(ByVal FolderPath As String, Optional ByVal Suffix As String = "")
    If Dir(FolderPath, vbDirectory) = "" Then Exit Sub
    Dim FSO As Object, objFolder As Object, objFile As Object, nCount As Long, T As Long, nPath As String, mFile As String, mSuffix As String
    mSuffix = IIf(Suffix = "", "", IIf(Left(Suffix, 1) = ".", Right(Suffix, Len(Suffix) - 1), Suffix))
    Set FSO = CreateObject("Scripting.FileSystemObject")
    nPath = FolderPath & IIf(Right(FolderPath, 1) = "\", "", "\")
    Set objFolder = FSO.GetFolder(nPath)
    If objFolder.Files.Count = 0 Then Exit Sub
    For Each objFile In objFolder.Files
        mFile = objFile.Name
        If mSuffix <> "" Then                                                   '指定类型
            If GetFileSuffix(mFile) = mSuffix Then NewGdipCommonObject Left(mFile, Len(mFile) - Len(mSuffix) - 1), GdiplusImage, LoadImage(nPath & mFile)
        ElseIf mSuffix = "" Then
            If GetFileSuffix(mFile) = "bmp" Or GetFileSuffix(mFile) = "png" Or GetFileSuffix(mFile) = "jpg" Or GetFileSuffix(mFile) = "jpeg" Or GetFileSuffix(mFile) = "gif" Then _
            NewGdipCommonObject Left(mFile, Len(mFile) - Len(mSuffix) - 1), GdiplusImage, LoadImage(nPath & mFile)
        End If
    Next objFile
End Sub

'保存图像文件
Public Sub SaveImageFile(ByVal hImage As Long, ByVal FilePath As String, Optional ByVal FileEncoder As EncoderValue = EncoderValueColorTypeRGB, Optional ByVal ReleaseHandles As Boolean = False)
    InitGDIPlus
    Dim mSuffix As String
    Dim Params As EncoderParameters
    mSuffix = GetFileSuffix(FilePath)
    Params.Count = 1
    CLSIDFromString StrPtr(EncoderQuality), Params.Parameter.Guid
    Params.Parameter.NumberOfValues = 1
    Params.Parameter.ValueType = EncoderParameterValueTypeLong
    Params.Parameter.value = FileEncoder
    Select Case mSuffix
    Case "jpg", "jpeg"
        GdipSaveImageToFile hImage, StrPtr(FilePath), GetImageEncoderClsid(ImageFileSuffix.JPG), Params
    Case "bmp"
        GdipSaveImageToFile hImage, StrPtr(FilePath), GetImageEncoderClsid(ImageFileSuffix.Bmp), Params
    Case "gif"
        GdipSaveImageToFile hImage, StrPtr(FilePath), GetImageEncoderClsid(ImageFileSuffix.GIF), Params
    Case "png"
        GdipSaveImageToFile hImage, StrPtr(FilePath), GetImageEncoderClsid(ImageFileSuffix.PNG), Params
    Case "emf"
        GdipSaveImageToFile hImage, StrPtr(FilePath), GetImageEncoderClsid(ImageFileSuffix.EMF), Params
    Case "wmf"
        GdipSaveImageToFile hImage, StrPtr(FilePath), GetImageEncoderClsid(ImageFileSuffix.WMF), Params
    Case "tiff"
        GdipSaveImageToFile hImage, StrPtr(FilePath), GetImageEncoderClsid(ImageFileSuffix.TIF), Params
    Case Else
        MsgBox "文件格式错误：不是有效的图片文件格式。", vbCritical, "错误"
    End Select
    If ReleaseHandles Then GdipDisposeImage hImage
End Sub

'获得图像的尺寸
Public Function GetImageSize(ByVal hImage As Long) As SizeL
    InitGDIPlus
    If hImage = 0 Then Exit Function
    Dim imgWidth As Long, imgHeight As Long
    GdipGetImageWidth hImage, imgWidth
    GdipGetImageHeight hImage, imgHeight
    GetImageSize = NewSize(imgWidth, imgHeight)
End Function

'获得文件后缀
Public Function GetFileSuffix(ByVal FilePath As String) As String
    Dim nPath As String
    nPath = FilePath
    GetFileSuffix = LCase(Right(nPath, Len(nPath) - InStrRev(nPath, ".")))
End Function

'图像拷贝（根据hDC，可以实现双缓冲）
Public Sub CopyGraphics(ByVal hSourceGraphics As Long, ByVal hDestinationGraphics As Long, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hSrcDC As Long, hDstDC As Long, mWidth As Single, mHeight As Single
    GdipGetDC hSourceGraphics, hSrcDC
    GdipGetDC hDestinationGraphics, hDstDC
    mWidth = GetDeviceCaps(hDstDC, HORZRES)
    mHeight = GetDeviceCaps(hDstDC, VERTRES)
    BitBlt hDstDC, 0, 0, CLng(mWidth), CLng(mHeight), hSrcDC, 0, 0, vbSrcCopy
End Sub

'在画布上绘制简单文本
Public Sub DrawSimpleText(ByVal hGraphics As Long, ByVal Text As String, ByRef mFont As FontType, ByRef DatumPoint As PointL, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hStringFormat As Long, hFontFamily As Long, hFont As Long, hPath As Long, mRect As RectL
    If Text = "" Then Exit Sub
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    GdipCreateStringFormat 0, 0, hStringFormat                                  '创建字符串格式
    GdipStringFormatGetGenericTypographic hStringFormat                         '设置通用字符串格式
    GdipCreateFontFamilyFromName StrPtr(mFont.Name), 0, hFontFamily             '创建字体族
    GdipCreateFont hFontFamily, mFont.Size, mFont.Style, UnitPixel, hFont       '创建字体
    With mRect
        .Left = DatumPoint.X
        .Top = DatumPoint.Y
    End With
    hPath = NewPath                                                             '创建路径
    GdipAddPathStringI hPath, StrPtr(Text), -1, hFontFamily, mFont.Style, CLng(mFont.Size), mRect, hStringFormat '添加字体路径
    GdipFillPath hGraphics, hFill, hPath                                        '填充
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '描边
    GdipDeletePath hPath                                                        '删除路径
    GdipDeleteStringFormat hStringFormat                                        '删除临时字符串格式句柄
    GdipDeleteFont hFont                                                        '删除临时字体
    GdipDeleteFontFamily hFontFamily                                            '删除临时字体族句柄
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteGraphics hGraphics
    End If
End Sub

'在画布上绘制遮罩文本(以百分比计算)
Public Sub DrawMaskedText(ByVal hGraphics As Long, ByVal Text As String, ByRef mFont As FontType, ByRef DatumPoint As PointL, ByVal Percentage As Single, ByVal hBorder As Long, ByVal hFill As Long, ByVal hMask As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hStringFormat As Long, hFontFamily As Long, hFont As Long, hPath As Long, mRect As RectL, oRect As RectF, nRect As RectF, tRect As RectL, hRegion As Long
    Dim tCodePointsFitted As Long, tLinesFilled As Long, nPercentage As Single  '临时需要的变量
    If Text = "" Then Exit Sub
    nPercentage = IIf(Percentage < 0, 0, IIf(Percentage > 100, 100, Percentage))
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    GdipCreateStringFormat 0, 0, hStringFormat                                  '创建字符串格式
    GdipStringFormatGetGenericTypographic hStringFormat                         '设置通用字符串格式
    GdipCreateFontFamilyFromName StrPtr(mFont.Name), 0, hFontFamily             '创建字体族
    GdipCreateFont hFontFamily, mFont.Size, mFont.Style, UnitPixel, hFont       '创建字体
    oRect = NewRectFloat(DatumPoint.X, DatumPoint.Y, Len(Text) * mFont.Size, mFont.Size) '原始矩形
    GdipMeasureString hGraphics, StrPtr(Text), Len(Text), hFont, oRect, hStringFormat, nRect, tCodePointsFitted, tLinesFilled
    tRect = RectF2RectL(nRect)                                                  '字符串所在矩形
    With tRect
        .Right = 0
        .Bottom = 0
    End With
    mRect = RectF2RectL(nRect)                                                  '遮罩矩形
    mRect.Right = mRect.Right * nPercentage / 100                               '计算遮罩百分比
    GdipCreateRegionRectI mRect, hRegion                                        '创建遮罩区域
    hPath = NewPath                                                             '创建路径
    GdipAddPathStringI hPath, StrPtr(Text), Len(Text), hFontFamily, mFont.Style, CLng(mFont.Size), tRect, hStringFormat '添加字体路径
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '描边
    GdipFillPath hGraphics, hFill, hPath                                        '填充字符串
    GdipSetClipRectI hGraphics, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, CombineModeReplace '裁剪遮罩区域
    GdipFillPath hGraphics, hMask, hPath                                        '填充遮罩区域
    GdipResetClip hGraphics                                                     '重置裁剪
    GdipDeleteRegion hRegion
    GdipDeletePath hPath
    GdipDeleteFont hFont
    GdipDeleteFontFamily hFontFamily
    GdipDeleteStringFormat hStringFormat
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteBrush hMask
        GdipDeleteGraphics hGraphics
    End If
End Sub

'在画布上绘制直线
Public Sub DrawLine(ByVal hGraphics As Long, ByRef Point1 As PointL, ByRef Point2 As PointL, ByVal hPen As Long, Optional ByVal ReleaseHandles As Boolean = False)
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    GdipDrawLineI hGraphics, hPen, Point1.X, Point1.Y, Point2.X, Point2.Y
    If ReleaseHandles Then
        GdipDeletePen hPen
        GdipDeleteGraphics hGraphics
    End If
End Sub

'在画布上绘制矩形
Public Sub DrawRectangle(ByVal hGraphics As Long, ByRef mRect As RectL, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hPath  As Long
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    hPath = NewPath
    GdipAddPathRectangleI hPath, mRect.Left, mRect.Top, mRect.Right - mRect.Left, mRect.Bottom - mRect.Top '添加矩形路径
    GdipFillPath hGraphics, hFill, hPath                                        '填充
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '描边
    GdipDeletePath hPath
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteGraphics hGraphics
    End If
End Sub

'在画布上绘制椭圆（包含圆形）
Public Sub DrawEllipse(ByVal hGraphics As Long, ByRef mRect As RectL, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hPath As Long
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias                      '抗锯齿处理
    hPath = NewPath
    GdipAddPathEllipseI hPath, mRect.Left, mRect.Top, mRect.Right - mRect.Left, mRect.Bottom - mRect.Top '添加矩形路径
    GdipFillPath hGraphics, hFill, hPath                                        '填充
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '描边
    GdipDeletePath hPath
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteGraphics hGraphics
    End If
End Sub

'在画布上绘制圆角矩形
Public Sub DrawRoundedRectangle(ByVal hGraphics As Long, ByRef mRect As RectL, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal FilletRadius As Long = -1, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hPath As Long, RoundSize As Long, tPath(7) As Long, T As Long
    Dim mLeft As Long, mTop As Long, mWidth As Long, mHeight As Long
    If FilletRadius <= 0 And FilletRadius <> -1 Then
        MsgBox "参数错误：FilletRadius的值必须大于0。", vbCritical, "错误"
        Exit Sub
    ElseIf FilletRadius > Min(mRect.Bottom - mRect.Top, mRect.Right - mRect.Left) / 2 Or FilletRadius = -1 Then
        RoundSize = Min(mRect.Bottom - mRect.Top, mRect.Right - mRect.Left) / 2
    Else
        RoundSize = FilletRadius
    End If
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    hPath = NewPath(FillMode.FillModeWinding)
    For T = 0 To 7
        tPath(T) = NewPath(FillMode.FillModeWinding)
    Next T
    mLeft = mRect.Left
    mTop = mRect.Top
    mWidth = mRect.Right - mRect.Left
    mHeight = mRect.Bottom - mRect.Top
    GdipAddPathArcI tPath(0), mLeft, mTop, RoundSize * 2, RoundSize * 2, 180, 90
    GdipAddPathArcI tPath(2), mLeft + mWidth - 2 * RoundSize, mTop, RoundSize * 2, RoundSize * 2, 270, 90
    GdipAddPathArcI tPath(4), mLeft + mWidth - 2 * RoundSize, mTop + mHeight - 2 * RoundSize, RoundSize * 2, RoundSize * 2, 0, 90
    GdipAddPathArcI tPath(6), mLeft, mTop + mHeight - 2 * RoundSize, RoundSize * 2, RoundSize * 2, 90, 90
    GdipAddPathLineI tPath(1), mLeft + RoundSize, mTop, mLeft + mWidth - RoundSize, mTop
    GdipAddPathLineI tPath(3), mLeft + mWidth, mTop + RoundSize, mLeft + mWidth, mTop + mHeight - RoundSize
    GdipAddPathLineI tPath(5), mLeft + mWidth - RoundSize, mTop + mHeight, mLeft + RoundSize, mTop + mHeight
    GdipAddPathLineI tPath(7), mLeft, mTop + RoundSize, mLeft, mTop + mHeight - RoundSize
    For T = 0 To 7
        GdipAddPathPath hPath, tPath(T), 1
    Next T
    GdipFillPath hGraphics, hFill, hPath                                        '填充
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '描边
    GdipDeletePath hPath
    For T = 0 To 7
        GdipDeletePath tPath(T)
    Next T
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteGraphics hGraphics
    End If
End Sub

'在画布上绘制多边形
Public Sub DrawPolygon(ByVal hGraphics As Long, ByRef mPoint() As PointL, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    If UBound(mPoint) - LBound(mPoint) < 2 Then
        MsgBox "参数错误：点集数组 mPoint 包含的数据不足，至少需要3个点才能构成一个封闭多边形。", vbCritical, "错误"
        Exit Sub
    End If
    Dim hPath As Long, tPath() As Long, T As Long
    ReDim tPath(UBound(mPoint) - LBound(mPoint)) As Long
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    hPath = NewPath(FillMode.FillModeWinding)
    For T = 0 To UBound(tPath) - 1
        tPath(T) = NewPath(FillMode.FillModeWinding)
        GdipAddPathLineI tPath(T), mPoint(T + LBound(mPoint)).X, mPoint(T + LBound(mPoint)).Y, mPoint(T + LBound(mPoint) + 1).X, mPoint(T + LBound(mPoint) + 1).Y
        GdipAddPathPath hPath, tPath(T), 1
    Next T
    tPath(UBound(tPath)) = NewPath(FillMode.FillModeWinding)
    GdipAddPathLineI tPath(UBound(tPath)), mPoint(UBound(mPoint)).X, mPoint(UBound(mPoint)).Y, mPoint(0).X, mPoint(0).Y
    GdipAddPathPath hPath, tPath(UBound(tPath)), 1
    GdipFillPath hGraphics, hFill, hPath                                        '填充
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '描边
    GdipDeletePath hPath
    For T = 0 To UBound(tPath)
        GdipDeletePath tPath(T)
    Next T
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteGraphics hGraphics
    End If
End Sub

'在画布上绘制图像
Public Sub DrawImage(ByVal hGraphics As Long, ByVal hImage As Long, ByRef DatumPoint As PointL, Optional ByVal TransformMode As RotateFlipType = RotateNoneFlipNone, Optional ByVal Zoom As Single = 1#, Optional ByVal ReleaseHandles As Boolean = False)
    Dim imgWidth As Long, imgHeight As Long
    If Zoom <= 0 Then
        MsgBox "参数错误： Zoom 的值应大于0。", vbCritical, "错误"
        Exit Sub
    End If
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipGetImageWidth hImage, imgWidth
    GdipGetImageHeight hImage, imgHeight
    GdipImageRotateFlip hImage, TransformMode
    If TransformMode Mod 2 = 0 Then
        GdipDrawImageRectI hGraphics, hImage, DatumPoint.X, DatumPoint.Y, imgWidth * Zoom, imgHeight * Zoom
    Else
        GdipDrawImageRectI hGraphics, hImage, DatumPoint.X, DatumPoint.Y, imgHeight * Zoom, imgWidth * Zoom
    End If
    If ReleaseHandles Then
        GdipDisposeImage hImage
        GdipDeleteGraphics hGraphics
    End If
End Sub

'在画布上绘制瓦片*。
'注：瓦片（Tile）指一整张图像中指定一个区域的图像。
Public Sub DrawTile(ByVal hGraphics As Long, ByVal hImage As Long, ByRef DatumPoint As PointL, ByRef SourceRect As RectL, Optional ByVal Zoom As Single = 1#, Optional ByVal ReleaseHandles As Boolean = False)
    Dim imgWidth As Long, imgHeight As Long, mSize As SizeL
    Dim hMemDC As Long, hMemGraphics As Long
    If Zoom <= 0 Then
        MsgBox "参数错误： Zoom 的值应大于0。", vbCritical, "错误"
        Exit Sub
    End If
    InitGDIPlus
    GdipDrawImageRectRectI hGraphics, hImage, DatumPoint.X, DatumPoint.Y, (SourceRect.Right - SourceRect.Left) * Zoom, (SourceRect.Bottom - SourceRect.Top) * Zoom, _
    SourceRect.Left, SourceRect.Top, SourceRect.Right - SourceRect.Left, SourceRect.Bottom - SourceRect.Top, UnitPixel
    If ReleaseHandles Then
        GdipDisposeImage hImage
        GdipDeleteGraphics hGraphics
    End If
End Sub

'在画布上绘制指定的路径
Public Sub DrawPath(ByVal hGraphics As Long, ByVal hPath As Long, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipFillPath hGraphics, hFill, hPath                                        '填充
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '描边
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteGraphics hGraphics
        GdipDeletePath hPath
    End If
End Sub

'用指定颜色涂满整个画布
Public Sub FillWholeGraphics(ByVal hGraphics As Long, Optional ByVal mColor As Long = &HFFFFFFFF, Optional ByVal ReleaseHandles As Boolean = False)
    InitGDIPlus
    GdipGraphicsClear hGraphics, mColor
    If ReleaseHandles Then GdipDeleteGraphics hGraphics
End Sub

'合并路径（把hPathAdditional合并至hPathOriginal中）
Public Sub CombinePath(ByVal hPathOriginal As Long, ByVal hPathAdditional As Long, Optional ByVal Connecting As Boolean = True)
    InitGDIPlus
    GdipAddPathPath hPathOriginal, hPathAdditional, Abs(CLng(Connecting))
End Sub

'判断点是否在区域上
Public Function IsPointOnRegion(ByVal hGraphics As Long, ByVal hRegion As Long, ByRef mPoint As PointL, Optional ByVal ReleaseHandles As Boolean = False) As Boolean
    InitGDIPlus
    Dim Retn As Long
    GdipIsVisibleRegionPointI hRegion, mPoint.X, mPoint.Y, hGraphics, Retn
    IsPointOnRegion = (Retn <> 0)
    If ReleaseHandles Then
        GdipDeleteGraphics hGraphics
        GdipDeleteRegion hRegion
    End If
End Function

'判断矩形是否与区域相交，或者矩形位于区域内部
'True - 矩形与区域相交，或者矩形位于区域内部；False - 矩形与区域相离
Public Function IsRectOnRegion(ByVal hGraphics As Long, ByVal hRegion As Long, ByRef mRect As RectL, Optional ByVal ReleaseHandles As Boolean = False) As Boolean
    InitGDIPlus
    Dim Retn As Long
    GdipIsVisibleRegionRectI hRegion, mRect.Left, mRect.Top, mRect.Right - mRect.Left, mRect.Bottom - mRect.Top, hGraphics, Retn
    IsRectOnRegion = (Retn <> 0)
    If ReleaseHandles Then
        GdipDeleteGraphics hGraphics
        GdipDeleteRegion hRegion
    End If
End Function
