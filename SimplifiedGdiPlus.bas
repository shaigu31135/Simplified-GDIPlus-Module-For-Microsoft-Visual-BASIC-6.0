Attribute VB_Name = "SimplifiedGdiPlus"
'SimplifiedGdiPlus
'����modGdipģ�飨ԭ���ߣ�vIstaSwx������
'����ע��������΢��msdn�ֲᣨlearn.microsoft.com���������к���ƫ��
'ע���а�����[?]��Ϊ�����ݲ���ȷ
'By ���ư��侩��

'1.0.0 (2024-07-31)
'�����˴�ģ�顣

'1.0.1��2024-08-28��
'�޸���DrawPath����ʵ�ֿ���ݣ�
'�Ķ������й���/������ReleaseHandles����ȱʡΪFalse��
'�Ķ�����CreateBasicShapePath����ΪNewBasicShapePath��
'������NewPolygonPath��

Option Explicit

'##################  ö  ��  ##################
'���ظ�ʽ
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

'ͼ�������λ
Public Enum GpUnit
    UnitWorld                                                                   '����
    UnitDisplay                                                                 '��ʾ��
    UnitPixel                                                                   '����
    UnitPoint                                                                   '��
    UnitInch                                                                    'Ӣ��
    UnitDocument                                                                '�ĵ�
    UnitMillimeter                                                              '����
End Enum

'·��������
Public Enum PathPointType
    PathPointTypeStart = 0                                                      '��ʼ��
    PathPointTypeLine = 1                                                       'ֱ��
    PathPointTypeBezier = 3                                                     '����������
    PathPointTypePathTypeMask = &H7                                             '·����������
    PathPointTypePathDashMode = &H10                                            '·������ģʽ[?]
    PathPointTypePathMarker = &H20                                              '·�����
    PathPointTypeCloseSubpath = &H80                                            '�ر���·��
    PathPointTypeBezier3 = PathPointTypeBezier
End Enum

'ͨ�������壨�����ã�
Public Enum GenericFontFamily
    GenericFontFamilySerif
    GenericFontFamilySansSerif
    GenericFontFamilyMonospace
End Enum

'���ͣ�������ʽ�����������ʹ�ã���󲻳���15
Public Enum FontStyle
    FontStyleRegular = 0                                                        '����
    FontStyleBold = 1                                                           '����
    FontStyleItalic = 2                                                         'б��
    FontStyleBoldItalic = FontStyleBold + FontStyleItalic
    FontStyleUnderline = 4                                                      '�»���
    FontStyleStrikeout = 8                                                      'ɾ����
End Enum

'����
Public Enum StringAlignment
    StringAlignmentNear = 0                                                     '����
    StringAlignmentCenter = 1                                                   '����
    StringAlignmentFar = 2                                                      'Զ��
End Enum

'���ģʽ
Public Enum FillMode
    FillModeAlternate                                                           '������䣨ȱʡ��
    FillModeWinding                                                             'Χ�����
    'FillModeAlternate���ӷ�������е�һ����������Զ��ˮƽ��һ�����ߣ��������ߴ�Խ�������߿���ʱ�����������
    '�������ߴ�Խż�����߿���ʱ�������������
    'FillModeWinding���ӷ�������е�һ����������Զ��ˮƽ��һ�����ߣ��������ߴ�Խ�������߿���ʱ����������򣻵�
    '�����ߴ�Խż�����߿���ʱ����Ҫ���ݱ߿��ߵķ������жϣ���������ı߿����ڲ�ͬ����ı߿�����Ŀ��ȣ������
    '��������粻��ȣ������������
End Enum

'ƽ��ģʽ��ע�ⲻ��Ť��ģʽWarpMode��
Public Enum WrapMode
    WrapModeTile                                                                '����ת��ƽ��
    WrapModeTileFlipX                                                           '��һ���д�һ�������ƶ�����һ������ʱˮƽ��ת����
    WrapModeTileFlipY                                                           '�����д�һ�������ƶ�����һ������ʱ��ֱ��ת����
    WrapModeTileFlipXY                                                          '�������ƶ�ʱˮƽ��ת�������������ƶ�ʱ��ֱ��ת����
    WrapModeClamp                                                               '������ƽ��
End Enum

'���Խ���ģʽ
Public Enum LinearGradientMode
    LinearGradientModeHorizontal                                                'ˮƽ���� ����
    LinearGradientModeVertical                                                  '��ֱ���� ��
    LinearGradientModeForwardDiagonal                                           '���ϵ����½��� ��
    LinearGradientModeBackwardDiagonal                                          '���ϵ����½��� ��
End Enum

'����ģʽ
Public Enum QualityMode
    QualityModeInvalid = -1                                                     '��Ч
    QualityModeDefault = 0                                                      'Ĭ��
    QualityModeLow = 1                                                          '������
    QualityModeHigh = 2                                                         '������
End Enum

'��ɫ�ϳ�ģʽ
Public Enum CompositingMode
    CompositingModeSourceOver                                                   '���ģʽ��������ɫ���뱳��ɫ��ϣ���ϱ����ɳ�����ɫ�Ħ���������
    CompositingModeSourceCopy                                                   '����ģʽ��������ɫ��ֱ�Ӹ��Ǳ���ɫ
End Enum

'��ɫ�ϳ�����
Public Enum CompositingQuality
    CompositingQualityInvalid = QualityModeInvalid
    CompositingQualityDefault = QualityModeDefault
    CompositingQualityHighSpeed = QualityModeLow
    CompositingQualityHighQuality = QualityModeHigh
    CompositingQualityGammaCorrected                                            'ʹ�æ�У��
    CompositingQualityAssumeLinear                                              '�ٶ�ģ��Ϊ����
End Enum

'ƽ��ģʽ
Public Enum SmoothingMode
    SmoothingModeInvalid = QualityModeInvalid
    SmoothingModeDefault = QualityModeDefault
    SmoothingModeHighSpeed = QualityModeLow
    SmoothingModeHighQuality = QualityModeHigh
    SmoothingModeNone                                                           '��ʹ��
    SmoothingModeAntiAlias                                                      '����ݣ�ʹ�� 8 �� 4 ��ɸѡ����
End Enum

'ͼ���ֵģʽ
Public Enum InterpolationMode
    InterpolationModeInvalid = QualityModeInvalid
    InterpolationModeDefault = QualityModeDefault
    InterpolationModeLowQuality = QualityModeLow
    InterpolationModeHighQuality = QualityModeHigh
    InterpolationModeBilinear                                                   '˫���Բ�ֵ
    InterpolationModeBicubic                                                    '˫���β�ֵ
    InterpolationModeNearestNeighbor                                            '���ٽ���ֵ
    InterpolationModeHighQualityBilinear                                        '��������˫���Բ�ֵ
    InterpolationModeHighQualityBicubic                                         '��������˫���β�ֵ
End Enum

'����ƫ��ģʽ
Public Enum PixelOffsetMode
    PixelOffsetModeInvalid = QualityModeInvalid
    PixelOffsetModeDefault = QualityModeDefault
    PixelOffsetModeHighSpeed = QualityModeLow
    PixelOffsetModeHighQuality = QualityModeHigh
    PixelOffsetModeNone                                                         '�������ľ����������꣨������ƫ�ƣ�
    PixelOffsetModeHalf                                                         '�������ĵ������������ֵ֮�䣨����ƫ�ƣ�
    '�ٶ�ͼ�����Ͻǵ�����Ϊ��0,0����
    'PixelOffsetModeNone�����ظ��� x �� y ���� �C0.5 �� 0.5 ֮������򣬼���������λ�ڣ�0��0����
    'PixelOffsetModeHalf�����ظ��� x �� y ���� 0 �� 1 ֮������򣬼���������λ�ڣ�0.5��0.5����
End Enum

'�ı���Ⱦ��ʾ
Public Enum TextRenderingHint
    TextRenderingHintSystemDefault = 0                                          'ʹ�õ�ǰ��ѡϵͳ����ƽ��ģʽ�����ַ�
    TextRenderingHintSingleBitPerPixelGridFit                                   'ʹ���ַ�����λͼ����ʾ�����ַ�
    TextRenderingHintSingleBitPerPixel                                          'ʹ���ַ�����λͼ�����ַ���������ʾ��ʾ
    TextRenderingHintAntiAliasGridFit                                           'ʹ���ַ����������λͼ����ʾ�����ַ�
    TextRenderingHintAntiAlias                                                  'ʹ���ַ����������λͼ�����ַ���������ʾ��ʾ
    TextRenderingHintClearTypeGridFit                                           'ʹ���ַ����� ClearType λͼ����ʾ�����ַ�
End Enum

'��ɫ����˳��
Public Enum MatrixOrder
    MatrixOrderPrepend = 0                                                      '�¾���λ�����о���λ���
    MatrixOrderAppend = 1                                                       '�¾���λ�����о���λ�Ҳ�
End Enum

'��ɫ��������
Public Enum ColorAdjustType
    ColorAdjustTypeDefault                                                      'Ĭ��
    ColorAdjustTypeBitmap                                                       'λͼ
    ColorAdjustTypeBrush                                                        '��ˢ
    ColorAdjustTypePen                                                          '����
    ColorAdjustTypeText                                                         '�ı�
    ColorAdjustTypeCount                                                        '����
    ColorAdjustTypeAny                                                          '��Ԥ����
End Enum

'��ɫ�����־
Public Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = 0                                                 '��������ɫֵ����ͬһ��ɫ�����������
    ColorMatrixFlagsSkipGrays = 1                                               '������ɫ������������ɫ����*
    ColorMatrixFlagsAltGray = 2                                                 '��ɫ��һ�������������ɫ��������һ���������
    'ע����ɫ������ָ���ɫ����ɫ����ɫ������ֵ����ͬ���κ���ɫ��
End Enum

'Ť��ģʽ��ע�ⲻ��ƽ��ģʽWrapMode��
Public Enum WarpMode
    WarpModePerspective                                                         '͸��Ť��
    WarpModeBilinear                                                            '˫����Ť��
End Enum

'�ϲ�ģʽ
Public Enum CombineMode
    CombineModeReplace                                                          '���������滻Ϊ������
    CombineModeIntersect                                                        '���������滻Ϊ�������������Ľ���
    CombineModeUnion                                                            '���������滻Ϊ�������������Ĳ���
    CombineModeXor                                                              '���������滻Ϊ�����������������
    CombineModeExclude                                                          '���������滻Ϊλ��������֮��ĸ����򲿷�
    CombineModeComplement                                                       '���������滻Ϊ�������ⲿ�������򲿷�
End Enum

'ͼ������ģʽ
Public Enum ImageLockMode
    ImageLockModeRead = &H1                                                     '����ͼ���һ�����Ա��ȡ
    ImageLockModeWrite = &H2                                                    '����ͼ���һ�����Ա�д��
    ImageLockModeUserInputBuf = &H4                                             '���û���ȡ��д��ͼ����ʹ�õĻ�����
End Enum

'��״����
Public Enum ShapeType
    ShapeTypeRectangle = 0                                                      '����
    ShapeTypeEllipse = 1                                                        '��Բ������Բ�Σ�
    ShapeTypeRoundedRectangle = 2                                               'Բ�Ǿ���
End Enum

'��״����ģʽ
Public Enum ShapeDrawingMode
    ShapeDrawingModeEdge = 0                                                    '���
    ShapeDrawingModeFill = 1                                                    '���
    ShapeDrawingModeEdgeAndFill = 2                                             '��ߺ����
End Enum

'ͼ�񱣴��ʽ
Public Enum GpImageSaveFormat
    GpSaveBMP = 0
    GpSaveJPEG = 1
    GpSaveGIF = 2
    GpSavePNG = 3
    GpSaveTIFF = 4
End Enum

'ͼ���ʽ��ʶ��
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

'ͼ������
Public Enum ImageType
    ImageTypeUnknown = 0                                                        'δ֪
    ImageTypeBitmap = 1                                                         'λͼ
    ImageTypeMetafile = 2                                                       'ͼԪ�ļ�
End Enum

'ͼ���������ͣ������ã�
Public Enum ImagePropertyType
    ImagePropertyTypeByte = 1
    ImagePropertyTypeASCII = 2
    ImagePropertyTypeShort = 3
    ImagePropertyTypeLong = 4
    ImagePropertyTypeRational = 5                                               '��������[?]
    ImagePropertyTypeUndefined = 7                                              'δ�����
    ImagePropertyTypeSLONG = 9
    ImagePropertyTypeSRational = 10
End Enum

'ͼ���/��������־
Public Enum ImageCodecFlags
    ImageCodecFlagsEncoder = &H1                                                '������
    ImageCodecFlagsDecoder = &H2                                                '������
    ImageCodecFlagsSupportBitmap = &H4                                          '֧��λͼ
    ImageCodecFlagsSupportVector = &H8                                          '֧������[?]
    ImageCodecFlagsSeekableEncode = &H10                                        '�ɶ�λ������
    ImageCodecFlagsBlockingDecode = &H20                                        '�����Ľ�����[?]
    ImageCodecFlagsBuiltin = &H10000                                            '�ڽ�ʽ[?]
    ImageCodecFlagsSystem = &H20000                                             '��ϵͳ
    ImageCodecFlagsUser = &H40000                                               '���û�
End Enum

'��ɫ���־
Public Enum PaletteFlags
    PaletteFlagsHasAlpha = &H1                                                  '�߱���������͸���ȣ�
    PaletteFlagsGrayScale = &H2                                                 '�������Ҷ�
    PaletteFlagsHalftone = &H4                                                  'Windows��ɫ����ɫ��
End Enum

'��ת/��ת����
Public Enum RotateFlipType
    RotateNoneFlipNone = 0                                                      '��
    Rotate90FlipNone = 1                                                        '��ת90��
    Rotate180FlipNone = 2                                                       '��ת180��
    Rotate270FlipNone = 3                                                       '��ת270��
    RotateNoneFlipX = 4                                                         'ˮƽ��ת
    Rotate90FlipX = 5                                                           '����ת90�㣬Ȼ����ˮƽ��ת
    Rotate180FlipX = 6                                                          '����ת180�㣬Ȼ����ˮƽ��ת
    Rotate270FlipX = 7                                                          '����ת270�㣬Ȼ����ˮƽ��ת
    RotateNoneFlipY = Rotate180FlipX
    Rotate90FlipY = Rotate270FlipX
    Rotate180FlipY = RotateNoneFlipX
    Rotate270FlipY = Rotate90FlipX
    RotateNoneFlipXY = Rotate180FlipNone
    Rotate90FlipXY = Rotate270FlipNone
    Rotate180FlipXY = RotateNoneFlipNone
    Rotate270FlipXY = Rotate90FlipNone
End Enum

'��ɫ���ģʽ
Public Enum ColorMode
    ColorModeARGB32 = 0
    ColorModeARGB64 = 1
End Enum

'CMYKģʽͨ����־�������ã�
Public Enum ColorChannelFlags
    ColorChannelFlagsC = 0                                                      '��ɫ
    ColorChannelFlagsM                                                          '���ɫ
    ColorChannelFlagsY                                                          '��ɫ
    ColorChannelFlagsK                                                          '��ɫ
    ColorChannelFlagsLast                                                       '��һ��[?]
End Enum

'ARGB����
Public Enum ColorShiftComponents
    AlphaShift = 24                                                             '������
    RedShift = 16                                                               '�����
    GreenShift = 8                                                              '�̷���
    BlueShift = 0                                                               '������
End Enum

'ARGB��ɫ����
Public Enum ColorMaskComponents
    AlphaMask = &HFF000000
    RedMask = &HFF0000
    GreenMask = &HFF00
    BlueMask = &HFF
End Enum

'���أ�ֵԽ���ֿ�����Խ�֣�
Public Enum FontWeight
    FW_DONTCARE = 0&                                                            'ʹ��Ĭ�ϵ�����
    FW_THIN = 100&                                                              '��ϸ��-3��
    FW_EXTRALIGHT = 200&                                                        '�ر�ϸ��-2��
    FW_ULTRALIGHT = FW_EXTRALIGHT                                               '�ر�ϸ��-2��
    FW_LIGHT = 300&                                                             '��ϸ��-1��
    FW_NORMAL = 400&                                                            '������ϸ��0��
    FW_REGULAR = FW_NORMAL                                                      '������ϸ��0��
    FW_MEDIUM = 500&                                                            '�Դ֣�+1��
    FW_SEMIBOLD = 600&                                                          '�еȴ֣�+2��
    FW_DEMIBOLD = FW_SEMIBOLD                                                   '
    FW_BOLD = 700&                                                              '�֣�+3��
    FW_EXTRABOLD = 800&                                                         '�ر�֣�+4��
    FW_ULTRABOLD = FW_EXTRABOLD                                                 '
    FW_HEAVY = 900&                                                             '��֣�+5��
    FW_BLACK = FW_HEAVY                                                         '
End Enum

'�ַ�������
Public Enum CharSetType
    ANSI_CHARSET = 0
    DEFAULT_CHARSET = 1                                                         '���ݵ�ǰϵͳ�������ã�Ĭ�ϣ�
    SYMBOL_CHARSET = 2
    SHIFTJIS_CHARSET = 128
    HANGEUL_CHARSET = 129
    HANGUL_CHARSET = 129
    GB2312_CHARSET = 134                                                        '�������ģ�����2312��
    CHINESEBIG5_CHARSET = 136                                                   '���w���ģ�����a��
    GREEK_CHARSET = 161
    TURKISH_CHARSET = 162
    HEBREW_CHARSET = 177
    ARABIC_CHARSET = 178
    BALTIC_CHARSET = 186
    RUSSIAN_CHARSET = 204
    THAI_CHARSET = 222
    EASTEUROPE_CHARSET = 238
    OEM_CHARSET = 255                                                           '�����ڲ���ϵͳ���ַ���
    JOHAB_CHARSET = 130
    VIETNAMESE_CHARSET = 163
    MAC_CHARSET = 77
End Enum

'���������������
Public Enum OutPrecisionType
    OUT_DEFAULT_PRECIS = 0                                                      'Ĭ������ӳ������Ϊ
    OUT_STRING_PRECIS = 1                                                       '����ӳ������ʹ�ô�ֵ������ö�ٹ�դ����ʱ�᷵�ش�ֵ
    OUT_CHARACTER_PRECIS = 2                                                    'δʹ��
    OUT_STROKE_PRECIS = 3                                                       '����ӳ������ʹ�ô�ֵ������ö��TrueType���������������������ʸ������ʱ���ش�ֵ
    OUT_TT_PRECIS = 4                                                           '��ϵͳ�������ͬ������ʱ��ָʾ����ӳ����ѡ�� TrueType ����
    OUT_DEVICE_PRECIS = 5                                                       '��ϵͳ�������ͬ������ʱ��ָʾ����ӳ����ѡ���豸����
    OUT_RASTER_PRECIS = 6                                                       '��ϵͳ�������ͬ������ʱ��ָʾ����ӳ����ѡ���դ����
    OUT_TT_ONLY_PRECIS = 7                                                      'ָʾ����ӳ��������TrueType�����н���ѡ�����ϵͳ��û�а�װTrueType���壬����ӳ���������ص�Ĭ����Ϊ��
    OUT_OUTLINE_PRECIS = 8                                                      'ָʾ����ӳ������TrueType���������ڴ�ٵ������н���ѡ��
End Enum

'������þ�������
Public Enum ClipPrecisionType
    CLIP_DEFAULT_PRECIS = 0                                                     'ָ��Ĭ�ϼ�����Ϊ
    CLIP_CHARACTER_PRECIS = 1                                                   'δʹ��
    CLIP_STROKE_PRECIS = 2                                                      '����ӳ������ʹ�ã�����ö�ٹ�դ��ʸ����TrueType����ʱ����*
    CLIP_MASK = 15                                                              'δʹ��
    CLIP_LH_ANGLES = 16                                                         '�����������תȡ��������ϵ�ķ��������ֻ�������*
    CLIP_TT_ALWAYS = 32                                                         'δʹ��
    CLIP_EMBEDDED = 128                                                         '����ָ���˱�־����ʹ��Ƕ���ֻ������
    'ע1��CLIP_STROKE_PRECIS - Ϊ�˼��ݣ�ö������ʱʼ�շ��ش�ֵ��
    'ע2��CLIP_LH_ANGLES - ���δʹ�ã��豸����ʼ����ʱ����ת���������������תȡ��������ϵ�ķ���
End Enum

'������������
Public Enum FontQualityType
    DEFAULT_QUALITY = 0                                                         'ʹ��Ĭ�ϵ�������������
    DRAFT_QUALITY = 1                                                           '�ݸ��������߼��������Եľ�ȷƥ������������������
    PROOF_QUALITY = 2                                                           '�������������������������߼��������Եľ�ȷƥ�䣩
    NONANTIALIASED_QUALITY = 3                                                  '����ʼ��Ϊ�ǿ����*
    ANTIALIASED_QUALITY = 4                                                     '�������֧�ָ����壬���������С����̫С��̫��������ʼ��Ϊ�����
    'ע�����ANTIALIASED_QUALITY��NONANTIALIASED_QUALITY��δѡ�У�������û��ڿ��������ѡ��ƽ����Ļ����ʱ������ŻΌ���
End Enum

'�ַ�����ʽ��־
Public Enum StringFormatFlags
    StringFormatFlagsNoUse = &H0                                                '��ʹ��
    StringFormatFlagsDirectionRightToLeft = &H1                                 '���ҵ����˳��
    StringFormatFlagsDirectionVertical = &H2                                    '��ֱ���Ƶ����ı���
    StringFormatFlagsNoFitBlackBox = &H4                                        '�������ַ���ͣ���ַ����Ĳ��־�����
    StringFormatFlagsDisplayFormatControl = &H20                                'ʹ�ô������ַ���ʾUnicode��ʽ�����ַ�
    StringFormatFlagsNoFontFallback = &H400                                     '�滻�ַ�������Ч���ַ���Ĭ�ϵ�ȱʧ�ַ�Ϊ��ض����ȥ������һƲ��
    StringFormatFlagsMeasureTrailingSpaces = &H800                              '��ĩ�ո�������ַ���������
    StringFormatFlagsNoWrap = &H1000                                            '�����ı�����
    StringFormatFlagsLineLimit = &H2000                                         '�ڲ��־��������Ʋ�������
    StringFormatFlagsNoClip = &H4000                                            '������ʾ���ڲ��־����Ϸ����ַ��Ͳ��־�����������ı�[?]
    StringFormatFlagsBypassGDI = &H80000000                                     '�ƹ�GDI����[?]
End Enum

'�ַ����ü�
Public Enum StringTrimming
    StringTrimmingNone = 0                                                      '���ü�
    StringTrimmingCharacter = 1                                                 '�ڲ��־��������һ���ַ��ı߽紦�Ͽ��ַ�����Ĭ�ϣ�
    StringTrimmingWord = 2                                                      '�ڲ��־��������һ�����ʵı߽紦�Ͽ��ַ���
    StringTrimmingEllipsisCharacter = 3                                         '�ڲ��־��������һ���ַ��ı߽紦�Ͽ��ַ����������ַ�������롰...��
    StringTrimmingEllipsisWord = 4                                              '�ڲ��־��������һ�����ʵı߽紦�Ͽ��������ַ�������롰...��
    StringTrimmingEllipsisPath = 5                                              '�ڲ��־��������һ��·���ı߽紦�Ͽ��������ַ�������롰...��[?]
End Enum

'�ַ��������滻�������ã�
Public Enum StringDigitSubstitute
    StringDigitSubstituteUser = 0                                               '�û�
    StringDigitSubstituteNone = 1                                               '����
    StringDigitSubstituteNational = 2                                           '���չ��ң��������滻
    StringDigitSubstituteTraditional = 3                                        '���ձ����趨�滻
End Enum

'��ˢ��Ӱ��ʽ�������ã�
Public Enum HatchStyle
    HatchStyleHorizontal                                                        'ˮƽ��
    HatchStyleVertical                                                          '��ֱ��
    HatchStyleForwardDiagonal                                                   '�ܣ�����ݣ�
    HatchStyleBackwardDiagonal                                                  '��������ݣ�
    HatchStyleCross                                                             'ˮƽ�ߺʹ�ֱ�߽���
    HatchStyleDiagonalCross                                                     'б�߽��棨����ݣ�
    HatchStyle05Percent                                                         '��Ӱ����Ϊ5%������ö��������
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
    HatchStyleLightDownwardDiagonal                                             '�ܣ���HatchStyleForwardDiagonal���ܣ�������ݣ�
    HatchStyleLightUpwardDiagonal                                               '������HatchStyleBackwardDiagonal���ܣ�������ݣ�
    HatchStyleDarkDownwardDiagonal                                              '��ϸΪHatchStyleLightDownwardDiagonal�Ķ�����������ͬ
    HatchStyleDarkUpwardDiagonal                                                '��ϸΪHatchStyleLightUpwardDiagonal�Ķ�����������ͬ
    HatchStyleWideDownwardDiagonal                                              '
    HatchStyleWideUpwardDiagonal                                                '
    HatchStyleLightVertical                                                     '
    HatchStyleLightHorizontal                                                   '
    HatchStyleNarrowVertical                                                    '
    HatchStyleNarrowHorizontal                                                  '
    HatchStyleDarkVertical                                                      '
    HatchStyleDarkHorizontal                                                    '
    HatchStyleDashedDownwardDiagonal                                            '�ɣ���ɵ�ˮƽ��
    HatchStyleDashedUpwardDiagonal                                              '�ɣ���ɵ�ˮƽ��
    HatchStyleDashedHorizontal                                                  'ˮƽ����
    HatchStyleDashedVertical                                                    '��ֱ����
    HatchStyleSmallConfetti                                                     'С�ߵ㡤
    HatchStyleLargeConfetti                                                     '��ߵ��
    HatchStyleZigZag                                                            '�����ˮƽ��
    HatchStyleWave                                                              '������ˮƽ��
    HatchStyleDiagonalBrick                                                     '��ש��
    HatchStyleHorizontalBrick                                                   'ˮƽש��
    HatchStyleWeave                                                             '������֯
    HatchStylePlaid                                                             '�ո������ӻ���
    HatchStyleDivot                                                             '��Ƥ��
    HatchStyleDottedGrid                                                        'ˮƽ��������
    HatchStyleDottedDiamond                                                     'б�߽�������
    HatchStyleShingle                                                           '��Ƭ
    HatchStyleTrellis                                                           '���
    HatchStyleSphere                                                            '��������
    HatchStyleSmallGrid                                                         'С����
    HatchStyleSmallCheckerBoard                                                 'С�������
    HatchStyleLargeCheckerBoard                                                 '���������
    HatchStyleOutlinedDiamond                                                   'б�߽��棨������ݣ�
    HatchStyleSolidDiamond                                                      'б����������
    HatchStyleTotal                                                             '����Ӱ��������͸����
    HatchStyleLargeGrid = HatchStyleCross
    HatchStyleMin = HatchStyleHorizontal
    HatchStyleMax = HatchStyleTotal - 1
End Enum

'���ʶ���
Public Enum PenAlignment
    PenAlignmentCenter = 0                                                      '�����ڻ��Ƶ����������Ķ���
    PenAlignmentInset = 1                                                       '�ڻ��ƶ����ʱ�������ڶ���α�Ե���ڲ�����
End Enum

'��ˢ����
Public Enum BrushType
    BrushTypeSolidColor = 0                                                     'ʵɫ
    BrushTypeHatchFill = 1                                                      '��Ӱ���ʣ��μ�HatchStyleö�٣�
    BrushTypeTextureFill = 2                                                    '������
    BrushTypePathGradient = 3                                                   '��·������
    BrushTypeLinearGradient = 4                                                 '���Խ���
End Enum

'������ʽ
Public Enum DashStyle
    DashStyleSolid                                                              '������������
    DashStyleDash                                                               '- - - -
    DashStyleDot                                                                '��������������
    DashStyleDashDot                                                            '-��-��-��-��
    DashStyleDashDotDot                                                         '-����-����-����
    DashStyleCustom                                                             '�û��Զ���
End Enum

'���߶˵���״
Public Enum DashCap
    DashCapFlat = 0                                                             'ƽͷ
    DashCapRound = 2                                                            '��
    DashCapTriangle = 3                                                         '�� *
    'ע��������ʶ�������ΪPenAlignmentInset������ʹ��DashCapTriangle��
End Enum

'ֱ�߶˵���״
Public Enum LineCap
    LineCapFlat = 0                                                             'ƽͷ
    LineCapSquare = 1                                                           '��
    LineCapRound = 2                                                            '��
    LineCapTriangle = 3                                                         '��
    LineCapNoAnchor = &H10                                                      '��ê��
    LineCapSquareAnchor = &H11                                                  '��ê��
    LineCapRoundAnchor = &H12                                                   '��ê��
    LineCapDiamondAnchor = &H13                                                 '��ê��
    LineCapArrowAnchor = &H14                                                   '��ê��
    LineCapCustom = &HFF                                                        '�Զ��壨�μ�CustomLineCapType��
    LineCapAnchorMask = &HF0                                                    '����ê��[?]
End Enum

'�Զ���ֱ�߶˵���״����
Public Enum CustomLineCapType
    CustomLineCapTypeDefault = 0                                                'Ĭ��
    CustomLineCapTypeAdjustableArrow = 1                                        '����Ӧ��ͷ
End Enum

'���߽�����ʽ
Public Enum LineJoin
    LineJoinMiter = 0                                                           'б������[?]
    LineJoinBevel = 1                                                           'б������
    LineJoinRound = 2                                                           'Բ������
    LineJoinMiterClipped = 3                                                    'б�����Ӳ��������ಿ��[?]
End Enum

'�������ͣ��μ�BrushType
Public Enum PenType
    PenTypeSolidColor = BrushTypeSolidColor
    PenTypeHatchFill = BrushTypeHatchFill
    PenTypeTextureFill = BrushTypeTextureFill
    PenTypePathGradient = BrushTypePathGradient
    PenTypeLinearGradient = BrushTypeLinearGradient
    PenTypeUnknown = -1                                                         'δ֪
End Enum

'ͼԪ�ļ����ͣ������ã�
Public Enum MetafileType
    MetafileTypeInvalid                                                         'Gdip�в���ʶ���ͼԪ�ļ���ʽ
    MetafileTypeWmf                                                             'WMF�ļ���ֻ����GDI��¼��
    MetafileTypeWmfPlaceable                                                    'WMF�ļ����ļ�ǰ����һ���ɷ��õ�ͼԪ�ļ���ͷ��[?]
    MetafileTypeEmf                                                             'EMF�ļ���ֻ����GDI��¼��
    MetafileTypeEmfPlusOnly                                                     'EMF�ļ���ֻ����GDI+��¼��
    MetafileTypeEmfPlusDual                                                     'EMF�ļ�������GDI+��¼��GDI��¼��ʹ��GDI���ƽ����������½���
    'WMF��WindowsͼԪ�ļ���Windows Metafile�����ɼ򵥵������ͷ��������ͼ�Σ���ɵ�ʸ��ͼ��
    'EMF����ǿ��ͼԪ�ļ���Enhanced Metafile���Ƕ�WMF�ĸĽ�����չ��֧�ָ������ɫ�͸����ӵ�ͼ���ʾ��
End Enum

'EMF�ļ����ͣ������ã�
Public Enum EmfType
    EmfTypeEmfOnly = MetafileTypeEmf
    EmfTypeEmfPlusOnly = MetafileTypeEmfPlusOnly
    EmfTypeEmfPlusDual = MetafileTypeEmfPlusDual
End Enum

'��������
Public Enum ObjectType
    ObjectTypeInvalid                                                           '��Ч�ģ�������
    ObjectTypeBrush                                                             '��ˢ
    ObjectTypePen                                                               '����
    ObjectTypePath                                                              '·��
    ObjectTypeRegion                                                            '����
    ObjectTypeImage                                                             'ͼ��
    ObjectTypeFont                                                              '����
    ObjectTypeStringFormat                                                      '�ַ�����ʽ
    ObjectTypeImageAttributes                                                   'ͼ������
    ObjectTypeCustomLineCap                                                     '�Զ���ֱ�߶˵�
    ObjectTypeGraphics                                                          'ͼ��
    ObjectTypeMax = ObjectTypeGraphics                                          '
    ObjectTypeMin = ObjectTypeBrush                                             '
End Enum

'ͼԪ�ļ���ܾ��ζ�����λ
Public Enum MetafileFrameUnit
    MetafileFrameUnitPixel = UnitPixel
    MetafileFrameUnitPoint = UnitPoint
    MetafileFrameUnitInch = UnitInch
    MetafileFrameUnitDocument = UnitDocument                                    '�ĵ����壨ͨ��Ϊ 1/300 Ӣ�磩
    MetafileFrameUnitMillimeter = UnitMillimeter
    MetafileFrameUnitGdi                                                        '1/100 ���ף���GDI���ݣ�
End Enum

'����ռ䣨�����ã�
Public Enum CoordinateSpace
    CoordinateSpaceWorld                                                        '���綨��[?]
    CoordinateSpacePage                                                         'ҳ�涨��
    CoordinateSpaceDevice                                                       '�豸����
End Enum

'�ȼ�ǰ׺[?]
Public Enum HotkeyPrefix
    HotkeyPrefixNone = 0                                                        '��ǰ׺
    HotkeyPrefixShow = 1                                                        '��ʾǰ׺
    HotkeyPrefixHide = 2                                                        '����ǰ׺
End Enum

'ˢ�»������������ã�
Public Enum FlushIntention
    FlushIntentionFlush = 0                                                     'ˢ��������������ֲ����������ڳ��ֲ������֮ǰ���أ�
    FlushIntentionSync = 1                                                      'ˢ��������������ֲ������ڳ��ֲ�����ɺ�Ż᷵�أ�
End Enum

'����������ֵ���ͣ��μ�ImagePropertyType��
Public Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1
    EncoderParameterValueTypeASCII = 2
    EncoderParameterValueTypeShort = 3
    EncoderParameterValueTypeLong = 4
    EncoderParameterValueTypeRational = 5
    EncoderParameterValueTypeLongRange = 6                                      '����Χ
    EncoderParameterValueTypeUndefined = 7                                      'δ����
    EncoderParameterValueTypeRationalRange = 8                                  'ʵ����Χ[?]
End Enum

'������ֵ[?]
Public Enum EncoderValue
    EncoderValueColorTypeCMYK                                                   'CMYK��ɫģʽ��GDIP 1.0 ����Ч��
    EncoderValueColorTypeYCCK                                                   'YCCK��ɫģʽ[?]��GDIP 1.0 ����Ч��
    EncoderValueCompressionLZW                                                  'ʹ��LZW�㷨*ѹ��Tiff��ʽͼ��
    EncoderValueCompressionCCITT3                                               'ʹ��CCTII3�㷨*ѹ��Tiff��ʽͼ��
    EncoderValueCompressionCCITT4                                               'ʹ��CCTII4�㷨ѹ��Tiff��ʽͼ��
    EncoderValueCompressionRle                                                  'ʹ��RLE�㷨*ѹ��Tiff��ʽͼ��
    EncoderValueCompressionNone                                                 '��ѹ��Tiff��ʽͼ��
    EncoderValueScanMethodInterlaced                                            '����ɨ�跽��[?]��GDIP 1.0 ����Ч��
    EncoderValueScanMethodNonInterlaced                                         '�ǽ���ɨ�跽��[?]��GDIP 1.0 ����Ч��
    EncoderValueVersionGif87                                                    '[?]
    EncoderValueVersionGif89                                                    '[?]
    EncoderValueRenderProgressive                                               '���з�ʽ��Ⱦ[?]��GDIP 1.0 ����Ч��
    EncoderValueRenderNonProgressive                                            '�����з�ʽ��Ⱦ[?]��GDIP 1.0 ����Ч��
    EncoderValueTransformRotate90                                               '��Jpegͼ��*˳ʱ����ת90�㣨����ʧ��
    EncoderValueTransformRotate180                                              '��Jpegͼ��˳ʱ����ת180�㣨����ʧ��
    EncoderValueTransformRotate270                                              '��Jpegͼ��˳ʱ����ת270�㣨����ʧ��
    EncoderValueTransformFlipHorizontal                                         '��Jpegͼ��ˮƽ��ת������ʧ��
    EncoderValueTransformFlipVertical                                           '��Jpegͼ��ֱ��ת������ʧ��
    EncoderValueMultiFrame                                                      'ͼ����ö�֡����
    EncoderValueLastFrame                                                       '��֡����ͼ������һ֡
    EncoderValueFlush                                                           '�رձ���������
    EncoderValueFrameDimensionTime                                              '��ʱ�䶨���֡ά��[?]��GDIP 1.0 ����Ч��
    EncoderValueFrameDimensionResolution                                        '�Էֱ��ʶ����֡ά��[?]��GDIP 1.0 ����Ч��
    EncoderValueFrameDimensionPage                                              '��ҳ�涨���֡ά��[?]��GDIP 1.0 ����Ч��
    EncoderValueColorTypeGray                                                   '�Ҷ���ɫ
    EncoderValueColorTypeRGB                                                    'RGB��ɫ
    'ע1��LZW�㷨����������ѹ���㷨����Lempel-Ziv-Welch Encoding�������㷨ͨ������һ���ַ������ý϶̵Ĵ�������ʾ�ϳ�
    '���ַ�����ʵ��ѹ����
    'ע2����δ�ҵ�CCTII3�㷨��CCTII4�㷨���ܡ�
    'ע3��RLE�㷨�������г̳���ѹ���㷨����Run Length Encoding�������㷨��һ����ʾ�������ȵ������ֽڼ���һ�����ݿ飬��
    '����ԭ�����������ɿ����ݣ��Ӷ��ﵽ��ʡ�洢�ռ��Ŀ�ġ�
    'ע4��Jpeg����������ͼ��ר���顱��Joint Photographic Experts Group����JPEGͼ���ʽ��������ѹ����ʽ������õ�ͼ
    '���ļ���ʽ֮һ��
End Enum

'λͼѹ��ģʽ
Public Enum BitmapCompressionMode
    BL_RGB = &H0&
    BI_RLE8 = &H1&
    BI_RLE4 = &H2&
    BI_BITFIELDS = &H3&
    BI_JPEG = &H4&
    BI_PNG = &H5&
End Enum

'Debug�¼����𣨲����ã�
Public Enum DebugEventLevel
    DebugEventLevelFatal                                                        '�ؼ�����[?]
    DebugEventLevelWarning                                                      '����
End Enum

'����Gdip�������ص�״̬
Public Enum GpStatus
    Ok = 0                                                                      '�ɹ�����Gdip����
    GenericError = 1                                                            '����Gdip����ʱ������һ���ԵĴ���
    InvalidParameter = 2                                                        '����Gdip����ʱ����Ĳ�����Ч
    OutOfMemory = 3                                                             '����Gdip����ʱ�ڴ治��
    ObjectBusy = 4                                                              '����Gdip����ʱĿ�����æµ����Ӧ
    InsufficientBuffer = 5                                                      '����Gdip����ʱ��������С����
    NotImplemented = 6                                                          '����Gdip����ʱ��δʵ�ֲ���
    Win32Error = 7                                                              '����Win32����
    WrongState = 8                                                              '״̬����
    Aborted = 9                                                                 '����Gdip����ʱ��������ֹ
    FileNotFound = 10                                                           '�Ҳ����ļ�
    ValueOverflow = 11                                                          '����Gdip����ʱ����ֵ���
    AccessDenied = 12                                                           '����Gdip����ʱ���ʱ��ܾ�
    UnknownImageFormat = 13                                                     'δ֪��ͼ���ʽ
    FontFamilyNotFound = 14                                                     '�Ҳ���������
    FontStyleNotFound = 15                                                      '�Ҳ�����������
    NotTrueTypeFont = 16                                                        '����TrueType��ʽ����
    UnsupportedGdiplusVersion = 17                                              'ʹ�õ��ǲ�֧�ֵ�GDIP�汾
    GdiplusNotInitialized = 18                                                  'GDIPδ��ʼ��
    PropertyNotFound = 19                                                       '�Ҳ�����Ӧ����
    PropertyNotSupported = 20                                                   '��֧�ֵĶ�Ӧ����
End Enum

'ͼ���ļ���׺
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

'Gdipͨ�ö���
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

'EMF+�ļ���¼����
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

'���㷽ʽ
Public Enum CalculatingMethod
    RoundDown                                                                   '����ȡ��
    RoundNear                                                                   '�ٽ�ȡ��
    RoundUp                                                                     '����ȡ��
End Enum

'################  ��  ��  ��  ################
'��
Public Type PointL
    X As Long
    Y As Long
End Type

Public Type PointF
    X As Single
    Y As Single
End Type

'����
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

'�ߴ�
Public Type SizeL
    Width As Long                                                               '��
    Height As Long                                                              '��
End Type

Public Type SizeF
    Width As Single
    Height As Single
End Type

'RGB��ɫ
Public Type RGBQuad
    rgbBlue As Byte                                                             '��ɫ����
    rgbGreen As Byte                                                            '��ɫ����
    rgbRed As Byte                                                              '��ɫ����
    rgbReserved As Byte                                                         '����ֵ������Ϊ0
End Type

'ARGB��ɫ
Public Type ARGBColor
    Alpha As Byte                                                               '��������͸����*��
    Red As Byte
    Green As Byte
    Blue As Byte
    'ע��͸����ֵ��0����ȫ͸������255����ȫ��͸������
End Type

'λͼ��Ϣͷ��
Public Type BitmapInfoHeader
    biSize As Long                                                              'λͼ��Ϣͷ���Ĵ�С�����ֽڼ��㣩
    biWidth As Long                                                             'λͼ��ȣ������ؼ��㣬��ͬ��
    biHeight As Long                                                            'λͼ�߶�
    biPlanes As Integer                                                         'Ŀ���豸����������Ϊ1
    biBitCount As Integer                                                       '��¼ÿ����������Ҫ��λ��Bit����
    biCompression As BitmapCompressionMode                                      'ͼƬ���õ�ѹ����ʽ��Ĭ��Ϊ��ѹ����BL_RGB��
    biSizeImage As Long                                                         'ͼ��Ĵ�С*�����ֽڼ��㣩
    biXPelsPerMeter As Long                                                     'λͼ��Ŀ���豸��ˮƽ�ֱ��ʣ�������/�׼��㣩
    biYPelsPerMeter As Long                                                     'λͼ��Ŀ���豸�Ĵ�ֱ�ֱ��ʣ�������/�׼��㣩
    biClrUsed As Long                                                           '��ɫ������ɫ������Ĭ��Ϊ0
    biClrImportant As Long                                                      '��Ҫ��ɫ��������Ĭ��Ϊ0*
    'ע1������δѹ����λͼ��biSizeImage��ֵĬ��Ϊ0��
    'ע2�����biClrImportant��ֵΪ0�����ʾ������ɫ������Ҫ��
End Type

'λͼ��Ϣ
Public Type BitmapInfo
    bmiHeader As BitmapInfoHeader
    bmiColors As RGBQuad
End Type

'λͼ����
Public Type BitmapData
    Width As Long
    Height As Long
    Stride As Long                                                              'λͼ����Ŀ����
    PixelFormat As Long                                                         '���ظ�ʽ
    Scan0 As Long                                                               'λͼ�е�һ���������ݵĵ�ַ
    Reserved As Long                                                            '����ֵ������Ϊ0
End Type

'��ɫ����
Public Type ColorMatrix
    Matrix(0 To 4, 0 To 4) As Double
End Type

'·������
Public Type PathData
    Count As Long                                                               '����
    Points As Long                                                              'ָ��PointL�����ָ��[?]
    Types As Long                                                               'ָ��Byte�����ָ��[?]
End Type

'��/����������ʶ��
Public Type ClsID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'��������������
Public Type EncoderParameter
    Guid As ClsID
    NumberOfValues As Long                                                      'ֵ������[?]
    ValueType As EncoderParameterValueType
    value As EncoderValue
End Type

'����������
Public Type EncoderParameters
    Count As Long                                                               '����
    Parameter As EncoderParameter
End Type

'��������
Public Type FontType
    Name As String
    Size As Single
    Weight As FontWeight
    Style As FontStyle
End Type

'�߼�����ṹ��Ascii��[?]
Public Type LogFontA
    lfHeight As Long                                                            '�ַ��߶�*
    lfWidth As Long                                                             '�ַ���ƽ�����*
    lfEscapement As Long                                                        'ת���������豸��x��֮��ĽǶȣ���1/10����㣩
    lfOrientation As Long                                                       'ÿ���ַ��Ļ��ߺ��豸x��֮��ĽǶȣ���1/10����㣩
    lfWeight As FontWeight                                                      '���أ��ֵĴ�ϸ��
    lfItalic As Byte                                                            'б��
    lfUnderline As Byte                                                         '�»���
    lfStrikeOut As Byte                                                         'ɾ����
    lfCharSet As CharSetType                                                    '�ַ���
    lfOutPrecision As OutPrecisionType                                          '�������*
    lfClipPrecision As ClipPrecisionType                                        '���þ���*
    lfQuality As FontQualityType                                                '�������
    lfPitchAndFamily As Byte                                                    '����ļ���ϵ��
    lfFaceName(31) As Byte                                                      '��NULL��β���ַ���ָ���������������
    'ע1�� ����ӳ���������·�ʽ����lfHeight��ָ����ֵ��
    '--lfHeight��ֵ----����----------------------------------------------------------------------------------
    '  >0              ����ӳ��������ֵת��Ϊ�豸��λ�����������������ĵ�Ԫ��߶�ƥ�䡣
    '  =0              ����ӳ����������ƥ����ʱʹ��Ĭ�ϸ߶�ֵ��
    '  <0              ����ӳ��������ֵת��Ϊ�豸��λ�����������ֵ�����������ַ��߶�ƥ�䡣
    'ע2�����lfWidthΪ�㣬���豸���ݺ�Ƚ��������������ֻ��ݺ�Ƚ���ƥ�䣬�Բ����ɲ�ֵ�ľ���ֵȷ������ӽ�ƥ���
    'ע3��lfOutPrecision���ڶ������������������ĸ߶ȡ���ȡ��ַ�����ת�塢�����������͵�ƥ��̶ȡ�
    'ע4��lfClipPrecision���ڶ�����μ��ò��ֳ�������������ַ���
End Type

'�߼�����ṹ��WideChar��[?]
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

'ͼ���/��������Ϣ
Public Type ImageCodecInfo
    ClassID As ClsID
    FormatID As ClsID
    CodecName As Long                                                           '��/����������
    DllName As Long
    FormatDescription As Long                                                   '��ʽ����
    FilenameExtension As Long
    MimeType As Long
    Flags As ImageCodecFlags
    Version As Long
    SigCount As Long
    SigSize As Long
    SigPattern As Long
    SigMask As Long
End Type

'��ɫ��
Public Type ColorPalette
    Flags As PaletteFlags
    Count As Long
    Entries(0 To 255) As Long
End Type

'WMF�ļ���ͷ��������
Public Type PwmfRect16
    Left As Integer
    Top As Integer
    Width As Integer
    Height As Integer
End Type

'WMF�ļ��ɷ��õ�ͼԪ�ļ���ͷ
Public Type WmfPlaceableFileHeader
    Key As Long
    Hmf As Integer
    BoundingBox As PwmfRect16
    Inch As Integer
    Reserved As Long
    Checksum As Integer
End Type

'ENHԪ���ݵ�ͷ������[?]
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

'Ԫ����ͷ��
Public Type METAHEADER
    mtType As Integer
    mtHeaderSize As Integer
    mtVersion As Integer
    mtSize As Long
    mtNoObjects As Integer
    mtMaxRecord As Long
    mtNoParameters As Integer
End Type

'ͼԪ�ļ�ͷ��[?]
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

'������Ŀ
Public Type PropertyItem
    iPropId As Long
    iLength As Long
    itype As Integer
    iValue As Long
End Type

'�ַ���Χ
Public Type CharacterRange
    First As Long
    Length As Long
End Type

'GDIP��ʼ������
Public Type GdiplusStartupInput
    GdiplusVersion As Long                                                      '�汾
    DebugEventCallback As Long                                                  'Debug�¼��ص�
    SuppressBackgroundThread As Long                                            '���ƺ�̨�߳�[?]
    SuppressExternalCodecs As Long                                              '�����ⲿ��/������[?]
End Type

'GDIP����
Public Type GdiplusObject
    GdiplusObjectName As String
    GdiplusObjectType As GdiplusCommonObject
    GdiplusObjectHandle As Long
End Type

'##############  ��  ��  ��  ��  ##############
'1���豸����������
Public Declare Function GdipGetDC Lib "gdiplus" (ByVal Graphics As Long, hDC As Long) As GpStatus
Public Declare Function GdipReleaseDC Lib "gdiplus" (ByVal Graphics As Long, ByVal hDC As Long) As GpStatus
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDCAPI Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

'2������
Public Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, Graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hWnd As Long, Graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHDC2 Lib "gdiplus" Alias "GdipCreateFromHdc2" (ByVal hDC As Long, ByVal hDevice As Long, Graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWNDICM Lib "gdiplus" Alias "GdipCreateFromHWndICM" (ByVal hWnd As Long, Graphics As Long) As GpStatus
Public Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, Graphics As Long) As GpStatus
Public Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Public Declare Function GdipGraphicsClear Lib "gdiplus" (ByVal Graphics As Long, ByVal lColor As Long) As GpStatus

'3�����ģʽ����Ⱦ��ƽ��ģʽ������ƫ�ơ��ı���Ⱦ��ʾ���ı�����Ͳ�ֵģʽ
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

'4������任��ҳ������
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

'5����ȡ�ֱ��ʡ��㼯�任
Public Declare Function GdipGetDpiX Lib "gdiplus" (ByVal Graphics As Long, DPI As Single) As GpStatus
Public Declare Function GdipGetDpiY Lib "gdiplus" (ByVal Graphics As Long, DPI As Single) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipTransformPointsI Lib "gdiplus" (ByVal Graphics As Long, ByVal DestSpace As CoordinateSpace, ByVal SrcSpace As CoordinateSpace, Points As PointL, ByVal Count As Long) As GpStatus

'6�������ӽ���ɫ[?]��������ɫ����ɫ��
Public Declare Function GdipGetNearestColor Lib "gdiplus" (ByVal Graphics As Long, Argb As Long) As GpStatus
Public Declare Function GdipCreateHalftonePalette Lib "gdiplus" () As Long

'7���������
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

'8�����ӻ����κͿ��ӻ����ж�
Public Declare Function GdipIsVisibleRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, Result As Long) As GpStatus
Public Declare Function GdipIsVisiblePointI Lib "gdiplus" (ByVal Graphics As Long, ByVal X As Long, ByVal Y As Long, Result As Long) As GpStatus

'9���洢�͸�ԭ����
Public Declare Function GdipSaveGraphics Lib "gdiplus" (ByVal Graphics As Long, State As Long) As GpStatus
Public Declare Function GdipRestoreGraphics Lib "gdiplus" (ByVal Graphics As Long, ByVal State As Long) As GpStatus

'10����������
Public Declare Function GdipBeginContainerI Lib "gdiplus" (ByVal Graphics As Long, dstRect As RectL, srcRect As RectL, ByVal Unit As GpUnit, State As Long) As GpStatus
Public Declare Function GdipBeginContainer2 Lib "gdiplus" (ByVal Graphics As Long, State As Long) As GpStatus
Public Declare Function GdipEndContainer Lib "gdiplus" (ByVal Graphics As Long, ByVal State As Long) As GpStatus

'11�������߶Ρ����ߡ����������ߡ����Ρ���Բ�����Ρ�����Ρ��������ߡ�����������ߡ�·��
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

'12�������Ρ���Բ�����Ρ�����Ρ�����������ߡ�·��������
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

'13��ͼ�����
Public Declare Function GdipDrawImageI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long) As GpStatus
Public Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawImagePointsI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, DstPoints As PointL, ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal X As Long, ByVal Y As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal SrcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal SrcUnit As GpUnit, Optional ByVal ImageAttributes As Long = 0, Optional ByVal CallBack As Long = 0, Optional ByVal CallBackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, Points As PointL, ByVal Count As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal SrcUnit As GpUnit, Optional ByVal ImageAttributes As Long = 0, Optional ByVal CallBack As Long = 0, Optional ByVal CallBackData As Long = 0) As GpStatus

'14��ͼ���/����������
Public Declare Function GdipGetImageDecoders Lib "gdiplus" (ByVal NumDecoders As Long, ByVal Size As Long, Decoders As Any) As GpStatus
Public Declare Function GdipGetImageDecodersSize Lib "gdiplus" (NumDecoders As Long, Size As Long) As GpStatus
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Declare Function GdipGetImageEncodersSize Lib "gdiplus" (NumEncoders As Long, Size As Long) As GpStatus
Public Declare Function GdipGetImageEncoders Lib "gdiplus" (ByVal NumEncoders As Long, ByVal Size As Long, encoders As Any) As GpStatus
Public Declare Function GdipGetEncoderParameterListSize Lib "gdiplus" (ByVal Image As Long, ClsIDEncoder As ClsID, Size As Long) As GpStatus
Public Declare Function GdipGetEncoderParameterList Lib "gdiplus" (ByVal Image As Long, ClsIDEncoder As ClsID, ByVal Size As Long, Buffer As EncoderParameters) As GpStatus

'15��ͼ����ء��ͷš����ơ�����
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

'16��ͼ�������Ϣ����ز���
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

'17�����Բ���
Public Declare Function GdipGetPropertyCount Lib "gdiplus" (ByVal Image As Long, NumOfProperty As Long) As GpStatus
Public Declare Function GdipGetPropertyIdList Lib "gdiplus" (ByVal Image As Long, ByVal NumOfProperty As Long, List As Long) As GpStatus
Public Declare Function GdipGetPropertyItemSize Lib "gdiplus" (ByVal Image As Long, ByVal PropId As Long, Size As Long) As GpStatus
Public Declare Function GdipGetPropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal PropId As Long, ByVal PropSize As Long, Buffer As PropertyItem) As GpStatus
Public Declare Function GdipGetPropertySize Lib "gdiplus" (ByVal Image As Long, TotalBufferSize As Long, NumProperties As Long) As GpStatus
Public Declare Function GdipGetAllPropertyItems Lib "gdiplus" (ByVal Image As Long, ByVal TotalBufferSize As Long, ByVal NumProperties As Long, AllItems As PropertyItem) As GpStatus
Public Declare Function GdipRemovePropertyItem Lib "gdiplus" (ByVal Image As Long, ByVal PropId As Long) As GpStatus
Public Declare Function GdipSetPropertyItem Lib "gdiplus" (ByVal Image As Long, Item As PropertyItem) As GpStatus

'18���������
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

'19��ֱ�߶˵㡢��ͷ�˵�
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

'20��λͼ���
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

'21����ˢ����ˢ��Ӱ�����
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

'22����ͼ���
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

'23��·������
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

'24������
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

'25������
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

'26��ͼ������
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

'27���ַ������߼�����
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

'28���ַ������ַ�����ʽ
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

'29��ͼԪ�ļ�
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

'30����������
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

'##################  ��  ��  ##################
'1��ͼ�������
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
'2��DC���
Private Const HORZRES As Long = 8&
Private Const VERTRES As Long = 10&
'##################  ��  ��  ##################

Private mToken As Long                                                          '�����ΪGDIP�Ƿ񱻳�ʼ�������ݣ�
Private Objects() As GdiplusObject, ObjCount As Long                            'Gdipͨ�ö���

'##################  ��  ��  /  ��  ��  ##################

'ȡ�����еĽϴ�ֵ����Сֵ
Private Function Max(ByVal A As Long, ByVal B As Long) As Long
    Max = IIf(A > B, A, B)
End Function

Private Function Min(ByVal A As Long, ByVal B As Long) As Long
    Min = IIf(A < B, A, B)
End Function

'�����ļ���׺��ö�Ӧ����������ʶ��
Public Function GetImageEncoderClsid(ByVal FileSuffix As ImageFileSuffix) As ClsID
    CLSIDFromString StrPtr(ImageEncoderPrefix & CInt(FileSuffix) & ImageEncoderSuffix), GetImageEncoderClsid
End Function

'��ʼ��GDIPlus
Public Sub InitGDIPlus(Optional ByVal ShowLog As Boolean = False)
    Dim uInput As GdiplusStartupInput, Retn As GpStatus
    If mToken <> 0 Then
        If ShowLog Then Debug.Print "GdiPlus�ѳ�ʼ����"
        Exit Sub
    End If
    uInput.GdiplusVersion = 1
    Retn = GdiplusStartup(mToken, uInput)
    If Retn <> Ok Then
        If ShowLog Then Debug.Print "GdiPlusδ�ܳɹ���ʼ��������ԭ��" & _
        Choose(Retn, "����Gdip����ʱ������һ���ԵĴ���", _
        "����Gdip����ʱ����Ĳ�����Ч", _
        "����Gdip����ʱ�ڴ治��", _
        "����Gdip����ʱĿ�����æµ����Ӧ", _
        "����Gdip����ʱ��������С����", _
        "����Gdip����ʱ��δʵ�ֲ���", _
        "����Win32����", _
        "״̬����", _
        "����Gdip����ʱ��������ֹ", _
        "�Ҳ����ļ�", _
        "����Gdip����ʱ����ֵ���", _
        "����Gdip����ʱ���ʱ��ܾ�", _
        "δ֪��ͼ���ʽ", _
        "�Ҳ���������", _
        "�Ҳ�����������", _
        "����TrueType��ʽ����", _
        "ʹ�õ��ǲ�֧�ֵ�GDIP�汾", _
        "GDIPδ��ʼ��", _
        "�Ҳ�����Ӧ����", _
        "��֧�ֵĶ�Ӧ����")
    Else
        If ShowLog Then Debug.Print "GdiPlus��ʼ����ɡ�"
    End If
End Sub

'�ر�GDIPlus
Public Sub CloseGDIPlus(Optional ByVal ShowLog As Boolean = False)
    If mToken = 0 Then Exit Sub
    DeleteAllGdipCommonObjects
    GdiplusShutdown mToken
    mToken = 0
    If ShowLog Then Debug.Print "GdiPlus�ѹرա�"
End Sub

'�½�����
Public Sub NewGdipCommonObject(ByVal ObjName As String, ByVal ObjType As GdiplusCommonObject, ByVal ObjHandle As Long)
    Dim T As Long
    If ObjCount > 0 Then                                                        '����GDIPͨ�ö�����Ҫ����������⣬��ȷ������Ψһ��
        For T = 0 To ObjCount - 1
            If Objects(T).GdiplusObjectName = ObjName Then
                MsgBox "���� NewGdipCommonObject ʱ�����Ѵ�������Ϊ��" & ObjName & "����Gdiplusͨ�ö���", vbCritical, "����"
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

'NewGdipCommonObject���̵ļ򻯵���
Public Sub AddGCO(ByVal ObjName As String, ByVal ObjType As GdiplusCommonObject, ByVal ObjHandle As Long)
    NewGdipCommonObject ObjName, ObjType, ObjHandle
End Sub

'ɾ�����ж���
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

'���������Ƴ�����
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

'RemoveGdipCommonObject���̵ļ򻯵���
Public Sub DelGCO(ByVal ObjName As String)
    RemoveGdipCommonObject ObjName
End Sub

'���������Ƴ������ڲ�ʹ�ã�
Private Sub RemoveGdipCommonObjectByIndex(ByVal Index As Long)
    InitGDIPlus
    If Index < 0 Or Index >= ObjCount Then Exit Sub
    Dim T As Long
    Select Case Objects(Index).GdiplusObjectType
    Case GdiplusCommonObject.GdiplusBrush                                       '��ˢ
        GdipDeleteBrush Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusMatrix                                      '����
        GdipDeleteMatrix Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusPen                                         '����
        GdipDeletePen Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusStringFormat                                '�ַ�����ʽ
        GdipDeleteStringFormat Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusFont                                        '����
        GdipDeleteFont Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusFontFamily                                  '�ַ���
        GdipDeleteFontFamily Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusGraphics                                    '����
        GdipDeleteGraphics Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusPath                                        '·��
        GdipDeletePath Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusRegion                                      '����
        GdipDeleteRegion Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusPathIter                                    '·��������
        GdipDeletePathIter Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusCachedBitmap                                '����λͼ
        GdipDeleteCachedBitmap Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusImage                                       'ͼ��
        GdipDisposeImage Objects(Index).GdiplusObjectHandle
    Case GdiplusCommonObject.GdiplusDeviceContext                               '�豸����������
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

'�������ֻ�ö�����������0��ʾ��ȡ�����Ч��
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

'GetGdipCommonObjectHandle�����ļ򻯵���
Public Function GetGCO(ByVal ObjectName As String) As Long
    GetGCO = GetGdipCommonObjectHandle(ObjectName)
End Function

'�½���
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

'���ֵ�任
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

'�½��������͵������
'PointsText�ĸ�ʽΪ��X1, Y1, X2, Y2 ��
Public Function NewArrayOfPointL(ByVal PointsText As String, Optional ByVal Delimiter As String = ",") As PointL()
    If Delimiter = "" Then
        MsgBox "��Ч�ķָ������ָ���Ϊһ���ַ���", vbCritical, "����"
        Exit Function
    ElseIf PointsText = "" Then
        MsgBox "PointsText����Ϊ�ա�", vbCritical, "����"
        Exit Function
    End If
    Dim Retn() As PointL, T As Long, s() As String
    s = Split(PointsText, Delimiter)
    If UBound(s) Mod 2 = 0 Then
        MsgBox "���� PointsText �������ĵ��X�����Y��������������һ�¡�", vbCritical, "����"
        Exit Function
    End If
    ReDim Retn((UBound(s) - 1) / 2) As PointL
    For T = 0 To UBound(Retn)
        Retn(T) = NewPoint(Val(s(T * 2)), Val(s(T * 2 + 1)))
    Next T
    NewArrayOfPointL = Retn
End Function

'�½�����
Public Function NewRect(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long) As RectL
    If Width <= 0 Or Height <= 0 Then
        MsgBox "��Ч������ֵ��", vbCritical, "����"
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
        MsgBox "��Ч������ֵ��", vbCritical, "����"
        Exit Function
    End If
    With NewRectFloat
        .Left = Left
        .Top = Top
        .Right = Width + Left
        .Bottom = Height + Top
    End With
End Function

'���־��α任
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

'�½��ߴ�
Public Function NewSize(ByVal Width As Long, ByVal Height As Long) As SizeL
    If Width <= 0 Or Height <= 0 Then
        MsgBox "��Ч������ֵ��", vbCritical, "����"
        Exit Function
    End If
    With NewSize
        .Width = Width
        .Height = Height
    End With
End Function

Public Function NewSizeFloat(ByVal Width As Single, ByVal Height As Single) As SizeF
    If Width <= 0 Or Height <= 0 Then
        MsgBox "��Ч������ֵ��", vbCritical, "����"
        Exit Function
    End If
    With NewSizeFloat
        .Width = Width
        .Height = Height
    End With
End Function

'���ֳߴ�任
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

'�½�ARGB��ɫ
Public Function NewARGBColor(ByVal Alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As ARGBColor
    With NewARGBColor
        .Alpha = Alpha
        .Red = Red
        .Green = Green
        .Blue = Blue
    End With
End Function

'��ARGB��ɫת��Ϊ��������ʽ
Public Function ARGBColor2Long(ByRef mARGBColor As ARGBColor) As Long
    Dim Retn As String
    Retn = Right("00" & Hex(mARGBColor.Alpha), 2) & Right("00" & Hex(mARGBColor.Red), 2) & Right("00" & Hex(mARGBColor.Green), 2) & Right("00" & Hex(mARGBColor.Blue), 2)
    ARGBColor2Long = CLng(Val("&H" & Retn))
End Function

'��vb6�Դ���OLE_COLOR��ΪARGB��ɫ
Public Function OleColor2ARGBColor(ByVal OleColor As Long, Optional ByVal Alpha As Byte = 255) As ARGBColor
    Dim R As Byte, G As Byte, B As Byte, C As String
    C = Right("000000" & Hex(OleColor), 6)
    R = CByte(Val("&h" & Right(C, 2)))
    G = CByte(Val("&h" & Mid(C, 3, 2)))
    B = CByte(Val("&h" & Left(C, 2)))
    OleColor2ARGBColor = NewARGBColor(Alpha, R, G, B)
End Function

'�½�����
Public Function NewPen(ByVal Color As Long, ByVal Width As Single) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreatePen1 Color, Width, GpUnit.UnitPixel, Retn
    NewPen = Retn
End Function

'�½���ɫ��ˢ
Public Function NewSolidBrush(ByVal Color As Long) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreateSolidFill Color, Retn
    NewSolidBrush = Retn
End Function

'�½����Խ��仭ˢ
Public Function NewGradientLineBrush(ByRef Point1 As PointL, ByRef Point2 As PointL, ByVal Color1 As Long, ByVal Color2 As Long, Optional ByVal WrapType As WrapMode = WrapModeTile) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreateLineBrushI Point1, Point2, Color1, Color2, WrapType, Retn
    NewGradientLineBrush = Retn
End Function

'�½�����ˢ
Public Function NewHatchBrush(ByVal BrushHatchStyle As HatchStyle, ByVal ForeColor As Long, ByVal BackColor As Long) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreateHatchBrush BrushHatchStyle, ForeColor, BackColor, Retn
    NewHatchBrush = Retn
End Function

'�½���ͼ��ˢ
Public Function NewChartletBrush(ByVal hImage As Long, ByRef DatumPoint As PointL, Optional ByVal mWrapMode As WrapMode = WrapModeTile, Optional ByVal ReleaseHandles As Boolean = False)
    InitGDIPlus
    Dim Retn As Long
    GdipCreateTexture hImage, mWrapMode, Retn
    GdipResetTextureTransform Retn
    GdipTranslateTextureTransform Retn, DatumPoint.X, DatumPoint.Y, MatrixOrderAppend
    NewChartletBrush = Retn
    If ReleaseHandles Then GdipDisposeImage hImage
End Function

'�½��ַ�����ʽ
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

'�½���������ͼ��ı任��
Public Function NewMatrix(ByVal m11 As Single, ByVal m12 As Single, ByVal m21 As Single, ByVal m22 As Single, ByVal dX As Single, ByVal dY As Single) As Long
    InitGDIPlus
    Dim Retn As Long
    GdipCreateMatrix Retn
    GdipSetMatrixElements Retn, m11, m12, m21, m22, dX, dY
    NewMatrix = Retn
End Function

'�½�ƽ�ƾ���
Public Function NewTranslationalMatrix(ByVal X As Long, ByVal Y As Long) As Long
    NewTranslationalMatrix = NewMatrix(1, 0, 0, 1, X, Y)
End Function

'�½����ž���
Public Function NewScalingMatrix(ByVal ScalingRatio As Single) As Long
    NewScalingMatrix = NewMatrix(ScalingRatio, 0, 0, ScalingRatio, 0, 0)
End Function

'�½���ת���󣨽Ƕȣ�
Public Function NewRotatingMatrix(ByVal Degree As Single) As Long
    Const PI As Single = 3.14159265358979
    NewRotatingMatrix = NewMatrix(Cos(Degree / 180 * PI), Sin(Degree / 180 * PI), -Sin(Degree / 180 * PI), Cos(Degree / 180 * PI), 0, 0)
End Function

'�½�����
Public Function NewFont(ByVal hFontFamily As Long, ByVal cFontSize As Single, Optional ByVal cFontStyle As FontStyle = FontStyleRegular) As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreateFont hFontFamily, cFontSize, cFontStyle, UnitPixel, Retn
    NewFont = Retn
End Function

'�½�������
Public Function NewFontFamily(ByVal cFontName As String, Optional ByVal cFontCollection As Long = 0) As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreateFontFamilyFromName StrPtr(cFontName), cFontCollection, Retn
    NewFontFamily = Retn
End Function

'�½�����
Public Function NewGraphics(ByVal hDC As Long) As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreateFromHDC hDC, Retn
    NewGraphics = Retn
End Function

'�½��ڴ�DC���������DC�ľ��
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

'�½�·��
Public Function NewPath(Optional ByVal BrushMode As FillMode = FillModeAlternate) As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreatePath BrushMode, Retn
    NewPath = Retn
End Function

'�½�������״��·��������FilletRadius���ҽ���mShape����ΪShapeTypeRoundedRectangleʱ��Ч
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
            MsgBox "��������FilletRadius��ֵ�������0��", vbCritical, "����"
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

'�½������·��
Public Function NewPolygonPath(ByRef mPoint() As PointL) As Long
    If UBound(mPoint) - LBound(mPoint) < 2 Then
        MsgBox "�������󣺵㼯���� mPoint ���������ݲ��㣬������Ҫ3������ܹ���һ����ն���Ρ�", vbCritical, "����"
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

'�½�·��������
Public Function NewPathIterator(ByVal hPath As Long) As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreatePathIter Retn, hPath
    NewPathIterator = Retn
End Function

'�½�����
Public Function NewRegion() As Long
    Dim Retn As Long
    InitGDIPlus
    GdipCreateRegion Retn
    NewRegion = Retn
End Function

'����·��������Ӧ������
Public Function NewRegionFromPath(ByVal hPath As Long, Optional ByVal mCombineMode As CombineMode = CombineModeReplace, Optional ByVal ReleaseHandles As Boolean = False) As Long
    Dim Retn As Long
    InitGDIPlus
    Retn = NewRegion
    GdipCombineRegionPath Retn, hPath, mCombineMode
    If ReleaseHandles Then GdipDeletePath hPath
    NewRegionFromPath = Retn
End Function

'���ݾ��δ�����Ӧ������
Public Function NewRegionFromRect(ByRef mRect As RectL) As Long
    Dim Retn As Long
    InitGDIPlus
    Retn = NewRegion
    GdipCreateRegionRectI mRect, Retn
    NewRegionFromRect = Retn
End Function

'�½���������
Public Function NewFontType(ByVal FontName As String, ByVal FontSize As Single, Optional ByVal Style As FontStyle = FontStyleRegular, Optional ByVal Weight As FontWeight = FW_NORMAL) As FontType
    With NewFontType
        .Name = FontName
        .Size = FontSize
        .Weight = Weight
        .Style = Style
    End With
End Function

'��StdFontת��ΪFontType
Public Function StdFont2FontType(ByRef sFont As StdFont, Optional ByVal FontSizeCalculatingMethod As CalculatingMethod = RoundDown) As FontType
    With StdFont2FontType
        .Name = sFont.Name
        .Size = CSng(Choose(FontSizeCalculatingMethod + 1, Int(sFont.Size * 4 / 3), Round(sFont.Size * 4 / 3), Abs(Int(0 - (sFont.Size * 4 / 3)))))
        .Weight = sFont.Weight
        .Style = IIf(sFont.Bold, 1, 0) + IIf(sFont.Italic, 2, 0) + IIf(sFont.Underline, 4, 0) + IIf(sFont.Strikethrough, 8, 0)
    End With
End Function

'��StdPicture�л��ͼ����hImage
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

'����ͼ�񣬷���ͼ����hImage
Public Function LoadImage(ByVal FilePath As String) As Long
    If Dir(FilePath) = "" Then Exit Function
    InitGDIPlus
    Dim Retn As Long
    GdipLoadImageFromFile StrPtr(FilePath), Retn
    LoadImage = Retn
End Function

'���ļ�������������ͼ�񣬲��洢��Gdipͨ�ö��������У����ļ�������Ϊ���ƣ�
'��׺����Suffix����������д��ʾ��������Gdip֧�ֵ�ͼ�񣬷���ֻ��������ͼ����Suffix����Ϊ.png����pngʱ��ʾֻ��������ļ����µ�����pngͼƬ�ļ���
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
        If mSuffix <> "" Then                                                   'ָ������
            If GetFileSuffix(mFile) = mSuffix Then NewGdipCommonObject Left(mFile, Len(mFile) - Len(mSuffix) - 1), GdiplusImage, LoadImage(nPath & mFile)
        ElseIf mSuffix = "" Then
            If GetFileSuffix(mFile) = "bmp" Or GetFileSuffix(mFile) = "png" Or GetFileSuffix(mFile) = "jpg" Or GetFileSuffix(mFile) = "jpeg" Or GetFileSuffix(mFile) = "gif" Then _
            NewGdipCommonObject Left(mFile, Len(mFile) - Len(mSuffix) - 1), GdiplusImage, LoadImage(nPath & mFile)
        End If
    Next objFile
End Sub

'����ͼ���ļ�
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
        MsgBox "�ļ���ʽ���󣺲�����Ч��ͼƬ�ļ���ʽ��", vbCritical, "����"
    End Select
    If ReleaseHandles Then GdipDisposeImage hImage
End Sub

'���ͼ��ĳߴ�
Public Function GetImageSize(ByVal hImage As Long) As SizeL
    InitGDIPlus
    If hImage = 0 Then Exit Function
    Dim imgWidth As Long, imgHeight As Long
    GdipGetImageWidth hImage, imgWidth
    GdipGetImageHeight hImage, imgHeight
    GetImageSize = NewSize(imgWidth, imgHeight)
End Function

'����ļ���׺
Public Function GetFileSuffix(ByVal FilePath As String) As String
    Dim nPath As String
    nPath = FilePath
    GetFileSuffix = LCase(Right(nPath, Len(nPath) - InStrRev(nPath, ".")))
End Function

'ͼ�񿽱�������hDC������ʵ��˫���壩
Public Sub CopyGraphics(ByVal hSourceGraphics As Long, ByVal hDestinationGraphics As Long, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hSrcDC As Long, hDstDC As Long, mWidth As Single, mHeight As Single
    GdipGetDC hSourceGraphics, hSrcDC
    GdipGetDC hDestinationGraphics, hDstDC
    mWidth = GetDeviceCaps(hDstDC, HORZRES)
    mHeight = GetDeviceCaps(hDstDC, VERTRES)
    BitBlt hDstDC, 0, 0, CLng(mWidth), CLng(mHeight), hSrcDC, 0, 0, vbSrcCopy
End Sub

'�ڻ����ϻ��Ƽ��ı�
Public Sub DrawSimpleText(ByVal hGraphics As Long, ByVal Text As String, ByRef mFont As FontType, ByRef DatumPoint As PointL, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hStringFormat As Long, hFontFamily As Long, hFont As Long, hPath As Long, mRect As RectL
    If Text = "" Then Exit Sub
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias                      '����ݴ���
    GdipCreateStringFormat 0, 0, hStringFormat                                  '�����ַ�����ʽ
    GdipStringFormatGetGenericTypographic hStringFormat                         '����ͨ���ַ�����ʽ
    GdipCreateFontFamilyFromName StrPtr(mFont.Name), 0, hFontFamily             '����������
    GdipCreateFont hFontFamily, mFont.Size, mFont.Style, UnitPixel, hFont       '��������
    With mRect
        .Left = DatumPoint.X
        .Top = DatumPoint.Y
    End With
    hPath = NewPath                                                             '����·��
    GdipAddPathStringI hPath, StrPtr(Text), -1, hFontFamily, mFont.Style, CLng(mFont.Size), mRect, hStringFormat '�������·��
    GdipFillPath hGraphics, hFill, hPath                                        '���
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '���
    GdipDeletePath hPath                                                        'ɾ��·��
    GdipDeleteStringFormat hStringFormat                                        'ɾ����ʱ�ַ�����ʽ���
    GdipDeleteFont hFont                                                        'ɾ����ʱ����
    GdipDeleteFontFamily hFontFamily                                            'ɾ����ʱ��������
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteGraphics hGraphics
    End If
End Sub

'�ڻ����ϻ��������ı�(�԰ٷֱȼ���)
Public Sub DrawMaskedText(ByVal hGraphics As Long, ByVal Text As String, ByRef mFont As FontType, ByRef DatumPoint As PointL, ByVal Percentage As Single, ByVal hBorder As Long, ByVal hFill As Long, ByVal hMask As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hStringFormat As Long, hFontFamily As Long, hFont As Long, hPath As Long, mRect As RectL, oRect As RectF, nRect As RectF, tRect As RectL, hRegion As Long
    Dim tCodePointsFitted As Long, tLinesFilled As Long, nPercentage As Single  '��ʱ��Ҫ�ı���
    If Text = "" Then Exit Sub
    nPercentage = IIf(Percentage < 0, 0, IIf(Percentage > 100, 100, Percentage))
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias                      '����ݴ���
    GdipCreateStringFormat 0, 0, hStringFormat                                  '�����ַ�����ʽ
    GdipStringFormatGetGenericTypographic hStringFormat                         '����ͨ���ַ�����ʽ
    GdipCreateFontFamilyFromName StrPtr(mFont.Name), 0, hFontFamily             '����������
    GdipCreateFont hFontFamily, mFont.Size, mFont.Style, UnitPixel, hFont       '��������
    oRect = NewRectFloat(DatumPoint.X, DatumPoint.Y, Len(Text) * mFont.Size, mFont.Size) 'ԭʼ����
    GdipMeasureString hGraphics, StrPtr(Text), Len(Text), hFont, oRect, hStringFormat, nRect, tCodePointsFitted, tLinesFilled
    tRect = RectF2RectL(nRect)                                                  '�ַ������ھ���
    With tRect
        .Right = 0
        .Bottom = 0
    End With
    mRect = RectF2RectL(nRect)                                                  '���־���
    mRect.Right = mRect.Right * nPercentage / 100                               '�������ְٷֱ�
    GdipCreateRegionRectI mRect, hRegion                                        '������������
    hPath = NewPath                                                             '����·��
    GdipAddPathStringI hPath, StrPtr(Text), Len(Text), hFontFamily, mFont.Style, CLng(mFont.Size), tRect, hStringFormat '�������·��
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '���
    GdipFillPath hGraphics, hFill, hPath                                        '����ַ���
    GdipSetClipRectI hGraphics, mRect.Left, mRect.Top, mRect.Right, mRect.Bottom, CombineModeReplace '�ü���������
    GdipFillPath hGraphics, hMask, hPath                                        '�����������
    GdipResetClip hGraphics                                                     '���òü�
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

'�ڻ����ϻ���ֱ��
Public Sub DrawLine(ByVal hGraphics As Long, ByRef Point1 As PointL, ByRef Point2 As PointL, ByVal hPen As Long, Optional ByVal ReleaseHandles As Boolean = False)
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias                      '����ݴ���
    GdipDrawLineI hGraphics, hPen, Point1.X, Point1.Y, Point2.X, Point2.Y
    If ReleaseHandles Then
        GdipDeletePen hPen
        GdipDeleteGraphics hGraphics
    End If
End Sub

'�ڻ����ϻ��ƾ���
Public Sub DrawRectangle(ByVal hGraphics As Long, ByRef mRect As RectL, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hPath  As Long
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias                      '����ݴ���
    hPath = NewPath
    GdipAddPathRectangleI hPath, mRect.Left, mRect.Top, mRect.Right - mRect.Left, mRect.Bottom - mRect.Top '��Ӿ���·��
    GdipFillPath hGraphics, hFill, hPath                                        '���
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '���
    GdipDeletePath hPath
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteGraphics hGraphics
    End If
End Sub

'�ڻ����ϻ�����Բ������Բ�Σ�
Public Sub DrawEllipse(ByVal hGraphics As Long, ByRef mRect As RectL, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hPath As Long
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias                      '����ݴ���
    hPath = NewPath
    GdipAddPathEllipseI hPath, mRect.Left, mRect.Top, mRect.Right - mRect.Left, mRect.Bottom - mRect.Top '��Ӿ���·��
    GdipFillPath hGraphics, hFill, hPath                                        '���
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '���
    GdipDeletePath hPath
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteGraphics hGraphics
    End If
End Sub

'�ڻ����ϻ���Բ�Ǿ���
Public Sub DrawRoundedRectangle(ByVal hGraphics As Long, ByRef mRect As RectL, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal FilletRadius As Long = -1, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    Dim hPath As Long, RoundSize As Long, tPath(7) As Long, T As Long
    Dim mLeft As Long, mTop As Long, mWidth As Long, mHeight As Long
    If FilletRadius <= 0 And FilletRadius <> -1 Then
        MsgBox "��������FilletRadius��ֵ�������0��", vbCritical, "����"
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
    GdipFillPath hGraphics, hFill, hPath                                        '���
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '���
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

'�ڻ����ϻ��ƶ����
Public Sub DrawPolygon(ByVal hGraphics As Long, ByRef mPoint() As PointL, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    If UBound(mPoint) - LBound(mPoint) < 2 Then
        MsgBox "�������󣺵㼯���� mPoint ���������ݲ��㣬������Ҫ3������ܹ���һ����ն���Ρ�", vbCritical, "����"
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
    GdipFillPath hGraphics, hFill, hPath                                        '���
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '���
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

'�ڻ����ϻ���ͼ��
Public Sub DrawImage(ByVal hGraphics As Long, ByVal hImage As Long, ByRef DatumPoint As PointL, Optional ByVal TransformMode As RotateFlipType = RotateNoneFlipNone, Optional ByVal Zoom As Single = 1#, Optional ByVal ReleaseHandles As Boolean = False)
    Dim imgWidth As Long, imgHeight As Long
    If Zoom <= 0 Then
        MsgBox "�������� Zoom ��ֵӦ����0��", vbCritical, "����"
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

'�ڻ����ϻ�����Ƭ*��
'ע����Ƭ��Tile��ָһ����ͼ����ָ��һ�������ͼ��
Public Sub DrawTile(ByVal hGraphics As Long, ByVal hImage As Long, ByRef DatumPoint As PointL, ByRef SourceRect As RectL, Optional ByVal Zoom As Single = 1#, Optional ByVal ReleaseHandles As Boolean = False)
    Dim imgWidth As Long, imgHeight As Long, mSize As SizeL
    Dim hMemDC As Long, hMemGraphics As Long
    If Zoom <= 0 Then
        MsgBox "�������� Zoom ��ֵӦ����0��", vbCritical, "����"
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

'�ڻ����ϻ���ָ����·��
Public Sub DrawPath(ByVal hGraphics As Long, ByVal hPath As Long, ByVal hBorder As Long, ByVal hFill As Long, Optional ByVal DrawBorder As Boolean = True, Optional ByVal ReleaseHandles As Boolean = False)
    InitGDIPlus
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    GdipFillPath hGraphics, hFill, hPath                                        '���
    If DrawBorder Then GdipDrawPath hGraphics, hBorder, hPath                   '���
    If ReleaseHandles Then
        GdipDeletePen hBorder
        GdipDeleteBrush hFill
        GdipDeleteGraphics hGraphics
        GdipDeletePath hPath
    End If
End Sub

'��ָ����ɫͿ����������
Public Sub FillWholeGraphics(ByVal hGraphics As Long, Optional ByVal mColor As Long = &HFFFFFFFF, Optional ByVal ReleaseHandles As Boolean = False)
    InitGDIPlus
    GdipGraphicsClear hGraphics, mColor
    If ReleaseHandles Then GdipDeleteGraphics hGraphics
End Sub

'�ϲ�·������hPathAdditional�ϲ���hPathOriginal�У�
Public Sub CombinePath(ByVal hPathOriginal As Long, ByVal hPathAdditional As Long, Optional ByVal Connecting As Boolean = True)
    InitGDIPlus
    GdipAddPathPath hPathOriginal, hPathAdditional, Abs(CLng(Connecting))
End Sub

'�жϵ��Ƿ���������
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

'�жϾ����Ƿ��������ཻ�����߾���λ�������ڲ�
'True - �����������ཻ�����߾���λ�������ڲ���False - ��������������
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
