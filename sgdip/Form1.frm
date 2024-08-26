VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "马云爱逛京东"
   ClientHeight    =   11355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16185
   BeginProperty Font 
      Name            =   "思源黑体 CN Heavy"
      Size            =   9
      Charset         =   134
      Weight          =   900
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   757
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1079
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    '初始化gdip
    InitGDIPlus
    Me.FontSize = 60
    Dim mPoint As PointL, mRect As RectL, mRegion As Long
    '新建点
    mPoint = NewPoint(440, 40)
    '画圆角矩形
    DrawRoundedRectangle NewGraphics(Me.hDC), NewRect(20, 20, 200, 100), NewPen(&HFF000000, 1), NewHatchBrush(HatchStyle.HatchStyleWeave, &HFF39C5BB, &HFFFFFFFF)
    '画多边形
    DrawPolygon NewGraphics(Me.hDC), NewArrayOfPointL("100,200,300,200,300,500,200,320,100,500"), NewPen(&HFF000000, 1), NewGradientLineBrush(NewPoint(100, 200), NewPoint(100, 500), &HFF33CCFF, &HFFFFCC33)
    '画遮罩文本
    DrawMaskedText NewGraphics(Me.hDC), "人生若只如初见", StdFont2FontType(Me.Font), NewPoint(440, 250), 35, NewPen(&HFF000000, 1), NewGradientLineBrush(NewPoint(440, 250), NewPoint(440, 250 + StdFont2FontType(Me.Font).Size), &HFF33CCFF, &HFFFFCC33), NewHatchBrush(HatchStyle.HatchStylePlaid, &HFFFFFFFF, &HFFFFCCCC)
    '加载工程目录下所有png格式图片
    LoadImagesFromFolder App.Path, "png"
    '画简单文本
    DrawSimpleText NewGraphics(Me.hDC), Me.Caption, StdFont2FontType(Me.Font), mPoint, NewPen(&HFF000000, 1), NewChartletBrush(GetGCO("2"), mPoint)
    '画图像
    DrawImage NewGraphics(Me.hDC), GetGCO("3"), NewPoint(400, 400), Rotate90FlipX, 0.8
    '新建矩形
    mRect = NewRect(100, 200, 200, 100)
    '设置点
    mPoint = NewPoint(120, 240)
    '测试点是否在区域内（可用于点击判断）
    Debug.Print IsPointOnRegion(NewGraphics(Me.hDC), NewRegionFromRect(mRect), mPoint) '点在区域内部
    '画瓦片
    DrawTile NewGraphics(Me.hDC), GetGCO("1"), NewPoint(500, 400), NewRect(113, 32, 66, 70), 1.8
    '刷新窗体
    Me.Refresh
    '关闭gdip
    CloseGDIPlus
End Sub
