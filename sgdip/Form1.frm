VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "���ư��侩��"
   ClientHeight    =   11355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16185
   BeginProperty Font 
      Name            =   "˼Դ���� CN Heavy"
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
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    '��ʼ��gdip
    InitGDIPlus
    Me.FontSize = 60
    Dim mPoint As PointL, mRect As RectL, mRegion As Long
    '�½���
    mPoint = NewPoint(440, 40)
    '��Բ�Ǿ���
    DrawRoundedRectangle NewGraphics(Me.hDC), NewRect(20, 20, 200, 100), NewPen(&HFF000000, 1), NewHatchBrush(HatchStyle.HatchStyleWeave, &HFF39C5BB, &HFFFFFFFF)
    '�������
    DrawPolygon NewGraphics(Me.hDC), NewArrayOfPointL("100,200,300,200,300,500,200,320,100,500"), NewPen(&HFF000000, 1), NewGradientLineBrush(NewPoint(100, 200), NewPoint(100, 500), &HFF33CCFF, &HFFFFCC33)
    '�������ı�
    DrawMaskedText NewGraphics(Me.hDC), "������ֻ�����", StdFont2FontType(Me.Font), NewPoint(440, 250), 35, NewPen(&HFF000000, 1), NewGradientLineBrush(NewPoint(440, 250), NewPoint(440, 250 + StdFont2FontType(Me.Font).Size), &HFF33CCFF, &HFFFFCC33), NewHatchBrush(HatchStyle.HatchStylePlaid, &HFFFFFFFF, &HFFFFCCCC)
    '���ع���Ŀ¼������png��ʽͼƬ
    LoadImagesFromFolder App.Path, "png"
    '�����ı�
    DrawSimpleText NewGraphics(Me.hDC), Me.Caption, StdFont2FontType(Me.Font), mPoint, NewPen(&HFF000000, 1), NewChartletBrush(GetGCO("2"), mPoint)
    '��ͼ��
    DrawImage NewGraphics(Me.hDC), GetGCO("3"), NewPoint(400, 400), Rotate90FlipX, 0.8
    '�½�����
    mRect = NewRect(100, 200, 200, 100)
    '���õ�
    mPoint = NewPoint(120, 240)
    '���Ե��Ƿ��������ڣ������ڵ���жϣ�
    Debug.Print IsPointOnRegion(NewGraphics(Me.hDC), NewRegionFromRect(mRect), mPoint) '���������ڲ�
    '����Ƭ
    DrawTile NewGraphics(Me.hDC), GetGCO("1"), NewPoint(500, 400), NewRect(113, 32, 66, 70), 1.8
    'ˢ�´���
    Me.Refresh
    '�ر�gdip
    CloseGDIPlus
End Sub
