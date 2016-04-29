VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CScreenCliper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////
'@@summary 支持全屏截屏
'@@require 无
'@@reference 代码部分来源网络，我进行了组合构建，以及删减调试
'@@license MIT
'@@author sunruiyeyipeng@163.com
'@@create 2016年4月28日14:10:43
'@@modify
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有声明
'//
'//////////////////////////////////////////////////////////////////////////////

'------------------------------------------------------------------------------
'       私有数据类型
'------------------------------------------------------------------------------
Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type
Private Type GdiplusStartupInput
  GdiplusVersion As Long
  DebugEventCallback As Long
  SuppressBackgroundThread As Long
  SuppressExternalCodecs As Long
End Type
Private Type EncoderParameter
  GUID As GUID
  NumberOfValues As Long
  Type As Long
  Value As Long
End Type
Private Type EncoderParameters
  Count As Long
  Parameter As EncoderParameter
End Type

Private Type PicBmp
  Size As Long
  Type As Long
  hbmp As Long
  hPal As Long
  Reserved As Long
End Type
'------------------------------------------------------------------------------
'       私有API
'------------------------------------------------------------------------------
Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Function GdiplusShutdown Lib "GDIPlus" (ByVal token As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" (ByVal hbm As Long, ByVal hPal As Long, Bitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal Image As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "GDIPlus" (ByVal Image As Long, ByVal filename As Long, clsidEncoder As GUID, encoderParams As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal str As Long, id As GUID) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, Bitmap As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" _
(ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" _
(PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, _
IPic As IPicture) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'//////////////////////////////////////////////////////////////////////////////
'//
'//      类
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       初始化
'------------------------------------------------------------------------------
Private Sub Class_Initialize()

End Sub

'------------------------------------------------------------------------------
'       销毁
'------------------------------------------------------------------------------
Private Sub Class_Terminate()

End Sub

'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Private Function PictureBoxSaveJPG(ByVal pict As StdPicture, ByVal filename As String, Optional ByVal quality As Byte = 80) As Boolean
  Dim tSI As GdiplusStartupInput
  Dim lRes As Long
  Dim lGDIP As Long
  Dim lBitmap As Long
    
  '初始化 GDI+
  tSI.GdiplusVersion = 1
  lRes = GdiplusStartup(lGDIP, tSI, 0)
   
  If lRes = 0 Then
    '从句柄创建 GDI+ 图像
    lRes = GdipCreateBitmapFromHBITMAP(pict.Handle, 0, lBitmap)
       
    If lRes = 0 Then
      Dim tJpgEncoder As GUID
      Dim tParams As EncoderParameters
       
      '初始化解码器的GUID标识
      CLSIDFromString StrPtr("{557CF401-1A04-11D3-9A73-0000F81EF32E}"), tJpgEncoder
       
      '设置解码器参数
      tParams.Count = 1
      With tParams.Parameter ' Quality
        '得到Quality参数的GUID标识
        CLSIDFromString StrPtr("{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"), .GUID
        .NumberOfValues = 1
        .Type = 4
        .Value = VarPtr(quality)
      End With
       
      '保存图像
      lRes = GdipSaveImageToFile(lBitmap, StrPtr(filename), tJpgEncoder, tParams)
       
      '销毁GDI+图像
      GdipDisposeImage lBitmap
    End If
     
    '销毁 GDI+
    GdiplusShutdown lGDIP
  End If
    
  If lRes Then
    PictureBoxSaveJPG = False
  Else
    PictureBoxSaveJPG = True
  End If
End Function

'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////
Public Sub ClipScreenToFile(ByVal FilePath As String)
  Dim pPicture As New StdPicture
  Set pPicture = ClipToImage
  PictureBoxSaveJPG pPicture, FilePath
End Sub

Private Function ClipToImage() As StdPicture
  Dim screenDc As Long, width As Long, height As Long
  Dim lPicDc As Long, lPicBmp As Long
  Dim G As GUID, p As PicBmp
  width = Screen.width / Screen.TwipsPerPixelX
  height = Screen.height / Screen.TwipsPerPixelY
  
  screenDc = GetDC(0)
  lPicDc = CreateCompatibleDC(screenDc)
  lPicBmp = CreateCompatibleBitmap(screenDc, width, height)
  SelectObject lPicDc, lPicBmp
  BitBlt lPicDc, 0, 0, width, height, screenDc, 0, 0, vbSrcCopy
  With G
    .Data1 = &H20400
    .Data4(0) = &HC0&
    .Data4(7) = &H46&
  End With
  With p
    .Size = Len(p)
    .Type = vbPicTypeBitmap
    .hbmp = lPicBmp
    .hPal = &H0&
  End With
  OleCreatePictureIndirect p, G, True, ClipToImage
  ReleaseDC 0, screenDc
  DeleteObject lPicBmp
End Function
