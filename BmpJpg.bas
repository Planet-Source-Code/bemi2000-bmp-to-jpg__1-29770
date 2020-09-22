Attribute VB_Name = "Module2"

'kopie ijl11.dll  naar WINDOWS\SYSTEM map

Type BITMAPINFOHEADER
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
 Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
 Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
 Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
 Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
 Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
''--------------------------------------------------------------------------------
 Type JPEG_CORE_PROPERTIES_VB
  UseJPEGPROPERTIES As Long
  DIBBytes As Long
  DIBWidth As Long
  DIBHeight As Long
  DIBPadBytes As Long
  DIBChannels As Long
  DIBColor As Long
  DIBSubsampling As Long
  JPGFile As Long
  JPGBytes As Long
  JPGSizeBytes As Long
  JPGWidth As Long
  JPGHeight As Long
  JPGChannels As Long
  JPGColor As Long
  JPGSubsampling As Long
  JPGThumbWidth As Long
  JPGThumbHeight As Long
  cconversion_reqd As Long
  upsampling_reqd As Long
  jquality As Long
  jprops(0 To 19999) As Byte
End Type
 Declare Function ijlInit Lib "ijl11.dll" (jcprops As Any) As Long
 Declare Function ijlFree Lib "ijl11.dll" (jcprops As Any) As Long
 Declare Function ijlWrite Lib "ijl11.dll" (jcprops As Any, ByVal ioType As Long) As Long
 Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Dim m_hDIb As Long, m_hBmpOld As Long
Public m_hDC As Long, m_lPtr As Long, camheight As Long, camwidth As Long
Sub bmpjpgzetten()
Dim m_tBI As BITMAPINFOHEADER
m_tBI.biSize = 40
m_tBI.biWidth = camwidth
m_tBI.biHeight = camheight
m_tBI.biPlanes = 1
m_tBI.biBitCount = 24
m_tBI.biCompression = 0
m_tBI.biSizeImage = ((camwidth * 3 + 3) And &HFFFFFFFC) * camheight
m_hDC = CreateCompatibleDC(0)
m_hDIb = CreateDIBSection(m_hDC, m_tBI, 0, m_lPtr, 0, 0)
m_hBmpOld = SelectObject(m_hDC, m_hDIb)
End Sub
Sub bmpjpgopnul()
ijlFree tJ
SelectObject m_hDC, m_hBmpOld
DeleteObject m_hDIb
DeleteDC m_hDC
End Sub
