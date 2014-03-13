Attribute VB_Name = "轉換圖片"
Option Explicit

Public Const LF_FACESIZE = 32
Public Const DEFAULT_CHARSET = 1
Public Const GGO_GRAY4_BITMAP = 5
Public Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Public Type CDPFONT
        Name As String
        Size As Long
        Bold As Boolean
        Italic As Boolean
        Underline As Boolean
        StrikeThrough As Boolean
        color As Long
End Type

Public Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(0 To LF_FACESIZE - 1) As Byte
End Type

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Type GLYPHMETRICS
        gmBlackBoxX As Long
        gmBlackBoxY As Long
        gmptGlyphOrigin As POINTAPI
        gmCellIncX As Integer
        gmCellIncY As Integer
End Type

Public Type FIXED
        fract As Integer
        Value As Integer
End Type

Public Type MAT2
        eM11 As FIXED
        eM12 As FIXED
        eM21 As FIXED
        eM22 As FIXED
End Type

Public Type BITMAPFILEHEADER '14 bytes
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER '40 bytes
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

Public Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Declare Function GetGlyphOutline Lib "gdi32" Alias "GetGlyphOutlineA" (ByVal hdc As Long, ByVal uChar As Long, ByVal fuFormat As Long, lpgm As GLYPHMETRICS, ByVal cbBuffer As Long, lpBuffer As Any, lpmat2 As MAT2) As Long

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public Declare Function GetRefPtr Lib "vbutil32" (lpVoid As Any) As Long
Public Declare Function RShiftDWord Lib "vbutil32" (ByVal dw As Long, ByVal c As Integer) As Long
Public Declare Function LShiftDWord Lib "vbutil32" (ByVal dw As Long, ByVal c As Integer) As Long

Public Sub 字形轉成圖片(字型 As CDPFONT, 字形 As String, 圖檔 As String, 解析度 As Long, Success As Boolean)
    
Dim hdc As Long, dpi As Long
Dim font As LOGFONT, hOldFont As Long, hFont As Long
Dim gm As GLYPHMETRICS, mat As MAT2
Dim buf() As Byte, bufsize As Long
Dim bf As BITMAPFILEHEADER, bi As BITMAPINFOHEADER
Dim rgb(0 To 16) As RGBQUAD, db As Single, dg As Single, dr As Single
Dim nscan As Long, WidthBytes As Long, i As Long, j As Long
Dim nfile As Integer
        
On Error GoTo BmpError

hdc = GetDC(0)
If 解析度 = 0 Then
    dpi = GetDeviceCaps(hdc, LOGPIXELSY)
Else
    dpi = 解析度
End If

' LOGFONT 資料結構的設定

'RtlMoveMemory font.lfFaceName(0), ByVal 字型.Name, LenB(字型.Name) + 1
RtlMoveMemory font.lfFaceName(0), ByVal 字型.Name, LenB(StrConv(字型.Name, vbFromUnicode)) + 1
font.lfHeight = (字型.Size * -20) * dpi / 1440
font.lfWidth = 0
font.lfWeight = IIf(字型.Bold, 700, 400)
font.lfItalic = IIf(字型.Italic, 1, 0)
font.lfUnderline = IIf(字型.Underline, 1, 0)
font.lfStrikeOut = IIf(字型.StrikeThrough, 1, 0)
font.lfCharSet = DEFAULT_CHARSET
     
' 建立字型物件
    
hFont = CreateFontIndirect(font)
    
hOldFont = SelectObject(hdc, hFont)
    
mat.eM11.fract = 0
mat.eM11.Value = 1
mat.eM12.fract = 0
mat.eM12.Value = 0
mat.eM21.fract = 0
mat.eM21.Value = 0
mat.eM22.fract = 0
mat.eM22.Value = 1
    
bufsize = GetGlyphOutline(hdc, CLng(Asc(字形)), GGO_GRAY4_BITMAP, gm, 0, Null, mat)
        
ReDim buf(bufsize - 1)
bufsize = GetGlyphOutline(hdc, CLng(Asc(字形)), GGO_GRAY4_BITMAP, gm, bufsize, buf(0), mat)
     
bf.bfType = &H4D42
bf.bfSize = 122 + bufsize '14 + 40 + 4 * 17 + bufsize (17 gray level)
bf.bfReserved1 = 0
bf.bfReserved2 = 0
bf.bfOffBits = 122 '14 + 40 + 4 * 17(17 gray level)
    
bi.biSize = 40
bi.biWidth = gm.gmBlackBoxX
bi.biHeight = gm.gmBlackBoxY
bi.biPlanes = 1
bi.biBitCount = 8
bi.biCompression = 0
bi.biSizeImage = bufsize
bi.biXPelsPerMeter = dpi * 39.37 '39.37=100/2.54
bi.biYPelsPerMeter = bi.biXPelsPerMeter
bi.biClrUsed = 17
bi.biClrImportant = 17
    
rgb(0).rgbBlue = &HFF
rgb(0).rgbGreen = &HFF
rgb(0).rgbRed = &HFF
rgb(0).rgbReserved = 0

rgb(16).rgbBlue = RShiftDWord(字型.color And &HFF0000, 16)
rgb(16).rgbGreen = RShiftDWord(字型.color And &HFF00&, 8)
rgb(16).rgbRed = 字型.color And &HFF&
rgb(16).rgbReserved = 0

db = (CLng(rgb(16).rgbBlue) - CLng(rgb(0).rgbBlue)) / 16
dg = (CLng(rgb(16).rgbGreen) - CLng(rgb(0).rgbGreen)) / 16
dr = (CLng(rgb(16).rgbRed) - CLng(rgb(0).rgbRed)) / 16
    
For i = 1 To 15
    rgb(i).rgbBlue = i * db + rgb(0).rgbBlue
    rgb(i).rgbGreen = i * dg + rgb(0).rgbGreen
    rgb(i).rgbRed = i * dr + rgb(0).rgbRed
    rgb(i).rgbReserved = 0
Next i
    
WidthBytes = LShiftDWord(RShiftDWord(gm.gmBlackBoxX * bi.biBitCount + 31, 5), 2)
nscan = bufsize / WidthBytes
    
nfile = FreeFile
Open 圖檔 For Binary Access Write As #nfile
Put #nfile, 1, bf
Put #nfile, , bi
Put #nfile, , rgb
For i = nscan - 1 To 0 Step -1
    For j = 0 To WidthBytes - 1
        Put #nfile, , buf(WidthBytes * i + j)
    Next j
Next i
Close #nfile
          
' 還原字型
    
SelectObject hdc, hOldFont
ReleaseDC 0, hdc
    
' 刪除字型
    
DeleteObject hFont

Success = True

Exit Sub

BmpError:
Success = False

End Sub

