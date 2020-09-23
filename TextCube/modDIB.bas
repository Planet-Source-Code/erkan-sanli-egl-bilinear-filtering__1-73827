Attribute VB_Name = "modDIB"
Option Explicit

Private Const DIB_RGB_COLORS = 0

Public Type COLORRGBA_BYTE
    R               As Byte
    G               As Byte
    B               As Byte
    A               As Byte
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
    Header          As BITMAPINFOHEADER
    Bits()          As COLORRGBA_BYTE
End Type

Public Type DIB
    hDC             As Long
    Width           As Long
    Height          As Long
    hBmp            As Long
    bi              As BITMAPINFO
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
'Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long

Public Sub CreateArray(di As DIB)
    'Dim hBmp As Long
'    Width = di.Width
'    Height = di.Height
    CreateDC di
    di.hBmp = CreateDIBSection(di.hDC, di.bi, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    SelectObject di.hDC, di.hBmp
    GetBits di '.hBmp, di.bi.Bits

End Sub

Public Sub CreateArrayFromPicbox(di As DIB)
  
    CreateDC di
    SelectObject di.hDC, di.hBmp
    GetBits di

End Sub

Private Sub CreateDC(di As DIB)

    With di.bi.Header
        .biBitCount = 32
        .biPlanes = 1
        .biSize = 40
        .biWidth = di.Width
        .biHeight = -di.Height
    End With
    di.hDC = CreateCompatibleDC(0)
    
End Sub

Public Sub Clear(di As DIB)
  
    Erase di.bi.Bits
    DeleteObject di.hBmp
    DeleteDC di.hDC

End Sub

Private Sub GetBits(di As DIB)
  
    ReDim di.bi.Bits(di.Width - 1, di.Height - 1)
    GetDIBits di.hDC, di.hBmp, 0, di.Height, di.bi.Bits(0, 0), di.bi, DIB_RGB_COLORS

End Sub

Public Sub SetBits(destDC As Long, di As DIB, Data() As COLORRGBA_BYTE)
  
    SetDIBits destDC, di.hBmp, 0, di.Height, Data(0, 0), di.bi, 0

End Sub

'Public Sub StrechBits(destDC As Long, di As DIB, Data() As COLORRGBA_BYTE)
'
'    StretchDIBits destDC, 0, 0, di.Width, di.Height, 0, 0, di.Width, di.Height, Data(0, 0), di.bi, DIB_RGB_COLORS, vbSrcCopy
'
'End Sub




