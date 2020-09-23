Attribute VB_Name = "Module1"
Private Const SRCCOPY = &HCC0020


Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long


Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long


Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long


Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long


Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long


Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Type POINTAPI
    X As Long
    Y As Long
End Type

Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Declare Function WindowFromPoint& Lib "user32" (ByVal lpPointX As Long, ByVal lpPointY As Long)






Sub DisableHDC(SourceDC As Long, SourceWidth As Long, SourceHeight As Long)
    Const BLACK = 0
    Const DARKGREY = &H808080
    Const WHITE = &HFFFFFF
    Dim i As Long
    Dim j As Long
    Dim PixelColor As Long
    Dim BackgroundColor As Long
    Dim MemoryDC As Long
    Dim MemoryBitmap As Long
    Dim OldBitmap As Long
    Dim BooleanArray() As Boolean
    ReDim BooleanArray(SourceWidth, SourceHeight)
    MemoryDC = CreateCompatibleDC(SourceDC)
    MemoryBitmap = CreateCompatibleBitmap(SourceDC, SourceWidth, SourceHeight)
    OldBitmap = SelectObject(MemoryDC, MemoryBitmap)
    BitBlt MemoryDC, 0, 0, SourceWidth, SourceHeight, SourceDC, 0, 0, SRCCOPY
    BackgroundColor = GetBkColor(SourceDC)
    ' Scan Pixels and if the pixel is black
    ' it is flagged as true and saved in Boo
    '     leanArray(x,y)
    ' then colored dark grey (disabled color
    '     )


    For i = 0 To SourceWidth


        For j = 0 To SourceHeight
            PixelColor = GetPixel(MemoryDC, i, j)


            If PixelColor <> BackgroundColor Then ' skip background color pixels


                If PixelColor = BLACK Or Not PixelColor = WHITE Then
                    BooleanArray(i, j) = True
                    SetPixel MemoryDC, i, j, DARKGREY
                Else
                    SetPixel MemoryDC, i, j, BackgroundColor
                End If
            End If
        Next
    Next
    ' For each Black pixel, draw a white sha


    '     dow 1 pixel down and
        ' 1 pixel to the right to create a shado
        '     w effect


        For i = 0 To SourceWidth - 1


            For j = 0 To SourceHeight - 1


                If BooleanArray(i, j) = True Then


                    If BooleanArray(i + 1, j + 1) = False Then
                        SetPixel MemoryDC, i + 1, j + 1, WHITE
                    End If
                End If
            Next
        Next
        BitBlt SourceDC, 0, 0, SourceWidth, SourceHeight, MemoryDC, 0, 0, SRCCOPY


        SelectObject MemoryDC, OldBitmap
            DeleteObject MemoryBitmap
            DeleteDC MemoryDC
        End Sub
