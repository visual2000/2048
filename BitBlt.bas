Attribute VB_Name = "BitBlt"
Public Declare Function BitBlt Lib "gdi32" _
 (ByVal hDestDC As Long, ByVal X As Long, _
 ByVal Y As Long, ByVal nWidth As Long, _
 ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc _
 As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


