Attribute VB_Name = "modGradiente"
Option Explicit

Public Function ColorBlend(ByVal RGB1 As Long, ByVal RGB2 As Long, ByVal Percent As Single) As Long

    Dim R As Integer, R1 As Integer, R2 As Integer, _
        G As Integer, G1 As Integer, G2 As Integer, _
        B As Integer, B1 As Integer, B2 As Integer
    
    If Percent >= 1 Then
        ColorBlend = RGB2
        Exit Function
    ElseIf Percent <= 0 Then
        ColorBlend = RGB1
        Exit Function
    End If
  
    R1 = RGBRed(RGB1)
    R2 = RGBRed(RGB2)
    G1 = RGBGreen(RGB1)
    G2 = RGBGreen(RGB2)
    B1 = RGBBlue(RGB1)
    B2 = RGBBlue(RGB2)
  
    R = ((R2 * Percent) + (R1 * (1 - Percent)))
    G = ((G2 * Percent) + (G1 * (1 - Percent)))
    B = ((B2 * Percent) + (B1 * (1 - Percent)))
    
    ColorBlend = RGB(R, G, B)
End Function

Private Function RGBRed(RGBColor As Long) As Integer
    RGBRed = RGBColor And &HFF
End Function

Private Function RGBGreen(RGBColor As Long) As Integer
    RGBGreen = ((RGBColor And &H100FF00) / &H100)
End Function

Private Function RGBBlue(RGBColor As Long) As Integer
    RGBBlue = (RGBColor And &HFF0000) / &H10000
End Function
