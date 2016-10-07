VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPalett"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ColourPaletUnique As Variant
Public ColourRGBUnique As Variant
Public ColourPalletText As Variant

Public Property Get Count() As Long
  Count = UBound(ColourPaletUnique) + 1
End Property

Private Function map(i As Long, Mapping As Variant) As Long
  map = Mapping(((i - 1) Mod UBound(Mapping)))
End Function
Public Property Get ColourRGB(c As Long) As Long
  ColourRGB = map(c, ColourRGBUnique)
End Property
Public Property Get Colour(c As Long) As Long
  Colour = map(c, ColourPaletUnique)
End Property
Public Property Get fontColour(c As Long) As Long
  fontColour = map(c, ColourPalletText)
End Property
Private Sub Class_Initialize()
'Colour system based on http://dmcritchie.mvps.org/excel/colors.htm
'Duplicate colours are removed from the sequencial list as well as RGB(255,255,255)
'White) and RGB(255,255,204) which are too pale.
'Text colour (black or white) are assigned to maxsimise readability
'Color Codes.xlsm

  ColourPaletUnique = Array(1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 20, 21, 22, 23, _
                            24, 33, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, _
                            52, 53, 55, 56)
  ColourRGBUnique = Array(0, 255, 65280, 16711680, 65535, 16711935, 16776960, 128, 32768, 8388608, 32896, _
                          8388736, 8421376, 12632256, 8421504, 16751001, 6697881, 16777164, 6684774, _
                          8421631, 13395456, 16764108, 16763904, 13434828, 10092543, 16764057, 13408767, _
                          16751052, 10079487, 16737843, 13421619, 52377, 52479, 39423, 26367, 10053222, _
                          9868950, 6697728, 6723891, 13056, 13107, 13209, 10040115, 3355443)
  ColourPalletText = Array(2, 1, 1, 2, 1, 1, 1, 2, 2, 2, 2, 2, 2, 1, 1, 1, 2, 1, 2, 1, 1, 1, 1, 1, 1, 1, _
                            1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 1, 2, 2, 2, 2, 2)

End Sub
