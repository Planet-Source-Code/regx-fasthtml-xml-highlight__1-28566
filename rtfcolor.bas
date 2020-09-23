Attribute VB_Name = "mNew"
Option Explicit

Function fcnGetRTFColor(ByVal Color As Variant) As String
  '***  this function accepts a VB color (long)
  '***  or a HTML color (string) and
  '***  returns a RTF color table def.
  
  Const sHEX = "0123456789ABCDEF"
  Dim lngRed As Long, lngGreen As Long, lngBlue As Long

  If VarType(Color) = vbLong Then
    lngRed = Color Mod 256&
    lngGreen = (Color Mod 65536) \ 256&
    lngBlue = Color \ 65536
  ElseIf VarType(Color) = vbString Then
    '***  the string should be something like this: #D0D5DF
    '***  strip of the right 6 chars
    Color = Right$(Color, 6)
    
    '***  find the position for each char in sHEX. Position is the value
    lngRed = 16& * (InStr(1, sHEX, Mid$(Color, 1, 1), vbTextCompare) - 1) + _
              1& * (InStr(1, sHEX, Mid$(Color, 2, 1), vbTextCompare) - 1)
    lngGreen = 16& * (InStr(1, sHEX, Mid$(Color, 3, 1), vbTextCompare) - 1) + _
                1& * (InStr(1, sHEX, Mid$(Color, 4, 1), vbTextCompare) - 1)
    lngBlue = 16& * (InStr(1, sHEX, Mid$(Color, 5, 1), vbTextCompare) - 1) + _
               1& * (InStr(1, sHEX, Mid$(Color, 6, 1), vbTextCompare) - 1)
  Else
    '***  this function accepts a VB color (long)
    '***  or a HTML color (string) only.
    Stop
  End If
  
  fcnGetRTFColor = "\red" & CStr(lngRed) & "\green" & CStr(lngGreen) & "\blue" & CStr(lngBlue) & ";"
End Function

