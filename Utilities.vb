Imports MOL.Form1
Imports netDxf

' Miscellaneous SHARED routines that don't belong in a class
Module Utilities
    Function Leetro2Float(b As UInteger) As Single
        ' Converts a custom Leetro floating-point format to an IEEE 754 floating-point number.
        ' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]
        If b = 0 Then Return 0.0          ' special case
        If b = &H81000000UL Then Return 1.0
        Dim vExp As Integer = (b >> 24)  ' exponent stored as -127 - 128
        Dim vMan As Integer = (b >> 9 And &H3FFF) Or 2 ^ 14   ' mantissa (adding implied leading bit)
        Dim vSgn As Integer = b >> 23 And 1         ' sign
        If vSgn > 0 Then vSgn = -1 Else vSgn = 1
        Dim significand As Single = vMan / 16384.0
        Dim result As Single = significand * Math.Pow(2, vExp) * vSgn
        Return result
    End Function

    Function Float2Int(f As Single) As UInteger
        ' Convert integer to Leetro float
        ' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]
        ' IEEE float format    [seeeeeee|emmmmmmm|mmmmmmmm|mmmmmmmm]
        Dim ieee As IntFloatUnion
        ieee.f = f
        ' Handle special cases
        If ieee.i = 0 Then Return 0 ' Zero case
        If ieee.i = &H3F800000 Then Return 1 ' One case
        Dim leetro As UInteger = 0          ' Leetro version
        If ieee.f < 0 Then leetro = leetro Or &H800000     ' transfer sign bit 
        Dim exp As Integer = ieee.i >> 23   ' ieee floats have a 127 exp bias, leetro does not
        exp = exp And &HFF    ' reduce exp to 8 bits
        exp -= 127
        exp = exp And &HFF
        Dim exponent As Long = (exp << 24) And &HFF000000L
        leetro = leetro Or CUInt(exponent)             ' add exponent
        Dim significand As Integer = ieee.i And &H7FFFFF
        leetro = leetro Or significand
        Return leetro
    End Function
    Public Function PowerSpeedColor(power As Single, speed As Single) As AciColor
        ' Convert speed/power to an approximate color
        'Return AciColor.FromHsl(30 / 360, 0.6, CSng(power / 100) * (1 - (CSng(speed) / My.Settings.SpeedMax)))
        Dim pwr = (power - My.Settings.PowerMin) / (My.Settings.PowerMax - My.Settings.PowerMin)
        Dim spd = 1 - ((speed - My.Settings.SpeedMin) / (My.Settings.SpeedMax - My.Settings.SpeedMin))
        Return AciColor.FromHsl(0, 0, pwr * spd)
    End Function

    Function Distance(p1 As Vector2, p2 As Vector2) As Single
        ' Calculate the distance between two points
        Return Math.Sqrt((p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2)
    End Function
End Module
