Imports System.Runtime.InteropServices
Imports netDxf

' Miscellaneous SHARED routines that don't belong in a class
Module Utilities
    ' Create a union so that a float can be accessed as a UInt
    <StructLayout(LayoutKind.Explicit)> Public Structure IntFloatUnion
        <FieldOffset(0)>
        Public i As Integer
        <FieldOffset(0)> Dim f As Single
    End Structure

    Public ScaleUnity = New Matrix3(1, 0, 0, 0, 1, 0, 0, 0, 1)             ' scale factor of 1
    Public ScaleMM = New Matrix3(1 / Form1.xScale, 0, 0, 0, 1 / Form1.yScale, 0, 0, 0, 0)    ' scale steps to mm
    Public Function Leetro2Float(b As Long) As Single
        ' Converts a custom Leetro floating-point format to an IEEE 754 floating-point number.
        ' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]
        Dim ieee As IntFloatUnion
        If b = 0 Then Return 0.0          ' special case
        Dim vExp As Integer = (b >> 24)
        vExp += 127 And &HFF
        Dim vMan As Integer = (b >> 9) And &H3FFF
        Dim vSgn As Integer = b >> 23 And 1         ' sign
        ieee.i = (vSgn << 31) Or (vExp << 23) Or (vMan << 9)
        Return ieee.f   ' return floating point
    End Function

    Function Float2Int(f As Single) As Integer
        ' Convert integer to Leetro float
        ' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]
        ' IEEE float format    [seeeeeee|emmmmmmm|mmmmmmmm|mmmmmmmm]
        Dim ieee As IntFloatUnion
        ieee.f = f
        ' Handle special cases
        If ieee.i = 0 Then Return 0 ' Zero case
        If ieee.i = &H3F800000 Then Return 1 ' One case
        Dim vSgn = (ieee.i >> 31) And 1     ' extract sign bit
        Dim exp As Integer = (ieee.i >> 23) And &HFF
        exp -= 127          ' ieee floats have a 127 exp bias, leetro does not
        exp = exp And &HFF
        Dim vMan = ieee.i And &H7FFFFF
        ieee.i = (exp << 24) Or (vSgn << 23) Or vMan
        Return ieee.i
    End Function

    Function LeetroToIEEE(leetroint As Integer) As Double
        ' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]
        ' IEEE float format    [seeeeeee|emmmmmmm|mmmmmmmm|mmmmmmmm]
        ' Even though Leetro numbers usually have the bottom 9 bits = 0, sometimes they don't. We will support a Leetro mantissa of 23 bits

        ' Extract Leetro float components
        Dim exponent As Integer = (leetroint >> 24) And &HFF ' Exponent (8 bits)
        Dim sign As Integer = (leetroint >> 23) And &H1 ' Sign bit (1 bit)
        Dim mantissa As Integer = leetroint And &H7FFFFF ' 23-bit mantissa

        ' Convert exponent from Leetro (biased by 128) to IEEE 754 (biased by 127)
        Dim ieeeExponent As Integer = exponent + 127 ' - 128 + 127

        ' Convert mantissa from 15 bits to 23 bits (shift left by 8 bits)
        Dim ieeeMantissa As Integer = mantissa

        ' Construct IEEE 754 format (sign | exponent | mantissa)
        Dim ieeeInt As Integer = (sign << 31) Or (ieeeExponent << 23) Or ieeeMantissa

        Dim ieee As IntFloatUnion
        ieee.i = ieeeInt
        ' Convert to IEEE 754 floating point number
        Return ieee.f
    End Function
    Function IEEEToLeetro(ieeeFloat As Double) As Integer
        ' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]
        ' IEEE float format    [seeeeeee|emmmmmmm|mmmmmmmm|mmmmmmmm]
        ' Even though Leetro numbers usually have the bottom 9 bits = 0, sometimes they don't. We will support a Leetro mantissa of 23 bits

        ' Convert IEEE float to integer representation
        Dim ieeeBytes As Byte() = BitConverter.GetBytes(ieeeFloat)
        Dim ieeeInt As Integer = BitConverter.ToInt32(ieeeBytes, 0)

        ' Extract IEEE 754 components
        Dim sign As Integer = (ieeeInt >> 31) And &H1 ' Sign bit (1 bit)
        Dim exponent As Integer = (ieeeInt >> 23) And &HFF ' Exponent (8 bits)
        Dim mantissa As Integer = ieeeInt And &H7FFFFF ' 23-bit mantissa

        ' Convert exponent from IEEE (biased by 127) to Leetro (biased by 128)
        Dim leetroExponent As Integer = exponent - 127 + 128

        ' Convert mantissa from 23 bits to 15 bits (shift right by 8 bits, rounding if necessary)
        Dim leetroMantissa As Integer = mantissa And &H7FFFFF  ' Keep only 23 bits

        ' Construct Leetro format (exponent | sign | mantissa | 0 padding)
        Dim leetroInt As Integer = (leetroExponent << 24) Or (sign << 23) Or leetroMantissa

        Return leetroInt
    End Function

    Public Function PowerSpeedColor(power As Double, speed As Double) As AciColor
        ' Convert speed/power to an approximate color
        ' Power = power in watts
        ' Speed = speed in mm/s
        'Return AciColor.FromHsl(30 / 360, 0.6, CSng(power / 100) * (1 - (CSng(speed) / My.Settings.SpeedMax)))
        Dim pwr = 1 - (power / My.Settings.PowerMax)
        Dim spd = speed / My.Settings.SpeedMax
        Return AciColor.FromHsl(0, 0, pwr * spd)
    End Function

    Function Distance(p1 As Vector2, p2 As Vector2) As Double
        ' Calculate the distance between two points
        Return Math.Sqrt((p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2)
    End Function

    Function Vect3(vect2 As Vector2) As Vector3
        ' Convert a Vector2 type to a Vector3 type
        Dim result As New Vector3
        With result
            .X = vect2.X
            .Y = vect2.Y
            .Z = 0
        End With
        Return result
    End Function

    Function PointToVector2(p As System.Windows.Point) As Vector2
        ' Convert a windows point to a DXF vector2
        Return New Vector2(p.X, p.Y)
    End Function
    Function Vector2ToPoint(p As Vector2) As System.Windows.Point
        ' Convert Vector2 to a system.windows.point
        Return New System.Windows.Point(p.X, p.Y)
    End Function
    Function VectorToPoint(p As System.Windows.Vector) As System.Windows.Point
        Return New System.Windows.Point(p.X, p.Y)
    End Function
End Module
