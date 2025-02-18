Imports System.Runtime.InteropServices
Imports netDxf

' Miscellaneous SHARED routines that don't belong in a class
Friend Module Utilities
    ' Create a union so that a float can be accessed as a UInt
    <StructLayout(LayoutKind.Explicit)> Public Structure IntFloatUnion
        <FieldOffset(0)>
        Public i As Integer
        <FieldOffset(0)>
        Public f As Single
    End Structure

    Public ScaleUnity = New Matrix3(1, 0, 0, 0, 1, 0, 0, 0, 1)             ' scale factor of 1
    Public ScaleMM = New Matrix3(1 / Form1.xScale, 0, 0, 0, 1 / Form1.yScale, 0, 0, 0, 0)    ' scale steps to mm
    Public Function Leetro2Float(b As Long) As Single
        ' Converts a custom Leetro floating-point format to an IEEE 754 floating-point number.
        ' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]

        Dim ieee As IntFloatUnion
        If b = 0 Then Return 0.0          ' special case
        Dim vExp As Integer = (b >> 24)
        vExp = (vExp + 127) And &HFF             ' IEEE exponent bias and remove sign extention
        Dim vMan As Integer = b And &H7FFFFF    ' Even though Leetro numbers usually have the bottom 9 bits = 0, sometimes they don't. We will support a Leetro mantissa of 23 bits
        Dim vSgn As Integer = b >> 23 And 1         ' sign
        ieee.i = (vSgn << 31) Or (vExp << 23) Or vMan
        Return ieee.f   ' return floating point
    End Function

    Public Function Float2Leetro(f As Single) As Integer
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
        exp = exp And &HFF  ' remove any sign extension
        Dim vMan = ieee.i And &H7FFFFF      ' Even though Leetro numbers usually have the bottom 9 bits = 0, sometimes they don't. We will support a Leetro mantissa of 23 bits
        ieee.i = (exp << 24) Or (vSgn << 23) Or vMan
        Return ieee.i
    End Function

    Public Function EngraveParameters(speed As Double) As (Acclen As Double, AccSpace As Double, StartSpd As Double, Acc As Double, YSpeed As Double)
        ' For a given speed, return dependant cutter parameters
        Select Case speed
            Case 0 To 150 : Return (12.0, 0, 60.0, 7000.0, 30.0)
            Case 150 To 250 : Return (14.0, -0.8, 60.0, 7000.0, 30.0)
            Case 250 To 350 : Return (16.0, -0.12, 60.0, 7000.0, 30.0)
            Case 350 To 450 : Return (18.0, -0.25, 60.0, 7000.0, 30.0)
            Case 450 To 550 : Return (20.0, -0.35, 60.0, 7000.0, 30.0)
            Case 550 To 650 : Return (24.0, -0.45, 60.0, 7000.0, 30.0)
            Case 650 To 750 : Return (28.0, -0.55, 60.0, 7000.0, 30.0)
            Case 750 To 1200 : Return (30.0, -0.65, 60.0, 7000.0, 30.0)
            Case Else
                Throw New System.Exception($"Cannot provide engrave parameters for speed of {speed}")
        End Select
    End Function
    Public Function PowerSpeedColor(power As Single, speed As Single) As AciColor
        ' Convert speed/power to an approximate color
        'Return AciColor.FromHsl(30 / 360, 0.6, CSng(power / 100) * (1 - (CSng(speed) / My.Settings.SpeedMax)))
        If power > My.Settings.PowerMax Then Throw New System.Exception($"{power} is greater than maximum {My.Settings.PowerMax}")
        If speed > My.Settings.SpeedMax Then Throw New System.Exception($"{speed} is greater than maximum {My.Settings.SpeedMax}")
        Dim pwr = 1 - (power / My.Settings.PowerMax)
        Dim spd = speed / My.Settings.SpeedMax
        Dim lum = pwr * spd
        ' Vary the luminance of HSL color. For White, L=1, for Black L=0
        Return AciColor.FromHsl(0, 0, lum)
    End Function

    Public Function Distance(p1 As IntPoint, p2 As IntPoint) As Single
        ' Calculate the distance between two points
        Return Math.Sqrt((p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2)
    End Function
    Public Function Distance(p1 As Vector2, p2 As Vector2) As Single
        ' Calculate the distance between two points
        Return Math.Sqrt((p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2)
    End Function

    Public Function SolveQuadratic(s As Double, u As Double, a As Double) As (Double, Double)
        ' Coefficients for the quadratic equation
        ' Solve s = ut + 0.5 a * t^2
        Dim AA As Double = 0.5 * a
        Dim B As Double = u
        Dim C As Double = -s

        ' Calculate the discriminant
        Dim discriminant As Double = B * B - 4 * AA * C

        ' Check if the discriminant is negative, zero, or positive
        If discriminant < 0 Then
            Throw New ArgumentException("No real solutions exist.")
        End If

        ' Calculate the two solutions using the quadratic formula
        Dim t1 As Double = (-B + Math.Sqrt(discriminant)) / (2 * AA)
        Dim t2 As Double = (-B - Math.Sqrt(discriminant)) / (2 * AA)

        ' Return the solutions as a tuple
        Return (t1, t2)
    End Function

End Module
