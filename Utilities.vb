﻿Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports netDxf
Imports netDxf.Entities

' Miscellaneous SHARED routines that don't belong in a class
Friend Module Utilities

    ''' <summary>
    ''' Represents a union that allows a float to be accessed as an integer and vice versa.
    ''' This is useful for low-level manipulation of floating-point numbers, such as converting
    ''' between custom floating-point formats and IEEE 754 format.
    ''' </summary>
    <StructLayout(LayoutKind.Explicit)>
    Public Structure IntFloatUnion
        ''' <summary>
        ''' The integer representation of the floating-point number.
        ''' </summary>
        <FieldOffset(0)>
        Public i As Integer

        ''' <summary>
        ''' The floating-point representation of the integer.
        ''' </summary>
        <FieldOffset(0)>
        Public f As Single
    End Structure

    ''' <summary>
    ''' Converts a custom Leetro floating-point format to an IEEE 754 floating-point number.
    ''' </summary>
    ''' <param name="b">The Leetro floating-point number represented as a Long.</param>
    ''' <returns>The equivalent IEEE 754 floating-point number as a Single.</returns>
    Public Function Float2Double(b As Long) As Single
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

    ''' <summary>
    ''' Converts an IEEE 754 floating-point number to a custom Leetro floating-point format.
    ''' </summary>
    ''' <param name="f">The IEEE 754 floating-point number as a Single.</param>
    ''' <returns>The equivalent Leetro floating-point number represented as an Integer.</returns>
    Public Function Double2Float(f As Single) As Integer
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

    ''' <summary>
    ''' For a given speed, returns the dependent cutter parameters.
    ''' </summary>
    ''' <param name="speed">The speed for which to get the engrave parameters.</param>
    ''' <returns>
    ''' A tuple containing the following parameters:
    ''' <list type="bullet">
    ''' <item><description>Acclen: Acceleration length.</description></item>
    ''' <item><description>AccSpace: Acceleration space.</description></item>
    ''' <item><description>StartSpd: Start speed.</description></item>
    ''' <item><description>Acc: Acceleration.</description></item>
    ''' <item><description>YSpeed: Y-axis speed.</description></item>
    ''' </list>
    ''' </returns>
    ''' <exception cref="Exception">Thrown when the speed is outside the supported range.</exception>
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
                Throw New Exception($"Cannot provide engrave parameters for speed of {speed}")
        End Select
    End Function
    ''' <summary>
    ''' Converts speed and power values to an approximate color.
    ''' </summary>
    ''' <param name="power">The power value as a percentage.</param>
    ''' <param name="speed">The speed value in mm/s.</param>
    ''' <returns>An AciColor representing the combined power and speed values.</returns>
    ''' <exception cref="Exception">Thrown when the power or speed exceeds the maximum allowed values.</exception>
    Public Function PowerSpeedColor(power As Single, speed As Single) As AciColor
        ' Convert speed/power to an approximate color
        'Return AciColor.FromHsl(30 / 360, 0.6, CSng(power / 100) * (1 - (CSng(speed) / My.Settings.SpeedMax)))
        If power > My.Settings.PowerMax Then Throw New Exception($"{power} is greater than maximum {My.Settings.PowerMax}")
        If speed > My.Settings.SpeedMax Then Throw New Exception($"{speed} is greater than maximum {My.Settings.SpeedMax}")
        Dim pwr = 1 - (power / My.Settings.PowerMax)
        Dim spd = speed / My.Settings.SpeedMax
        Dim lum = pwr * spd
        ' Vary the luminance of HSL color. For White, L=1, for Black L=0
        Return AciColor.FromHsl(0, 0, lum)
    End Function

    ''' <summary>
    ''' Calculates the distance between two points represented by IntPoint structures.
    ''' </summary>
    ''' <param name="p1">The first point.</param>
    ''' <param name="p2">The second point.</param>
    ''' <returns>The distance between the two points as a Single.</returns>
    Public Function Distance(p1 As IntPoint, p2 As IntPoint) As Single
        ' Calculate the distance between two points
        Return Math.Sqrt((p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2)
    End Function

    ''' <summary>
    ''' Calculates the distance between two points represented by Vector2 structures.
    ''' </summary>
    ''' <param name="p1">The first point.</param>
    ''' <param name="p2">The second point.</param>
    ''' <returns>The distance between the two points as a Single.</returns>
    Public Function Distance(p1 As Vector2, p2 As Vector2) As Single
        ' Calculate the distance between two points
        Return Math.Sqrt((p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2)
    End Function

    ''' <summary>
    ''' Converts an angle from degrees to radians.
    ''' </summary>
    ''' <param name="deg">The angle in degrees.</param>
    ''' <returns>The angle in radians.</returns>
    Public Function DegToRad(deg As Double) As Double
        ' convert degrees to radians
        Return deg * Math.PI / 180.0
    End Function

    ''' <summary>
    ''' Given a vector startpoint, endpoint, and bulge, creates a Polyline2D representing the bulge points.
    ''' </summary>
    ''' <param name="startpoint">The starting point of the bulge.</param>
    ''' <param name="Endpoint">The ending point of the bulge.</param>
    ''' <param name="bulge">The bulge value, which determines the curvature of the arc.</param>
    ''' <param name="numpoints">The number of points to generate along the arc. Default is 10.</param>
    ''' <returns>A Polyline2D representing the bulge points.</returns>
    ''' <exception cref="Exception">Thrown when the bulge value is infeasibly large.</exception>
    Function GenerateBulge(startpoint As Vector2, Endpoint As Vector2, bulge As Double, Optional numpoints As Integer = 10) As Polyline2D
        ' Given a vector startpoint, endpoint and bulge, create a Polyline2D representing the bulge points
        ' The arc will contain numpoints points, default 10
        Dim result As New Polyline2D
        Select Case Math.Abs(bulge)
            Case Is <= 1
                ' OK
            Case Is <= 2
                MsgBox($"A bulge value of {bulge} is very large.", vbAbort + vbOKOnly, "Warning: large bulge")
            Case Else
                MsgBox($"A bulge value of {bulge} is infeasibly large.", vbAbort + vbOKOnly, "Infeasible bulge")
        End Select

        ' Item1=center, Item2=radius, Item3=StartAngle, Item4=EndAngle
        ' StartAngle and EndAngle are wrt to line joining StartPoint and EndPoint

        Dim arc = MathHelper.ArcFromBulge(startpoint, Endpoint, bulge)
        Dim StartAngleRad = DegToRad(arc.Item3)
        Dim EndAngleRad = DegToRad(arc.Item4)
        ' Ensure the angles are in the correct order
        If EndAngleRad < StartAngleRad Then
            EndAngleRad += 2 * Math.PI
        End If
        Dim angleIncrement = (EndAngleRad - StartAngleRad) / (numpoints - 1)
        For i = 0 To numpoints - 1
            Dim angle = StartAngleRad + i * angleIncrement
            Dim x = arc.Item1.X + arc.Item2 * Math.Cos(angle)
            Dim y = arc.Item1.Y + arc.Item2 * Math.Sin(angle)
            result.Vertexes.Add(New Polyline2DVertex(x, y))
        Next
        Return result
    End Function
    ' Generic function to check if a variable has a value from any enum
    Public Function IsValidEnumValue(Of T As Structure)(value As Integer) As Boolean
        If Not GetType(T).IsEnum Then
            Throw New ArgumentException("T must be an enumerated type")
        End If
        Return [Enum].IsDefined(GetType(T), value)
    End Function

    ''' <summary>
    ''' Retrieves the comments for an enum type and its members from the XML documentation file.
    ''' </summary>
    ''' <param name="enumType">The enum type for which to retrieve comments.</param>
    ''' <returns>
    ''' A dictionary where the keys are the names and values of the enum members, and the values are the corresponding comments.
    ''' </returns>
    ''' <exception cref="FileNotFoundException">Thrown when the XML documentation file is not found.</exception>
    ''' <remarks>
    ''' This function reads the XML documentation file generated by the compiler to extract comments for the specified enum type and its members.
    ''' The XML documentation file must be located in the same directory as the assembly and have the same name as the assembly with a .xml extension.
    ''' </remarks>
    Public Function GetEnumComments(ByVal enumType As Type) As Dictionary(Of String, String)
        Dim comments As New Dictionary(Of String, String)

        ' Load the XML documentation file
        Dim assemblyLocation As String = Assembly.GetExecutingAssembly().Location
        Dim xmlPath As String = Path.ChangeExtension(assemblyLocation, ".xml")
        If Not File.Exists(xmlPath) Then
            Throw New FileNotFoundException("XML documentation file not found.")
        End If

        Dim xmlDoc As XDocument = XDocument.Load(xmlPath)

        ' Get the comments for the enum type
        Dim enumComment = xmlDoc.Descendants("member").
        Where(Function(m) m.Attribute("name").Value = $"T:{enumType.FullName}").
        Select(Function(m) m.Element("summary").Value.Trim()).
        FirstOrDefault()

        If enumComment IsNot Nothing Then
            comments.Add(enumType.Name, enumComment)
        End If

        ' Get the comments for the enum members
        For Each member In enumType.GetFields(BindingFlags.Public Or BindingFlags.Static)
            Dim memberComment = xmlDoc.Descendants("member").
            Where(Function(m) m.Attribute("name").Value = $"F:{enumType.FullName}.{member.Name}").
            Select(Function(m) m.Element("summary").Value.Trim()).
            FirstOrDefault()

            If memberComment IsNot Nothing Then
                comments.Add($"{member.Name}={CInt(member.GetValue(Nothing))}", memberComment)
            End If
        Next

        Return comments
    End Function

    Public Function GetStructureComments(ByVal structType As Type) As Dictionary(Of String, String)
        Dim comments As New Dictionary(Of String, String)

        Try
            ' Load the XML documentation file
            Dim assemblyLocation As String = Assembly.GetExecutingAssembly().Location
            Dim xmlPath As String = Path.ChangeExtension(assemblyLocation, ".xml")
            If Not File.Exists(xmlPath) Then
                Throw New FileNotFoundException("XML documentation file not found.")
            End If

            Dim xmlDoc As XDocument = XDocument.Load(xmlPath)

            ' Get the comments for the structure type
            Dim structComment = xmlDoc.Descendants("member").
                Where(Function(m) m.Attribute("name").Value = $"T:{structType.FullName}").
                Select(Function(m) m.Element("summary").Value.Trim()).
                FirstOrDefault()

            If structComment IsNot Nothing Then
                comments.Add(structType.Name, structComment)
            End If

            ' Get the comments for the structure members
            For Each member In structType.GetFields(BindingFlags.Public Or BindingFlags.Instance)
                Dim memberComment = xmlDoc.Descendants("member").
                    Where(Function(m) m.Attribute("name").Value = $"F:{structType.FullName}.{member.Name}").
                    Select(Function(m) m.Element("summary").Value.Trim()).
                    FirstOrDefault()

                If memberComment IsNot Nothing Then
                    comments.Add(member.Name, memberComment)
                End If
            Next

        Catch ex As Exception
            ' Log the exception message
            Console.WriteLine($"Error: {ex.Message}")
        End Try

        Return comments
    End Function
End Module
