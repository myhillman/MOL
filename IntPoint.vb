Imports netDxf

Public Class IntPoint
    ' class represents a Point with integer X and Y
    Public Property X As Integer
    Public Property Y As Integer
    Public ReadOnly Property Length As Double   ' length of vector assuming a (0,0) origin
        Get
            Return Math.Sqrt(X * X + Y * Y)
        End Get
    End Property

    Public Sub New(ByVal x As Integer, ByVal y As Integer)
        Me.X = x
        Me.Y = y
    End Sub

    ' p1 + p2
    Public Shared Operator +(ByVal p1 As IntPoint, p2 As IntPoint) As IntPoint
        Return New IntPoint(p1.X + p2.X, p1.Y + p2.Y)
    End Operator

    ' p1 - p2
    Public Shared Operator -(ByVal p1 As IntPoint, p2 As IntPoint) As IntPoint
        Return New IntPoint(p1.X - p2.X, p1.Y - p2.Y)
    End Operator

    ' Test for equality
    Public Shared Operator =(ByVal p1 As IntPoint, p2 As IntPoint) As Boolean
        Return p1.X = p2.X AndAlso p1.Y = p2.Y
    End Operator

    ' Test for inequality
    Public Shared Operator <>(ByVal p1 As IntPoint, p2 As IntPoint) As Boolean
        Return Not p1 = p2
    End Operator

    ' Multiply by scaler
    Public Shared Operator *(ByVal p1 As IntPoint, mult As Single) As IntPoint
        Return New IntPoint(p1.X * mult, p1.Y * mult)
    End Operator

    ' Divide by scaler
    Public Shared Operator /(ByVal p1 As IntPoint, mult As Single) As IntPoint
        Return New IntPoint(p1.X / mult, p1.Y / mult)
    End Operator

    ' system.windows.point are in mm, IntPoint are in steps. Convert using scale
    Public Shared Widening Operator CType(ByVal p1 As IntPoint) As System.Windows.Point
        Return New System.Windows.Point(p1.X / Form1.xScale, p1.Y / Form1.yScale)
    End Operator

    ' system.windows.point are in mm, IntPoint are in steps. Convert using scale
    Public Shared Widening Operator CType(ByVal p1 As System.Windows.Point) As IntPoint
        Return New IntPoint(p1.X * Form1.xScale, p1.Y * Form1.yScale)
    End Operator

    ' Convert an IntPoint to a vector2 for DXF. Converted point is scaled to mm
    Public Shared Widening Operator CType(ByVal p1 As IntPoint) As Vector2
        Return New Vector2(p1.X / Form1.xScale, p1.Y / Form1.yScale)
    End Operator
End Class