Imports netDxf

''' <summary>
''' Represents a point with integer X and Y coordinates.
''' </summary>
Public Class IntPoint
    ''' <summary>
    ''' Gets or sets the X coordinate of the point.
    ''' </summary>
    ''' <value>The X coordinate of the point.</value>
    Public Property X As Integer

    ''' <summary>
    ''' Gets or sets the Y coordinate of the point.
    ''' </summary>
    ''' <value>The Y coordinate of the point.</value>
    Public Property Y As Integer

    ''' <summary>
    ''' Gets the length of the vector from the origin (0, 0) to this point.
    ''' </summary>
    ''' <value>The length of the vector.</value>
    Public ReadOnly Property Length As Double
        Get
            Return Math.Sqrt(CLng(X) * CLng(X) + CLng(Y) * CLng(Y))
        End Get
    End Property

    ''' <summary>
    ''' Initializes a new instance of the <see cref="IntPoint"/> class with the specified coordinates.
    ''' </summary>
    ''' <param name="x">The X coordinate of the point.</param>
    ''' <param name="y">The Y coordinate of the point.</param>
    Public Sub New(ByVal x As Integer, ByVal y As Integer)
        Me.X = x
        Me.Y = y
    End Sub

    ''' <summary>
    ''' Adds two <see cref="IntPoint"/> instances.
    ''' </summary>
    ''' <param name="p1">The first point.</param>
    ''' <param name="p2">The second point.</param>
    ''' <returns>A new <see cref="IntPoint"/> that is the sum of the two points.</returns>
    Public Shared Operator +(ByVal p1 As IntPoint, p2 As IntPoint) As IntPoint
        Return New IntPoint(p1.X + p2.X, p1.Y + p2.Y)
    End Operator

    ''' <summary>
    ''' Subtracts one <see cref="IntPoint"/> from another.
    ''' </summary>
    ''' <param name="p1">The point to subtract from.</param>
    ''' <param name="p2">The point to subtract.</param>
    ''' <returns>A new <see cref="IntPoint"/> that is the difference of the two points.</returns>
    Public Shared Operator -(ByVal p1 As IntPoint, p2 As IntPoint) As IntPoint
        Return New IntPoint(p1.X - p2.X, p1.Y - p2.Y)
    End Operator

    ''' <summary>
    ''' Determines whether two <see cref="IntPoint"/> instances are equal.
    ''' </summary>
    ''' <param name="p1">The first point.</param>
    ''' <param name="p2">The second point.</param>
    ''' <returns><c>True</c> if the points are equal; otherwise, <c>False</c>.</returns>
    Public Shared Operator =(ByVal p1 As IntPoint, p2 As IntPoint) As Boolean
        Return p1.X = p2.X AndAlso p1.Y = p2.Y
    End Operator

    ''' <summary>
    ''' Determines whether two <see cref="IntPoint"/> instances are not equal.
    ''' </summary>
    ''' <param name="p1">The first point.</param>
    ''' <param name="p2">The second point.</param>
    ''' <returns><c>True</c> if the points are not equal; otherwise, <c>False</c>.</returns>
    Public Shared Operator <>(ByVal p1 As IntPoint, p2 As IntPoint) As Boolean
        Return Not p1 = p2
    End Operator

    ''' <summary>
    ''' Multiplies an <see cref="IntPoint"/> by a scalar value.
    ''' </summary>
    ''' <param name="p1">The point to multiply.</param>
    ''' <param name="mult">The scalar value.</param>
    ''' <returns>A new <see cref="IntPoint"/> that is the product of the point and the scalar.</returns>
    Public Shared Operator *(ByVal p1 As IntPoint, mult As Single) As IntPoint
        Return New IntPoint(p1.X * mult, p1.Y * mult)
    End Operator

    ''' <summary>
    ''' Divides an <see cref="IntPoint"/> by a scalar value.
    ''' </summary>
    ''' <param name="p1">The point to divide.</param>
    ''' <param name="mult">The scalar value.</param>
    ''' <returns>A new <see cref="IntPoint"/> that is the quotient of the point and the scalar.</returns>
    Public Shared Operator /(ByVal p1 As IntPoint, mult As Single) As IntPoint
        Return New IntPoint(p1.X / mult, p1.Y / mult)
    End Operator

    ''' <summary>
    ''' Converts an <see cref="IntPoint"/> to a <see cref="System.Windows.Point"/>.
    ''' </summary>
    ''' <param name="p1">The <see cref="IntPoint"/> to convert.</param>
    ''' <returns>A <see cref="System.Windows.Point"/> that represents the same point.</returns>
    Public Shared Widening Operator CType(ByVal p1 As IntPoint) As System.Windows.Point
        Return New System.Windows.Point(p1.X / Form1.xScale, p1.Y / Form1.yScale)
    End Operator

    ''' <summary>
    ''' Converts a <see cref="System.Windows.Point"/> to an <see cref="IntPoint"/>.
    ''' </summary>
    ''' <param name="p1">The <see cref="System.Windows.Point"/> to convert.</param>
    ''' <returns>An <see cref="IntPoint"/> that represents the same point.</returns>
    Public Shared Widening Operator CType(ByVal p1 As System.Windows.Point) As IntPoint
        Return New IntPoint(p1.X * Form1.xScale, p1.Y * Form1.yScale)
    End Operator

    ''' <summary>
    ''' Converts an <see cref="IntPoint"/> to a <see cref="Vector2"/> for DXF.
    ''' </summary>
    ''' <param name="p1">The <see cref="IntPoint"/> to convert.</param>
    ''' <returns>A <see cref="Vector2"/> that represents the same point, scaled to millimeters.</returns>
    Public Shared Widening Operator CType(ByVal p1 As IntPoint) As Vector2
        Return New Vector2(p1.X / Form1.xScale, p1.Y / Form1.yScale)
    End Operator
End Class
