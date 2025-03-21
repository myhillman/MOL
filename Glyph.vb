Imports netDxf.Entities

''' <summary>
''' Represents a glyph, which is a single character made up of strokes.
''' </summary>
Public Class Glyph

    ''' <summary>
    ''' Gets or sets the list of strokes that make up the glyph.
    ''' Each stroke is represented by a <see cref="Polyline2D"/>.
    ''' </summary>
    ''' <value>The list of strokes that make up the glyph.</value>
    Public Property Strokes As List(Of Polyline2D)

    ''' <summary>
    ''' Gets the width of the glyph, which is the maximum X position of all vertices in the strokes.
    ''' </summary>
    ''' <value>The width of the glyph.</value>
    Public ReadOnly Property Width As Double
        Get
            Dim w As Double = 0
            For Each s In Strokes
                For Each v In s.Vertexes
                    w = Math.Max(w, v.Position.X)
                Next
            Next
            Width = w
            Exit Property
        End Get
    End Property

    ''' <summary>
    ''' Initializes a new instance of the <see cref="Glyph"/> class with an empty list of strokes.
    ''' </summary>
    Public Sub New()
        Me.Strokes = New List(Of Polyline2D)()
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="Glyph"/> class with the specified list of strokes.
    ''' </summary>
    ''' <param name="strokes">The list of strokes that make up the glyph.</param>
    ''' <exception cref="ArgumentNullException">Thrown when the <paramref name="strokes"/> parameter is null.</exception>
    Public Sub New(strokes As List(Of Polyline2D))
        ArgumentNullException.ThrowIfNull(strokes)
        Me.Strokes = strokes
    End Sub
End Class


