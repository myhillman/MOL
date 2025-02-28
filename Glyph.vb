Imports netDxf.Entities

Public Class Glyph

    ' A glyph is a single character made up of strokes
    Public Property Strokes As List(Of Polyline2D)
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
    Public Sub New()
        Me.Strokes = New List(Of Polyline2D)()
    End Sub

    Public Sub New(strokes As List(Of Polyline2D))
        If strokes Is Nothing Then
            Throw New ArgumentNullException(NameOf(strokes))
        End If

        Me.Strokes = strokes
    End Sub
End Class
