Imports System.IO
Imports System.Runtime.InteropServices
Imports netDxf
Imports netDxf.Entities
Imports netDxf.Tables
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar
Imports System.Windows.Media
Imports System.Windows
Imports System.Globalization
Imports System.Windows.Shapes

' Define the different block types
Enum BLOCK_TYPE
    HEADER
    CONFIG
    TEST
    CUT
    DRAW
End Enum

' Define a custom type
Enum OnOff_type
    Off = 0
    _On = 1
End Enum
Public Class Form1
    Const BLOCK_SIZE = 512
    Dim debugflag As Boolean = True
    Dim hexdump As Boolean = False
    Dim dlg = New OpenFileDialog
    Dim stream As FileStream, reader As BinaryReader, writer As BinaryWriter
    Dim ConfigChunk As Integer, TestChunk As Integer, CutChunk As Integer
    Dim DrawChunks As New List(Of Integer)      ' list of draw chunks
    Dim TopRight As System.Drawing.Point, BottomLeft As System.Drawing.Point
    Dim StartPosns As New Dictionary(Of Integer, Vector2)   ' start position for drawing
    Dim startposn As Vector2
    Dim position As Vector2    ' current laser head position
    Dim delta As Vector2       ' a delta from position
    Dim xScale As Single = 160.07    ' X axis steps/mm
    Dim yScale As Single = 160.07    ' Y axis steps/mm
    Dim CurrentBlockType
    Const NumColors = 255
    Dim colors(NumColors) As AciColor   ' array of colors
    Dim ColorIndex As Integer = 1       ' index into colors array
    Dim CommandUsage As New Dictionary(Of Integer, Integer)
    Dim dxf As New DxfDocument()
    ' Define some layers for DXF file
    Dim TextLayer As New Layer("Text") With {       ' text
                .Color = AciColor.Default
            }
    Dim EngraveLayer As New Layer("Engrave") With { ' engraving
                .Color = AciColor.Default,
                .Lineweight = 10     ' 0.1 mm
            }
    Dim DrawLayer = New Layer("Draw") With {        ' drawing
            .Color = AciColor.Red,
            .Linetype = Linetype.Continuous,
            .Lineweight = Lineweight.Default
        }
    Dim CFGLayer = New Layer("Config") With {   ' config
                .Color = AciColor.Blue,
                .Linetype = Linetype.DashDot,
                .Lineweight = Lineweight.Default
        }
    Dim TestLayer = New Layer("Test") With {    ' Test bounding box
            .Color = AciColor.Yellow,
            .Linetype = Linetype.DashDot,
            .Lineweight = Lineweight.Default
        }
    Dim CutLayer = New Layer("Cut") With {      ' Cut bounding box
            .Color = AciColor.Red,
            .Linetype = Linetype.Dashed,
            .Lineweight = Lineweight.Default
        }
    Dim MoveLayer = New Layer("Move") With {    ' head movement without laser on
            .Color = AciColor.Default,
            .Linetype = Linetype.Continuous,
            .Lineweight = 1,
            .IsVisible = False
        }
    Dim EngraveCutLayer = New Layer("Engrave Cut") With {       ' Engrave move laser on
            .Lineweight = Lineweight.Default
        }
    Dim EngraveMoveLayer = New Layer("Engrave Move") With {     ' Engrave laser off
            .Lineweight = Lineweight.Default
        }


    Class parameter
        ' a class for a parameter
        Property name As String     ' the name of the parameter
        Property type As Type       ' the type of the parameter
        Sub New(name As String, type As Type)
            Me.name = name
            Me.type = type
        End Sub

    End Class
    Class LASERcmd
        Property mnemonic As String
        Property parameters As New List(Of parameter)    ' the type of the parameter
        Sub New(mnemonic As String)
            Me.mnemonic = mnemonic
        End Sub
        Sub New(mnemonic As String, p As List(Of parameter))
            Me.mnemonic = mnemonic
            Me.parameters = p
        End Sub
    End Class
    ' Definitions of full (param count + command) MOL file commands
    Const MOL_MOVEREL = &H3026000
    Const MOL_START = &H3026040
    Const MOL_ORIGIN = &H346040
    Const MOL_MOTION_COMMAND_COUNT = &H3090080
    Const MOL_BEGSUB = &H1300008
    Const MOL_BEGSUBa = &H1300048
    Const MOL_ENDSUB = &H1400048
    Const MOL_MCB = &H80000946
    Const MOL_SETSPD = &H3000301
    Const MOL_MOTION = &H3000341
    Const MOL_LASER = &H1000606
    Const MOL_GOSUB = &H1500048
    Const MOL_BLOWER = &H1004601
    Const MOL_BLOWERa = &H1004A41
    Const MOL_BLOWERb = &H1004B41
    Const MOL_SEGMENT = &H500008
    Const MOL_ENGMVX = &H210040
    Const MOL_ENGMVY = &H214040
    Const MOL_END = 0

    ' Dictionary of all commands
    Dim LASERcmds As New Dictionary(Of Integer, LASERcmd) From {
        {MOL_MOVEREL, New LASERcmd("MOVEREL", New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))},
                            {New parameter("dx", GetType(Int32))},
                            {New parameter("dy", GetType(Int32))}
                            }
                           )},
        {MOL_START, New LASERcmd("START", New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))},
                            {New parameter("x", GetType(Int32))},
                            {New parameter("y", GetType(Int32))}
                            }
                           )},
        {MOL_ORIGIN, New LASERcmd("ORIGIN", New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))},
                            {New parameter("x", GetType(Int32))},
                            {New parameter("y", GetType(Int32))}
                            }
                           )},
        {MOL_MOTION_COMMAND_COUNT, New LASERcmd("MOTION_COMMAND_COUNT", New List(Of parameter) From {{New parameter("count", GetType(Int32))}})},
        {MOL_BEGSUB, New LASERcmd("BEGSUB", New List(Of parameter) From {{New parameter("n", GetType(Int32))}})},
        {MOL_BEGSUBa, New LASERcmd("BEGSUBa", New List(Of parameter) From {{New parameter("n", GetType(Int32))}})},
        {MOL_ENDSUB, New LASERcmd("ENDSUB", New List(Of parameter) From {{New parameter("n", GetType(Int32))}})},
        {MOL_MCB, New LASERcmd("MCB", New List(Of parameter) From {{New parameter("Size", GetType(Int32))}})},
        {MOL_MOTION, New LASERcmd("MOTION", New List(Of parameter) From {
                            {New parameter("a", GetType(Int32))},
                            {New parameter("b", GetType(Int32))},
                            {New parameter("c", GetType(Int32))}
                            }
                           )},
        {MOL_SETSPD, New LASERcmd("SETSPD", New List(Of parameter) From {
                            {New parameter("a", GetType(Single))},
                            {New parameter("b", GetType(Single))},
                            {New parameter("c", GetType(Single))}
                            }
                           )},
        {MOL_LASER, New LASERcmd("LASER", New List(Of parameter) From {{New parameter("On/Off", GetType(OnOff_type))}})},
        {MOL_GOSUB, New LASERcmd("GOSUB", New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))},
                            {New parameter("x", GetType(Single))},
                            {New parameter("y", GetType(Single))}
                            }
                           )},
        {MOL_SEGMENT, New LASERcmd("SEGMENT", New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))},
                            {New parameter("x", GetType(Single))},
                            {New parameter("y", GetType(Single))}
                            }
                           )},
        {MOL_BLOWER, New LASERcmd("BLOWER", New List(Of parameter) From {{New parameter("On/Off", GetType(OnOff_type))}})},
        {MOL_BLOWERa, New LASERcmd("BLOWERa", New List(Of parameter) From {{New parameter("On/Off", GetType(OnOff_type))}})},
        {MOL_BLOWERb, New LASERcmd("BLOWERb", New List(Of parameter) From {{New parameter("On/Off", GetType(OnOff_type))}})},
        {MOL_ENGMVX, New LASERcmd("ENGMVX", New List(Of parameter) From {{New parameter("Axis", GetType(Integer))}, {New parameter("n", GetType(Integer))}})},
        {MOL_ENGMVY, New LASERcmd("ENGMVX", New List(Of parameter) From {{New parameter("x", GetType(Integer))}, {New parameter("OnOff", GetType(Boolean))}})},
        {MOL_END, New LASERcmd("END")}
        }

    Private Function OpenToolStripMenuItem1_ClickAsync(sender As Object, e As EventArgs) As Task Handles OpenToolStripMenuItem1.Click
        Dim entity As Text
        ' Create a list of 255 colors
        Const saturation As Double = 1
        Const lightness As Double = 0.5
        Dim stp As Double = 1.0 / NumColors
        For i = 0 To NumColors
            Dim hue As Double = i * stp
            colors(i) = AciColor.FromHsl(hue, saturation, lightness)
        Next

        With dlg
            .Filter = "MOL file|*.mol|All files|*.*"
        End With
        Dim result = dlg.ShowDialog()
        If result.OK Then
            ' Init internal variables
            dxf = New DxfDocument()
            DrawChunks.Clear()
            StartPosns.Clear()
            position = New Vector2(0, 0)
            stream = File.Open(dlg.filename, FileMode.Open)
            reader = New BinaryReader(stream, System.Text.Encoding.Unicode, False)

            Dim Size As Long = reader.BaseStream.Length / 4     ' size of the inputfile in 4 byte words
            Dim MOLfile(Size) As Int32                          ' buffer for whole file
            ' Read the whole file
            Dim i = 0
            While reader.BaseStream.Position < reader.BaseStream.Length
                MOLfile(i) = reader.ReadInt32
                i += 1
            End While

            DisplayHeader()

            DisplayConfig(dxf)

            DisplayTest(dxf)

            DisplayCut(dxf)

            DisplayDraw(dxf)

            Dim BaseDXF = System.IO.Path.GetFileNameWithoutExtension(dlg.filename)
            Dim DXFfile = $"{BaseDXF}_MOL.dxf"

            ' Add the filename at the bottom of the image
            Dim width = TopRight.X - BottomLeft.X
            Dim height = TopRight.Y - BottomLeft.Y
            entity = New Text($"File: {DXFfile}") With {
                .Position = New Vector3(BottomLeft.X + width / 2, BottomLeft.Y - height / 10, 0),
                .Alignment = Entities.TextAlignment.TopCenter,
                .Layer = TextLayer,
                .Height = Math.Max(width, height) / 20
                }
            dxf.Entities.Add(entity)
            ' Save the file
            dxf.Save(DXFfile)
            stream.Close()
            ' display the list of start points
            TextBox1.AppendText($"Start points: {StartPosns.Count}{vbCrLf}")
            For Each sp In StartPosns
                TextBox1.AppendText($" {sp.Key} : x={sp.Value.X},y={sp.Value.Y}{vbCrLf}")
            Next
            ' Open dxf file
            Dim myProcess As New Process
            With myProcess
                .StartInfo.FileName = DXFfile
                .StartInfo.UseShellExecute = True
                .StartInfo.RedirectStandardOutput = False
                .Start()
                .Dispose()
            End With

        End If
    End Function
    Function GetInt() As Integer
        ' read 4 byte integer from input stream
        Return reader.ReadInt32
    End Function
    Function GetUInt() As UInteger
        ' read 4 byte uinteger from input stream
        Return reader.ReadUInt32
    End Function
    Function GetInt(n As Integer) As Integer
        reader.BaseStream.Seek(n, SeekOrigin.Begin)    ' reposition to offset n
        Return GetInt()
    End Function

    Function GetFloat() As Single
        ' read 4 byte float from data
        Return Leetro2Float(GetUInt())
    End Function

    Public Function GetFloat(n As Integer) As Single
        reader.BaseStream.Seek(n, SeekOrigin.Begin)    ' reposition to offset n
        Return GetFloat()
    End Function

    Sub PutFloat(f As Single)
        writer.Write(Float2Int(f))
    End Sub

    Sub PutInt(n As Integer)
        writer.Write(n)
    End Sub

    ' Create a union so that a float can be accessed as a UInt
    <StructLayout(LayoutKind.Explicit)> Public Structure IntFloatUnion
        <FieldOffset(0)>
        Public i As UInteger
        <FieldOffset(0)> Dim f As Single
    End Structure

    Sub DisplayHeader()
        ' Display header info
        Dim chunk As Integer

        CurrentBlockType = BLOCK_TYPE.HEADER
        reader.BaseStream.Position = 0
        TextBox1.AppendText($"HEADER &0{vbCrLf}")
        TextBox1.AppendText($"Size of file: &h{GetInt():x}{vbCrLf}")
        Dim MotionBlocks = GetInt()
        ' Create a progress bar based on Motion Blocks
        With ProgressBar1
            .Minimum = 0
            .Maximum = MotionBlocks - 33
            .Value = 0
        End With
        TextBox1.AppendText($"Motion blocks: {MotionBlocks}{vbCrLf}")
        Dim v() As Byte = reader.ReadBytes(4)
        TextBox1.AppendText($"Version: {CInt(v(3))}.{CInt(v(2))}.{CInt(v(1))}.{CInt(v(0))}{vbCrLf}")
        TopRight = New System.Drawing.Point(GetInt(&H18), GetInt())
        TextBox1.AppendText($"Origin: {TopRight.X},{TopRight.Y}{vbCrLf}")
        BottomLeft = New System.Drawing.Point(GetInt(&H20), GetInt())
        TextBox1.AppendText($"Bottom Left: ({BottomLeft.X},{BottomLeft.Y}){vbCrLf}")
        ' Scale factors
        Dim n = GetInt(&H24C)   ' ??
        xScale = GetFloat()     ' X scale
        yscale = GetFloat()     ' Y scale
        TextBox1.AppendText($"X scale: {xScale}, Y scale {yscale}{vbCrLf}")
        ' Speed/accel
        Dim arg1 = GetFloat(&H274)
        Dim arg2 = GetFloat()
        Dim arg3 = GetFloat()
        TextBox1.AppendText($"Initial speed: {arg1}, Max speed {arg2}, Accel {arg3}{vbCrLf}")
        'Object size
        n = GetInt(&H284)
        Dim x = GetInt() / xScale
        Dim y = GetInt() / yscale
        TextBox1.AppendText($"Object size {x}mm x {y}mm{vbCrLf}")

        ConfigChunk = GetInt(&H70)
        TestChunk = GetInt(&H74)
        CutChunk = GetInt(&H78)
        chunk = GetInt(&H7C)        ' first draw chunk
        While chunk <> 0
            DrawChunks.Add(chunk)
            chunk = GetInt()
        End While
        TextBox1.AppendText($"Config chunk: {ConfigChunk}{vbCrLf}")
        TextBox1.AppendText($"Test chunk: {TestChunk}{vbCrLf}")
        TextBox1.AppendText($"Cut chunk: {CutChunk}{vbCrLf}")
        TextBox1.AppendText($"Draw chunks: {String.Join(",", DrawChunks.ToArray)}{vbCrLf}")
        TextBox1.AppendText($"{vbCrLf}")
    End Sub

    Sub DisplayConfig(dxf As DxfDocument)
        ' Display config block info
        CurrentBlockType = BLOCK_TYPE.CONFIG
        Dim StartOfBlock = ConfigChunk * BLOCK_SIZE
        TextBox1.AppendText($"{vbCrLf}CONFIG @&{StartOfBlock:x}{vbCrLf}")
        reader.BaseStream.Seek(StartOfBlock, SeekOrigin.Begin)    ' reposition to start of block
        DecodeStream(dxf, CFGLayer, False, False)
        TextBox1.AppendText($"{vbCrLf}")
    End Sub
    Sub DisplayTest(ByRef dxf As DxfDocument)
        ' Display test block info
        CurrentBlockType = BLOCK_TYPE.TEST
        Dim StartOfBlock = TestChunk * BLOCK_SIZE

        TextBox1.AppendText($"{vbCrLf}TEST @&{StartOfBlock:x}{vbCrLf}")
        If hexdump Then
            HexDumpBlock(StartOfBlock)
            TextBox1.AppendText($"{vbCrLf}")
        End If
        reader.BaseStream.Seek(StartOfBlock, SeekOrigin.Begin)    ' reposition to offset n
        position = New Vector2(0, 0)
        DecodeStream(dxf, TestLayer, False, True)
    End Sub

    Sub DisplayCut(ByRef dxf As DxfDocument)
        ' Display cut block info
        CurrentBlockType = BLOCK_TYPE.CUT
        Dim StartOfBlock = CutChunk * BLOCK_SIZE
        If StartOfBlock > 0 Then
            TextBox1.AppendText($"{vbCrLf}CUT @&{StartOfBlock:x}{vbCrLf}")
            If hexdump Then
                HexDumpBlock(StartOfBlock)
                TextBox1.AppendText($"{vbCrLf}")
            End If
            TextBox1.AppendText($"{vbCrLf}")
            reader.BaseStream.Seek(StartOfBlock, SeekOrigin.Begin)    ' reposition to offset n
            DecodeStream(dxf, CutLayer, False, False)
        Else
            TextBox1.AppendText($"There is no CUT chunk{vbCrLf}")
        End If
    End Sub
    Sub DisplayDraw(ByRef dxf As DxfDocument)
        ' Display draw block info
        CurrentBlockType = BLOCK_TYPE.DRAW

        ' Draw all chunks
        For Each chunk In DrawChunks
            Dim StartOfBlock = chunk * BLOCK_SIZE
            TextBox1.AppendText($"{vbCrLf}DRAW @&{StartOfBlock:x}{vbCrLf}")
            If hexdump Then
                HexDumpBlock(StartOfBlock)
                TextBox1.AppendText($"{vbCrLf}")
            End If
            reader.BaseStream.Seek(StartOfBlock, SeekOrigin.Begin)    ' reposition to offset n
            DecodeStream(dxf, DrawLayer, True, False)
            TextBox1.AppendText($"{vbCrLf}")
        Next
    End Sub

    Sub DecodeStream(dxf As DxfDocument, BaseLayer As Layer, CanChangeLayer As Boolean, Optional ShowMovement As Boolean = False)
        Dim done As Boolean = False, cmd As Integer, nWords As Integer
        Dim LaserIsOn As Boolean = False

        Dim spdMin As Single = 100.0
        Dim spdMax As Single = 100.0
        Dim i As UInteger
        Dim layer As Layer = BaseLayer
        Dim motion As New Polyline2D With {.Layer = layer}      ' new motion block
        If CurrentBlockType = BLOCK_TYPE.CONFIG Or CurrentBlockType = BLOCK_TYPE.CUT Or CurrentBlockType = BLOCK_TYPE.TEST Then
            ' These blocks have an implicit start position of (0,0) 
            position = New Vector2(0, 0)
        End If
        If CurrentBlockType = BLOCK_TYPE.TEST Then motion.Vertexes.Add(New Polyline2DVertex(position))

        While Not done And stream.Position < stream.Length
            If debugflag Then TextBox1.AppendText($"@{stream.Position:x}")
            Dim cmd_posn = reader.BaseStream.Position   ' position of start of command
            cmd = GetInt()
            Dim command = cmd And &HFFFFFF    ' bottom 24 bits is command
            nWords = (cmd >> 24) And &HFF  ' command length is top 8 bits
            If nWords = &H80 Then nWords = GetInt() And &HFFFF
            Dim param_posn = stream.Position  ' remember where parameters (if any) start
            ' Keep track of command usage
            If Not CommandUsage.TryAdd(cmd, 1) Then
                CommandUsage(cmd) += 1
            End If
            If debugflag Then TextBox1.AppendText($" CMD &h{cmd:x8} ")
            Select Case command
                Case &H0        ' END
                    done = True
                    If debugflag Then TextBox1.AppendText("END")
                    If motion.Vertexes.Count > 0 Then
                        ' add the motion to the dxf file    
                        dxf.Entities.Add(motion)
                        motion = New Polyline2D With {.Layer = layer}
                    End If
                Case MOL_SETSPD And &HFFFFFF, MOL_MOTION And &HFFFFFF   ' SETSPD/MOTION
                    decodeCmd(cmd_posn)
                Case MOL_LASER And &HFFFFFF     ' switch laser
                    decodeCmd(cmd_posn)
                    reader.BaseStream.Seek(param_posn, SeekOrigin.Begin)    ' backup to read the parameters again
                    i = GetUInt()
                    Select Case i
                        Case 0 : LaserIsOn = False
                            layer = BaseLayer
                        Case 1 : LaserIsOn = True
                            layer = MoveLayer
                        Case Else
                            If debugflag Then TextBox1.AppendText($"Unknown LaserIsOn parameter &h{i:x8}")
                    End Select
                    If motion.Vertexes.Count > 0 Then
                        ' Display any motion
                        motion.Layer = layer
                        dxf.Entities.Add(motion)
                        motion = New Polyline2D With {.Layer = layer}
                    End If
                Case MOL_MCB And &HFFFFFF   ' Begin Motion Block
                    decodeCmd(cmd_posn)  ' trick for MCB only to ensure the MCB is decoded
                    param_posn -= 4
                    nWords = 1
                    With ProgressBar1
                        If .Value < .Maximum Then .Value += 1
                    End With
                Case &H4601, &HB06  ' Blower
                    decodeCmd(cmd_posn)
                Case &HE46
                    Dim params As New List(Of Integer)
                    ' Collect parameters. Could be 3, 5 or 7
                    For i = 1 To nWords
                        params.Add(GetFloat)
                    Next
                    If debugflag Then TextBox1.AppendText($"Laser parameters: {String.Join(",", params)}")
                Case MOL_MOVEREL And &HFFFFFF
                    decodeCmd(cmd_posn)
                    reader.BaseStream.Seek(param_posn, SeekOrigin.Begin)    ' backup to read the parameters again
                    Dim n = GetUInt()
                    Dim delta = New Vector2(GetInt(), GetInt())      ' move relative command
                    If CurrentBlockType = BLOCK_TYPE.CONFIG Then
                        startposn = delta ' a MOVEREL in the CONFIG section is a START
                    Else
                        If motion.Vertexes.Count = 0 Then motion.Vertexes.Add(New Polyline2DVertex(position))   ' first move - add start point
                        position += delta
                        motion.Vertexes.Add(New Polyline2DVertex(position))
                    End If
                Case MOL_START And &HFFFFFF
                    decodeCmd(cmd_posn)
                    reader.BaseStream.Seek(param_posn, SeekOrigin.Begin)    ' backup to read the parameters again
                    Dim n = GetUInt()
                    If motion.Vertexes.Count = 0 Then motion.Vertexes.Add(New Polyline2DVertex(position))   ' first move - add start point
                    Dim posn = New Vector2(GetInt(), GetInt())      ' move relative command
                    startposn = posn       ' remember start position for this block
                Case &H4601       ' Accelerate/Decelerate
                    Dim acc = GetInt()
                    Select Case acc
                        Case &H1 : If debugflag Then TextBox1.AppendText($"Accelerate")
                        Case &H2 : If debugflag Then TextBox1.AppendText($"Decelerate")
                        Case Else
                            If debugflag Then TextBox1.AppendText($"Unknown acceleration parameter &h{acc:x8}")
                    End Select
                Case MOL_ENGMVX And &HFFFFFF    ' Engrave move X (laser off)
                    decodeCmd(cmd_posn)
                    Dim axis = GetInt()
                    Dim AxisStr As String
                    Select Case axis
                        Case 3
                            AxisStr = "Y"
                        Case 4
                            AxisStr = "X"
                        Case Else
                            AxisStr = "Unknown"
                    End Select
                    Dim n = GetInt()
                    If debugflag Then TextBox1.AppendText($"Engrave move axis={AxisStr},n={n}")
                Case MOL_ENGMVY And &HFFFFFF  ' Engrave move Y (laser on/off)
                    Dim x = GetInt()
                    Dim OnOff = GetInt()
                    If debugflag Then TextBox1.AppendText($"Engrave move x={x}, laser={OnOff}")
                Case MOL_ORIGIN And &HFFFFFF    ' origin
                    decodeCmd(cmd_posn)
                Case MOL_MOTION_COMMAND_COUNT
                    decodeCmd(cmd_posn)
                Case MOL_BEGSUB And &HFFFFFF, MOL_BEGSUBa And &HFFFFFF ' begin subroutine
                    decodeCmd(cmd_posn)
                    reader.BaseStream.Seek(param_posn, SeekOrigin.Begin)    ' backup to read the parameters again
                    Dim n = GetUInt()
                    If Not StartPosns.ContainsKey(n) Then
                        position = New Vector2(0, 0)
                    Else
                        If n = 3 Then
                            position = StartPosns(n)
                        Else
                            position += StartPosns(n)
                        End If
                        motion.Vertexes.Add(New Polyline2DVertex(position))
                    End If
                    ' Create layer for draw
                    If CanChangeLayer Then
                        layer = New Layer($"Draw {n}") With
                    {
                        .Color = colors(ColorIndex),
                        .Linetype = Linetype.Continuous,
                        .Lineweight = Lineweight.W100
                    }
                        ColorIndex += 20
                    End If
                Case MOL_ENDSUB And &HFFFFFF ' end subroutine
                    decodeCmd(cmd_posn)
                    reader.BaseStream.Seek(param_posn, SeekOrigin.Begin)    ' backup to read the parameters again
                    i = GetUInt()
                    If motion.Vertexes.Count > 0 Then
                        ' add the motion to the dxf file    
                        dxf.Entities.Add(motion)
                        motion = New Polyline2D With {.Layer = layer}
                    End If
                Case MOL_GOSUB And &HFFFFFF, MOL_SEGMENT And &HFFFFFF ' GOSUB/SEGMENT
                    decodeCmd(cmd_posn)
                    reader.BaseStream.Seek(param_posn, SeekOrigin.Begin)    ' backup to read the parameters again
                    Select Case nWords
                        Case 1
                            If debugflag Then TextBox1.AppendText($"End of segment list")
                        Case 3
                            Dim n = GetInt()
                            Dim x = CInt(GetFloat() * xScale)
                            Dim y = CInt(GetFloat() * yScale)
                            StartPosns.Add(n, startposn)
                        Case Else
                            If debugflag Then TextBox1.AppendText($"Bad SEGMENT command length {nWords}")
                    End Select
                Case MOL_GOSUB ' GOSUB
                    decodeCmd(cmd_posn)
                Case Else
                    ' Unrecognised command
                    decodeCmd(cmd_posn)
            End Select
            My.Application.DoEvents()
            ' move position to next command
            reader.BaseStream.Seek(param_posn + nWords * 4, SeekOrigin.Begin)
            If debugflag Then TextBox1.AppendText($"{vbCrLf}")
        End While
        If debugflag Then TextBox1.AppendText($"{vbCrLf}")
    End Sub
    Sub decodeCmd(cmd_posn As Integer)
        ' Lookup command in known commands table
        ' Reader is positioned at parameters
        Dim value As LASERcmd = Nothing, cmd As Integer, cmd_len As Integer

        reader.BaseStream.Seek(cmd_posn, SeekOrigin.Begin)      ' find command
        cmd = GetInt()
        cmd_len = cmd >> 24 And &HFF
        If cmd_len = &H80 Then
            cmd_len = GetInt() And &HFF
        End If
        If cmd = MOL_MCB Then
            ' MCB special case. First parameter is length of whole MCB, but we want to decode it, so just treat length as a simple parameter
            reader.BaseStream.Seek(cmd_posn + 4, SeekOrigin.Begin)      ' back up one word to read the parameter count as data
        End If
        If LASERcmds.TryGetValue(cmd, value) Then
            If cmd_len <> value.parameters.Count And cmd <> MOL_MCB Then
                MsgBox($"Command {value.mnemonic}: length is {cmd_len}, but data table says {value.parameters.Count}", vbAbort + vbOKOnly, "Data error")
            End If
            TextBox1.AppendText($" {value.mnemonic}")
            For Each p In value.parameters
                TextBox1.AppendText($" {p.name}=")
                Select Case p.type
                    Case GetType(Boolean) : TextBox1.AppendText(CType(GetInt(), Boolean))
                    Case GetType(Int32) : TextBox1.AppendText(CType(GetInt(), Int32))
                    Case GetType(Single) : TextBox1.AppendText(CType(GetFloat(), Single))
                    Case GetType(OnOff_type) : If GetInt() Mod 2 = 0 Then TextBox1.AppendText("Off") Else TextBox1.AppendText("On")
                    Case Else
                        MsgBox($"{value.mnemonic}: Unrecognised parameter type of {p.type}", vbCritical + vbOKOnly, "data error")
                End Select
            Next
        Else
            ' UNKNOWN command. Just show parameters
            TextBox1.AppendText($"Unknown: Params {cmd_len}: ")
            For i = 1 To cmd_len
                TextBox1.AppendText($" {GetInt():x8}")
            Next
        End If
    End Sub
    Sub HexDumpBlock(addr As Integer)
        ' dump a block in hex
        Dim done As Boolean = False, nWords As UInteger
        reader.BaseStream.Seek(addr, SeekOrigin.Begin)    ' reposition to offset n
        While Not done And stream.Position < stream.Length
            TextBox1.AppendText($"@&h{stream.Position:x8}: ")
            Dim cmd As UInteger = GetUInt()
            TextBox1.AppendText($" &h{cmd:x8}")
            If cmd = 0 Then Exit While
            nWords = (cmd >> 24) And &HFF
            If nWords = &H80 Then
                nWords = GetUInt()
                TextBox1.AppendText($" &h{nWords:x8}")
            End If
            For i As Integer = 1 To nWords
                cmd = GetUInt()
                TextBox1.AppendText($" &h{cmd:x8}")
            Next
            ' Display mnemonic if known
            Dim value As LASERcmd = Nothing
            If LASERcmds.TryGetValue(cmd, value) Then TextBox1.AppendText($" ({value.mnemonic})")
            TextBox1.AppendText($"{vbCrLf}")
        End While
    End Sub

    Private Sub WordAccessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WordAccessToolStripMenuItem.Click
        Dim value As Integer, value1() As Byte
        Dim stream = File.Open(dlg.filename, FileMode.Open)
        Dim reader = New BinaryReader(stream, System.Text.Encoding.Unicode, False)
        value = reader.ReadInt32
        value = reader.ReadInt32
        value1 = reader.ReadBytes(4)
        reader.BaseStream.Seek(4, SeekOrigin.Begin)    ' reposition to offset 4
        value = reader.ReadInt32
    End Sub

    Private Sub HEXDUMPONToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HEXDUMPONToolStripMenuItem.Click
        hexdump = True
    End Sub

    Private Sub HEXDUMPOFFToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HEXDUMPOFFToolStripMenuItem.Click
        hexdump = False
    End Sub

    Private Sub DEBUGONToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DEBUGONToolStripMenuItem.Click
        debugflag = True
    End Sub

    Private Sub DEBUGOFFToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DEBUGOFFToolStripMenuItem.Click
        debugflag = False
    End Sub

    Private Sub MakeTestCardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeTestCardToolStripMenuItem.Click
        Dim buffer() As Byte
        Dim Outline = New Rect(-85, -100, 85, 100)     ' outline of test card
        dxf = New DxfDocument()     ' create empty DXF file
        Dim origin = New System.Windows.Point(-67, -80)   ' bottom right of test card
        Dim CellSize = New System.Windows.Size(5, 5)       ' size of test card cells
        Dim CellMargin = New Size(1.5, 1.5)                 ' margin around each cell
        Dim cellDimension = New Size(CellSize.Width + CellMargin.Width, CellSize.Height + CellMargin.Height)    ' Size of overall cell
        DrawText(dxf, My.Settings.Material, System.Windows.TextAlignment.Center, New System.Windows.Point(origin.X - origin.X / 2, -5), 5, 0)
        DrawText(dxf, $"Interval: {My.Settings.Interval} mm", System.Windows.TextAlignment.Center, New System.Windows.Point(origin.X - origin.X / 2, -9.5), 5, 0)
        DrawText(dxf, $"Passes: { My.Settings.Passes}", System.Windows.TextAlignment.Center, New System.Windows.Point(origin.X - origin.X / 2, -14), 5, 0)

        Dim rows(9) As Integer    ' labels for rows
        For i = 0 To UBound(rows)
            rows(i) = (My.Settings.SpeedMax - My.Settings.SpeedMin) / UBound(rows) * i + My.Settings.SpeedMin
        Next
        Dim cols(9) As Integer      ' labels for columns
        For i = 0 To UBound(cols)
            cols(i) = (My.Settings.PowerMax - My.Settings.PowerMin) / UBound(cols) * i + My.Settings.PowerMin
        Next
        ' Draw a grid of boxes
        For row = 0 To UBound(rows)
            For col = 0 To UBound(cols)
                Dim speed = rows(row)
                Dim power = cols(col)
                DrawBox(dxf, New System.Windows.Point(origin.X + row * cellDimension.Width, origin.Y + col * cellDimension.Height), CellSize.Width, CellSize.Height, My.Settings.Engrave, power, speed)
            Next
        Next
        ' Put labels on the axes
        For row = 0 To UBound(rows)
            DrawText(dxf, $"{rows(row)}", System.Windows.TextAlignment.Left, New System.Windows.Point(origin.X - 8, origin.Y + row * cellDimension.Height + 2), CellSize.Width, 0)
        Next
        For col = 0 To UBound(cols)
            DrawText(dxf, $"{cols(col)}", System.Windows.TextAlignment.Left, New System.Windows.Point(origin.X + col * cellDimension.Width + 4, origin.Y - 8), 6, 90)
        Next
        DrawText(dxf, "Power (%)", System.Windows.TextAlignment.Center, New System.Windows.Point(origin.X - origin.X / 2, -95), 10, 0)
        DrawText(dxf, "Speed (mm/s)", System.Windows.TextAlignment.Center, New System.Windows.Point(origin.X - 10, origin.Y - origin.Y / 2), 10, 90)
        DrawBox(dxf, Outline.TopLeft, Outline.Width, Outline.Height, False, My.Settings.PowerMax * 0.8, My.Settings.SpeedMax * 0.8)       ' the cut outline
        dxf.Save("Test card.dxf")

        Dim ReadStream = File.Open("FOUR.MOL", FileMode.Open)           ' file containing template blocks
        reader = New BinaryReader(ReadStream, System.Text.Encoding.Unicode, False)
        Dim WriteStream = File.Open("TESTCARD.MOL", FileMode.Create)    ' file containing test card
        writer = New BinaryWriter(WriteStream, System.Text.Encoding.Unicode, False)
        buffer = reader.ReadBytes(BLOCK_SIZE)      ' read header
        writer.Write(buffer)                    ' write to new file
        buffer = reader.ReadBytes(BLOCK_SIZE)      ' read config
        writer.Write(buffer)                    ' write to new file
        buffer = reader.ReadBytes(BLOCK_SIZE)      ' read TEST
        writer.Write(buffer)                    ' write to new file
        buffer = reader.ReadBytes(BLOCK_SIZE)      ' read CUT
        writer.Write(buffer)                    ' write to new file
        buffer = reader.ReadBytes(BLOCK_SIZE)      ' read DRAW
        writer.Write(buffer)                    ' write to new file
        reader.Close()


        ' Now update template values
        ' HEADER
        ' Bottom left
        writer.Seek(&H20, SeekOrigin.Begin)
        writer.Write(CInt(Outline.X * xScale))
        writer.Write(CInt(Outline.Y * yscale))

        'TEST
        Dim position = New Vector2(0, 0)    ' TEST block start
        writer.Seek(&H430, SeekOrigin.Begin)       ' start of move commands
        Dim moves As New List(Of Vector2)
        With Outline
            moves.Add(New Vector2(.Left, 0))
            moves.Add(New Vector2(0, .Top))
            moves.Add(New Vector2(- .Left, 0))
            moves.Add(New Vector2(0, - .Top))
        End With

        'CUT
        position = New Vector2(0, 0)    ' CUT block start
        writer.Seek(&H430, SeekOrigin.Begin)       ' start of move commands
        moves = New List(Of Vector2)
        With Outline
            moves.Add(New Vector2(.Left, 0))
            moves.Add(New Vector2(0, .Top))
            moves.Add(New Vector2(- .Left, 0))
            moves.Add(New Vector2(0, - .Top))
        End With

        For Each Mv In moves
            MOVERELSPLIT(writer, Mv)
        Next
        PutInt(MOL_ENDSUB) : PutInt(1)      ' end subroutine #1
        PutInt(MOL_END)                     ' end of block

        writer.Close()

        TextBox1.AppendText("Done")
    End Sub
    Sub MOVERELSPLIT(writer As BinaryWriter, p As Vector2)
        ' Create a move command from the current position to p, in 3 phases
        ' Phase 1 is slow, Phase 2 fast, phase 3 slow
        Const Acc = 0.1       ' percentage acceleration/deceleration time
        Const SPD_SLOW = 12.0    ' slow cutting speed
        Const SPD_FAST = 100.0   ' fast cutting speed

        ' now divide line into 3 pieces. Calculate 1st and last line, then calculate middle line from 1st and last.
        Dim break1 = New Vector2(position.X + p.X * Acc, position.Y + p.Y * Acc)
        Dim break2 = New Vector2(position.X + p.X * (1 - Acc), position.Y + p.Y * (1 - Acc))
        Dim delta1 = break1 - position
        Dim delta2 = break2 - break1
        Dim delta3 = delta1
        'Dim l1 = New Entities.Line(position, New Vector2(position.X + p.X * Acc, position.Y + p.Y * Acc))
        'Dim l2 As New Entities.Line(New Vector2(p.X * (1 - (1 - 2 * Acc)), p.Y * (1 - (1 - 2 * Acc))), p)
        'Dim l3 As New Entities.Line(New Vector2(p.X * (1 - Acc), p.Y * (1 - Acc)), p)
        MOVERELCMD(writer, delta1, SPD_SLOW)
        MOVERELCMD(writer, delta2, SPD_FAST)
        MOVERELCMD(writer, delta3, SPD_SLOW)
    End Sub
    Sub MOVERELCMD(writer As BinaryWriter, p As Vector2, speed As Single)
        ' send a MOVEREL command to MOL file
        PutInt(MOL_SETSPD) : PutFloat(speed) : PutInt(64032) : PutInt(64032)
        PutInt(MOL_MOVEREL) : PutInt(772) : PutInt(CInt(p.X * xScale)) : PutInt(CInt(p.Y * yscale))
        position += p
    End Sub
    Sub DrawBox(dxf As DxfDocument, origin As System.Windows.Point, width As Single, height As Single, shaded As Boolean, power As Integer, speed As Integer)
        ' Draw a box
        ' origin in mm
        ' width in mm
        ' height in mm
        ' power as %
        ' speed as mm/sec
        Dim motion As New Polyline2D With {
            .Layer = DrawLayer
        }
        TextBox1.AppendText($"Drawing box at ({origin.X},{origin.Y}) width {width} height {height} power {power} speed {speed}{vbCrLf}")
        My.Application.DoEvents()
        Dim rect As New System.Windows.Rect(0, 0, width, height)
        Dim mat As New Matrix(xScale, 0, 0, yscale, origin.X * xScale, origin.Y * yscale)
        rect.Transform(mat)
        ' create an hls color to represent power and speed as a shade of brown (=30 degrees)
        Dim shading As AciColor = PowerSpeedColor(power, speed)
        If Not shaded Then
            motion.Color = shading
            ' draw 4 lines surrounding the cell
            Dim x1 = rect.Left, x2 = rect.Right, y1 = rect.Bottom, y2 = rect.Top
            With motion.Vertexes
                .Add(New Polyline2DVertex(x1, y1))
                .Add(New Polyline2DVertex(x2, y1))
                .Add(New Polyline2DVertex(x2, y2))
                .Add(New Polyline2DVertex(x1, y2))
                .Add(New Polyline2DVertex(x1, y1))
            End With
            dxf.Entities.Add(motion)
        Else
            Dim position = New System.Windows.Point(rect.Left, rect.Top)     ' note Top <--> Bottom
            Dim RasterHeight = My.Settings.Interval * yscale   ' separation of raster lines
            Dim RasterWidth = rect.Width                     ' width of raster lines
            Dim line As Integer
            motion = New Polyline2D
            Dim direction As Integer
            ' Draw the first line
            motion.Vertexes.Add(New Polyline2DVertex(position.X, position.Y))
            position.X += RasterWidth
            motion.Vertexes.Add(New Polyline2DVertex(position.X, position.Y))
            line = 1
            While position.Y < rect.Bottom      ' note Top <--> Bottom
                ' Not finished yet, so add a line. Add vertical and horizontal lines
                motion.Vertexes.Add(New Polyline2DVertex(position.X, position.Y + RasterHeight))    ' add small vertical movement
                position.Y += RasterHeight
                If line Mod 2 = 0 Then direction = 1 Else direction = -1
                motion.Vertexes.Add(New Polyline2DVertex(position.X + RasterWidth * direction, position.Y))
                position.X += RasterWidth * direction
                line += 1
            End While
            motion.Color = shading
            motion.Layer = EngraveLayer
            dxf.Entities.Add(motion)
        End If
    End Sub
    Sub DrawText(dxf As DxfDocument, text As String, alignment As System.Windows.TextAlignment, origin As System.Windows.Point, fontsize As Single, angle As Single)
        ' Write some text at specified size, position and angle
        ' Initially this writes to a DXF file for testing, but will eventually write to a MOL file

        Dim font = New FontFamily("1CamBam_Stick_0")      ' pick a font
        Dim typeface = New Typeface(font, FontStyles.Normal, FontWeights.Medium, FontStretches.Normal)
        Dim emSize = fontsize * 96.0 / 72.0
        Const PixelsPerDip = 1
        Dim FontStyle = FontStyles.Normal
        Dim fontWeight = FontWeights.Light
        Dim position As System.Windows.Point
        Dim shading = PowerSpeedColor(My.Settings.PowerMax / 2, My.Settings.SpeedMax / 2)      ' use 50% power and speed for text
        ' Create the formatted text based on the properties set.
        Dim FormattedText As New FormattedText(text,
                    CultureInfo.GetCultureInfo("en-us"),
                    FlowDirection.LeftToRight,
                    typeface,
                    emSize,
                    Brushes.DimGray,
                    PixelsPerDip) With {
                        .TextAlignment = alignment
                    }
        ' Build the geometry object that represents the text, and extract vectors.
        Dim _textGeometry As Geometry = FormattedText.BuildGeometry(New System.Windows.Point(0, 0))
        ' massage the geometry to get it in the right place
        Dim transforms As New TransformGroup
        With transforms
            .Children.Add(New TranslateTransform(0, -_textGeometry.Bounds.Top)) ' move the text to the origin
            .Children.Add(New ScaleTransform(1, -1))   ' reflect in y axis
            .Children.Add(New TranslateTransform(0, _textGeometry.Bounds.Height))    ' move geometry back to bottom left
            .Children.Add(New RotateTransform(angle))                    ' rotate  
            Dim TextScale = 0.352778    ' conversion of points to mm
            .Children.Add(New ScaleTransform(TextScale, TextScale))       ' scale points to mm
            .Children.Add(New ScaleTransform(xScale, yscale))       ' scale mm to steps
            .Children.Add(New TranslateTransform(origin.X * xScale, origin.Y * yscale))    ' move geometry to origin
        End With
        _textGeometry.Transform = transforms

        ' Build the geometry object that represents the text, and extract vectors.
        Dim Flattened As PathGeometry = _textGeometry.GetFlattenedPathGeometry(fontsize / 8, ToleranceType.Relative)
        ' draw the text
        For Each figure As PathFigure In Flattened.Figures
            Dim motion = New Polyline2D With {.Layer = TextLayer, .Color = shading}
            motion.Vertexes.Add(New Polyline2DVertex(figure.StartPoint.X, figure.StartPoint.Y))
            position = figure.StartPoint
            For Each seg As PathSegment In figure.Segments
                Dim pnts = CType(seg, PolyLineSegment).Points
                For Each pnt In pnts
                    motion.Vertexes.Add(New Polyline2DVertex(pnt.X, pnt.Y))
                Next
            Next
            dxf.Entities.Add(motion)
        Next
    End Sub
    Private Sub TestIEEEToLeetroFpConversionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestIEEEToLeetroFpConversionToolStripMenuItem.Click
        Dim TestCases() = {208.3218, 800, 0, 1000.0, -1000.0, -1.0, 1.0, 0.1, -0.1}
        TextBox1.AppendText($"07505555 is {Leetro2Float(&H7505555)}{vbCrLf}")
        TextBox1.AppendText($"09480000 is {Leetro2Float(&H9480000)}{vbCrLf}")
        TextBox1.AppendText($"03000e46 is {Leetro2Float(&H3000E46)}{vbCrLf}")

        For Each t In TestCases
            TextBox1.AppendText($"Testcase {t}: ")
            Dim ToInt As UInteger = Float2Int(t)
            Dim FromInt As Single = Leetro2Float(ToInt)
            If t = FromInt Then TextBox1.AppendText("Success") Else TextBox1.AppendText($"Failure: result {FromInt}")
            TextBox1.AppendText($"{vbCrLf}")
        Next
    End Sub

    Private Sub TestCardParametersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestCardParametersToolStripMenuItem.Click
        TestPanel.ShowDialog()
    End Sub

    Private Sub TextTestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TextTestToolStripMenuItem.Click
        Dim motion As New Polyline2D
        Dim font = New FontFamily("1CamBam_Stick_0")      ' pick a font. This one from http://www.atelier-des-fougeres.fr/Cambam/Aide/Plugins/stickfonts.html
        'Dim font = New FontFamily("Arial")      ' pick a font. This one from http://www.atelier-des-fougeres.fr/Cambam/Aide/Plugins/stickfonts.html
        Dim typeface = New Typeface(font, FontStyles.Normal, FontWeights.Medium, FontStretches.Normal)
        Const fontSize = 8  ' points
        Const emSize = fontSize * 96.0 / 72.0
        Const PixelsPerDip = 1
        Dim FontStyle = FontStyles.Normal
        Dim fontWeight = FontWeights.Light
        Dim position As System.Windows.Point
        dxf = New DxfDocument()     ' create empty DXF file
        ' Create the formatted text based on the properties set.
        Dim FormattedText As New FormattedText("Test text",
                    CultureInfo.GetCultureInfo("en-us"),
                    FlowDirection.LeftToRight,
                    typeface,
                    emSize,
                    Brushes.DimGray,
                    PixelsPerDip) With {
                        .TextAlignment = System.Windows.TextAlignment.Left
                    }
        ' Build the geometry object that represents the text, and extract vectors.
        Dim _textGeometry As Geometry = FormattedText.BuildGeometry(New System.Windows.Point(0, 0))
        Dim transforms As New TransformGroup
        With transforms
            .Children.Add(New TranslateTransform(0, -_textGeometry.Bounds.Top)) ' move the text to the origin
            .Children.Add(New ScaleTransform(1, -1))   ' reflect in y axis
            .Children.Add(New TranslateTransform(0, _textGeometry.Bounds.Height))    ' move geometry back to bottom left
        End With
        _textGeometry.Transform = transforms
        Dim Flattened As PathGeometry = _textGeometry.GetFlattenedPathGeometry(fontSize / 8, ToleranceType.Relative)
        With motion.Vertexes
            .Add(New Polyline2DVertex(Flattened.Bounds.Left, Flattened.Bounds.Top))
            .Add(New Polyline2DVertex(Flattened.Bounds.Right, Flattened.Bounds.Top))
            .Add(New Polyline2DVertex(Flattened.Bounds.Right, Flattened.Bounds.Bottom))
            .Add(New Polyline2DVertex(Flattened.Bounds.Left, Flattened.Bounds.Bottom))
            .Add(New Polyline2DVertex(Flattened.Bounds.Left, Flattened.Bounds.Top))
        End With
        dxf.Entities.Add(motion)
        ' draw the text
        For Each figure As PathFigure In Flattened.Figures
            motion = New Polyline2D
            motion.Vertexes.Add(New Polyline2DVertex(figure.StartPoint.X, figure.StartPoint.Y))
            position = figure.StartPoint
            For Each seg As PathSegment In figure.Segments
                Dim pnts = CType(seg, PolyLineSegment).Points
                For Each pnt In pnts
                    motion.Vertexes.Add(New Polyline2DVertex(pnt.X, pnt.Y))
                Next
            Next
            dxf.Entities.Add(motion)
        Next
        dxf.Save("Text test.dxf")
        TextBox1.AppendText("Done")
    End Sub

    Private Sub DXFMetricsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DXFMetricsToolStripMenuItem.Click
        ' Calculate some metrics for a DXF file
        Dim layers As New List(Of String), travel As Single = 0, vertexes As Integer = 0
        TextBox1.AppendText($"{vbCrLf}DXF Metrics{vbCrLf}")
        With dxf
            TextBox1.AppendText($"Layers: {dxf.Layers.Count}")
            For Each l In .Layers
                layers.Add(l.Name)
            Next
            TextBox1.AppendText($" {String.Join(",", layers)}{vbCrLf}")
            For Each ln In .Entities.Lines
                travel += Distance(New Vector2(ln.StartPoint.X, ln.StartPoint.Y), New Vector2(ln.EndPoint.X, ln.EndPoint.Y))
                vertexes += 2
            Next
            TextBox1.AppendText($"Lines: { .Entities.Lines.Count}  Vertexes: {vertexes}{vbCrLf}")
            TextBox1.AppendText($"MLines: { .Entities.MLines.Count}{vbCrLf}")
            TextBox1.AppendText($"Points: { .Entities.Points.Count}{vbCrLf}")
            vertexes = 0
            For Each seg In .Entities.Polylines2D
                vertexes += seg.Vertexes.Count
                For i = 0 To seg.Vertexes.Count - 2
                    travel += Distance(seg.Vertexes(i).Position, seg.Vertexes(i + 1).Position)
                Next
            Next
            TextBox1.AppendText($"Polylines2D: { .Entities.Polylines2D.Count}  Vertexes: {vertexes}{vbCrLf}")
            TextBox1.AppendText($"Texts: { .Entities.Texts.Count}{vbCrLf}")
            TextBox1.AppendText($"Travel: { travel / xScale / 100:f1}m{vbCrLf}")
            ' Display command usage
            TextBox1.AppendText($"{vbCrLf}Command frequency{vbCrLf}")
            For Each cmd In CommandUsage
                Dim value As LASERcmd = Nothing, decode As String
                If LASERcmds.TryGetValue(cmd.Key, value) Then decode = value.mnemonic Else decode = $"0x{cmd.Key:x8}"
                TextBox1.AppendText($"{decode}  {cmd.Value}{vbCrLf}")
            Next
        End With
    End Sub
End Class
Class EngraveLine
        ' Captures the data for a single line of an engraving
        Property Start As Vector2   ' the start position of this line of engraving
        Property Segments As New List(Of EngraveSegment)    ' list of segments of the engrave

        Public Sub New(start As Vector2)
            Me.Start = start
            Me.Segments = New List(Of EngraveSegment)
        End Sub
    End Class
    Class EngraveSegment
        ' Captures the data for a segment of an engraving line. An EngraveLine is a contiguous list of EngraveSegments
        ' Segments are horizontal, so only X changes.
        Property Length As Integer       ' length of this segment (steps). Can be +ve or -ve
        Property Speed As Single        ' speed of laser movement (mm/sec)
        Property Power As Single        ' power of laser ( 0 - 100%)
        Property Laser As Boolean       ' laser On or Off
        ReadOnly Property Color As AciColor     ' the color to use to represent this speed/power
            Get
                Return PowerSpeedColor(Me.Power, Me.Speed)
            End Get
        End Property

    Public Sub New(speed As Single, power As Single, laser As Boolean, length As Single)
        Me.Speed = speed
        Me.Power = power
        Me.Laser = laser
        Me.Length = length
    End Sub
End Class
