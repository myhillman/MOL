Imports System.IO
Imports netDxf
Imports netDxf.Entities
Imports netDxf.Tables
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar
Imports System.Windows.Media
Imports System.Windows
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Net.WebRequestMethods

' Define the different block types
Public Enum BLOCK_TYPE
    HEADER = 0
    CONFIG = 1
    TEST = 2
    CUT = 3
    DRAW = 5
End Enum

' It is necessary to do two passes through the MOL file.
' The first is a simple start to finish pass that simply provides a disasemambly of everything it sees
' The second pass executes the code in the config section (only), and follows GOSUB calls as required.
Public Enum PASS_TYPE
    DISASSEMBLE = 1       ' instructions are decoded, but not executed
    EXECUTE = 2     ' instructions are not decoded, but executed
End Enum
Public Enum ParameterCount
    FIXED       ' the number of parameters for this command is fixed
    VARIABLE    ' the number of parameters for this command is variable
End Enum
' Define some custom types
Enum OnOff_type
    Off = 0
    _On = 1
End Enum
Enum Acceleration_type
    Accelerate = 1
    Decelerate = 2
End Enum

' A structure to hold 2 x 16 bit values. All 32 bits can be written through the Steps property. Individual 16 bit values can be read via OnSteps and OffSteps property
<StructLayout(LayoutKind.Explicit)> Public Structure OnOffSteps
    <FieldOffset(0)>
    Public Steps As Integer
    <FieldOffset(0)>
    Public OnSteps As Short        ' number of steps laser on
    <FieldOffset(2)>
    Public OffSteps As Short        ' number of steps laser off
End Structure

Public Class Form1
    Private Const BLOCK_SIZE = 512
    Private debugflag As Boolean = True
    Private hexdump As Boolean = False
    Private dlg = New OpenFileDialog
    Private stream As FileStream, reader As BinaryReader, writer As BinaryWriter
    Private TopRight As System.Windows.Point, BottomLeft As System.Windows.Point

    Private startposn As Vector2
    Private delta As Vector2       ' a delta from position
    Private Const NumColors = 255
    Private colors(NumColors) As AciColor   ' array of colors
    Private ColorIndex As Integer = 1       ' index into colors array
    Private CommandUsage As New Dictionary(Of Integer, Integer)
    Private dxf As New DxfDocument()
    Private StepsArray() As OnOffSteps              ' parameter to ENGLSR command
    Private Stack As New Stack()                    ' used by GOSUB to hold return address
    Private motion As New Polyline2D                ' list of all laser moves
    Private layer As Layer                          ' currenr drawing layer
    Private MotionControlBlock As New List(Of Integer)  ' buffer for commands which need to be in MCB
    Private UseMCB As Boolean = False                 ' if ture write to MCB, else write to file
    Private MCBCount As Integer = 0                   ' count of MCB generated
    Private Const MCBMax = 508                           ' maximum words in a MCB

    ' Variables that reflect the state of the laser cutter
    Private position As Vector2    ' current laser head position
    Private LaserIsOn As Boolean = False
    Public xScale As Single = 160.07    ' X axis steps/mm
    Public yScale As Single = 160.07    ' Y axis steps/mm
    Private AccelLength As Integer      ' # of steps for Acceleration and Deceleration
    Private CurrentBlockType As BLOCK_TYPE
    Private CurrentSubr As Integer = 0                ' current subroutine we are in
    Private StartPosns As New Dictionary(Of Integer, Vector2)      ' start position for drawing by subroutine #
    Private SubrAddrs As New Dictionary(Of Integer, Integer)       ' list of subroutine numbers and their start address
    Private ConfigChunk As Integer, TestChunk As Integer, CutChunk As Integer
    Private DrawChunks As New List(Of Integer)      ' list of draw chunks
    Private FirstStart As New Vector2(0, 0), EmptyVector As New Vector2(0, 0)
    Private ENGLSRsteps As New List(Of OnOffSteps)

    ' Define some layers for DXF file
    Private TextLayer As New Layer("Text") With {       ' text
                .Color = AciColor.Default
            }
    Private EngraveLayer As New Layer("Engrave") With { ' engraving
                .Color = AciColor.Cyan,
                .Lineweight = 10     ' 0.1 mm
            }
    Private DrawLayer = New Layer("Draw") With {        ' drawing
            .Color = AciColor.Red,
            .Linetype = Linetype.Continuous,
            .Lineweight = Lineweight.Default
        }
    Private ConfigLayer = New Layer("Config") With {   ' config
                .Color = AciColor.Magenta,
                .Linetype = Linetype.Continuous,
                .Lineweight = Lineweight.Default
        }
    Private TestLayer = New Layer("Test") With {    ' Test bounding box
            .Color = AciColor.Yellow,
            .Linetype = Linetype.DashDot,
            .Lineweight = Lineweight.Default
        }
    Private CutLayer = New Layer("Cut") With {      ' Cut bounding box
            .Color = AciColor.Red,
            .Linetype = Linetype.Dashed,
            .Lineweight = Lineweight.Default
        }
    Private MoveLayer = New Layer("Move") With {    ' head movement without laser on
            .Color = AciColor.Default,
            .Linetype = Linetype.Continuous,
            .Lineweight = Lineweight.Default,
            .IsVisible = False
        }
    Private EngraveCutLayer = New Layer("Engrave Cut") With {       ' Engrave move laser on
            .Lineweight = 10        ' 0.1 mm
        }
    Private EngraveMoveLayer = New Layer("Engrave Move") With {     ' Engrave laser off
            .Lineweight = Lineweight.Default
        }

    Class parameter
        ' a class for a parameter
        Property name As String     ' the name of the parameter
        Property typ As Type       ' the type of the parameter
        Sub New(name As String, type As Type)
            Me.name = name
            Me.typ = type
        End Sub

    End Class
    Class LASERcmd
        Property mnemonic As String
        Property parameterType As ParameterCount = ParameterCount.FIXED       ' VARIABLE OR FIXED
        Property parameters As New List(Of parameter)    ' the type of the parameter
        Sub New(mnemonic As String)
            Me.mnemonic = mnemonic
        End Sub

        Sub New(mnemonic As String, pc As ParameterCount)
            Me.mnemonic = mnemonic
            Me.parameterType = pc
        End Sub
        Sub New(mnemonic As String, pc As ParameterCount, p As List(Of parameter))
            Me.mnemonic = mnemonic
            Me.parameters = p
            Me.parameterType = pc
        End Sub
    End Class
    ' Definitions of full (param count + command) MOL file commands
    Const MOL_MOVEREL = &H3026000
    Const MOL_START = &H3026040
    Const MOL_ORIGIN = &H346040
    Const MOTION_CMD_COUNT = &H3090080
    Const MOL_BEGSUB = &H1300008
    Const MOL_BEGSUBa = &H1300048
    Const MOL_ENDSUB = &H1400048
    Const MOL_MCB = &H80000946
    Const MOL_SETSPD = &H3000301
    Const MOL_MOTION = &H3000341
    Const MOL_LASER = &H1000606
    Const MOL_GOSUB = &H1500048
    Const MOL_GOSUB3 = &H3500048
    Const MOL_X5_FIRST = &H200548
    Const MOL_X6_LAST = &H200648
    Const MOL_ACCELERATION = &H1004601
    Const MOL_BLOWER = &H1004A41
    Const MOL_BLOWERa = &H1004B41
    Const MOL_SEGMENT = &H500008
    Const MOL_GOSUBn = &H80500048
    Const MOL_ENGACD = &H1000346
    Const MOL_ENGMVY = &H2010040
    Const MOL_ENGMVX = &H2014040
    Const MOL_SCALE = &H3000E46
    Const MOL_PWRSPD5 = &H5000E46
    Const MOL_PWRSPD7 = &H7000E46
    Const MOL_ENGLSR = &H80000146
    Const MOL_END = 0

    ' Dictionary of all commands
    Private LASERcmds As New SortedDictionary(Of Integer, LASERcmd) From {
        {MOL_MOVEREL, New LASERcmd("MOVEREL", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))},
                            {New parameter("dx", GetType(Int32))},
                            {New parameter("dy", GetType(Int32))}
                            }
                           )},
        {MOL_START, New LASERcmd("START", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))},
                            {New parameter("x", GetType(Int32))},
                            {New parameter("y", GetType(Int32))}
                            }
                           )},
        {MOL_SCALE, New LASERcmd("SCALE", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("x scale", GetType(Single))},
                            {New parameter("y scale", GetType(Single))},
                            {New parameter("z scale", GetType(Single))}
                            }
                           )},
        {MOL_PWRSPD5, New LASERcmd("PWRSPD5", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("Corner PWR % x100", GetType(Int32))},
                            {New parameter("Max PWR % x100", GetType(Int32))},
                            {New parameter("Start cutter speed/speedmult", GetType(Single))},
                            {New parameter("Start cutter speed/speedmult", GetType(Single))},
                            {New parameter("Start cutter speed/speedmult", GetType(Single))}
                            }
                           )},
        {MOL_PWRSPD7, New LASERcmd("PWRSPD7", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("Corner PWR % x100", GetType(Int32))},
                            {New parameter("Max PWR % x100", GetType(Int32))},
                            {New parameter("Start cutter speed/speedmult", GetType(Single))},
                            {New parameter("Start cutter speed/speedmult", GetType(Single))},
                            {New parameter("Start cutter speed/speedmult", GetType(Single))},
                            {New parameter("Laser2 Corner Power", GetType(Int32))},
                            {New parameter("Laser2 Max Power", GetType(Int32))}
                            }
                           )},
        {MOL_ORIGIN, New LASERcmd("ORIGIN", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))},
                            {New parameter("x", GetType(Int32))},
                            {New parameter("y", GetType(Int32))}
                            }
                           )},
        {MOTION_CMD_COUNT, New LASERcmd("MOTION_COMMAND_COUNT", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("count", GetType(Int32))}})},
        {MOL_BEGSUB, New LASERcmd("BEGSUB", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("n", GetType(Int32))}})},
        {MOL_BEGSUBa, New LASERcmd("BEGSUBa", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("n", GetType(Int32))}})},
        {MOL_ENDSUB, New LASERcmd("ENDSUB", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("n", GetType(Int32))}})},
        {MOL_MCB, New LASERcmd("MCB", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("Size", GetType(Int32))}})},
        {MOL_MOTION, New LASERcmd("MOTION", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("Initial speed", GetType(Single))},
                            {New parameter("Max speed", GetType(Single))},
                            {New parameter("Acceleration", GetType(Single))}
                            }
                           )},
        {MOL_SETSPD, New LASERcmd("SETSPD", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("Initial speed", GetType(Single))},
                            {New parameter("Max speed", GetType(Single))},
                            {New parameter("Acceleration", GetType(Single))}
                            }
                           )},
        {MOL_LASER, New LASERcmd("LASER", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("On/Off", GetType(OnOff_type))}})},
        {MOL_GOSUB, New LASERcmd("GOSUB", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))}
                            }
                           )},
        {MOL_GOSUB3, New LASERcmd("GOSUB3", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))},
                            {New parameter("x", GetType(Single))},
                            {New parameter("y", GetType(Single))}
                            }
                           )},
        {MOL_X5_FIRST, New LASERcmd("X5_FIRST", ParameterCount.FIXED)},
        {MOL_X6_LAST, New LASERcmd("X6_LAST", ParameterCount.FIXED)},
        {MOL_SEGMENT, New LASERcmd("SEGMENT", ParameterCount.FIXED, New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))}
                            }
                           )},
        {MOL_GOSUBn, New LASERcmd("GOSUBn", ParameterCount.VARIABLE, New List(Of parameter) From {
                            {New parameter("n", GetType(Int32))},
                            {New parameter("x", GetType(Single))},
                            {New parameter("y", GetType(Single))}
                            }
                           )},
        {MOL_ACCELERATION, New LASERcmd("ACCELERATION", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("Acceleration", GetType(Acceleration_type))}})},
        {MOL_BLOWER, New LASERcmd("BLOWER", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("On/Off", GetType(OnOff_type))}})},
        {MOL_BLOWERa, New LASERcmd("BLOWERa", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("On/Off", GetType(OnOff_type))}})},
        {MOL_ENGMVX, New LASERcmd("ENGMVX", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("Axis", GetType(Integer))}, {New parameter("n", GetType(Integer))}})},
        {MOL_ENGMVY, New LASERcmd("ENGMVY", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("dy", GetType(Integer))}, {New parameter("??", GetType(Integer))}})},
        {MOL_ENGACD, New LASERcmd("ENGACD", ParameterCount.FIXED, New List(Of parameter) From {{New parameter("x", GetType(Int32))}})},
        {MOL_ENGLSR, New LASERcmd("ENGLSR", ParameterCount.VARIABLE, New List(Of parameter) From {
                            {New parameter("List of steps", GetType(OnOffSteps))}
                            }
                           )},
        {MOL_END, New LASERcmd("END")}
        }

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
        ' read 4 byte float from current offset
        Return Leetro2Float(GetUInt())
    End Function

    Public Function GetFloat(n As Integer) As Single
        ' read float from specified offset
        reader.BaseStream.Seek(n, SeekOrigin.Begin)    ' reposition to offset n
        Return GetFloat()
    End Function

    Sub PutFloat(f As Single)
        ' write float at current offset
        If UseMCB Then
            MotionControlBlock.Add(Float2Int(f))
            If MotionControlBlock.Count >= MCBMax Then FlushMCB()      ' MCB limited in size
        Else
            writer.Write(Float2Int(f))
        End If
    End Sub

    Sub PutInt(n As Integer)
        ' write n at current offset
        If UseMCB Then
            MotionControlBlock.Add(n)
            If MotionControlBlock.Count >= MCBMax Then FlushMCB()      ' MCB limited in size
        Else
            writer.Write(n)
        End If
    End Sub
    Sub PutInt(n As Integer, addr As Integer)
        ' write n at specifed offset
        If UseMCB Then
            MsgBox($"You can't write explicitly to address {addr:x} as an MCB in in operation", vbAbort + vbOKOnly, "Explicit write when MCB on")
        Else
            writer.BaseStream.Seek(addr, 0)     ' go to offset
            writer.Write(n)                     ' write data
        End If
    End Sub

    Sub DisplayHeader(textfile As TextWriter)
        ' Display header info
        Dim chunk As Integer

        CurrentBlockType = BLOCK_TYPE.HEADER
        reader.BaseStream.Position = 0
        textfile.WriteLine($"HEADER &0")
        textfile.WriteLine($"Size of file: &h{GetInt():x}")
        Dim MotionBlocks = GetInt()
        textfile.WriteLine($"Motion blocks: {MotionBlocks}")
        Dim v() As Byte = reader.ReadBytes(4)
        textfile.WriteLine($"Version: {CInt(v(3))}.{CInt(v(2))}.{CInt(v(1))}.{CInt(v(0))}")
        TopRight = New System.Windows.Point(GetInt(&H18), GetInt())
        textfile.WriteLine($"Origin: {TopRight.X},{TopRight.Y}")
        BottomLeft = New System.Windows.Point(GetInt(&H20), GetInt())
        textfile.WriteLine($"Bottom Left: ({BottomLeft.X},{BottomLeft.Y})")

        ConfigChunk = GetInt(&H70)
        TestChunk = GetInt(&H74)
        CutChunk = GetInt(&H78)
        chunk = GetInt(&H7C)        ' first draw chunk
        While chunk <> 0
            DrawChunks.Add(chunk)
            chunk = GetInt()
        End While
        textfile.WriteLine($"Config chunk: {ConfigChunk}")
        textfile.WriteLine($"Test chunk: {TestChunk}")
        textfile.WriteLine($"Cut chunk: {CutChunk}")
        textfile.WriteLine($"Draw chunks: {String.Join(",", DrawChunks.ToArray)}")
        textfile.WriteLine()
    End Sub

    Sub DecodeStream(writer As System.IO.StreamWriter, block As BLOCK_TYPE, StartAddress As Integer)
        ' Decode a stream of commands in either DISASSEMBLE or EXECUTE mode
        ' StartAddress - start of commands
        ' Set start position for this block
        Dim subr As Integer
        Select Case block       ' convert block to subroutine
            Case BLOCK_TYPE.TEST : subr = 1
            Case BLOCK_TYPE.CUT : subr = 2
            Case Else
                subr = 0
        End Select
        If StartPosns.ContainsKey(subr) Then
            position = StartPosns(subr)
        Else position = New Vector2(0, 0)
        End If
        writer.WriteLine()
        writer.WriteLine($"Decode of commands at 0x{StartAddress:x} in block {block}")
        writer.WriteLine()
        ' position = StartPosns(block)
        reader.BaseStream.Seek(StartAddress, 0)      ' Start of block
        Do
        Loop Until Not decodeCmd(writer)
    End Sub

    Function decodeCmd(writer As System.IO.StreamWriter) As Boolean
        ' Display a decoded version of the current command
        ' Lookup command in known commands table
        ' returns false if at end of stream

        Dim value As LASERcmd = Nothing, cmd As Integer, cmd_len As Integer

        Dim cmdBegin = reader.BaseStream.Position         ' remember start of command
        writer.Write($"{cmdBegin:x}: ")
        cmd = GetInt()                  ' get command
        If Not CommandUsage.TryAdd(cmd, 1) Then CommandUsage(cmd) += 1      ' count commands used
        If cmd = MOL_END Then
            writer.WriteLine("END")
            Return False
        End If      ' We are done
        cmd_len = cmd >> 24 And &HFF
        If cmd_len = &H80 Then
            cmd_len = GetInt() And &HFF
        End If
        Select Case cmd
            Case MOL_MCB
                cmd_len = 1        ' allow contents of MCB to decode
                reader.BaseStream.Position = cmdBegin + 4     ' backup so we can read length as parameter
            Case MOL_BEGSUB, MOL_BEGSUBa
                ' Harvest subroutine start addresses during decode
                Dim n = GetUInt()
                Try
                    SubrAddrs.Add(n, reader.BaseStream.Position - 8)      ' remember address of this subroutine
                Catch ex As Exception
                    MsgBox(ex.Message, vbAbort + vbOKOnly, $"trying to add address for subr {n} at {reader.BaseStream.Position - 8}")
                End Try
                reader.BaseStream.Position -= 4          ' backup over parameter
        End Select

        If LASERcmds.TryGetValue(cmd, value) Then
            If value.parameterType = ParameterCount.FIXED And cmd_len <> value.parameters.Count And cmd <> MOL_MCB Then
                MsgBox($"Command {value.mnemonic}: length is {cmd_len}, but data table says {value.parameters.Count}", vbAbort + vbOKOnly, "Data error")
            End If
            writer.Write($" {value.mnemonic}")
            For Each p In value.parameters
                writer.Write($" {p.name}=")
                Select Case p.typ
                    Case GetType(Boolean) : writer.Write(CType(GetInt(), Boolean))
                    Case GetType(Int32) : writer.Write(CType(GetInt(), Int32))
                    Case GetType(Single) : writer.Write(CType(GetFloat(), Single))
                    Case GetType(OnOff_type) : Dim par = GetInt() : If par Mod 2 = 0 Then writer.Write($"Off({par:x})") Else writer.Write($"On({par:x})")
                    Case GetType(Acceleration_type) : writer.Write($"{CType(GetInt(), Acceleration_type)}")
                    Case GetType(OnOffSteps)    ' a list of On/Off steps
                        writer.Write($"List of {cmd_len} On/Off steps ")
                        Dim OneStep As OnOffSteps
                        For i = 1 To cmd_len    ' one structure for each word
                            OneStep.Steps = GetUInt()     ' get 32 bit word
                            writer.Write($" {OneStep.OnSteps}/{OneStep.OffSteps}")
                        Next
                    Case Else
                        MsgBox($"{value.mnemonic}: Unrecognised parameter type of {p.typ}", vbCritical + vbOKOnly, "data error")
                End Select
            Next
        Else
            ' UNKNOWN command. Just show parameters
            writer.Write($"Unknown: {cmd:x8} Params {cmd_len}: ")
            For i = 1 To cmd_len
                writer.Write($" {GetInt():x8}")
            Next
        End If
        writer.WriteLine()
        'reader.BaseStream.Seek(cmdBegin + cmd_len * 4 + 4, 0)       ' move to next command
        Return True         ' more commands follow
    End Function

    Sub ExecuteStream(block As BLOCK_TYPE, StartAddress As Integer, dxf As DxfDocument, DefaultLayer As Layer)
        ' Execute a stream of commands, rendering in dxf as we go
        layer = DefaultLayer
        Dim AddrTrace As New List(Of Integer)
        ' Set start position for this block
        Dim subr As Integer
        If StartPosns.ContainsKey(subr) Then
            position = StartPosns(subr)
        Else position = New Vector2(0, 0)
        End If
        LaserIsOn = False
        Dim addr = StartAddress
        Do
            AddrTrace.Add(addr)
            addr = ExecuteCmd(block, addr, dxf, layer)
        Loop Until addr = 0
    End Sub
    Function ExecuteCmd(block As BLOCK_TYPE, ByVal addr As Integer, dxf As DxfDocument, Baselayer As Layer) As Integer
        ' Execute a command at addr. Return addr as start of next instruction

        Dim cmd As Integer, nWords As Integer

        Dim spdMin As Single = 100.0
        Dim spdMax As Single = 100.0
        Dim i As UInteger
        layer = Baselayer
        motion.Layer = layer      ' set layer of motion block
        cmd = GetInt(addr)      ' get the command

        Dim command = cmd And &HFFFFFF    ' bottom 24 bits is command
        nWords = (cmd >> 24) And &HFF  ' command length is top 8 bits
        If nWords = &H80 Then nWords = GetInt() And &HFFFF
        If cmd = MOL_MCB Then nWords = 0    ' allow contents of MCB to execute
        Dim param_posn = stream.Position  ' remember where parameters (if any) start

        Select Case cmd
            Case MOL_END
                If motion.Vertexes.Count > 0 Then
                    ' add the motion to the dxf file    
                    dxf_polyline2d(motion, layer, 1 / xScale)
                    motion.Vertexes.Clear()
                End If
                Return 0
            Case MOL_MCB
                nWords = 0      ' allow MCB content to be executed

            Case MOL_LASER      ' switch laser
                ' Laser is changing state. Output any pending motion
                If motion.Vertexes.Count > 0 Then
                    ' Display any motion
                    dxf_polyline2d(motion, layer, 1 / xScale)
                    motion.Vertexes.Clear()
                End If
                i = GetUInt()
                Select Case i
                    Case 0 : LaserIsOn = False
                        motion.Layer = MoveLayer
                    Case 1 : LaserIsOn = True
                        motion.Layer = Baselayer
                    Case Else
                        If debugflag Then TextBox1.AppendText($"Unknown LaserIsOn parameter &h{i:x8}")
                End Select

            Case MOL_SCALE      ' also MOL_SPDPWR, MOL_SPDPWRx
                xScale = GetFloat() : yScale = GetFloat() : GetFloat()        ' x,y,z scale command

            Case MOL_MOVEREL
                Dim n = GetUInt()
                Dim delta = New Vector2(GetInt(), GetInt())      ' move relative command
                If motion.Vertexes.Count = 0 Then motion.Vertexes.Add(New Polyline2DVertex(position))   ' first move - add start point
                position += delta
                motion.Vertexes.Add(New Polyline2DVertex(position))

            Case MOL_START
                Dim n = GetUInt()
                If motion.Vertexes.Count > 0 Then
                    dxf.Entities.Add(motion)
                    motion.Vertexes.Clear()
                End If
                'If motion.Vertexes.Count = 0 Then   ' first move - add start point
                Dim posn = New Vector2(GetInt(), GetInt())      ' start position
                If FirstStart = EmptyVector Then
                    ' First move in start is absolute position
                    FirstStart = posn
                    position = posn
                Else
                    ' show a move to a new position
                    DXF_line(position, position + posn, MoveLayer, 1 / xScale)
                End If
                If n = 772 Then

                End If

            Case MOL_ENGMVY     ' Engrave move Y (laser off)
                Dim axis = GetInt()      ' consume Axis parameter
                Dim dy = GetInt()
                ' Move in X direction (laser on/off)
                Dim move As New Vector2(0, dy)    ' move in Y direction
                DXF_line(position, position + move, MoveLayer, 1 / xScale)

            Case MOL_ENGACD
                AccelLength = GetInt()

            Case MOL_ENGMVX   ' Engrave move X  (laser on and off)
                GetInt()      ' consume Axis parameter
                Dim TotalSteps = GetInt()
                Dim startposn = position
                Dim direction = Math.Sign(TotalSteps)     ' 1=LtoR, -1=RtoL
                Dim move As Vector2
                ' Move for the initial acceleration
                move = New Vector2(AccelLength * direction, 0)
                DXF_line(position, position + move, MoveLayer, 1 / xScale)
                ' do On/Off steps
                For Each stp In ENGLSRsteps
                    ' The On portion of the move
                    move = New Vector2(stp.OnSteps * direction, 0)    ' move in X direction
                    DXF_line(position, position + move, EngraveLayer, 1 / xScale)
                    ' The Off portion of the move
                    move = New Vector2(stp.OffSteps * direction, 0)    ' move in X direction
                    DXF_line(position, position + move, MoveLayer, 1 / xScale)
                Next
                ' Move for the deceleration
                move = New Vector2(AccelLength * direction, 0)
                DXF_line(position, position + move, MoveLayer, 1 / xScale)

            Case MOL_ENGLSR       ' engraving movement
                ENGLSRsteps.Clear()
                For i = 1 To nWords
                    Dim Steps As OnOffSteps
                    Steps.Steps = GetInt()      ' get 2 16 bit values, accessable through OnSteps & OffSteps
                    ENGLSRsteps.Add(Steps)
                Next

            Case MOL_BEGSUB, MOL_BEGSUBa  ' begin subroutine
                Dim n = GetUInt()
                If n >= 3 Then
                    ' Create layer for this subroutine, but not config test or cut
                    layer = New Layer($"SUBR {n}") With
                    {
                        .Color = colors(ColorIndex),
                        .Linetype = Linetype.Continuous,
                        .Lineweight = Lineweight.Default
                    }
                    motion.Layer = layer
                    ColorIndex += 20
                End If

            Case MOL_ENDSUB  ' end subroutine
                Dim n = GetUInt()
                If motion.Vertexes.Count > 0 Then
                    ' add the motion to the dxf file    
                    dxf_polyline2d(motion, layer, 1 / xScale)
                    motion.Vertexes.Clear()
                End If
                If Stack.Count > 0 Then
                    reader.BaseStream.Seek(Stack.Pop, 0)     ' return to the saved return address
                    Return reader.BaseStream.Position
                Else
                    MsgBox($"Stack is exhausted - no return address for ENDSUB {n}", vbAbort + vbOKOnly, "Stack exhausted")
                End If

            Case MOL_GOSUBn       ' Call subroutine with parameters
                Dim n = GetInt()     ' get subroutine number
                Dim delta = New Vector2(GetFloat(), GetFloat())     ' subroutine parameter
                Dim startSub = FirstStart + delta
                DXF_line(position, startSub, MoveLayer, 1 / xScale)
                If n < 100 Then
                    TextBox1.AppendText($"Calling subr {n} with position ({delta.X},{delta.Y}){vbCrLf}")
                    Stack.Push(reader.BaseStream.Position)              ' address of next instruction
                    Return SubrAddrs(n)                     ' jump to start address and continue
                End If

            Case MOL_GOSUB ' GOSUB
                Dim n As Integer = GetInt()
                Select Case nWords
                    Case 1
                    Case 3
                        Dim x = GetInt()
                        Dim y = GetInt()
                End Select

            Case Else
                ' consume all parameters
                For i = 1 To nWords
                    GetInt()    ' consume parameters
                Next
        End Select
        Return param_posn + nWords * 4       ' address of next instruction
    End Function
    Sub DXF_line(startpoint As Vector2, endpoint As Vector2, layer As Layer, Optional ByVal scale As Single = 1.0)
        ' Add line to dxf file specified by startpoint, endpoint and layer. position is updated.
        Dim line As New Line(startpoint, endpoint) With {
            .Layer = layer
        }
        If scale <> 1.0 Then
            Dim transform As New Matrix3(scale, 0, 0, 0, scale, 0, 0, 0, scale)
            line.TransformBy(transform, New Vector3(0, 0, 0))
        End If
        dxf.Entities.Add(line)
        position = endpoint
    End Sub
    Sub dxf_polyline2d(Polyline As Polyline2D, Layer As Layer, Optional ByVal scale As Single = 1.0)
        ' Add a polyline to the specified layer
        Dim ply As Polyline2D = Polyline.Clone
        ply.Layer = Layer
        If scale <> 1.0 Then
            Dim transform As New Matrix3(scale, 0, 0, 0, scale, 0, 0, 0, scale)
            ply.TransformBy(transform, New Vector3(0, 0, 0))
        End If
        dxf.Entities.Add(ply)
        position = Polyline.Vertexes.Last.Position      ' update position to end of polyline
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
        Dim stream = System.IO.File.Open(dlg.filename, FileMode.Open)
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

        writer = New BinaryWriter(System.IO.File.Open("TESTCARD.MOL", FileMode.Create), System.Text.Encoding.Unicode, False)
        Dim Outline = New Rect(-85, -100, 85, 100)     ' outline of test card in mm
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
        Const ENGACD = 1763       ' Acceleration/Deceleration distance
        ' All boxes are the same except for power/speed setting. On/Off same for all
        Dim Steps As OnOffSteps             ' we only need one step
        Steps.OnSteps = CInt(CellSize.Width * xScale)        ' on for full cell width
        Steps.OffSteps = 0             ' go staright to decelerate
        Dim OnOffTotal = ENGACD * 2 + Steps.OnSteps + Steps.OffSteps
        Dim YStep = CInt(0.1 * yScale) '  Y increment for each line
        WriteMOL(MOL_ENGACD, {ENGACD}, &HA18)     ' define the acceleration start distance
        writer.BaseStream.Position = &HAA8          ' start of ENGLSR section

        ' Now create the engraved box for each setting
        For row = 0 To UBound(rows)
            For col = 0 To UBound(cols)
                Dim speed = rows(row)
                Dim power = cols(col)
                DrawBox(dxf, New System.Windows.Point(origin.X + row * cellDimension.Width, origin.Y + col * cellDimension.Height), CellSize.Width, CellSize.Height, My.Settings.Engrave, power, speed)
                ' Now create equivalent MOL
                Dim direction As Integer = 1
                ' goto origin. Don't know how
                Dim height As Single = 0      ' height of engraving so far
                While height < CellSize.Height      ' do until we reach cell height
                    WriteMOL(MOL_ENGLSR, {1, New List(Of OnOffSteps) From {{Steps}}})      ' one full line of on/off
                    WriteMOL(MOL_ENGMVX, {4, OnOffTotal})
                    WriteMOL(MOL_ENGMVY, {YStep, 63})
                    OnOffTotal *= -1           ' reverse direction for next line
                    height += YStep
                End While
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

        Dim ReadStream = System.IO.File.Open("FOUR.MOL", FileMode.Open)           ' file containing template blocks
        reader = New BinaryReader(ReadStream, System.Text.Encoding.Unicode, False)
        writer.Seek(0, SeekOrigin.Begin)            ' set reader to start of file
        buffer = reader.ReadBytes(BLOCK_SIZE)      ' read header
        writer.Seek(0, SeekOrigin.Begin)            ' set writer to start of file
        writer.Write(buffer)                    ' write to new file
        buffer = reader.ReadBytes(BLOCK_SIZE)      ' read config
        writer.Write(buffer)                    ' write to new file
        buffer = reader.ReadBytes(BLOCK_SIZE)      ' read TEST
        writer.Write(buffer)                    ' write to new file
        buffer = reader.ReadBytes(BLOCK_SIZE)      ' read CUT
        writer.Write(buffer)                    ' write to new file
        buffer = reader.ReadBytes(BLOCK_SIZE)      ' read empty block
        writer.Write(buffer)                    ' write to new file
        reader.Close()

        'TEST
        ' Moves the head around the boundary of the outline
        WriteMOL(MOL_BEGSUBa, {1}, &H400)    ' begin SUB 1
        WriteMOL(&H80600148, {6, &H25B, &H91D4000, &HF13A300, &H1113A340, &H101D7A80, &HF716C00})   ' Unknown command
        UseMCB = True     ' start buffer command output
        Dim moves As New List(Of Vector2)
        With Outline
            moves.Add(New Vector2(.Left, 0))
            moves.Add(New Vector2(0, .Top))
            moves.Add(New Vector2(- .Left, 0))
            moves.Add(New Vector2(0, - .Top))
        End With
        For Each Mv In moves
            ' do moves to next point in halves
            MOVERELSPLIT(Mv, 2)
        Next
        FlushMCB()
        WriteMOL(MOL_ENDSUB, {1})
        WriteMOL(MOL_END)

        'CUT
        ' Cut a line around a 2mm margin around the boundary 
        Dim CutLine As New Rect()   ' grow the outline to be the cutline
        CutLine = Outline
        CutLine.Inflate(New Size(2, 2))        ' Inflate cutline by 2mm all round
        WriteMOL(MOL_BEGSUBa, {2}, &H600)    ' begin SUB 2
        WriteMOL(&H80600148, {6, &H25B, &H91D4000, &HF13A300, &H1113A340, &H101D7A80, &HF716C00})   ' Unknown command
        WriteMOL(MOL_PWRSPD5, {4000, 4000, 629.0F, 3149.0F, 0F})    ' set power & speed
        UseMCB = True         ' start buffering commands
        moves.Clear()

        With CutLine
            moves.Add(New Vector2(.Left, 0))
            moves.Add(New Vector2(0, .Top))
            moves.Add(New Vector2(- .Left, 0))
            moves.Add(New Vector2(0, - .Top))
        End With

        ' We are at (0,0). Move to start of CUT line
        MOVERELSPLIT(New Vector2(CutLine.Right, CutLine.Bottom), 2)

        WriteMOL(MOL_LASER, {OnOff_type._On})
        For Each Mv In moves
            ' do moves to next point in 3 pieces
            MOVERELSPLIT(Mv, 3)
        Next
        WriteMOL(MOL_LASER, {OnOff_type.Off})
        FlushMCB()
        WriteMOL(MOL_ENDSUB, {2})      ' end subroutine #2
        WriteMOL(MOL_END)              ' end of block

        ' Create engraving as SUBR 3 and Text as SUBR 4
        ' Engraving layer
        Dim StartofBlock = 5 * BLOCK_SIZE   ' address of start of next block. First subr at A00
        Dim BlockNumber = 5
        PutInt(BlockNumber, &H7C)
        writer.BaseStream.Position = StartofBlock       ' position at start of block
        WriteMOL(MOL_BEGSUBa, {3}, StartofBlock)    ' begin SUB 
        FlushMCB()
        WriteMOL(MOL_ENDSUB, {3})    ' end SUB 
        WriteMOL(MOL_END)

        ' Text layer
        StartofBlock = (writer.BaseStream.Position \ BLOCK_SIZE + 1) * BLOCK_SIZE   ' address of start of next block
        BlockNumber = StartofBlock \ BLOCK_SIZE
        PutInt(BlockNumber, &H7C + 4)
        WriteMOL(MOL_BEGSUBa, {4}, StartofBlock)    ' begin SUB 
        FlushMCB()
        WriteMOL(MOL_ENDSUB, {4})    ' end SUB 
        WriteMOL(MOL_END)

        ' make binary file a multiple of BLOCK_SIZE
        Dim LastBlock = CInt(writer.BaseStream.Length \ BLOCK_SIZE)       '  current length of file in blocks
        writer.BaseStream.SetLength((LastBlock + 1) * BLOCK_SIZE)       ' round up to next block
        ' Now update template values
        ' HEADER
        writer.BaseStream.Seek(0, SeekOrigin.Begin)
        PutInt(writer.BaseStream.Length)            ' file size
        PutInt(MCBCount + 30)                         ' number of MCB
        ' Bottom left
        writer.Seek(&H20, SeekOrigin.Begin)
        writer.Write(CInt(Outline.X * xScale))
        writer.Write(CInt(Outline.Y * yScale))
        writer.Close()
        TextBox1.AppendText("Done")
    End Sub
    Sub WriteMOL(command As Integer, Optional ByVal Parameters() As Object = Nothing, Optional posn As Integer = -1)
        ' Write a MOL command, with parameters to MOL file
        ' Parameters will be written (but not length) if present
        ' writing occurs at current writer position, or "posn" if present

        ' Check the correct number of parameters
        If LASERcmds.ContainsKey(command) Then
            ' It is a known command, therefore do some checks
            Dim nWords = LASERcmds(command).parameters.Count
            If (Parameters Is Nothing) Then
                If nWords <> 0 Then
                    MsgBox($"WRITEMOL {LASERcmds(command).mnemonic}: parameter error. Parameter is Nothing, yet command requires {nWords}", vbAbort + vbOKOnly, "WRITEMOL parameter error")
                End If
            Else
                If LASERcmds(command).parameterType = ParameterCount.FIXED And Parameters.Length <> nWords Then
                    MsgBox($"WRITEMOL {LASERcmds(command).mnemonic}: parameter error. There are {Parameters.Length} parameters, yet command requires {nWords}", vbAbort + vbOKOnly, "WRITEMOL parameter error")
                End If
            End If

            ' Check correct type of parameters
            If nWords > 0 And LASERcmds(command).parameterType = ParameterCount.FIXED Then
                For i = 0 To nWords - 1
                    Dim a = LASERcmds(command).parameters(i).typ
                    Dim b = Parameters(i).GetType
                    If a <> b Then
                        MsgBox($"WRITEMOL {LASERcmds(command).mnemonic}: parameter type mismatch - spec is {a}, call is {b}", vbAbort + vbOKOnly, "Parameter type mismatch")
                    End If
                Next
            End If
        End If
        If posn >= 0 Then   ' explicit position
            writer.BaseStream.Position = posn       ' position the writer
        End If
        PutInt(command)
        If Parameters IsNot Nothing Then
            For Each p In Parameters
                Select Case p.GetType
                    Case GetType(Int32) : PutInt(p)
                    Case GetType(Single) : PutFloat(p)
                    Case GetType(System.Collections.Generic.List(Of OnOffSteps))
                        For Each s In p
                            PutInt(s.steps)
                        Next
                    Case GetType(Acceleration_type)
                        PutInt(p)
                    Case GetType(OnOff_type)
                        PutInt(p)
                    Case Else
                        MsgBox($"WRITEMOL Parameter of unsupported type {p.GetType}", vbAbort + vbOKOnly, "Unsupported type")
                End Select
            Next
        End If
    End Sub

    Sub FlushMCB()
        ' Flush the contents of the MotionControlBlock
        UseMCB = False    ' turn off MCB buffer
        If MotionControlBlock.Count > 0 Then
            WriteMOL(MOL_MCB, {MotionControlBlock.Count})   ' write MCB command and count
        ' Write the contents of the MCB
        For Each item In MotionControlBlock
            PutInt(item)
        Next
        MotionControlBlock.Clear()          ' clear the buffer
            MCBCount += 1                        ' count MCB
        End If
    End Sub
    Sub MOVERELSPLIT(p As Vector2, pieces As Integer)
        ' Create a move command from the current position to p
        ' Pieces = 2 or 3
        ' Accelerate before first one
        ' Decelerate before last one
        ' There may be a middle piece
        ' Phase 1 is slow, Phase 2 fast, phase 3 slow
        Const Acc = 0.1       ' percentage acceleration/deceleration time
        Dim moves As New List(Of System.Windows.Point)
        If pieces = 2 Then
            ' move in 2 equal pieces
            moves.Add(New System.Windows.Point(p.X / 2, p.Y / 2))
            moves.Add(New System.Windows.Point(p.X / 2, p.Y / 2))
            ' Move in short, long, short pieces
        Else
            Dim delta1 = New System.Windows.Point(p.X * Acc, p.Y * Acc)
            Dim delta2 = New System.Windows.Point(p.X * (1 - Acc), p.Y * (1 - Acc))
            moves.Add(delta1)
            moves.Add(delta2)
            moves.Add(delta1)
        End If

        Const SPD_SLOW = 12.0F    ' slow cutting speed
        Const SPD_FAST = 100.0F  ' fast cutting speed

        For Each m In moves
            If m = moves.First Then
                WriteMOL(MOL_ACCELERATION, {Acceleration_type.Accelerate})
                WriteMOL(MOL_SETSPD, {SPD_SLOW, 36518.0F, 151176.0F})
            ElseIf m = moves.Last Then
                WriteMOL(MOL_ACCELERATION, {Acceleration_type.Decelerate})
                WriteMOL(MOL_SETSPD, {SPD_SLOW, 36518.0F, 151176.0F})
            Else
                WriteMOL(MOL_SETSPD, {SPD_FAST, 36518.0F, 151176.0F})
            End If
            WriteMOL(MOL_MOVEREL, {772, CInt(m.X * xScale), CInt(m.Y * yScale)})
        Next
    End Sub
    Sub DrawBox(dxf As DxfDocument, origin As System.Windows.Point, width As Single, height As Single, shaded As Boolean, power As Integer, speed As Integer)
        ' Draw a box
        ' origin in mm
        ' width in mm
        ' height in mm
        ' power as %
        ' speed as mm/sec
        Dim motion As New Polyline2D
        TextBox1.AppendText($"Drawing box at ({origin.X},{origin.Y}) width {width} height {height} power {power} speed {speed}{vbCrLf}")
        My.Application.DoEvents()
        Dim rect As New System.Windows.Rect(0, 0, width, height)
        'Dim mat As New Matrix(xScale, 0, 0, yScale, origin.X * xScale, origin.Y * yScale)
        'rect.Transform(mat)
        rect.Offset(origin.X, origin.Y)
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
            dxf_polyline2d(motion, DrawLayer)
        Else
            Dim position = New System.Windows.Point(rect.Left, rect.Top)     ' note Top <--> Bottom
            Dim RasterHeight = My.Settings.Interval    ' separation of raster lines
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
            dxf_polyline2d(motion, EngraveLayer)
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
            '.Children.Add(New ScaleTransform(xScale, yScale))       ' scale mm to steps
            .Children.Add(New TranslateTransform(origin.X, origin.Y))    ' move geometry to origin
        End With
        _textGeometry.Transform = transforms

        ' Build the geometry object that represents the text, and extract vectors.
        Dim Flattened As PathGeometry = _textGeometry.GetFlattenedPathGeometry(fontsize / 8, ToleranceType.Relative)
        ' draw the text
        For Each figure As PathFigure In Flattened.Figures
            Dim motion = New Polyline2D With {.Color = shading}
            motion.Vertexes.Add(New Polyline2DVertex(figure.StartPoint.X, figure.StartPoint.Y))
            position = figure.StartPoint
            For Each seg As PathSegment In figure.Segments
                Dim pnts = CType(seg, PolyLineSegment).Points
                For Each pnt In pnts
                    motion.Vertexes.Add(New Polyline2DVertex(pnt.X, pnt.Y))
                Next
            Next
            dxf_polyline2d(motion, TextLayer)
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
        dxf_polyline2d(motion, TextLayer)
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
            dxf_polyline2d(motion, TextLayer)
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
            ' MOL file data
            TextBox1.AppendText($"{vbCrLf}{vbCrLf}MOL file details{vbCrLf}{vbCrLf}")
            TextBox1.AppendText($"List of SubrAddrs{vbCrLf}")
            For Each s In SubrAddrs
                TextBox1.AppendText($"{s.Key}: {s.Value:x}{vbCrLf}")
            Next

        End With
    End Sub

    Private Sub ReconstructMINIMARIOMOLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReconstructMINIMARIOMOLToolStripMenuItem.Click
        ' Reconstruct the MINIMARIO.MOL file from the hex dump on the London Hackspace page
        ' Retrieved from https://wiki.london.hackspace.org.uk/view/Project:RELaserSoftware/MOL_file_format

        Dim bytes(&H1000 - 1) As Byte
        Using reader As New Microsoft.VisualBasic.FileIO.TextFieldParser("minimario.txt")
            reader.SetDelimiters(" ")
            Dim currentRow As String()
            While Not reader.EndOfData
                Try
                    currentRow = reader.ReadFields()
                    If currentRow.Length = 17 Then
                        ' extract the address
                        Dim addr As Integer
                        If Integer.TryParse(currentRow(0), Globalization.NumberStyles.HexNumber, Globalization.CultureInfo.InvariantCulture, addr) Then
                        Else
                            Throw New ArgumentException("Argument is invalid")
                        End If
                        ' extract row of bytes
                        For i = 1 To 16
                            Dim data As Byte
                            If Byte.TryParse(currentRow(i), Globalization.NumberStyles.HexNumber, Globalization.CultureInfo.InvariantCulture, data) Then
                                bytes(addr + i - 1) = data
                            Else
                                Throw New ArgumentException("Argument is invalid")
                            End If
                        Next
                    End If
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                End Try
            End While
        End Using
        ' now dump to binary file
        My.Computer.FileSystem.WriteAllBytes("minimario.MOL", bytes, False)
    End Sub

    Private Sub CommandSummaryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CommandSummaryToolStripMenuItem.Click
        ' produce a summary of all command
        TextBox1.AppendText($"Command summary{vbCrLf}{vbCrLf}")
        TextBox1.AppendText($"{"HEX value",-12}{"Mnemonic",-15}  Parameters (name is type){vbCrLf}")
        ' Sort dictionary by low 24 bits
        Dim sorted = From item In LASERcmds
                     Order By item.Key And &HFFFFFF
                     Select item
        For Each cmd In sorted
            TextBox1.AppendText($"0x{cmd.Key:x8}  {cmd.Value.mnemonic,-15}  ")
            'now parameters
            Dim parameters = New List(Of String)
            For Each p In cmd.Value.parameters
                parameters.Add($"{p.name} is {p.typ.Name}")
            Next
            TextBox1.AppendText(String.Join(", ", parameters.ToArray))
            TextBox1.AppendText($"{vbCrLf}")
        Next
        ' Custom types. Display details of any parameters that are defined using ENUM. These are used when a basetype won't do
        Dim CustomTypes As New List(Of Type)
        For Each cmd In LASERcmds       ' search all commands
            For Each p In cmd.Value.parameters  ' searcg all parameters
                If p.typ.BaseType.Name = "Enum" Then    ' Add any non standand types to list
                    If Not CustomTypes.Contains(p.typ) Then CustomTypes.Add(p.typ)
                End If
            Next
        Next
        ' Display list of custom types
        TextBox1.AppendText($"{vbCrLf}Custom types{vbCrLf}")
        For Each ct In CustomTypes
            Dim enums As New List(Of String)
            For Each value In System.Enum.GetValues(ct)     ' extract name and value of type
                Dim name = value.ToString
                Dim valu = CInt(value).ToString
                enums.Add($"{name}={valu}")
            Next
            TextBox1.AppendText($"{ct.Name,-25} {String.Join(",", enums)}{vbCrLf}")
        Next
    End Sub

    Private Sub DisassembleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisassembleToolStripMenuItem.Click
        With dlg
            .Filter = "MOL file|*.mol|All files|*.*"
        End With
        Dim result = dlg.ShowDialog()
        If result.OK Then
            ' Init internal variables

            DrawChunks.Clear()
            ' Create initial start positions
            StartPosns.Clear()
            SubrAddrs.Clear()
            ' Subr 1 & 2 are TEST & CUT
            For n = 1 To 2
                StartPosns.Add(n, New Vector2(0, 0))
            Next
            position = EmptyVector
            FirstStart = EmptyVector
            TextBox1.Clear()
            stream = System.IO.File.Open(dlg.filename, FileMode.Open)
            reader = New BinaryReader(stream, System.Text.Encoding.Unicode, False)

            Dim Size As Long = reader.BaseStream.Length / 4     ' size of the inputfile in 4 byte words
            ' Open a text file to write to
            Dim TextFile As System.IO.StreamWriter
            Dim Basename = System.IO.Path.GetFileNameWithoutExtension(dlg.filename)
            Dim textname = $"{Basename}.txt"
            TextFile = My.Computer.FileSystem.OpenTextFileWriter(textname, False)
            ' DISASSEMBLE everything
            DisplayHeader(TextFile)
            DecodeStream(TextFile, BLOCK_TYPE.CONFIG, BLOCK_TYPE.CONFIG * BLOCK_SIZE)
            DecodeStream(TextFile, BLOCK_TYPE.TEST, BLOCK_TYPE.TEST * BLOCK_SIZE)
            DecodeStream(TextFile, BLOCK_TYPE.CUT, BLOCK_TYPE.CUT * BLOCK_SIZE)
            For Each chunk In DrawChunks
                DecodeStream(TextFile, BLOCK_TYPE.DRAW, chunk * BLOCK_SIZE)
            Next
            ' Display command usage
            TextFile.WriteLine()
            TextFile.WriteLine("Command frequency")
            TextFile.WriteLine()
            For Each cmd In CommandUsage
                Dim value As LASERcmd = Nothing, decode As String
                If LASERcmds.TryGetValue(cmd.Key, value) Then decode = value.mnemonic Else decode = $"0x{cmd.Key:x8}"
                TextFile.WriteLine($"{decode}  {cmd.Value}")
            Next

            TextFile.Close()
            reader.Close()
            stream.Close()
            ' Open text file
            Dim myProcess As New Process
            With myProcess
                .StartInfo.FileName = textname
                .StartInfo.UseShellExecute = True
                .StartInfo.RedirectStandardOutput = False
                .Start()
                .Dispose()
            End With
        End If
    End Sub

    Private Sub RenderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenderToolStripMenuItem.Click
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

            DrawChunks.Clear()
            ' Create initial start positions
            StartPosns.Clear()
            SubrAddrs.Clear()
            ' Subr 1 & 2 are TEST & CUT
            For n = 1 To 2
                StartPosns.Add(n, New Vector2(0, 0))
            Next
            position = EmptyVector
            FirstStart = EmptyVector
            TextBox1.Clear()
            stream = System.IO.File.Open(dlg.filename, FileMode.Open)
            reader = New BinaryReader(stream, System.Text.Encoding.Unicode, False)

            Dim Size As Long = reader.BaseStream.Length / 4     ' size of the inputfile in 4 byte words
            Dim MOLfile(Size) As Int32                          ' buffer for whole file
            ' Read the whole file
            'Dim i = 0
            'While reader.BaseStream.Position < reader.BaseStream.Length
            '    MOLfile(i) = reader.ReadInt32
            '    i += 1
            'End While

            ' Now execute real code
            dxf = New DxfDocument()         ' render commands in dxf
            ExecuteStream(BLOCK_TYPE.CONFIG, ConfigChunk * BLOCK_SIZE, dxf, ConfigLayer)   ' config
            Stack.Push(0)   ' any old rubbish to keep ENDSUB happy
            ExecuteStream(BLOCK_TYPE.TEST, TestChunk * BLOCK_SIZE, dxf, TestLayer)   ' test
            Stack.Push(0)   ' any old rubbish to keep ENDSUB happy
            ExecuteStream(BLOCK_TYPE.CUT, CutChunk * BLOCK_SIZE, dxf, CutLayer)   ' cut
            'For Each chunk In DrawChunks
            '    Stack.Push(0)   ' any old rubbish to keep ENDSUB happy
            '    ExecuteStream(BLOCK_TYPE.DRAW, chunk * BLOCK_SIZE, dxf, DrawLayer)   ' draw
            'Next

            ' Add the filename at the bottom of the image
            Dim BaseDXF = System.IO.Path.GetFileNameWithoutExtension(dlg.filename)
            Dim DXFfile = $"{BaseDXF}_MOL.dxf"
            Dim rect = New System.Windows.Rect(TopRight, BottomLeft)    ' bounding rectangle
            rect.Scale(1 / xScale, 1 / yScale)                      ' convert to mm
            entity = New Text($"File: {DXFfile}") With {
                .Position = New Vector3(rect.X + rect.Width / 2, rect.Y - rect.Height / 10, 0),
                .Alignment = Entities.TextAlignment.TopCenter,
                .Layer = TextLayer,
                .Height = Math.Max(rect.Width, rect.Height) / 20
                }
            dxf.Entities.Add(entity)
            ' Save dxf in file
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
