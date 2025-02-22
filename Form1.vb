Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Documents
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar
Imports System.Windows.Media
Imports netDxf
Imports netDxf.Entities
Imports netDxf.Tables
Imports Windows.Win32.System
Imports YamlDotNet.Serialization
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.Design.AxImporter
Imports System.Text
Imports MOL.Form1
Imports System.DirectoryServices.ActiveDirectory
Imports System.Text.Unicode
Imports System.Windows.Ink
Imports System.Windows.Media.Imaging
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip

' Define the different block types
Public Enum BLOCKTYPE
    HEADER = 0
    CONFIG = 1
    TEST = 2
    CUT = 3
    DRAW = 5
End Enum

' It is necessary to do two passes through the MOL file.
' The first is a simple start to finish pass that simply provides a disasemambly of everything it sees
' The second pass executes the code in the config section (only), and follows GOSUB calls as required.
Public Enum ParameterCount
    FIXED       ' the number of parameters for this command is fixed
    VARIABLE    ' the number of parameters for this command is variable
End Enum
' Define some custom types
Friend Enum OnOff_type
    Off = 0
    [On] = 1
End Enum

Friend Enum Axis_type
    X = 4
    Y = 3
End Enum

Friend Enum Acceleration_type
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
    Private Const BLOCK_SIZE = 512          ' MOL is divided into blocks or chunks
    Private debugflag As Boolean = True
    Private hexdump As Boolean = False
    Private dlg = New OpenFileDialog
    Private stream As FileStream, reader As BinaryReader, writer As BinaryWriter
    Private TopRight As IntPoint, BottomLeft As IntPoint

    Private startposn As IntPoint
    Private delta As IntPoint       ' a delta from position
    Private ZERO = New IntPoint(0, 0)   ' a point with value (0,0)
    Private Const NumColors = 255
    Private colors(NumColors) As AciColor   ' array of colors
    Private ColorIndex As Integer = 1       ' index into colors array
    Private CommandUsage As New Dictionary(Of Integer, Integer)
    Private dxf As New DxfDocument()
    Private StepsArray() As OnOffSteps              ' parameter to ENGLSR command
    Private Stack As New Stack()                    ' used by GOSUB to hold return address
    Private motion As New Polyline2D                ' list of all laser moves
    Private layer As Layer                          ' current drawing layer
    Private MCBLK As New List(Of Integer)           ' buffer for commands which need to be in MCBLK
    Private UseMCBLK As Boolean = False                 ' if true write to MCBLK, else write to file
    Private MCBLKCount As Integer = 0                   ' count of MCBLK generated
    Private MVRELCnt As Integer = 0                    ' Count of MVREL + &h264 + &h284
    Private Const MCBLKMax = 508                           ' maximum words in a MCBLK
    Private ChunksWithCode As New Dictionary(Of Integer, Boolean)      ' list of chunks known to contain code. Boolean is true if has been executed
    Private FontData As New Dictionary(Of Integer, Glyph)       ' stroke fonts


    ' Variables that reflect the state of the laser cutter
    ' These parameters from Machine Options - Worktable dialog
    Private StartSpeed As Double = 5        ' Initial speed when moving
    Private QuickSpeed As Double = 300      ' Max speed when moving
    Private WorkAcc As Double = 500         ' Acceleration whilst cutting
    Private SpaceAcc As Double = 1200       ' Acceleration whilst just moving

    Private position As IntPoint    ' current laser head position in mm
    Private LaserIsOn As Boolean = False        ' tracks state of laser
    Public xScale As Double = 125.9842519685039    ' X axis steps/mm.   From the Worktable config dialog [Pulse Unit] parameter
    Public yScale As Double = 125.9842519685039    ' Y axis steps/mm
    Friend ScaleToSteps = New Matrix(xScale, 0.0, 0.0, yScale, 0.0, 0.0)       ' matrix to scale mm to steps
    Private AccelLength As Double      ' Acceleration and Deceleration distance
    Private CurrentBlockType As BLOCKTYPE
    Private CurrentSubr As Integer = 0                ' current subroutine we are in
    Private StartPosns As New Dictionary(Of Integer, (Absolute As Boolean, position As IntPoint))      ' start position for drawing by subroutine #
    Private SubrAddrs As New Dictionary(Of Integer, Integer)       ' list of subroutine numbers and their start address
    Private ConfigChunk As Integer, TestChunk As Integer, CutChunk As Integer
    Private DrawChunks As New List(Of Integer)      ' list of draw chunks
    Private EmptyVector As New Vector2(0, 0)
    Private ENGLSRsteps As New List(Of OnOffSteps)
    Private FirstMove As Boolean = True           ' true if next MVREL will be the first
    Private EngPower As Integer             ' Engrave power (%)
    Private EngSpeed As Double              ' Engrave speed (mm/s)
    Private MVRELInConfig As Integer      ' number of MVREL encountered so far in CONFIG section

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

    Public Class Parameter
        ' a class for a parameter
        Public Property Name As String     ' the name of the parameter
        Public Property Description As String
        Public Property Typ As Type       ' the type of the parameter
        Public Property Scale As Double    ' scale applied to value
        Public Property Units As String    ' the units of the value, e.g. mm


        Public Sub New(name As String, description As String, type As Type, Optional ByVal scale As Double = 1, Optional ByVal units As String = "")
            Me.Name = name
            Me.Description = description
            Me.Typ = type
            Me.Scale = scale
            Me.Units = units
        End Sub

    End Class

    Public Class LASERcmd
        Public Property Mnemonic As String
        Public Property Description As String
        Public Property ParameterType As ParameterCount = ParameterCount.FIXED       ' VARIABLE OR FIXED
        Public Property Parameters As New List(Of Parameter)    ' the type of the parameter

        Public Sub New(mnemonic As String, description As String)
            Me.Mnemonic = mnemonic
            Me.Description = description
        End Sub

        Public Sub New(Mnemonic As String, description As String, pc As ParameterCount)
            Me.Mnemonic = Mnemonic
            Me.ParameterType = pc
            Me.Description = description
        End Sub

        Public Sub New(Mnemonic As String, description As String, pc As ParameterCount, p As List(Of Parameter))
            Me.Mnemonic = Mnemonic
            Me.Parameters = p
            Me.ParameterType = pc
            Me.Description = description
        End Sub
    End Class
    ' Definitions of full (param count + command) MOL file commands
    Private Const MOL_MVREL = &H3026000
    Private Const MOL_START = &H3026040
    Private Const MOL_ORIGIN = &H346040
    Private Const MOTION_CMD_COUNT = &H3090080
    Private Const MOL_BEGSUB = &H1300008
    Private Const MOL_BEGSUBa = &H1300048
    Private Const MOL_ENDSUB = &H1400048
    Private Const MOL_MCBLK = &H80000946
    Private Const MOL_SETSPD = &H3000301
    Private Const MOL_MOTION = &H3000341
    Private Const MOL_LASER = &H1000606
    Private Const MOL_LASER1 = &H1000646
    Private Const MOL_LASER2 = &H1000806
    Private Const MOL_LASER3 = &H1000846
    Private Const MOL_GOSUB = &H1500048
    Private Const MOL_GOSUB3 = &H3500048
    Private Const MOL_X5_FIRST = &H200548
    Private Const MOL_X6_LAST = &H200648
    Private Const MOL_ACCELERATION = &H1004601
    Private Const MOL_BLWR = &H1004A41
    Private Const MOL_BLWRa = &H1004B41
    Private Const MOL_SEGMENT = &H500008
    Private Const MOL_GOSUBn = &H80500048
    Private Const MOL_ENGPWR = &H1000746
    Private Const MOL_ENGPWR1 = &H2000746
    Private Const MOL_ENGSPD = &H2014341
    Private Const MOL_ENGSPD1 = &H4010141
    Private Const MOL_ENGACD = &H1000346
    Private Const MOL_ENGMVY = &H2010040
    Private Const MOL_ENGMVX = &H2014040
    Private Const MOL_SCALE = &H3000E46
    Private Const MOL_PWRSPD5 = &H5000E46
    Private Const MOL_PWRSPD7 = &H7000E46
    Private Const MOL_ENGLSR = &H80000146
    Private Const MOL_END = 0

    Private Const MOL_UNKNOWN07 = &H3046040      ' unknown command refered to in London Hackspace documents
    Private Const MOL_UNKNOWN09 = &H326040      ' unknown command refered to in London Hackspace documents

    ' Dictionary of all commands
    Private Commands As New SortedDictionary(Of Integer, LASERcmd) From {
        {MOL_MVREL, New LASERcmd("MVREL", "Move the cutter by dx,dy", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "Equivalent to 0x0304, i.e. both x and y", GetType(Int32))},
                            {New Parameter("dx", "delta to move in X direction", GetType(Int32), 1 / xScale, "mm")},
                            {New Parameter("dy", "delta to move in Y direction", GetType(Int32), 1 / yScale, "mm")}
                            }
                           )},
        {MOL_START, New LASERcmd("START", "A start location for a following subroutine", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "Equivalent to 0x0304, i.e. both x and y", GetType(Int32))},
                            {New Parameter("x", "", GetType(Int32), 1 / xScale, "mm")},
                            {New Parameter("y", "", GetType(Int32), 1 / yScale, "mm")}
                            }
                           )},
        {MOL_SCALE, New LASERcmd("SCALE", "The scale used on all 3 axis", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("x scale", "the X scale", GetType(Double),, "steps/mm")},
                            {New Parameter("y scale", "the Y scale", GetType(Double),, "steps/mm")},
                            {New Parameter("z scale", "the Z scale", GetType(Double),, "steps/mm")}
                            }
                           )},
        {MOL_PWRSPD5, New LASERcmd("PWRSPD5", "Not yet fully understood", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("Corner PWR", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Max PWR", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Start cutter speed/speedmult", "", GetType(Double))},
                            {New Parameter("Start cutter speed/speedmult", "", GetType(Double))},
                            {New Parameter("Start cutter speed/speedmult", "", GetType(Double))}
                            }
                           )},
        {MOL_PWRSPD7, New LASERcmd("PWRSPD7", "Not yet fully understood", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("Corner PWR", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Max PWR", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Start cutter speed/speedmult", "", GetType(Double))},
                            {New Parameter("Start cutter speed/speedmult", "", GetType(Double))},
                            {New Parameter("Start cutter speed/speedmult", "", GetType(Double))},
                            {New Parameter("Laser2 Corner Power", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Laser2 Max Power", "", GetType(Int32), 0.01, "%")}
                            }
                           )},
        {MOL_ORIGIN, New LASERcmd("ORIGIN", "", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "", GetType(Int32))},
                            {New Parameter("x", "", GetType(Int32))},
                            {New Parameter("y", "", GetType(Int32))}
                            }
                           )},
        {MOTION_CMD_COUNT, New LASERcmd("MOTION_COMMAND_COUNT", "", ParameterCount.FIXED, New List(Of Parameter) From {
                        {New Parameter("count", "", GetType(Int32))}}
                        )},
        {MOL_BEGSUB, New LASERcmd("BEGSUB", "Begin of subroutine", ParameterCount.FIXED, New List(Of Parameter) From {
                        {New Parameter("n", "Subroutine number", GetType(Int32))}}
                        )},
        {MOL_BEGSUBa, New LASERcmd("BEGSUBa", "Begin of subroutine", ParameterCount.FIXED, New List(Of Parameter) From {
                        {New Parameter("n", "Subroutine number", GetType(Int32))}}
                        )},
        {MOL_ENDSUB, New LASERcmd("ENDSUB", "End of subroutine", ParameterCount.FIXED, New List(Of Parameter) From {
                        {New Parameter("n", "Subroutine number", GetType(Int32))}}
                        )},
        {MOL_MCBLK, New LASERcmd("MCBLK", "Motion Control Block", ParameterCount.FIXED, New List(Of Parameter) From {
                        {New Parameter("Size", "Number of words (limited to 510)", GetType(Int32),, "Words")}}
                        )},
        {MOL_MOTION, New LASERcmd("MOTION", "Set min/max speed & acceleration", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("Initial speed", "", GetType(Double), 1 / xScale, "mm/s")},
                            {New Parameter("Max speed", "", GetType(Double), 1 / xScale, "mm/s")},
                            {New Parameter("Acceleration", "", GetType(Double), 1 / xScale, "mm/s²")}
                            }
                           )},
        {MOL_SETSPD, New LASERcmd("SETSPD", "Set min/max speed & acceleration", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("Initial speed", "", GetType(Double), 1 / xScale, "m/s")},
                            {New Parameter("Max speed", "", GetType(Double), 1 / xScale, "m/s")},
                            {New Parameter("Acceleration", "", GetType(Double), 1 / xScale, "m/s²")}
                            }
                           )},
        {MOL_LASER, New LASERcmd("LASER", "Laser on or off control", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("On/Off", "On or Off", GetType(OnOff_type))}})},
        {MOL_LASER1, New LASERcmd("LASER1", "", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("On/Off", "", GetType(Int32))}})},
        {MOL_LASER2, New LASERcmd("LASER2", "", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("On/Off", "", GetType(Int32))}})},
        {MOL_LASER3, New LASERcmd("LASER3", "", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("On/Off", "", GetType(Int32))}})},
        {MOL_GOSUB, New LASERcmd("GOSUB", "Call a subroutine", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "Number of the subroutine", GetType(Int32))}
                            }
                           )},
        {MOL_GOSUB3, New LASERcmd("GOSUB3", "", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "", GetType(Int32))},
                            {New Parameter("x", "", GetType(Double))},
                            {New Parameter("y", "", GetType(Double))}
                            }
                           )},
        {MOL_X5_FIRST, New LASERcmd("X5_FIRST", "", ParameterCount.FIXED)},
        {MOL_X6_LAST, New LASERcmd("X6_LAST", "", ParameterCount.FIXED)},
        {MOL_SEGMENT, New LASERcmd("SEGMENT", "", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "", GetType(Int32))}
                            }
                           )},
        {MOL_GOSUBn, New LASERcmd("GOSUBn", "", ParameterCount.VARIABLE, New List(Of Parameter) From {
                            {New Parameter("n", "", GetType(Int32))},
                            {New Parameter("x", "", GetType(Double))},
                            {New Parameter("y", "", GetType(Double))}
                            }
                           )},
        {MOL_ACCELERATION, New LASERcmd("ACCELERATION", "Accelerate or Deccelerate", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("Acceleration", "", GetType(Acceleration_type))}})},
        {MOL_BLWR, New LASERcmd("BLWR", "Blower control", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("On/Off", "", GetType(OnOff_type))}})},
        {MOL_BLWRa, New LASERcmd("BLWRa", "Blower control", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("On/Off", "", GetType(OnOff_type))}})},
        {MOL_ENGPWR, New LASERcmd("ENGPWR", "Engrave Power", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("Engrave power", "", GetType(Integer), 0.01, "%")}})},
        {MOL_ENGPWR1, New LASERcmd("ENGPWR1", "", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("Power", "", GetType(Integer), 0.01, "%")},
                {New Parameter("??", "", GetType(Integer))}
                }
              )},
        {MOL_ENGSPD, New LASERcmd("ENGSPD", "Engrave speed", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("Axis", "", GetType(Axis_type))},
                {New Parameter("Speed", "", GetType(Double), 1 / xScale, "mm/s")}
                }
              )},
        {MOL_ENGSPD1, New LASERcmd("ENGSPD1", "", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("Axis", "", GetType(Axis_type))},
                {New Parameter("??", "", GetType(Double))},
                {New Parameter("Speed", "", GetType(Double), 1 / xScale, "mm/s")},
                {New Parameter("??", "", GetType(Double))}
                }
              )},
        {MOL_ENGMVX, New LASERcmd("ENGMVX", "Engrave one line in the X direction", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("Axis", "", GetType(Axis_type))}, {New Parameter("dx", "", GetType(Integer), 1 / yScale, "mm")}})},
                {MOL_ENGMVY, New LASERcmd("ENGMVY", "Engraving - Move in the Y direction", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("Axis", "", GetType(Axis_type))},
                {New Parameter("dy", "Distance to move", GetType(Integer), 1 / yScale, "mm")}}
        )},
        {MOL_ENGACD, New LASERcmd("ENGACD", "Engraving (Ac)(De)cceleration distance", ParameterCount.FIXED, New List(Of Parameter) From {
                    {New Parameter("x", "distance", GetType(Int32), 1 / xScale, "mm")}}
                    )},
        {MOL_ENGLSR, New LASERcmd("ENGLSR", "Engraving On/Off pattern", ParameterCount.VARIABLE, New List(Of Parameter) From {
                            {New Parameter("List of steps", "List of On/Off patterns", GetType(OnOffSteps), 1 / xScale, "mm")}
                            }
                           )},
        {MOL_END, New LASERcmd("END", "End of code")}
        }

    Public Function GetInt() As Integer
        ' read 4 byte integer from input stream
        Return reader.ReadInt32
    End Function

    Public Function GetUInt() As UInteger
        ' read 4 byte uinteger from input stream
        Return reader.ReadUInt32
    End Function

    Public Function GetInt(n As Integer) As Integer
        reader.BaseStream.Seek(n, SeekOrigin.Begin)    ' reposition to offset n
        Return GetInt()
    End Function

    Public Function GetFloat() As Double
        ' read 4 byte float from current offset
        Return Leetro2Float(GetUInt())
    End Function

    Public Function GetFloat(n As Integer) As Double
        ' read float from specified offset
        reader.BaseStream.Seek(n, SeekOrigin.Begin)    ' reposition to offset n
        Return GetFloat()
    End Function

    Public Sub PutFloat(f As Double)
        ' write float at current offset
        If UseMCBLK Then
            MCBLK.Add(Float2Leetro(f))
        Else
            writer.Write(Float2Leetro(f))
        End If
    End Sub

    Public Sub PutInt(n As Integer)
        ' write n at current offset
        If UseMCBLK Then
            MCBLK.Add(n)
        Else
            writer.Write(n)
        End If
    End Sub

    Public Sub PutInt(n As Integer, addr As Integer)
        ' write n at specifed offset
        If UseMCBLK Then
            Throw New System.Exception($"You can't write explicitly to address {addr:x} as an MCBLK in in operation")
        Else
            writer.BaseStream.Seek(addr, 0)     ' go to offset
            writer.Write(n)                     ' write data
        End If
    End Sub

    Public Sub DisplayHeader(textfile As TextWriter)
        ' Display header info
        Dim chunk As Integer

        CurrentBlockType = BLOCKTYPE.HEADER
        reader.BaseStream.Position = 0
        textfile.WriteLine($"HEADER &0")
        textfile.WriteLine($"Size of file: &h{GetInt():x} bytes")
        Dim MotionBlocks = GetInt()
        textfile.WriteLine($"Motion blocks: {MotionBlocks}")
        Dim v() As Byte = reader.ReadBytes(4)
        textfile.WriteLine($"Version: {CInt(v(3))}.{CInt(v(2))}.{CInt(v(1))}.{CInt(v(0))}")
        TopRight = New IntPoint(GetInt(&H18), GetInt())
        textfile.WriteLine($"Origin: ({TopRight.X},{TopRight.Y}) ({TopRight.X / xScale:f2},{TopRight.Y / yScale:f2})mm")
        BottomLeft = New IntPoint(GetInt(&H20), GetInt())
        textfile.WriteLine($"Bottom Left: ({BottomLeft.X},{BottomLeft.Y}) ({BottomLeft.X / xScale:f2},{BottomLeft.Y / yScale:f2})mm")

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

    Public Sub DecodeStream(writer As System.IO.StreamWriter, block As BLOCKTYPE, StartAddress As Integer)
        ' Decode a stream of commands
        ' StartAddress - start of commands
        ' Set start position for this block
        Dim subr As Integer
        Select Case block       ' convert block to subroutine
            Case BLOCKTYPE.TEST : subr = 1
            Case BLOCKTYPE.CUT : subr = 2
            Case Else
                subr = 0
        End Select
        If StartPosns.ContainsKey(subr) Then
            If StartPosns(subr).Absolute Then position = StartPosns(subr).position Else position += StartPosns(subr).position
        Else position = New IntPoint(0, 0)
        End If
        writer.WriteLine()
        writer.WriteLine($"Decode of commands at 0x{StartAddress:x} in block {block}")
        writer.WriteLine()
        ' position = StartPosns(block)
        reader.BaseStream.Seek(StartAddress, 0)      ' Start of block
        Try
            Do
            Loop Until Not DecodeCmd(writer)
        Catch ex As Exception
            Throw New System.Exception($"DecodeStream failed at addr {reader.BaseStream.Position:x8}{vbCrLf}{ex.Message}")
        End Try
    End Sub

    Public Function DecodeCmd(writer As System.IO.StreamWriter) As Boolean
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
            Case MOL_MCBLK
                cmd_len = 1        ' allow contents of MCBLK to decode
                reader.BaseStream.Position = cmdBegin + 4     ' backup so we can read length as parameter
        End Select

        If Commands.TryGetValue(cmd, value) Then
            If value.ParameterType = ParameterCount.FIXED And cmd_len <> value.Parameters.Count And cmd <> MOL_MCBLK Then
                Throw New System.Exception($"Command {value.Mnemonic}: length is {cmd_len}, but data table says {value.Parameters.Count}")
            End If
            writer.Write($" {value.Mnemonic}")
            For Each p In value.Parameters
                writer.Write($" {p.Name}=")
                Select Case p.Typ
                    Case GetType(Boolean) : writer.Write(CType(GetInt(), Boolean))
                    Case GetType(Int32) : If p.Scale = 1.0 Then writer.Write($"{GetInt()} {p.Units}") Else writer.Write($"{GetInt() * p.Scale:f2} {p.Units}")
                    Case GetType(Double) : writer.Write($"{GetFloat() * p.Scale:f2} {p.Units}")
                    Case GetType(OnOff_type) : Dim par = GetInt() : If par Mod 2 = 0 Then writer.Write($"Off({par:x})") Else writer.Write($"On({par:x})")
                    Case GetType(Acceleration_type) : writer.Write($"{CType(GetInt(), Acceleration_type)}")
                    Case GetType(Axis_type) : writer.Write($"{CType(GetInt(), Axis_type)}")
                    Case GetType(OnOffSteps)    ' a list of On/Off steps
                        writer.Write($"List of {cmd_len} On/Off steps ")
                        Dim OneStep As OnOffSteps
                        For i = 1 To cmd_len    ' one structure for each word
                            OneStep.Steps = GetUInt()     ' get 32 bit word
                            writer.Write($" {OneStep.OnSteps * p.Scale:f2}/{OneStep.OffSteps * p.Scale:f2} {p.Units}")
                        Next
                    Case Else
                        Throw New System.Exception($"{value.Mnemonic}: Unrecognised parameter type of {p.Typ}")
                End Select
            Next
        Else
            ' UNKNOWN command. Just show parameters
            writer.Write($"Unknown: {cmd:x8} Params {cmd_len}: ")
            For i = 1 To cmd_len
                Dim n As Integer = GetInt()
                writer.Write($" {n:x8}")
            Next
        End If
        writer.WriteLine()
        'reader.BaseStream.Seek(cmdBegin + cmd_len * 4 + 4, 0)       ' move to next command
        Return True         ' more commands follow
    End Function

    Public Sub ExecuteStream(block As BLOCKTYPE, StartAddress As Integer, dxf As DxfDocument, DefaultLayer As Layer)
        ' Execute a stream of commands, rendering in dxf as we go
        Dim AddrTrace As New List(Of Integer)
        Dim layer = DefaultLayer
        Dim CurrentBlock = block
        ' Set start position for this block
        ' Translate start address to subroutine
        ' Look in start address

        LaserIsOn = False
        Dim addr = StartAddress
        Do
            'AddrTrace.Add(addr)
            addr = ExecuteCmd(CurrentBlock, addr, dxf, layer)       ' execute command at addr
        Loop Until addr = 0         ' loop until an END statement
    End Sub

    Public Function ExecuteCmd(ByRef block As BLOCKTYPE, ByVal addr As Integer, dxf As DxfDocument, ByRef Layer As Layer) As Integer
        ' Execute a command at addr. Return addr as start of next instruction
        Dim cmd As Integer, nWords As Integer

        Dim i As UInteger
        cmd = GetInt(addr)      ' get the command

        Dim command = cmd And &HFFFFFF    ' bottom 24 bits is command
        nWords = (cmd >> 24) And &HFF  ' command length is top 8 bits
        If nWords = &H80 Then nWords = GetInt() And &HFFFF
        If cmd = MOL_MCBLK Then nWords = 0    ' allow contents of MCBLK to execute
        Dim param_posn = stream.Position  ' remember where parameters (if any) start

        Select Case cmd
            Case MOL_END
                If motion.Vertexes.Count > 0 Then
                    ' add the motion to the dxf file    
                    DXF_polyline2d(motion, Layer, 1 / xScale)
                    motion.Vertexes.Clear()
                End If
                Return 0
            Case MOL_MCBLK
                nWords = 0      ' allow MCBLK content to be executed

            Case MOL_LASER, MOL_LASER1     ' switch laser
                ' Laser is changing state. Output any pending motion
                If motion.Vertexes.Count > 0 Then
                    ' Display any motion
                    DXF_polyline2d(motion, Layer, 1 / xScale)
                    motion.Vertexes.Clear()
                End If
                i = GetUInt()
                Select Case i
                    Case 0 : LaserIsOn = False
                    Case 1, 2, 3 : LaserIsOn = True
                    Case Else
                        If debugflag Then TextBox1.AppendText($"Unknown LaserIsOn parameter &h{i:x8}")
                End Select

            Case MOL_SCALE      ' also MOL_SPDPWR, MOL_SPDPWRx
                xScale = GetFloat() : yScale = GetFloat() : GetFloat()        ' x,y,z scale command

            Case MOL_MVREL, MOL_START
                Dim n = GetUInt()           ' always 772
                Dim delta = New IntPoint(GetInt(), GetInt())      ' move relative command

                Select Case block
                    Case BLOCKTYPE.CONFIG
                        ' MVREL in the config block are start positions for each draw block
                        Dim SubBlock = DrawChunks(MVRELInConfig)      ' get chunk for this draw block
                        Dim bookmark = reader.BaseStream.Position         ' remember posn in decode stream
                        Dim subr = GetInt(SubBlock * BLOCK_SIZE + 4)          ' get the subroutine number
                        Dim abs = (MVRELInConfig = 0)             ' first is absolute, other are relative
                        StartPosns.Add(subr, (abs, delta))              ' add subroutine start address
                        reader.BaseStream.Position = bookmark                 ' restore reader position
                        MVRELInConfig += 1

                    Case BLOCKTYPE.TEST, BLOCKTYPE.CUT
                        DXF_line(position, position + delta, Layer)     ' render on this layer ignoring laser commands

                    Case BLOCKTYPE.DRAW
                        If LaserIsOn Then
                            Dim colour As AciColor
                            Select Case CurrentSubr
                                Case 3 : colour = PowerSpeedColor(EngPower, EngSpeed)
                                Case 4 : colour = AciColor.Green
                                Case Else
                                    colour = AciColor.Default
                            End Select
                            DXF_line(position, position + delta, Layer, colour)
                        Else
                            DXF_line(position, position + delta, MoveLayer)
                        End If
                    Case Else
                        Throw New System.Exception($"Unexpected block type: {block}")
                End Select
                If block <> BLOCKTYPE.CONFIG Then position += delta       ' update head position

            Case MOL_ENGACD
                AccelLength = GetInt()

            Case MOL_ENGPWR
                EngPower = GetInt() / 100

            Case MOL_ENGSPD
                GetInt()    ' axis
                EngSpeed = GetFloat() / xScale

            Case MOL_ENGMVX   ' Engrave move X  (laser on and off)
                GetInt()      ' consume Axis parameter
                ' ENGMVX is in 3 phases, Accelerate, Engrave, Decelerate
                Dim TotalSteps = GetInt()
                Dim direction = Math.Sign(TotalSteps)     ' 1=LtoR, -1=RtoL
                ' Move for the initial acceleration
                Dim delta As New IntPoint(AccelLength * direction, 0)
                DXF_line(position, position + delta, MoveLayer)
                position += delta
                ' do On/Off steps
                For Each stp In ENGLSRsteps
                    ' The On portion of the delta
                    delta = New IntPoint(stp.OnSteps * direction, 0)    ' delta in X direction
                    DXF_line(position, position + delta, EngraveLayer, PowerSpeedColor(EngPower, EngSpeed))     ' color set to represent engrave shade
                    position += delta
                    ' The Off portion of the delta
                    delta = New IntPoint(stp.OffSteps * direction, 0)    ' delta in X direction
                    If delta <> ZERO Then DXF_line(position, position + delta, MoveLayer)
                    position += delta
                Next
                ' delta for the deceleration
                delta = New IntPoint(AccelLength * direction, 0)
                DXF_line(position, position + delta, MoveLayer)
                position += delta       ' move the position along

            Case MOL_ENGMVY     ' Engrave move Y (laser off)
                Dim axis = GetInt()      ' consume Axis parameter
                Dim dy = GetInt()
                Dim delta As New IntPoint(0, dy)    ' move in Y direction
                DXF_line(position, position + delta, MoveLayer)
                position += delta      ' move position along

            Case MOL_ENGLSR       ' engraving movement
                ENGLSRsteps.Clear()
                For i = 1 To nWords
                    Dim Steps As OnOffSteps
                    Steps.Steps = GetInt()      ' get 2 16 bit values, accessable through OnSteps & OffSteps
                    ENGLSRsteps.Add(Steps)
                Next

            Case MOL_BEGSUB, MOL_BEGSUBa  ' begin subroutine
                Dim n = GetUInt()
                CurrentSubr = n
                If n < 100 Then
                    If StartPosns.ContainsKey(n) Then
                        If StartPosns(n).Absolute Then position = StartPosns(n).position Else position += StartPosns(n).position
                        TextBox1.AppendText($"Starting subr {n} with position ({position.X},{position.Y}){vbCrLf}")
                    Else
                        Throw New SystemException($"GOSUB {n} has no start position")
                    End If
                End If
                Select Case n
                    Case 1
                        Layer = TestLayer
                        block = BLOCKTYPE.TEST
                    Case 2
                        Layer = CutLayer
                        block = BLOCKTYPE.CUT
                    Case 3
                        Layer = EngraveLayer
                        block = BLOCKTYPE.DRAW
                    Case 4
                        Layer = TextLayer
                        block = BLOCKTYPE.DRAW
                    Case 5 To 100
                        ' Create layer for this subroutine
                        Layer = New Layer($"SUBR {n}") With
                                {
                                    .Color = colors(ColorIndex),
                                    .Linetype = Linetype.Continuous,
                                    .Lineweight = Lineweight.Default
                                }
                        motion.Layer = Layer
                        ColorIndex += 20
                End Select

                ' Remove from ChunksWithCode as they have been executed
                Dim chunk = addr / BLOCK_SIZE
                ChunksWithCode(chunk) = True    ' flag as executed

            Case MOL_ENDSUB  ' end subroutine
                Dim n = GetUInt()
                ' Output any buffered polyline
                If motion.Vertexes.Count > 0 Then
                    ' add the motion to the dxf file    
                    DXF_polyline2d(motion, Layer, 1 / xScale)
                    motion.Vertexes.Clear()
                End If
                TextBox1.AppendText($"Position at end of subr {n} is ({position.X},{position.Y}){vbCrLf}")
                If Stack.Count > 0 Then
                    Dim popped = Stack.Pop
                    block = popped.item2    'blk
                    Layer = popped.item3    'lyr
                    reader.BaseStream.Seek(popped.item1, 0)     ' return to the saved return address
                    Return reader.BaseStream.Position
                Else
                    Throw New System.Exception($"Stack is exhausted - no return address for ENDSUB {n}")
                End If

            Case MOL_GOSUBn       ' Call subroutine with parameters
                Dim n = GetInt()     ' get subroutine number
                Dim posn = New IntPoint(GetFloat(), GetFloat())     ' subroutine parameter ignored
                If n < 100 Then
                    Stack.Push((addr:=reader.BaseStream.Position, blk:=block, lyr:=Layer))              ' address of next instruction + current 
                    Return SubrAddrs(n)                     ' jump to start address and continue
                Else
                    Return reader.BaseStream.Position       ' ignore call. Goto next instruction
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

    ' DXF_line sub with Vector2 parameters
    'Sub DXF_line(startpoint As Vector2, endpoint As Vector2, layer As Layer, ByVal Optional color As AciColor = Nothing)
    '    ' Add line to dxf file specified by startpoint, endpoint and layer. position is updated.
    '    DXF_line(Vector2ToPoint(startpoint), Vector2ToPoint(endpoint), layer, color)
    'End Sub

    ' DXF_line sub with IntPoint parameters
    Public Sub DXF_line(startpoint As IntPoint, endpoint As IntPoint, layer As Layer, ByVal Optional color As AciColor = Nothing)
        ' Add line to dxf file specified by startpoint, endpoint and layer. 
        Dim line As New Line(startpoint, endpoint) With {
            .Layer = layer
        }
        If color IsNot Nothing Then
            line.Color = color
        End If
        dxf.Entities.Add(line)
    End Sub

    Public Sub DXF_polyline2d(Polyline As Polyline2D, Layer As Layer, Optional ByVal scale As Double = 1.0)
        ' Add a polyline to the specified layer
        Dim ply As Polyline2D = Polyline.Clone
        ply.Layer = Layer
        If scale <> 1.0 Then
            Dim transform As New Matrix3(scale, 0, 0, 0, scale, 0, 0, 0, scale)
            ply.TransformBy(transform, New Vector3(0, 0, 0))
        End If
        dxf.Entities.Add(ply)
    End Sub

    Public Sub HexDumpBlock(addr As Integer)
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
            If Commands.TryGetValue(cmd, value) Then TextBox1.AppendText($" ({value.Mnemonic})")
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
        ' Creates a DXF drawing representing a test card
        ' Add MOL commands to MOL file representing a test card
        ' SUBR 3 in Engrave commands
        ' SUBR 4 is text labels

        Dim Speeds(9) As Integer    ' labels for rows (speed)
        Dim Powers(9) As Integer    '  labels for columns (power)
        Dim buffer() As Byte
        ' All Rect structures are in mm. They will be scaled during the draw process
        Dim Outline = New Rect(-85, -100, 85, 100)     ' outline of test card in mm
        Dim CutLine As New Rect()   ' grow the outline to be the cutline
        CutLine = Outline
        CutLine.Inflate(New Size(2, 2))        ' Inflate cutline by 2mm all round
        ' Scale outline and cutline to steps
        'Outline.Transform(ScaleToSteps)      ' convert outline to steps
        'CutLine.Transform(ScaleToSteps)
        Dim CellSize = New System.Windows.Size(5, 5)       ' size of test card cells, not including margin
        Dim CellMargin = New Size(1.5, 1.5)                 ' margin around each cell
        Dim GridLine = New System.Windows.Rect(-67, -80, (CellSize.Width + CellMargin.Width) * 10, (CellSize.Height + CellMargin.Height) * 10)   ' rectangle for the grid of test card grid
        Dim cellDimension = New Size(CellSize.Width + CellMargin.Width, CellSize.Height + CellMargin.Height)    ' Size of overall cell

        ' Make an array of speed and power settings for the test card
        For i = 0 To Speeds.Length - 1
            Speeds(i) = (My.Settings.SpeedMax - My.Settings.SpeedMin) / UBound(Speeds) * i + My.Settings.SpeedMin
        Next
        For i = 0 To Powers.Length - 1
            Powers(i) = (My.Settings.PowerMax - My.Settings.PowerMin) / UBound(Powers) * i + My.Settings.PowerMin
        Next

        Dim mode = IIf(My.Settings.Engrave, "ENG", "CUT")
        ' Filename comprises ENG/CUT, Material, Interval & passes
        Dim filename = $"TESTCARD_{mode}_{My.Settings.Material}_{My.Settings.Interval:f1}_{My.Settings.Passes}.MOL"     ' output file name
        writer = New BinaryWriter(System.IO.File.Open(filename, FileMode.Create), System.Text.Encoding.Unicode, False)
        dxf = New DxfDocument()     ' create empty DXF file

        ' Copy the first 5 blocks of a template file to initialise the new MOL file
        Dim ReadStream = System.IO.File.Open("line_d_10.MOL", FileMode.Open)           ' file containing template blocks
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

        MakeTestBlock(writer, dxf, Outline, TestLayer)
        MakeCutBlock(writer, dxf, CutLine, CutLayer)
        MakeEngraveBlock(writer, dxf, GridLine, EngraveLayer, Speeds, Powers)
        MakeTextBlock(writer, dxf, GridLine, TextLayer, Speeds, Powers)

        dxf.Save("Test card.dxf")       ' save the DXF file

        ' make binary file a multiple of BLOCK_SIZE
        Dim LastBlock = CInt(writer.BaseStream.Length \ BLOCK_SIZE)       '  current length of file in blocks
        writer.BaseStream.SetLength((LastBlock + 1) * BLOCK_SIZE)       ' round up to next block
        ' Now update template values
        FlushMCBLK(False)     ' just in case
        ' HEADER
        writer.BaseStream.Seek(0, SeekOrigin.Begin)
        PutInt(writer.BaseStream.Length)         ' file size
        PutInt(MCBLKCount + 30)                    ' number of MCBLK
        PutInt(MVRELCnt, &H10)                 ' Count of MVREL, UNKNOWN07 & UNKNOWN09
        ' Bottom left
        writer.Seek(&H20, SeekOrigin.Begin)
        writer.Write(CInt(Outline.X * xScale))
        writer.Write(CInt(Outline.Y * yScale))
        ' CONFIG
        ' Update to add subroutines
        writer.BaseStream.Position = &H274
        For i = 3 To 4
            WriteMOL(MOL_MOTION, {StartSpeed * xScale, QuickSpeed * xScale, WorkAcc * xScale})
            WriteMOL(MOL_START, {772, StartPosns(i).position.X, StartPosns(i).position.Y})
            WriteMOL(MOL_GOSUBn, {3, i, 0, 0})
        Next
        ' add unknown command
        WriteMOL(&H1000846, {0})
        ' add GOSUBn 603
        WriteMOL(MOL_GOSUBn, {3, 603, 0, 0})
        ' end the config block
        PutInt(MOL_END)

        writer.Close()
        TextBox1.AppendText($"Done - output in {filename}")
    End Sub

    Public Sub MakeTestBlock(writer As BinaryWriter, dxf As DxfDocument, outline As Rect, layer As Layer)
        ' Make the test block
        position = New IntPoint(0, 0)       ' Layer starts at (0,0)
        WriteMOL(MOL_BEGSUBa, {1}, &H400)    ' begin SUB 1
        WriteMOL(&H80600148, {6, &H25B, &H91D4000, &HF13A300, &H1113A340, &H101D7A80, &HF716C00})   ' Unknown command
        UseMCBLK = True
        DrawBox(writer, dxf, outline, layer)
        FlushMCBLK(False)
        WriteMOL(MOL_ENDSUB, {1})
        WriteMOL(MOL_END)
    End Sub
    ''' <summary>
    ''' Make the Cut block
    ''' </summary>
    Public Sub MakeCutBlock(writer As BinaryWriter, dxf As DxfDocument, outline As Rect, layer As Layer)
        position = New IntPoint(0, 0)       ' Layer starts at (0,0)
        WriteMOL(MOL_BEGSUBa, {2}, &H600)    ' begin SUB 1
        WriteMOL(&H80600148, {6, &H25B, &H91D4000, &HF13A300, &H1113A340, &H101D7A80, &HF716C00})   ' Unknown command
        WriteMOL(MOL_PWRSPD5, {4000, 4000, 629.0, 3149.0, 0.0})    ' set power & speed
        UseMCBLK = True
        Dim delta = CType(outline.BottomRight, IntPoint)     ' convert to steps
        DXF_line(position, position + delta, CutLayer)
        MoveRelativeSplit(outline.BottomRight)  ' Move to start position in two goes
        DrawBox(writer, dxf, outline, layer, False, 40, 25)
        FlushMCBLK(False)
        WriteMOL(MOL_ENDSUB, {2})
        WriteMOL(MOL_END)
    End Sub

    Public Sub MakeEngraveBlock(writer As BinaryWriter, dxf As DxfDocument, outline As Rect, layer As Layer, speeds() As Integer, powers() As Integer)

        ' Create engraving as SUBR 3. Subroutine engraves a single cell
        Dim StartofBlock = 5 * BLOCK_SIZE   ' address of start of next block. First subr at A00
        Dim BlockNumber = 5
        PutInt(BlockNumber, &H7C)
        writer.BaseStream.Position = StartofBlock       ' position at start of block
        WriteMOL(MOL_BEGSUBa, {3}, StartofBlock)    ' begin SUB 
        WriteMOL(&H1000D46, {&HBB8})         ' unknown command


        Dim cellsize = New Size(outline.Width / 10, outline.Height / 10)    ' there are 10 x 10 cells 

        ' Now create the engraved box for each setting
        ' Power goes left to right across page
        ' Speed goes top to bottom down the page
        ' boxes are drawn top down

        For power = 0 To powers.Length - 1
            For speed = speeds.Length - 1 To 0 Step -1
                Dim cell = New Rect(New System.Windows.Point(outline.Left + power * cellsize.Width, outline.Top + speed * cellsize.Height), cellsize)  ' one 10x10 cell
                cell = Rect.Inflate(cell, -0.75, -0.75)       ' shrink it a bit to create a margin
                DrawBox(writer, dxf, cell, layer, My.Settings.Engrave, powers(power), speeds(speed))
                WriteMOL(MOL_LASER, {OnOff_type.Off})    ' turn laser off
            Next
        Next
        WriteMOL(MOL_ENDSUB, {3})    ' end SUB 
        WriteMOL(MOL_END)
    End Sub

    Public Sub MakeTextBlock(writer As BinaryWriter, dxf As DxfDocument, Outline As Rect, layer As Layer, speeds() As Integer, powers() As Integer)
        ' Create Text layer as SUBR 4
        Dim StartofBlock = (writer.BaseStream.Position \ BLOCK_SIZE + 1) * BLOCK_SIZE   ' address of start of next block
        Dim BlockNumber = StartofBlock \ BLOCK_SIZE
        Dim cellsize = New Size(Outline.Width / 10, Outline.Height / 10)
        PutInt(BlockNumber, &H7C + 4)
        WriteMOL(MOL_BEGSUBa, {4}, StartofBlock)    ' begin SUB
        UseMCBLK = True
        DrawText(writer, dxf, My.Settings.Material, System.Windows.TextAlignment.Center, New System.Windows.Point(Outline.Left + (Outline.Width / 2), -5), 3, 0)
        DrawText(writer, dxf, $"Interval: {My.Settings.Interval} mm", System.Windows.TextAlignment.Center, New System.Windows.Point(Outline.Left + Outline.Width / 2, -9.5), 3, 0)
        DrawText(writer, dxf, $"Passes: { My.Settings.Passes}", System.Windows.TextAlignment.Center, New System.Windows.Point(Outline.Left + Outline.Width / 2, -14), 3, 0)
        ' Put labels on the axes
        For speed = 0 To speeds.Length - 1
            DrawText(writer, dxf, $"{speeds(speed)}", System.Windows.TextAlignment.Center, New System.Windows.Point(Outline.Left - 4, Outline.Top + speed * cellsize.Height + 2), 3, 0)
        Next
        For power = 0 To powers.Length - 1
            DrawText(writer, dxf, $"{powers(power)}", System.Windows.TextAlignment.Center, New System.Windows.Point(Outline.Left + power * cellsize.Width + 6, Outline.Top - 4), 3, 90)
        Next
        DrawText(writer, dxf, "Power (%)", System.Windows.TextAlignment.Center, New System.Windows.Point(Outline.Left + Outline.Width / 2, Outline.Top - 16), 5, 0)
        DrawText(writer, dxf, "Speed (mm/s)", System.Windows.TextAlignment.Center, New System.Windows.Point(Outline.Left - 9, Outline.Bottom - Outline.Height / 2), 5, 90)
        ' Finish block
        FlushMCBLK(False)
        WriteMOL(MOL_ENDSUB, {4})    ' end SUB 
        WriteMOL(MOL_END)
    End Sub

    Public Sub WriteMOL(command As Integer, Optional ByVal Parameters() As Object = Nothing, Optional posn As Integer = -1)
        ' Write a MOL command, with parameters to MOL file
        ' Parameters are written in scaled values
        ' Parameters will be written (but not length) if present
        ' writing occurs at current writer position, or "posn" if present
        If UseMCBLK And posn <> -1 Then Throw New System.Exception($"You can't write explicitly to address {posn:x} as an MCBLK in in operation")

        Dim value As LASERcmd = Nothing        ' Check the correct number of parameters
        If Commands.TryGetValue(command, value) Then
            ' It is a known command, therefore do some checks
            Dim nWords = value.Parameters.Count
            If Parameters Is Nothing Then
                If nWords <> 0 Then
                    Throw New System.Exception($"WRITEMOL {value.Mnemonic}: parameter error. Parameter is Nothing, yet command requires {nWords}")
                End If
            Else
                If value.ParameterType = ParameterCount.FIXED And Parameters.Length <> nWords Then
                    Throw New System.Exception($"{value.Mnemonic}: parameter error. There are {Parameters.Length} parameters, yet command requires {nWords}")
                End If
            End If

            ' Check correct type of parameters
            If nWords > 0 And value.ParameterType = ParameterCount.FIXED Then
                For i = 0 To nWords - 1
                    Dim a = value.Parameters(i).Typ
                    Dim b = Parameters(i).GetType
                    If a <> b Then
                        Throw New System.Exception($"{value.Mnemonic}: parameter {value.Parameters(i).Name}: type mismatch - spec is {a}, call is {b}")
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
                    Case GetType(Double) : PutFloat(p)
                    Case GetType(System.Collections.Generic.List(Of OnOffSteps))
                        For Each s In p
                            PutInt(s.steps)
                        Next
                    Case GetType(Acceleration_type), GetType(OnOff_type), GetType(Axis_type)
                        PutInt(p)
                    Case Else
                        Throw New System.Exception($"WRITEMOL Parameter of unsupported type {p.GetType}")
                End Select
            Next
        End If
        ' Count the number of MOL_MVREL, MOL_UNKNOWN07 & MOL_UNKNOWN09 commands
        Select Case command
            Case MOL_MVREL, MOL_UNKNOWN07, MOL_UNKNOWN09
                MVRELCnt += 1
        End Select
        If MCBLK.Count >= MCBLKMax Then
            FlushMCBLK(True)      ' MCBLK limited in size
        End If
    End Sub

    Public Sub FlushMCBLK(KeepOpen As Boolean)
        ' Flush the contents of the MCBLK.
        ' The state of UseMCBLK will be updated to KeepOpen when complete
        If UseMCBLK Then
            If MCBLK.Count > 0 Then
                ' Disable MCBLK buffering
                UseMCBLK = False
                PutInt(MOL_MCBLK) : PutInt(MCBLK.Count)   ' write MCBLK command and count
                ' Write the contents of the MCBLK
                For Each item In MCBLK
                    PutInt(item)
                Next
                MCBLK.Clear()          ' clear the buffer
                MCBLKCount += 1                        ' count MCBLK
            End If
            UseMCBLK = KeepOpen       ' Turn MCBLK buffering back on
        End If
    End Sub


    Public Sub MoveRelativeSplit(p As System.Windows.Point)
        MoveRelativeSplit(CType(p, IntPoint))       ' convert mm to steps
    End Sub
    Public Sub MoveRelativeSplit(p As IntPoint)
        ' Create a move command from the current position to p
        ' Pieces = 2 or 3
        ' Accelerate before first one
        ' Decelerate before last one
        ' There may be a middle piece
        ' Phase 1 is slow, Phase 2 fast, phase 3 slow
        Dim moves As New List(Of IntPoint)
        ' Work out whether we can do a 2 part move, or it's too long and we need 3
        Dim Accel = IIf(LaserIsOn, WorkAcc, SpaceAcc)   ' Acceration difers if laser is on or not
        Dim T = (QuickSpeed - StartSpeed) / Accel       ' T =(max speed-initial speed)/acceleration time taken to reach QuickSpeed
        Dim S As Integer = CInt(StartSpeed * T + xScale * 0.5 * Accel * T ^ 2)  ' s=ut+0.5t*t  distance (steps) travelled whilst reaching this speed
        ' Calculate 2 or 3 part move
        Dim Dist = Math.Sqrt(p.X ^ 2 + p.Y ^ 2)         ' distance to move

        If S > Dist / 2 Then
            ' we can get more than halfway if we needed, so 2 pieces enough
            ' We do it this way to avoid rounding errors
            Dim delta1 = New IntPoint(p.X / 2, p.Y / 2)
            moves.Add(delta1)
            Dim delta2 = p - delta1
            moves.Add(delta2)
        Else
            ' we can't make it halfway whilst accelerating, so will need to coast
            Dim ratio = S / Dist        ' percentage of accelerate/decelerate phases
            Dim delta1 = New IntPoint(p.X * ratio, p.Y * ratio)
            Dim delta2 = New IntPoint(p.X * (1 - 2 * ratio), p.Y * (1 - 2 * ratio))
            Dim delta3 = p - delta1 - delta2
            moves.Add(delta1)       ' initial move
            moves.Add(delta2)       ' middle move
            moves.Add(delta3)       ' final move
        End If

        For m = 0 To moves.Count - 1
            If m = 0 Then    ' the first one
                WriteMOL(MOL_ACCELERATION, {Acceleration_type.Accelerate})
                WriteMOL(MOL_SETSPD, {StartSpeed * xScale, QuickSpeed * xScale, Accel * xScale})
            ElseIf m = moves.Count - 1 Then    ' the last one
                WriteMOL(MOL_ACCELERATION, {Acceleration_type.Decelerate})
                WriteMOL(MOL_SETSPD, {StartSpeed * xScale, QuickSpeed * xScale, Accel * xScale})
            Else
                WriteMOL(MOL_SETSPD, {QuickSpeed * xScale, QuickSpeed * xScale, Accel * xScale})     ' Quick speed in the middle section
            End If
            WriteMOL(MOL_MVREL, {772, moves(m).X, moves(m).Y})
        Next
        position += p       ' update position

    End Sub

    Public Sub DrawBox(writer As BinaryWriter, dxf As DxfDocument, outline As Rect, Layer As Layer, Optional shaded As Boolean = False, Optional power As Integer = 0, Optional speed As Integer = 0)
        ' Draw a box
        ' outline is a rect defining the box in steps
        ' power as %
        ' speed as mm/sec
        ' Boxes start at the BottomRight.
        ' Non shaded are drawn anti-clockwise, ending at the start point
        ' Shaded boxes end where the engraving ends
        ' if power or speed are 0, don't turn laser on

        TextBox1.AppendText($"Drawing box at ({CInt(outline.X)},{CInt(outline.Y)}) width {CInt(outline.Width)} height {CInt(outline.Height)} power {power} speed {speed} engraved {shaded}{vbCrLf}")
        My.Application.DoEvents()
        Dim motion As New Polyline2D With {.Layer = Layer}

        Dim shading As AciColor = PowerSpeedColor(power, speed)     ' color to represent engrave shade
        If Layer.Equals(EngraveLayer) Or Layer.Equals(TextLayer) Then
            Layer.Color = shading
        End If
        If Not shaded Then
            ' start position is "top right"
            If Layer.Equals(EngraveLayer) Then
                If Not StartPosns.ContainsKey(3) Then
                    position = outline.BottomRight      ' is converted to steps
                    StartPosns.Add(3, (True, position))     ' this is the first point in SUBR 3 add a starting position
                End If
            End If
            ' If we are not at the start position, move to start
            Dim delta As IntPoint = CType(outline.BottomRight, IntPoint) - position    ' really top right
            ' MOL commands
            If delta <> ZERO Then
                DXF_line(position, position + delta, MoveLayer)
                MoveRelativeSplit(delta)        ' Move to start of box (BottomRight) in 2 pieces
            End If
            'WriteMOL(MOL_pwrspd,{})
            ' Ordinary box with 4 sides
            If power > 0 Or speed > 0 Then
                WriteMOL(MOL_LASER, {OnOff_type.[On]})
                LaserIsOn = True
            End If
            MoveRelativeSplit(New System.Windows.Point(-outline.Width, 0))
            MoveRelativeSplit(New System.Windows.Point(0, -outline.Height))
            MoveRelativeSplit(New System.Windows.Point(outline.Width, 0))
            MoveRelativeSplit(New System.Windows.Point(0, outline.Height))
            If power > 0 Or speed > 0 Then WriteMOL(MOL_LASER, {OnOff_type.Off}) : LaserIsOn = False
            ' DXF commands
            ' draw 4 lines surrounding the cell
            Dim x1 = outline.Right, x2 = outline.Left, y1 = outline.Bottom, y2 = outline.Top
            With motion.Vertexes
                .Add(New Polyline2DVertex(x1, y1))
                .Add(New Polyline2DVertex(x2, y1))
                .Add(New Polyline2DVertex(x2, y2))
                .Add(New Polyline2DVertex(x1, y2))
                .Add(New Polyline2DVertex(x1, y1))
            End With
            DXF_polyline2d(motion, Layer)
        Else
            ' Engraved box
            ' Create equivalent MOL & DXF in parallel
            ' Move to start of engraving (TopLeft)
            ' create an hls color to represent power and speed as a shade of brown (=30 degrees)
            Dim startposn As IntPoint = position

            Layer.Color = shading
            Dim EngParameters = EngraveParameters(speed)
            Dim engacd = CInt(EngParameters.Acclen * xScale)     ' Acceleration distance in steps
            Dim deltaacd = New IntPoint(engacd, 0)       ' we need to backup the acceleration distance
            ' start position is "bottom left" - ACD
            If Layer.Equals(EngraveLayer) Then
                If Not StartPosns.ContainsKey(3) Then
                    position = outline.TopLeft      ' is converted to steps
                    position -= deltaacd
                    StartPosns.Add(3, (True, position))     ' this is the first point in SUBR 3 add a starting position
                End If
            End If
            ' get to start position. First get to bottom left of box, and then backup the acceleration distance
            Dim delta As IntPoint = CType(outline.TopLeft, IntPoint) - position     ' really bottom left
            delta -= deltaacd
            ' Move to start of engraving
            DXF_line(position, position + delta, MoveLayer)     ' move to start
            MoveRelativeSplit(delta)
            WriteMOL(MOL_LASER1, {2})     ' turn laser on engrave mode?
            Dim Steps As OnOffSteps                 ' construct steps as on for cellwidth, and off for 0
            Steps.OnSteps = outline.Width * xScale
            Steps.OffSteps = 0
            WriteMOL(MOL_ENGACD, {engacd})     ' define the acceleration start distance
            WriteMOL(MOL_ENGPWR, {power * 100})             ' define power
            WriteMOL(MOL_ENGSPD, {Axis_type.X, speed * xScale})          ' define speed
            Dim OnOffTotal As Integer = engacd * 2 + Steps.OnSteps + Steps.OffSteps
            Dim OnOffOnly As Integer = Steps.OnSteps + Steps.OffSteps   ' only on/off step distance
            Dim direction As Integer = 1        ' moving L to R
            Dim height As Integer = 0    ' count lines as integer and multiply to get height to avoid rounding errors
            Dim YInc As New IntPoint(0, My.Settings.Interval * yScale)  ' Y increment steps
            While height < outline.Height * xScale     ' do until we reach cell height
                WriteMOL(MOL_ENGLSR, {1, New List(Of OnOffSteps) From {{Steps}}})      ' one full line of on/off
                WriteMOL(MOL_ENGMVX, {Axis_type.X, OnOffTotal * direction})
                ' The ENGMVX movement occurs in 3 phases. Accelerate, OnOff steps, Decelerate. Need to backup for the Accelerate phase
                Dim accdelta = New IntPoint(engacd * direction, 0)
                DXF_line(position, position + accdelta, EngraveMoveLayer)      ' Accelerate
                position += accdelta
                Dim engdelta = New IntPoint(OnOffOnly * direction, 0)    ' distance we move in steps
                DXF_line(position, position + engdelta, EngraveCutLayer, shading)     ' Engrave
                position += engdelta
                DXF_line(position, position + accdelta, EngraveMoveLayer)      ' Decelerate same as accelerate
                position += accdelta
                ' Handle Y movement
                WriteMOL(MOL_ENGMVY, {Axis_type.Y, YInc.Y})
                DXF_line(position, position + YInc, EngraveMoveLayer)
                position += YInc
                direction *= -1           ' reverse direction for next line
                height += YInc.Y     ' move up to next line
            End While
        End If
    End Sub

    Public Sub DrawText(writer As BinaryWriter, dxf As DxfDocument, text As String, alignment As System.Windows.TextAlignment, origin As System.Windows.Point, fontsize As Double, angle As Double)
        ' Write some text at specified size, position and angle

        Dim Strokes = DisplayString(text, alignment, origin, fontsize, angle)       ' Convert the text to strokes
        Dim shading = PowerSpeedColor(My.Settings.PowerMax / 2, My.Settings.SpeedMax / 2)      ' use 50% power and speed for text

        ' draw the text
        ' DXF first
        UseMCBLK = True
        ' First render each stroke to DXF
        For Each stroke In Strokes
            stroke.Color = shading
            DXF_polyline2d(stroke, TextLayer)       ' Delay output of polyline to prevent position corruption
        Next

        ' render to MOL
        Dim matrix As New Matrix4()
        ' Initialize the matrix elements to do scaling only
        With matrix
            .M11 = xScale : .M12 = 0 : .M13 = 0 : .M14 = 0
            .M21 = 0 : .M22 = yScale : .M23 = 0 : .M24 = 0
            .M31 = 0 : .M32 = 0 : .M33 = 1 : .M34 = 0
            .M41 = 0 : .M42 = 0 : .M43 = 0 : .M44 = 1
        End With
        For Each stroke In Strokes
            stroke.TransformBy(matrix)    ' convert mm data to steps
            ' if this is the first figure in SUBR 4, record in StartPosns
            Dim StartPoint As New IntPoint(stroke.Vertexes.First.Position.X, stroke.Vertexes.First.Position.Y)
            If Not StartPosns.ContainsKey(4) Then
                Dim offset As IntPoint = StartPoint - position
                TextBox1.AppendText($"Setting subr 4 offset to ({offset.X},{offset.Y}){vbCrLf}")
                StartPosns.Add(4, (False, offset))            ' start offset for this subroutine
                position = StartPoint     ' position will be StartPoint when subr executes
            End If

            '
            ' Get to the start position of this stroke
            Dim delta As IntPoint = StartPoint - position
            If delta <> ZERO Then
                DXF_line(position, position + delta, MoveLayer)
                MoveRelativeSplit(delta)      ' move in 2 segments
            End If
            WriteMOL(MOL_LASER, {OnOff_type.[On]})
            ' Process all vertexes, except the first
            For p = 1 To stroke.Vertexes.Count - 1
                Dim thispoint = New IntPoint(stroke.Vertexes(p).Position.X, stroke.Vertexes(p).Position.Y)
                delta = thispoint - position
                MoveRelativeSplit(delta)          ' draw stroke
            Next
            WriteMOL(MOL_LASER, {OnOff_type.Off})

        Next
        FlushMCBLK(False)
    End Sub

    Private Sub TestIEEEToLeetroFpConversionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestIEEEToLeetroFpConversionToolStripMenuItem.Click
        ' Denormalized: When the exponent Is all zeros, but the mantissa Is Not
        ' Infinity: Represented by an exponent of all ones And a mantissa of all zeros
        ' Not a Number (NaN): Represented when the exponent field Is all ones with a zero sign bit

        ' List of Leetro and IEEE floats known to be equivalent
        Dim KnownEquivalents As New List(Of (IEEE As Double, Leetro As Integer)) From {
            {(0, 0)},
            {(208.328, &H7505555)},
            {(4, &H2000000)},
            {(3, &H1400000)},
            {(8.003, &H3000E46)},
            {(800, &H9480000)},
            {(125.98047, &H67BF7F0)},
            {(313.71875, &H81CDCDD)},
            {(-100, &H9FAEFFFF)},
            {(1041, &HA022000)},
            {(41666, &HF22C200)},
            {(145833, &H110E6A40)},
            {(1473.12, &HA3823B3)},
            {(0.1, 0)}
            }
        Dim TestCases() = {208.3218, 800, 0, 1000.0, -1000.0, -1.0, 1.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, -0.1, -2, 2, 10, -10, 0.9, 10.0 ^ 4, 10.0 ^ 5, 10.0 ^ 7, 10.0 ^ 8}

        TextBox1.AppendText($"Round trip discrepancy for 0.1 is {Leetro2Float(Float2Leetro(0.1))}{vbCrLf}")
        ' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]
        For i = 0.1 To 1 Step 0.1
            Dim n = Float2Leetro(i)
            TextBox1.AppendText($"{i:f1}  0x{n:x8} Sign={(n >> 23) And 1} exp={(n >> 24) And &HFF - 256} Mantissa={n And &H7FFFFF}{vbCrLf}")
        Next
        ' Test known cases
        TextBox1.AppendText($"Known Equivalents test{vbCrLf}")
        For Each t In KnownEquivalents
            Dim IEEE = t.IEEE
            Dim Leetro = t.Leetro 'And &HFFFFFE00
            Dim toIEEE = Leetro2Float(Leetro)
            Dim toLeetro = Float2Leetro(IEEE)
            'Dim toIEEE = LeetroToIEEE(Leetro)
            'Dim toLeetro = IEEEToLeetro(IEEE)
            TextBox1.AppendText($"IEEE {IEEE:f4} <=> Leetro 0x{Leetro:x8}: ")
            If Math.Abs(IEEE - toIEEE) < 0.1 Then TextBox1.AppendText("Leetro->IEEE Success ") Else TextBox1.AppendText($"Leetro->IEEE **Failure** {toIEEE:f4} result ")
            If Leetro = toLeetro Then TextBox1.AppendText("IEEE->Leetro Success ") Else TextBox1.AppendText($" IEEE->Leetro **Failure** 0x{toLeetro:x8} result")
            TextBox1.AppendText(vbCrLf)
        Next

        TextBox1.AppendText($"Random test cases{vbCrLf}")
        For Each t In TestCases
            TextBox1.AppendText($"Testcase {t}: ")
            Dim ToInt = Float2Leetro(t)
            Dim FromInt As Double = Leetro2Float(ToInt)
            If Math.Abs(t - FromInt) < 0.1 Then
                TextBox1.AppendText("Success")
            Else
                TextBox1.AppendText($"Failure: result {FromInt}")
            End If
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
        DXF_polyline2d(motion, TextLayer)
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
            DXF_polyline2d(motion, TextLayer)
        Next
        dxf.Save("Text test.dxf")
        TextBox1.AppendText("Done")
    End Sub

    Private Sub DXFMetricsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DXFMetricsToolStripMenuItem.Click
        ' Calculate some metrics for a DXF file
        Dim layers As New List(Of String), travel As Double = 0, vertexes As Integer = 0
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
        Dim sorted = From item In Commands
                     Order By item.Key And &HFFFFFF
                     Select item
        For Each cmd In sorted
            TextBox1.AppendText($"0x{cmd.Key:x8}  {cmd.Value.Mnemonic,-15}  ")
            'now parameters
            Dim parameters = New List(Of String)
            For Each p In cmd.Value.Parameters
                parameters.Add($"{p.Name} is {p.Typ.Name}")
            Next
            TextBox1.AppendText(String.Join(", ", parameters.ToArray))
            TextBox1.AppendText($"{vbCrLf}")
        Next
        ' Custom types. Display details of any parameters that are defined using ENUM. These are used when a basetype won't do
        Dim CustomTypes As New List(Of Type)
        For Each cmd In Commands       ' search all commands
            For Each p In cmd.Value.Parameters  ' searcg all parameters
                If p.Typ.BaseType.Name = "Enum" Then    ' Add any non standand types to list
                    If Not CustomTypes.Contains(p.Typ) Then CustomTypes.Add(p.Typ)
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
            Initialise()
            stream = System.IO.File.Open(dlg.filename, FileMode.Open)
            reader = New BinaryReader(stream, System.Text.Encoding.Unicode, False)

            Dim Size As Long = reader.BaseStream.Length / 4     ' size of the inputfile in 4 byte words
            ' Open a text file to write to
            Dim TextFile As System.IO.StreamWriter
            Dim Basename = System.IO.Path.GetFileNameWithoutExtension(dlg.filename)
            Dim textname = $"{Basename}.txt"
            TextFile = My.Computer.FileSystem.OpenTextFileWriter(textname, False)
            ' DISASSEMBLE everything
            TextFile.WriteLine($"Disassembly of {dlg.filename} on {Now}")
            DisplayHeader(TextFile)
            DecodeStream(TextFile, BLOCKTYPE.CONFIG, BLOCKTYPE.CONFIG * BLOCK_SIZE)
            DecodeStream(TextFile, BLOCKTYPE.TEST, BLOCKTYPE.TEST * BLOCK_SIZE)
            DecodeStream(TextFile, BLOCKTYPE.CUT, BLOCKTYPE.CUT * BLOCK_SIZE)
            For Each chunk In DrawChunks
                DecodeStream(TextFile, BLOCKTYPE.DRAW, chunk * BLOCK_SIZE)
            Next
            ' Display command usage
            TextFile.WriteLine()
            TextFile.WriteLine("Command frequency")
            TextFile.WriteLine()
            For Each cmd In CommandUsage
                Dim value As LASERcmd = Nothing, decode As String
                If Commands.TryGetValue(cmd.Key, value) Then decode = value.Mnemonic Else decode = $"0x{cmd.Key:x8}"
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
            Initialise()
            DrawChunks.Clear()

            position = ZERO
            TextBox1.Clear()
            stream = System.IO.File.Open(dlg.filename, FileMode.Open)
            reader = New BinaryReader(stream, System.Text.Encoding.Unicode, False)

            ' Get critical parameters from header
            TopRight = New IntPoint(GetInt(&H18), GetInt())
            BottomLeft = New IntPoint(GetInt(&H20), GetInt())
            ConfigChunk = GetInt(&H70)
            TestChunk = GetInt()
            CutChunk = GetInt()
            DrawChunks.Clear()
            Dim Chunk = GetInt(&H7C)        ' first draw chunk
            While Chunk <> 0
                DrawChunks.Add(Chunk)
                Chunk = GetInt()
            End While

            ' Harvest subroutine start addresses
            ' Look at the first instruction of each block. BEGSUBa instruction there has the SUBR number
            Dim blocks As New List(Of Integer) From {
                2,      ' TEST
                3       ' CUT
            }
            For Each b In DrawChunks
                blocks.Add(b)
            Next
            For Each block In blocks
                Dim addr = block * BLOCK_SIZE
                Dim n = GetInt(addr)
                If n <> MOL_BEGSUB And n <> MOL_BEGSUBa Then Throw New System.Exception($"Block {block} does not start with a BEGSUB or BEGSUBa command")
                n = GetInt()    ' the subroutine number
                SubrAddrs.Add(n, addr)
            Next

            ' Now execute real code
            dxf = New DxfDocument()         ' render commands in dxf

            ' Create a list of all chucks known to contain code. These will be removed from the list after execution.
            ' If there are any left, i.e. because the config code doesn't call them, they will be called
            ChunksWithCode.Clear()
            ChunksWithCode.Add(TestChunk, False)
            ChunksWithCode.Add(CutChunk, False)
            For Each s In DrawChunks
                ChunksWithCode.Add(s, False)
            Next
            MVRELInConfig = 0
            Stack.Push((addr:=0, blk:=BLOCKTYPE.TEST, lyr:=TestLayer))   ' any old rubbish to keep ENDSUB happy
            position = ZERO      ' this block has an origin of (0,0)
            ExecuteStream(BLOCKTYPE.TEST, TestChunk * BLOCK_SIZE, dxf, TestLayer)   ' test

            Stack.Push((addr:=0, blk:=BLOCKTYPE.CUT, lyr:=CutLayer))   ' any old rubbish to keep ENDSUB happy
            position = ZERO      ' this block has an origin of (0,0)
            ExecuteStream(BLOCKTYPE.CUT, CutChunk * BLOCK_SIZE, dxf, CutLayer)   ' cut

            Stack.Push((addr:=0, blk:=BLOCKTYPE.CONFIG, lyr:=ConfigLayer))   ' any old rubbish to keep ENDSUB happy
            position = ZERO      ' this block has an origin of (0,0)
            ExecuteStream(BLOCKTYPE.CONFIG, ConfigChunk * BLOCK_SIZE, dxf, ConfigLayer)   ' config (which calls everything else)

            ' Now execute any that weren't called from config
            For Each s In ChunksWithCode
                If Not s.Value Then
                    Stack.Push((addr:=0, blk:=BLOCKTYPE.DRAW, lyr:=DrawLayer))   ' any old rubbish to keep ENDSUB happy
                    position = ZERO      ' this block has an origin of (0,0)
                    ExecuteStream(BLOCKTYPE.DRAW, s.Key * BLOCK_SIZE, dxf, DrawLayer)   ' config (which calls everything else)
                End If
            Next

            ' Add the filename at the bottom of the image
            Dim BaseDXF = System.IO.Path.GetFileNameWithoutExtension(dlg.filename)
            Dim DXFfile = $"{BaseDXF}_MOL.dxf"
            Dim rect = New System.Windows.Rect(TopRight, BottomLeft)    ' bounding rectangle
            entity = New Text($"File: {DXFfile}") With {
                .Position = New Vector3(rect.X + rect.Width / 2, rect.Y - rect.Height / 10, 0),
                .Alignment = Entities.TextAlignment.TopCenter,
                .Layer = TextLayer,
                .Height = Math.Max(rect.Width, rect.Height) / 25
                }
            dxf.Entities.Add(entity)
            ' Save dxf in file
            dxf.Save(DXFfile)
            stream.Close()
            ' display the list of start points
            TextBox1.AppendText($"Start points: {StartPosns.Count}{vbCrLf}")
            For Each sp In StartPosns
                TextBox1.AppendText($" {sp.Key} : Absolute {sp.Value.Absolute}, x={sp.Value.position.X}, y={sp.Value.position.Y}{vbCrLf}")
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

    Public Sub Initialise()
        ' Initialise all variables
        ' Setup some start positions
        StartPosns.Clear()
        DrawChunks.Clear()
        ' Create initial start positions
        SubrAddrs.Clear()
        ' Subr 1 & 2 are TEST & CUT
        StartPosns.Add(1, (True, New IntPoint(0, 0)))
        StartPosns.Add(2, (True, New IntPoint(0, 0)))
        position = ZERO
        ChunksWithCode.Clear()

        TextBox1.Clear()

    End Sub
    Private Sub BoxTestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BoxTestToolStripMenuItem.Click
        writer = New BinaryWriter(System.IO.File.Open("BoxTest.bin", FileMode.Create), System.Text.Encoding.Unicode, False)
        Dim outline = New Rect(-200, -100, 60, 60)
        DrawBox(writer, dxf, outline, TestLayer)
        outline = New Rect(-180, -80, 20, 20)
        DrawBox(writer, dxf, outline, CutLayer)
        dxf.Save("BoxTest.dxf")
    End Sub

    Private Sub LineTestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LineTestToolStripMenuItem.Click
        ' draw a 10mm horizontal line
        dxf = New DxfDocument
        'dxf.Entities.Add(New Line(New Vector2(0, 0), New Vector2(0, 10)))
        DXF_line(New System.Windows.Point(0, 0), New System.Windows.Point(0, 10), MoveLayer)
        dxf.Save("LineTest.dxf")
    End Sub

    Private Sub ConvertLeetroToIEEEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConvertLeetroToIEEEToolStripMenuItem.Click
        Dim hex As String = InputBox("Input Leetro float in hex format", "")
        Dim value As UInt32 = Convert.ToUInt32(hex, 16)
        Dim float = Leetro2Float(value)
        MsgBox($"The value of 0x{hex} as a Leetro float is {float}", vbInformation + vbOKOnly, "Conversion to float")
    End Sub

    Private Sub CreateYAMLFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateYAMLFileToolStripMenuItem.Click

        Dim serializer = New SerializerBuilder() _
            .Build()

        Dim yaml = serializer.Serialize(Commands)
        TextBox1.AppendText($"{yaml}{vbCrLf}{vbCrLf}")

    End Sub

    Enum ReaderState
        Idle
        ReadVertexes
    End Enum
    Private Sub FontTestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontTestToolStripMenuItem.Click
        Dim dxf As New DxfDocument
        Dim st As String, ply As List(Of Polyline2D)

        st = "the quick brown fox jumps over the lazy dog! 0123456789"
        ply = DisplayString(st, System.Windows.TextAlignment.Left, New System.Windows.Point(10, 10), 9, 0)

        For Each pl In ply
            dxf.Entities.Add(pl)
        Next

        ' Draw a crosshair
        Dim Cross As New System.Windows.Point(300, 300)
        Dim w As New System.Windows.Size(10, 10)
        Dim ln As New Line(New Vector2(Cross.X - w.Width, Cross.Y), New Vector2(Cross.X + w.Width, Cross.Y)) With {.Color = AciColor.Red}
        dxf.Entities.Add(ln)
        ln = New Line(New Vector2(Cross.X, Cross.Y - w.Height), New Vector2(Cross.X, Cross.Y + w.Height)) With {.Color = AciColor.Red}
        dxf.Entities.Add(ln)
        ' 5,2.5;0,2.5,A-1;0,6.5;5,6.5,A-1;5,2.5 = "O"
        ' 1,10;1,-1,A0.181818   = "("

        Dim polyl As New Polyline2D
        polyl.Vertexes.Add(New Polyline2DVertex(New Vector2(1, 10)))
        polyl.Vertexes.Add(New Polyline2DVertex(New Vector2(1, -1), 0.181818))

        st = "String at 0 center"
        ply = DisplayString(st, System.Windows.TextAlignment.Center, Cross, 10, 0)
        For Each pl In ply
            dxf.Entities.Add(pl)
        Next
        st = "String at 90 left"
        ply = DisplayString(st, System.Windows.TextAlignment.Left, Cross, 10, 90)
        For Each pl In ply
            dxf.Entities.Add(pl)
        Next
        st = "String at 45 right"
        ply = DisplayString(st, System.Windows.TextAlignment.Right, Cross, 10, 45)
        For Each pl In ply
            dxf.Entities.Add(pl)
        Next

        st = "Power (%) - 10mm"
        ply = DisplayString(st, System.Windows.TextAlignment.Left, New System.Windows.Point(50, 50), 10, 45)
        For Each pl In ply
            dxf.Entities.Add(pl)
        Next

        st = "This is left justified - 5mm"
        ply = DisplayString(st, System.Windows.TextAlignment.Left, New System.Windows.Point(200, 60), 5, 0)
        For Each pl In ply
            dxf.Entities.Add(pl)
        Next
        st = "This is center justified - 5mm"
        ply = DisplayString(st, System.Windows.TextAlignment.Center, New System.Windows.Point(200, 70), 5, 0)
        For Each pl In ply
            dxf.Entities.Add(pl)
        Next
        st = "This is right justified - 5mm"
        ply = DisplayString(st, System.Windows.TextAlignment.Right, New System.Windows.Point(200, 80), 5, 0)
        For Each pl In ply
            dxf.Entities.Add(pl)
        Next
        st = "This is center justified - 20mm"
        ply = DisplayString(st, System.Windows.TextAlignment.Center, New System.Windows.Point(200, 90), 20, 0)
        For Each pl In ply
            dxf.Entities.Add(pl)
        Next

        dxf.Save("GlyphTest.dxf")
        TextBox1.AppendText("Done")
    End Sub

    Function DisplayString(text As String, alignment As System.Windows.TextAlignment, origin As System.Windows.Point, fontsize As Double, rotation As Double) As List(Of Polyline2D)
        ' Generate a list of Polyline2D for a string at origin, with scale and rotation
        ' text - string to be rendered
        ' alignment - left, center or right about origin
        ' fontsize - in mm
        ' rotation - about origin in degrees
        ' Returns a polyline for each stroke

        Const RawFontSize = 9     ' the fonts are defined 9 units high
        Const LetterSpacing = 3       ' space between letters
        Dim Width As Double
        Dim Strokes As List(Of Polyline2D)
        Dim utf8Encoding As New System.Text.UTF8Encoding()
        Dim encodedString() As Byte
        Dim baseline As Integer = 0             ' X position of next character
        Dim result As New List(Of Polyline2D)
        Dim ply As Polyline2D

        encodedString = utf8Encoding.GetBytes(text)       ' encode string to UTF-8
        ' Create a matrix to perform scaling, rotation and translation
        Dim TranslateMatrix As New Matrix4()        ' Matrix to perform a translation only
        With TranslateMatrix
            .M11 = 1 : .M12 = 0 : .M13 = 0 : .M14 = 0
            .M21 = 0 : .M22 = 1 : .M23 = 0 : .M24 = 0
            .M31 = 0 : .M32 = 0 : .M33 = 1 : .M34 = 0
            .M41 = 0 : .M42 = 0 : .M43 = 0 : .M44 = 1
        End With

        Dim TransformMatrix As New Matrix4()           ' Matrix to perform a rotation, scaling and traslation
        Dim cosTheta = Math.Cos(rotation * Math.PI / 180.0)
        Dim sinTheta = Math.Sin(rotation * Math.PI / 180.0)
        Dim scale = fontsize / RawFontSize          ' scale font to required fontsize mm
        With TransformMatrix
            .M11 = scale * cosTheta : .M12 = scale * -sinTheta : .M13 = 0 : .M14 = origin.X
            .M21 = scale * sinTheta : .M22 = scale * cosTheta : .M23 = 0 : .M24 = origin.Y
            .M31 = 0 : .M32 = 0 : .M33 = 1 : .M34 = 0
            .M41 = 0 : .M42 = 0 : .M43 = 0 : .M44 = 1
        End With

        For Each ch In text
            Dim x = Asc(ch)
            If ch = " " Then
                Width = 0       ' no strokes for space
            Else
                Strokes = FontData(Asc(ch)).Strokes
                Width = FontData(Asc(ch)).Width ' get width of this character
                TranslateMatrix.M14 = baseline      ' move stroke to start of letter
                For Each Stroke In Strokes
                    ply = Stroke.Clone          ' Create a polyline which is a deep copy of the data
                    ply.TransformBy(TranslateMatrix)    ' move to correct position
                    result.Add(ply)             ' add to result
                Next
            End If
            baseline += Width + LetterSpacing       ' leave gap between characters
        Next
        ' Now transform result to account for scale, rotation, origin and alignment
        Dim offset As Double
        Select Case alignment
            Case System.Windows.TextAlignment.Left
                offset = 0      ' Nothing to do. Is left aligned
            Case System.Windows.TextAlignment.Center
                offset = -(baseline - LetterSpacing) / 2              ' shift origin to middle of string
            Case System.Windows.TextAlignment.Right
                offset = -(baseline - LetterSpacing)               ' shift origin to right of string
            Case Else
                Throw New System.Exception($"Unrecognised alignment value: {alignment}")
        End Select

        ' Polylines are based at (0,0)
        For Each p In result
            If offset <> 0 Then
                With TranslateMatrix
                    .M14 = offset   ' translate origin to left, center or right of string
                    .M24 = 0
                End With
                p.TransformBy(TranslateMatrix)
            End If
            p.TransformBy(TransformMatrix)
        Next
        Return result
    End Function
    Private Sub LoadFonts()
        ' Load LibreCAD single stoke font

        ' The LibreCAD font format is described below
        '       Line 1 >= utf - 8 code + letter (same as QCAD)
        '       Line 2 & 3 >= sequence like Polyline vertex with ";" seperating vertex and "," separating x,y coords

        '       [0041] A
        '       0.0000,0.0000;3.0000,9.0000;6.0000,9.000
        '       1.0800,2.5500;4.7300,2.5500


        '       Line 2 >= sequence like Polyline vertex with ";" seperating vertex And
        '       "," separating x,y coords, if vertex Is prefixed with "A" the first field is a bulge
        '       [0066] f
        '       1.2873,0;1.2873,7.2945;3.4327,9.0000,A0.5590
        '       0.000000,6.0000,3.0000,6.0000

        '       A "C" preceding a UTF-8 code means copy these strokes.
        '       As below, a Q is made from an O, with one stroke added
        '       [0051] Q
        '       C004f
        '       6,0;4,2


        '       Font vertexes are loaded, and if required additional vertexes are added representing the bulge

        ' The font is 9 units high. Widths vary.

        Dim State As ReaderState = ReaderState.Idle
        Dim Strokes As List(Of Polyline2D) = Nothing
        Dim stroke As Polyline2D, LastVertex As Polyline2DVertex

        Dim FontPath = My.Settings.FontPath
        If Not System.IO.Directory.Exists(FontPath) Then
            MsgBox($"The font path {FontPath} cannot be found. Please use the parameters dialog to set one.", vbAbort + vbOK, "Font folder not found")
            Exit Sub
        Else
            Dim FontFile = My.Settings.FontFile
            If Not System.IO.File.Exists($"{FontPath}\{FontFile}") Then
                MsgBox($"The font file {FontFile} cannot be found. Please use the parameters dialog to set one.", vbAbort + vbOK, "Font file not found")
                Exit Sub
            Else
                FontData.Clear()        ' remove existing data
                Dim Key As Integer

                Dim fileReader As System.IO.StreamReader
                fileReader =
                My.Computer.FileSystem.OpenTextFileReader($"{FontPath}\{FontFile}")
                Dim LineNumber As Integer = 0
                While Not fileReader.EndOfStream        ' read all lines
                    Dim line = fileReader.ReadLine
                    LineNumber += 1
                    Select Case State
                        Case ReaderState.Idle
                            Dim options As RegexOptions = RegexOptions.IgnoreCase
                            Dim r As New Regex("^\[([0-9ABCDEF]+?)\]", options)    ' looking for UTF-8 code
                            Dim matches = r.Matches(line)
                            If matches.Count > 0 Then
                                Key = Convert.ToInt32(matches(0).Groups(1).Value, 16)
                                Strokes = New List(Of Polyline2D)    ' ready to store strokes
                                State = ReaderState.ReadVertexes     ' we have read the key
                            End If
                        Case ReaderState.ReadVertexes
                            If line = "" Then
                                FontData.Add(Key, New Glyph(Strokes))  ' end of stroke list. Add strokes to glyph, and glyph to dictionary
                                State = ReaderState.Idle    ' blank line. End of glyph
                            Else
                                Dim verts = Split(line, ";")    ' split line into vertexes
                                stroke = New Polyline2D
                                For Each Vertex In verts
                                    Dim Coords = Split(Vertex, ",")    ' split into X,Y.   Could be X,Y or X,Y,Bulge
                                    Select Case Coords.Length
                                        Case 1
                                            If Coords(0).StartsWith("C"c) Then
                                                ' It's a copy of another glyph + extras
                                                Coords(0) = Coords(0).Remove(0, 1)     ' remove the "C"
                                                Dim k = Convert.ToInt32(Coords(0), 16)
                                                For Each s In FontData(k).Strokes
                                                    Strokes.Add(s)   ' copy the character strokes
                                                Next
                                            Else
                                                Throw New System.Exception($"Illegal data on line {LineNumber}: {line}")
                                            End If
                                        Case 2
                                            LastVertex = New Polyline2DVertex(New Polyline2DVertex(CDbl(Coords(0)), CDbl(Coords(1))))
                                            stroke.Vertexes.Add(LastVertex)
                                        Case 3
                                            ' Vertex has a bulge. Apply bulge to vector and add
                                            Dim StartPoint = LastVertex.Position
                                            LastVertex = New Polyline2DVertex(New Polyline2DVertex(CDbl(Coords(0)), CDbl(Coords(1))))
                                            Dim Endpoint = LastVertex.Position
                                            Dim bulge As Double = CDbl(Coords(2).Remove(0, 1))   ' remove the "A"
                                            Dim arc = GenerateBulge(StartPoint, Endpoint, bulge)
                                            If arc.Vertexes(0).Position = Endpoint Then arc.Vertexes.Reverse()       ' the arc is reversed
                                            arc.Vertexes.RemoveAt(0)        ' the first vertex is the last of the previous line
                                            For Each v In arc.Vertexes
                                                stroke.Vertexes.Add(v)      ' add arc to polyline
                                            Next
                                        Case Else
                                            Throw New System.Exception($"Illegal number of coordinates ({Coords.Length} on line {LineNumber}: {line}")
                                    End Select
                                Next
                                Strokes.Add(stroke)
                            End If
                    End Select
                End While
                fileReader.Close()
                'TextBox1.AppendText($"{FontData.Count} glyphs read from {FontPath}\{FontFile}{vbCrLf}")
                'Dim MaxWidth = 0
                'For Each glyph In FontData
                '    Dim width = glyph.Value.Width
                '    MaxWidth = Math.Max(MaxWidth, width)
                '    Dim widstr = $"{width:f1}"
                '    If (width > 11 Or width = 0) Then widstr = $"*** {widstr} ***"
                '    TextBox1.AppendText($"{glyph.Key}(0x{glyph.Key:x})[{widstr}] = ")
                '    For Each s In glyph.Value.Strokes
                '        For Each v In s.Vertexes
                '            TextBox1.AppendText($"({v.Position.X:f1},{v.Position.Y:f1}) ")
                '        Next
                '    Next
                '    TextBox1.AppendText(vbCrLf)
                'Next
                'TextBox1.AppendText($"Maximum width is {MaxWidth:f1}{vbCrLf}")
            End If
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadFonts()
    End Sub
End Class

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
        Me.Strokes = strokes
    End Sub
End Class

Public Class IntPoint
    Public Property X As Integer
    Public Property Y As Integer

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