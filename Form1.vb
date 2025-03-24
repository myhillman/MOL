Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Media
Imports netDxf
Imports netDxf.Entities
Imports netDxf.Tables
Imports Windows.Win32.System
Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Windows.Ink
Imports GemBox.Spreadsheet

' Create a pseudo type to handle Leetro floating point numbers in a more natural way
' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]
''' <summary>
''' Represents a Leetro floating-point number and provides methods for encoding and decoding.
''' </summary>
Public Structure Float
    ''' <summary>
    ''' The Leetro floating-point value represented as an integer.
    ''' </summary>
    Public ReadOnly value As Integer

    ''' <summary>
    ''' Initializes a new instance of the <see cref="Float"/> structure from a double value.
    ''' </summary>
    ''' <param name="d">The double value to encode as a Leetro floating-point number.</param>
    Public Sub New(d As Double)
        value = Double2Float(d)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the <see cref="Float"/> structure from an integer value.
    ''' </summary>
    ''' <param name="leetro">The integer value representing a Leetro floating-point number.</param>
    Public Sub New(leetro As Integer)
        value = leetro
    End Sub

    ''' <summary>
    ''' Gets the double value equivalent of the Leetro floating-point number.
    ''' </summary>
    ''' <returns>The double value equivalent of the Leetro floating-point number.</returns>
    Public ReadOnly Property AsDouble() As Double
        Get
            Return Float2Double(value)
        End Get
    End Property

    ''' <summary>
    ''' Encodes a double value as a Leetro floating-point number.
    ''' </summary>
    ''' <param name="d">The double value to encode.</param>
    ''' <returns>A <see cref="Float"/> structure representing the encoded value.</returns>
    Public Shared Function Encode(d As Double) As Float
        Return New Float(d)
    End Function

    ''' <summary>
    ''' Encodes an integer value as a Leetro floating-point number.
    ''' </summary>
    ''' <param name="i">The integer value to encode.</param>
    ''' <returns>A <see cref="Float"/> structure representing the encoded value.</returns>
    Public Shared Function Encode(i As Integer) As Float
        Return New Float(CDbl(i))
    End Function

    ''' <summary>
    ''' Decodes a Leetro floating-point number to a double value.
    ''' </summary>
    ''' <param name="l">The <see cref="Float"/> structure to decode.</param>
    ''' <returns>The double value equivalent of the Leetro floating-point number.</returns>
    Public Shared Function Decode(l As Float) As Double
        Return l.AsDouble
    End Function

    ''' <summary>
    ''' Returns a string that represents the current object.
    ''' </summary>
    ''' <returns>A string that represents the current object.</returns>
    Public Overrides Function ToString() As String
        Return AsDouble().ToString()
    End Function
End Structure

''' <summary>
''' Represents the different types of blocks used in the application.
''' </summary>
Public Enum BLOCKTYPE
    ''' <summary>
    ''' The header block, containing metadata and configuration information.
    ''' </summary>
    HEADER = 0

    ''' <summary>
    ''' The configuration block, containing settings and parameters for the laser cutter.
    ''' </summary>
    CONFIG = 1

    ''' <summary>
    ''' The test block, used for testing and calibration purposes.
    ''' </summary>
    TEST = 2

    ''' <summary>
    ''' The cut block, containing instructions for cutting operations.
    ''' </summary>
    CUT = 3

    ''' <summary>
    ''' The draw block, containing instructions for drawing operations.
    ''' </summary>
    DRAW = 5
End Enum
''' <summary>
''' Represents the parameter count type for a command, indicating whether the number of parameters is fixed or variable.
''' </summary>
Public Enum ParameterCount
    ''' <summary>
    ''' The number of parameters for this command is fixed.
    ''' </summary>
    FIXED

    ''' <summary>
    ''' The number of parameters for this command is variable.
    ''' </summary>
    VARIABLE
End Enum

''' <summary>
''' Represents the different laser modes used in the application.
''' </summary>
Friend Enum LSRMode_Enum
    ''' <summary>
    ''' Laser is turned off.
    ''' </summary>
    Off = 0

    ''' <summary>
    ''' Laser is in cut mode, used for cutting through materials.
    ''' </summary>
    Cut = 1

    ''' <summary>
    ''' Laser is in engrave mode, used for normal engraving operations.
    ''' </summary>
    Engrave = 2

    ''' <summary>
    ''' Laser is in grade engrave mode, used for beveling the corners of engravings.
    ''' </summary>
    GradeEngrave = 3

    ''' <summary>
    ''' Laser is in hole mode, used for creating perforations.
    ''' </summary>
    Hole = 4

    ''' <summary>
    ''' Laser is in pen cut mode, rumored to crash LaserCut.
    ''' </summary>
    PenCut = 5
End Enum

''' <summary>
''' Represents the different axes used in the application.
''' </summary>
Friend Enum Axis_Enum
    ''' <summary>
    ''' The X axis.
    ''' </summary>
    X = 4

    ''' <summary>
    ''' The Y axis.
    ''' </summary>
    Y = 3
End Enum

''' <summary>
''' Represents the acceleration and deceleration states for the laser cutter.
''' </summary>
Friend Enum Acceleration_Enum
    ''' <summary>
    ''' The laser cutter is accelerating.
    ''' </summary>
    Accelerate = 1

    ''' <summary>
    ''' The laser cutter is decelerating.
    ''' </summary>
    Decelerate = 2
End Enum

''' <summary>
''' Represents the on and off states for the laser.
''' </summary>
Friend Enum OnOff_Enum
    ''' <summary>
    ''' The laser is turned off.
    ''' </summary>
    Off = 0

    ''' <summary>
    ''' The laser is turned on.
    ''' </summary>
    [On] = 1
End Enum

''' <summary>
''' A structure to hold two 16-bit values representing the number of steps the laser is on and off.
''' </summary>
''' <remarks>
''' The structure allows access to the combined 32-bit value as well as the individual 16-bit values.
''' </remarks>
<StructLayout(LayoutKind.Explicit)> Public Structure OnOffSteps
    ''' <summary>
    ''' The combined 32-bit value representing both on and off steps.
    ''' </summary>
    <FieldOffset(0)>
    Public Steps As Integer

    ''' <summary>
    ''' The number of steps the laser is on.
    ''' </summary>
    <FieldOffset(0)>
    Public OnSteps As UShort

    ''' <summary>
    ''' The number of steps the laser is off.
    ''' </summary>
    <FieldOffset(2)>
    Public OffSteps As UShort
End Structure

''' <summary>
''' The main form class for the WPF application, responsible for handling laser cutter operations, 
''' including reading and writing MOL files, rendering DXF files, and managing laser cutter settings and commands.
''' </summary>
Public Class Form1
    Private Const BLOCK_SIZE = 512          ' bytes. MOL is divided into blocks or chunks
    Private Const MOLTemplate = "16SQ.MOL"  ' template on which to base created MOL files
    Private debugflag As Boolean = True
    Private hexdump As Boolean = False
    Private dlg = New OpenFileDialog
    Private stream As FileStream
    Private reader As BinaryReader
    Private writer As BinaryWriter

    Private startposn As IntPoint
    Private delta As IntPoint       ' a delta from position
    Private ZERO = New IntPoint(0, 0)   ' a point with value (0,0)
    Private Const NumColors = 255
    Private colors(NumColors) As AciColor   ' array of colors
    Private ColorIndex As Integer = 1       ' index into colors array
    Private CommandUsage As New Dictionary(Of Integer, Integer)(50)
    Private dxf As New DxfDocument()
    Private dxfBlock As netDxf.Blocks.Block     ' current block we are writing to
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

    ' Header information
    Private FileSize As Integer         ' length of file in bytes
    Private MotionBlocks As Integer
    Private TopRight As IntPoint
    Private BottomLeft As IntPoint
    Private Version As String
    Private ConfigChunk As Integer
    Private TestChunk As Integer
    Private CutChunk As Integer
    Private DrawChunks As New List(Of Integer)(10)     ' list of draw chunks

    ' Variables that reflect the state of the laser cutter
    ' These parameters from Machine Options - Worktable dialog
    Private PWMFrequency As Integer = 20000  ' PWM frequency
    Private StartSpeed As Double = 5        ' Initial speed when moving
    Private QuickSpeed As Double = 300      ' Max speed when moving
    Private WorkAcc As Double = 500         ' Acceleration whilst cutting
    Private SpaceAcc As Double = 1200       ' Acceleration whilst just moving
    Private ClipBoxPower As Integer = 40    ' power setting for Cut box
    Private ClipBoxSpeed As Integer = 25    ' speed setting for Cut box
    Private LaserMode As LSRMode_Enum       ' current laser beam mode
    Private block As BLOCKTYPE              ' current block type

    Private position As IntPoint            ' current laser head position in steps
    Public xScale As Double = 125.9842519685039    ' X axis steps/mm.   From the Worktable config dialog [Pulse Unit] parameter
    Public yScale As Double = 125.9842519685039    ' Y axis steps/mm
    Public zScale As Double = 125.9842519685039    ' Z axis steps/mm
    Private AccelLength As Double      ' Acceleration and Deceleration distance
    Private CurrentSubr As Integer = 0                ' current subroutine we are in
    Private StartPosns As New Dictionary(Of Integer, (Absolute As Boolean, position As IntPoint))      ' start position for drawing by subroutine #
    Private SubrAddrs As New Dictionary(Of Integer, Integer)       ' list of subroutine numbers and their start address
    Private EmptyVector As New Vector2(0, 0)
    Private ENGLSRsteps As New List(Of OnOffSteps)
    Private FirstMove As Boolean = True           ' true if next MVREL will be the first
    Private EngPower As Integer             ' Engrave power (%)
    Private EngSpeed As Double              ' Engrave speed (mm/s)
    Private MVRELInConfig As Integer      ' number of MVREL encountered so far in CONFIG section
    Private MCBLKCounter As Integer = 0     ' counter when inside MCBLK

    Private MaxLength As Double = 0

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
        ''' <summary>
        ''' Gets or sets the name of the parameter.
        ''' </summary>
        ''' <value>The name of the parameter.</value>
        Public Property Name As String

        ''' <summary>
        ''' Gets or sets the description of the parameter.
        ''' </summary>
        ''' <value>The description of the parameter.</value>
        Public Property Description As String

        ''' <summary>
        ''' Gets or sets the type of the parameter.
        ''' </summary>
        ''' <value>The type of the parameter.</value>
        Public Property Typ As Type

        ''' <summary>
        ''' Gets or sets the scale applied to the parameter value.
        ''' </summary>
        ''' <value>The scale applied to the parameter value.</value>
        Public Property Scale As Double

        ''' <summary>
        ''' Gets or sets the units of the parameter value, e.g., mm.
        ''' </summary>
        ''' <value>The units of the parameter value.</value>
        Public Property Units As String

        ''' <summary>
        ''' Initializes a new instance of the <see cref="Parameter"/> class with the specified name, description, type, scale, and units.
        ''' </summary>
        ''' <param name="name">The name of the parameter.</param>
        ''' <param name="description">The description of the parameter.</param>
        ''' <param name="type">The type of the parameter.</param>
        ''' <param name="scale">The scale applied to the parameter value. Default is 1.</param>
        ''' <param name="units">The units of the parameter value. Default is an empty string.</param>
        Public Sub New(name As String, description As String, type As Type, Optional ByVal scale As Double = 1, Optional ByVal units As String = "")
            Me.Name = name
            Me.Description = description
            Me.Typ = type
            Me.Scale = scale
            Me.Units = units
        End Sub
    End Class

    Public Class MOLcmd
        ''' <summary>
        ''' Gets or sets the mnemonic of the command.
        ''' </summary>
        ''' <value>The mnemonic of the command.</value>
        Public Property Mnemonic As String

        ''' <summary>
        ''' Gets or sets the description of the command.
        ''' </summary>
        ''' <value>The description of the command.</value>
        Public Property Description As String

        ''' <summary>
        ''' Gets or sets the parameter count type of the command (FIXED or VARIABLE).
        ''' </summary>
        ''' <value>The parameter count type of the command.</value>
        Public Property ParameterType As ParameterCount = ParameterCount.FIXED

        ''' <summary>
        ''' Gets or sets the list of parameters for the command.
        ''' </summary>
        ''' <value>The list of parameters for the command.</value>
        Public Property Parameters As New List(Of Parameter)

        ''' <summary>
        ''' Initializes a new instance of the <see cref="MOLcmd"/> class with the specified mnemonic and description.
        ''' </summary>
        ''' <param name="mnemonic">The mnemonic of the command.</param>
        ''' <param name="description">The description of the command.</param>
        Public Sub New(mnemonic As String, description As String)
            Me.Mnemonic = mnemonic
            Me.Description = description
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="MOLcmd"/> class with the specified mnemonic, description, and parameter count type.
        ''' </summary>
        ''' <param name="mnemonic">The mnemonic of the command.</param>
        ''' <param name="description">The description of the command.</param>
        ''' <param name="pc">The parameter count type of the command.</param>
        Public Sub New(mnemonic As String, description As String, pc As ParameterCount)
            Me.Mnemonic = mnemonic
            Me.ParameterType = pc
            Me.Description = description
        End Sub

        ''' <summary>
        ''' Initializes a new instance of the <see cref="MOLcmd"/> class with the specified mnemonic, description, parameter count type, and parameters.
        ''' </summary>
        ''' <param name="mnemonic">The mnemonic of the command.</param>
        ''' <param name="description">The description of the command.</param>
        ''' <param name="pc">The parameter count type of the command.</param>
        ''' <param name="p">The list of parameters for the command.</param>
        Public Sub New(mnemonic As String, description As String, pc As ParameterCount, p As List(Of Parameter))
            Me.Mnemonic = mnemonic
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
    Private Const MOL_HOLE = &H1000846
    Private Const MOL_GOSUB = &H1500048
    Private Const MOL_GOSUBb = &H80500008
    Private Const MOL_GOSUB3 = &H3500048
    Private Const MOL_GOSUBn = &H80500048
    Private Const MOL_X5_FIRST = &H200548
    Private Const MOL_X6_LAST = &H200648
    Private Const MOL_ACCELERATION = &H1004601
    Private Const MOL_BLWR = &H1004A41
    Private Const MOL_BLWRa = &H1004B41
    Private Const MOL_BLWRb = &H1000B46
    Private Const MOL_SEGMENT = &H500008

    Private Const MOL_ENGPWR = &H1000746
    Private Const MOL_ENGPWR1 = &H2000746
    Private Const MOL_ENGSPD = &H2014341
    Private Const MOL_ENGSPD1 = &H4010141
    Private Const MOL_ENGACD = &H1000346
    Private Const MOL_ENGMVY = &H2010040
    Private Const MOL_ENGMVX = &H2014040
    Private Const MOL_GRDADJ = &H80000B46
    Private Const MOL_SCALE = &H3000E46
    Private Const MOL_PWRSPD5 = &H5000E46
    Private Const MOL_PWRSPD7 = &H7000E46
    Private Const MOL_ENGLSR = &H80000146
    Private Const MOL_END = 0

    Private Const MOL_UNKNOWN07 = &H3046040      ' unknown command refered to in London Hackspace documents
    Private Const MOL_UNKNOWN09 = &H326040      ' unknown command refered to in London Hackspace documents

    ''' <summary>
    ''' Dictionary of all MOL commands with their corresponding details.
    ''' </summary>
    ''' <remarks>
    ''' Each command is represented by a unique integer key and a MOLcmd object containing the mnemonic, description, parameter count, and parameters.
    ''' </remarks>
    Private Commands As New SortedDictionary(Of Integer, MOLcmd) From {
        {MOL_MVREL, New MOLcmd("MVREL", "Move the cutter by dx,dy", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "Equivalent to 0x0304, i.e. both x and y", GetType(Int32))},
                            {New Parameter("dx", "delta to move in X direction", GetType(Int32), 1 / xScale, "mm")},
                            {New Parameter("dy", "delta to move in Y direction", GetType(Int32), 1 / yScale, "mm")}
                            }
                           )},
        {MOL_START, New MOLcmd("START", "A start location for a following subroutine", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "Equivalent to 0x0304, i.e. both x and y", GetType(Int32))},
                            {New Parameter("x", "", GetType(Int32), 1 / xScale, "mm")},
                            {New Parameter("y", "", GetType(Int32), 1 / yScale, "mm")}
                            }
                           )},
        {MOL_SCALE, New MOLcmd("SCALE", "The scale used on all 3 axis", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("x scale", "the X scale", GetType(Float),, "steps/mm")},
                            {New Parameter("y scale", "the Y scale", GetType(Float),, "steps/mm")},
                            {New Parameter("z scale", "the Z scale", GetType(Float),, "steps/mm")}
                            }
                           )},
        {MOL_ORIGIN, New MOLcmd("ORIGIN", "", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "", GetType(Int32))},
                            {New Parameter("x", "", GetType(Int32))},
                            {New Parameter("y", "", GetType(Int32))}
                            }
                           )},
        {MOTION_CMD_COUNT, New MOLcmd("MOTION_COMMAND_COUNT", "", ParameterCount.FIXED, New List(Of Parameter) From {
                        {New Parameter("count", "", GetType(Int32))}}
                        )},
        {MOL_BEGSUB, New MOLcmd("BEGSUB", "Begin of subroutine", ParameterCount.FIXED, New List(Of Parameter) From {
                        {New Parameter("n", "Subroutine number", GetType(Int32))}}
                        )},
        {MOL_BEGSUBa, New MOLcmd("BEGSUBa", "Begin of subroutine", ParameterCount.FIXED, New List(Of Parameter) From {
                        {New Parameter("n", "Subroutine number", GetType(Int32))}}
                        )},
        {MOL_ENDSUB, New MOLcmd("ENDSUB", "End of subroutine", ParameterCount.FIXED, New List(Of Parameter) From {
                        {New Parameter("n", "Subroutine number", GetType(Int32))}}
                        )},
        {MOL_MCBLK, New MOLcmd("MCBLK", "Motion Control Block", ParameterCount.FIXED, New List(Of Parameter) From {
                        {New Parameter("Size", "Number of words (limited to 510)", GetType(Int32),, "Words")}}
                        )},
        {MOL_MOTION, New MOLcmd("MOTION", "Set min/max speed & acceleration", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("Initial speed", "", GetType(Float), 1 / xScale, "mm/s")},
                            {New Parameter("Max speed", "", GetType(Float), 1 / xScale, "mm/s")},
                            {New Parameter("Acceleration", "", GetType(Float), 1 / xScale, "mm/s²")}
                            }
                           )},
        {MOL_SETSPD, New MOLcmd("SETSPD", "Set min/max speed & acceleration", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("Initial speed", "", GetType(Float), 1 / xScale, "mm/s")},
                            {New Parameter("Max speed", "", GetType(Float), 1 / xScale, "mm/s")},
                            {New Parameter("Acceleration", "", GetType(Float), 1 / xScale, "mm/s²")}
                            }
                           )},
        {MOL_PWRSPD5, New MOLcmd("PWRSPD5", "Set Power & Speed", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("Corner PWR", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Max PWR", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Start speed", "", GetType(Float), 1 / xScale, "mm/s")},
                            {New Parameter("Max speed", "", GetType(Float), 1 / xScale, "mm/s")},
                            {New Parameter("Unknown", "", GetType(Float))}
                          }
                         )},
        {MOL_PWRSPD7, New MOLcmd("PWRSPD7", "Set Power & Speed (2 heads)", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("Corner PWR", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Max PWR", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Laser 2 Corner PWR", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Laser 2 Max PWR", "", GetType(Int32), 0.01, "%")},
                            {New Parameter("Start speed", "", GetType(Float), 1 / xScale, "mm/s")},
                            {New Parameter("Max speed", "", GetType(Float), 1 / xScale, "mm/s")},
                            {New Parameter("Unknown", "", GetType(Float))}
                            }
                         )},
        {MOL_LASER, New MOLcmd("LASER", "Laser mode control", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("Mode", "LASER mode", GetType(LSRMode_Enum))}})},
        {MOL_LASER1, New MOLcmd("LASER1", "LASER mode", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("Mode", "", GetType(LSRMode_Enum))}})},
        {MOL_LASER2, New MOLcmd("LASER2", "LASER mode", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("Mode", "", GetType(LSRMode_Enum))}})},
        {MOL_HOLE, New MOLcmd("HOLE", "Burn hole", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("Radiation Duration", "", GetType(Int32), 1 / PWMFrequency, "s")}})},
        {MOL_GOSUB, New MOLcmd("GOSUB", "Call a subroutine with no parameters", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "Number of the subroutine", GetType(Int32))}
                            }
                           )},
        {MOL_GOSUBb, New MOLcmd("GOSUBb", "Call subroutine with 3 parameters", ParameterCount.VARIABLE, New List(Of Parameter) From {
                        {New Parameter("n", "Subroutine number", GetType(Int32))},
                        {New Parameter("x", "x parameter", GetType(Float))},
                        {New Parameter("y", "y parameter", GetType(Float))}}
                        )},
        {MOL_GOSUB3, New MOLcmd("GOSUB3", "Can have 1 or 3 parameters", ParameterCount.VARIABLE, New List(Of Parameter) From {
                            {New Parameter("n", "Subroutine number", GetType(Int32))},
                            {New Parameter("x", "x parameter", GetType(Float))},
                            {New Parameter("y", "y parameter", GetType(Float))}
                            }
                           )},
         {MOL_GOSUBn, New MOLcmd("GOSUBn", "Can have 1 or 3 parameters", ParameterCount.VARIABLE, New List(Of Parameter) From {
                            {New Parameter("n", "Subroutine number", GetType(Int32))},
                            {New Parameter("x", "x parameter", GetType(Float))},
                            {New Parameter("y", "y parameter", GetType(Float))}
                            }
                           )},
        {MOL_X5_FIRST, New MOLcmd("X5_FIRST", "Always first in CONFIG block", ParameterCount.FIXED)},
        {MOL_X6_LAST, New MOLcmd("X6_LAST", "Always last in CONFIG block", ParameterCount.FIXED)},
        {MOL_SEGMENT, New MOLcmd("SEGMENT", "", ParameterCount.FIXED, New List(Of Parameter) From {
                            {New Parameter("n", "", GetType(Int32))}
                            }
                           )},
        {MOL_ACCELERATION, New MOLcmd("ACCELERATION", "Accelerate or Decelerate", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("Acceleration", "", GetType(Acceleration_Enum))}})},
        {MOL_BLWR, New MOLcmd("BLWR", "Blower control", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("On/Off", "", GetType(OnOff_Enum))}})},
        {MOL_BLWRa, New MOLcmd("BLWRa", "Blower control", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("On/Off", "", GetType(OnOff_Enum))}})},
        {MOL_BLWRb, New MOLcmd("BLWRb", "Blower control", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("On/Off", "", GetType(OnOff_Enum))}})},
        {MOL_ENGPWR, New MOLcmd("ENGPWR", "Engrave Power", ParameterCount.FIXED, New List(Of Parameter) From {{New Parameter("Engrave power", "", GetType(Integer), 0.01, "%")}})},
        {MOL_ENGPWR1, New MOLcmd("ENGPWR1", "", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("Power", "", GetType(Integer), 0.01, "%")},
                {New Parameter("??", "", GetType(Integer))}
                }
              )},
        {MOL_ENGSPD, New MOLcmd("ENGSPD", "Engrave speed", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("Axis", "", GetType(Axis_Enum))},
                {New Parameter("Speed", "", GetType(Float), 1 / xScale, "mm/s")}
                }
              )},
        {MOL_ENGSPD1, New MOLcmd("ENGSPD1", "", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("Axis", "", GetType(Axis_Enum))},
                {New Parameter("??", "", GetType(Float))},
                {New Parameter("Speed", "", GetType(Float), 1 / xScale, "mm/s")},
                {New Parameter("??", "", GetType(Float))}
                }
              )},
        {MOL_ENGMVX, New MOLcmd("ENGMVX", "Engrave one line in the X direction", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("Axis", "", GetType(Axis_Enum))}, {New Parameter("dx", "", GetType(Integer), 1 / xScale, "mm")}}
            )},
        {MOL_ENGMVY, New MOLcmd("ENGMVY", "Engraving - Move in the Y direction", ParameterCount.FIXED, New List(Of Parameter) From {
                {New Parameter("Axis", "", GetType(Axis_Enum))},
                {New Parameter("dy", "Distance to move", GetType(Integer), 1 / yScale, "mm")}}
                )},
        {MOL_GRDADJ, New MOLcmd("GRDADJ", "Grade Adjustment table", ParameterCount.VARIABLE, New List(Of Parameter) From {
                {New Parameter("??", "", GetType(Int32))},
                {New Parameter("ramp", "Grade engrave power steps", GetType(List(Of Integer)), 0.01, "W")}}
        )},
        {MOL_ENGACD, New MOLcmd("ENGACD", "Engraving (Ac)(De)celeration distance", ParameterCount.FIXED, New List(Of Parameter) From {
                    {New Parameter("x", "distance", GetType(Int32), 1 / xScale, "mm")}}
                    )},
        {MOL_ENGLSR, New MOLcmd("ENGLSR", "Engraving On/Off pattern", ParameterCount.VARIABLE, New List(Of Parameter) From {
                            {New Parameter("List of steps", "List of On/Off patterns", GetType(List(Of OnOffSteps)), 1 / xScale, "mm")}}
                           )},
        {MOL_END, New MOLcmd("END", "End of code")}
        }

    ''' <summary>
    ''' Reads a 4-byte integer from the input stream.
    ''' </summary>
    ''' <returns>The 4-byte integer read from the input stream.</returns>
    Public Function GetInt() As Integer
        ' read 4 byte integer from input stream
        Return reader.ReadInt32
    End Function

    ''' <summary>
    ''' Reads a 4-byte unsigned integer from the input stream.
    ''' </summary>
    ''' <returns>The 4-byte unsigned integer read from the input stream.</returns>
    Public Function GetUInt() As UInteger
        ' read 4 byte uinteger from input stream
        Return reader.ReadUInt32
    End Function

    ''' <summary>
    ''' Reads a 4-byte integer from the input stream.
    ''' </summary>
    ''' <returns>The 4-byte integer read from the input stream.</returns>
    Public Function GetInt(n As Integer) As Integer
        reader.BaseStream.Seek(n, SeekOrigin.Begin)    ' reposition to offset n
        Return GetInt()
    End Function

    ''' <summary>
    ''' Reads a 4-byte float from the current offset in the input stream.
    ''' </summary>
    ''' <returns>The 4-byte float read from the input stream.</returns>
    Public Function GetFloat() As Double
        ' read 4 byte float from current offset
        Return Float2Double(GetUInt())
    End Function

    Public Function GetFloat(n As Integer) As Double
        ' read float from specified offset
        reader.BaseStream.Seek(n, SeekOrigin.Begin)    ' reposition to offset n
        Return GetFloat()
    End Function

    ''' <summary>
    ''' Writes a float to the current offset in the output stream.
    ''' </summary>
    ''' <param name="f">The float value to write.</param>
    Public Sub PutFloat(f As Double)
        ' write float at current offset
        If UseMCBLK Then
            MCBLK.Add(Double2Float(f))
        Else
            writer.Write(Double2Float(f))
        End If
    End Sub
    ''' <summary>
    ''' Writes an integer to the current offset in the output stream.
    ''' </summary>
    ''' <param name="n">The integer value to write.</param>
    Public Sub PutInt(n As Integer)
        ' write n at current offset
        If UseMCBLK Then
            MCBLK.Add(n)
        Else
            writer.Write(n)
        End If
    End Sub
    ''' <summary>
    ''' Writes an integer to the specified offset in the output stream.
    ''' </summary>
    ''' <param name="n">The integer value to write.</param>
    ''' <param name="addr">The offset to write the integer to.</param>
    Public Sub PutInt(n As Integer, addr As Integer)
        ' write n at specifed offset
        If UseMCBLK Then
            Throw New System.Exception($"You can't write explicitly to address {addr:x} as an MCBLK in in operation")
        Else
            writer.BaseStream.Seek(addr, SeekOrigin.Begin)     ' go to offset
            writer.Write(n)                     ' write data
        End If
    End Sub
    ''' <summary>
    ''' Reads the header information from the input stream and initializes the corresponding fields.
    ''' </summary>
    Private Sub GetHeader()
        Dim chunk As Integer

        ' Check if the reader stream is open and can be read
        If Not reader.BaseStream.CanRead Then
            Throw New System.Exception("reader is not open")
        End If

        ' Set the stream position to the start of the file
        reader.BaseStream.Position = 0

        ' Read the file size in bytes
        FileSize = GetInt(0)

        ' Read the number of motion blocks
        MotionBlocks = GetInt()

        ' Read the version information
        Dim v() As Byte = reader.ReadBytes(4)
        Version = $"{v(3)}.{v(2)}.{v(1)}.{v(0)}"

        ' Read the top-right and bottom-left coordinates
        TopRight = New IntPoint(GetInt(&H18), GetInt())
        BottomLeft = New IntPoint(GetInt(&H20), GetInt())

        ' Read the chunk addresses for config, test, and cut sections
        ConfigChunk = GetInt(&H70)
        If ConfigChunk = 0 Then Throw New Exception("Config chunk is 0")
        TestChunk = GetInt()
        If TestChunk = 0 Then Throw New Exception("Test chunk is 0")
        CutChunk = GetInt()
        If CutChunk = 0 Then Throw New Exception("Cut chunk is 0")

        ' Read the draw chunks and add them to the DrawChunks list
        DrawChunks.Clear()
        chunk = GetInt()        ' first draw chunk
        While chunk <> 0
            DrawChunks.Add(chunk)
            chunk = GetInt()
        End While
        If DrawChunks(0) = 0 Then Throw New Exception("Draw chunk is 0")
    End Sub

    ''' <summary>
    ''' Displays header information from the input stream.
    ''' </summary>
    ''' <param name="textfile">The TextWriter to write the header information to.</param>
    Public Sub DisplayHeader(textfile As TextWriter)
        ' Display header info

        GetHeader()         ' get header info
        textfile.WriteLine($"HEADER 0")
        textfile.WriteLine($"Size of file: 0x{FileSize:x} bytes")
        textfile.WriteLine($"Motion blocks: {MotionBlocks}")
        textfile.WriteLine($"Version: {Version}")
        textfile.WriteLine($"Origin: ({TopRight.X},{TopRight.Y}) ({TopRight.X / xScale:f1},{TopRight.Y / yScale:f1})mm")
        textfile.WriteLine($"Bottom Left: ({BottomLeft.X},{BottomLeft.Y}) ({BottomLeft.X / xScale:f1},{BottomLeft.Y / yScale:f1})mm")
        textfile.WriteLine($"Config chunk: {ConfigChunk}")
        textfile.WriteLine($"Test chunk: {TestChunk}")
        textfile.WriteLine($"Cut chunk: {CutChunk}")
        textfile.WriteLine($"Draw chunks: {String.Join(",", DrawChunks.ToArray)}")
        textfile.WriteLine()
    End Sub
    ''' <summary>
    ''' Decodes a stream of commands from the input stream.
    ''' </summary>
    ''' <param name="writer">The StreamWriter to write the decoded commands to.</param>
    ''' <param name="StartAddress">The start address of the commands to decode.</param>
    Public Sub DecodeStream(writer As System.IO.StreamWriter, StartAddress As Integer)
        ' Decode a stream of commands
        ' StartAddress - start of commands
        ' Set start position for this block
        block = GetBlockType(StartAddress)
        Dim cmd = GetInt(StartAddress)
        If cmd = MOL_BEGSUB Or cmd = MOL_BEGSUBa Then
            Dim subr As Integer = GetInt(StartAddress + 4)      ' extract subroutine number
            If StartPosns.ContainsKey(subr) Then
                If StartPosns(subr).Absolute Then position = StartPosns(subr).position Else position += StartPosns(subr).position
            Else position = New IntPoint(0, 0)
            End If
        End If
        writer.WriteLine()
        writer.WriteLine($"Decode of commands at 0x{StartAddress:x} in block {block}")
        writer.WriteLine()
        ' position = StartPosns(block)
        reader.BaseStream.Seek(StartAddress, SeekOrigin.Begin)      ' Start of block
        'Try
        Do
            Loop Until Not DecodeCmd(writer)
        'Catch ex As Exception
        '    Throw New Exception($"DecodeStream failed at addr {reader.BaseStream.Position:x8}{vbCrLf}{ex.Message}")
        'End Try
    End Sub
    ''' <summary>
    ''' Decodes a command from the current position in the input stream.
    ''' </summary>
    ''' <param name="writer">The StreamWriter to write the decoded command to.</param>
    ''' <returns>True if more commands follow; otherwise, false.</returns>
    Public Function DecodeCmd(writer As System.IO.StreamWriter) As Boolean
        ' Display a decoded version of the current command
        ' Lookup command in known commands table
        ' returns false if at end of stream

        Dim value As MOLcmd = Nothing, cmd As Integer, cmd_len As Integer, OneStep As OnOffSteps, cmdBegin As Integer, p As Parameter, n As Integer
        Dim i1 As Integer, i2 As Integer, i3 As Integer, par As Integer

        Try
            cmdBegin = reader.BaseStream.Position         ' remember start of command
            writer.Write($"0x{cmdBegin:x}: ")
            cmd = GetInt()                  ' get command
            If Not CommandUsage.TryAdd(cmd, 1) Then CommandUsage(cmd) += 1      ' count commands used
            If cmd = MOL_END Then
                writer.WriteLine("END")
                Return False
            End If      ' We are done
            cmd_len = cmd >> 24 And &HFF
            If cmd_len = &H80 Then
                cmd_len = GetInt() And &H1FF
            End If
            Select Case cmd
                Case MOL_MCBLK
                    MCBLKCounter = cmd_len                  ' size of MCBLK in words
                    cmd_len = 1        ' allow contents of MCBLK to decode
                    reader.BaseStream.Position = cmdBegin + 4     ' backup so we can read length as parameter
            End Select

            If MCBLKCounter > 0 And cmd <> MOL_MCBLK Then writer.Write(" >")     ' prefix MCBLK with indicator
            If Commands.TryGetValue(cmd, value) Then
                If value.ParameterType = ParameterCount.FIXED And cmd_len <> value.Parameters.Count And cmd <> MOL_MCBLK Then
                    Throw New Exception($"Command {value.Mnemonic}: length is {cmd_len}, but data table says {value.Parameters.Count}")
                End If

                writer.Write($" {value.Mnemonic}")

                For n = 1 To cmd_len  ' number of parameters. Some commands have a variable number of parameters
                    p = value.Parameters(n - 1)     ' get the parameter
                    writer.Write($" {p.Name}=")
                    Select Case p.Typ
                        Case GetType(Boolean) : writer.Write(CType(GetInt(), Boolean))

                        Case GetType(Int32) : If p.Scale = 1.0 Then writer.Write($"{GetInt()} {p.Units}") Else writer.Write($"{GetInt() * p.Scale:f1} {p.Units}")

                        Case GetType(Float)
                            Dim l As New Float(GetInt())
                            Dim val = l.AsDouble      ' get float value
                            If p.Scale = 1.0 Then
                                writer.Write($"{val:f1} {p.Units}")
                            Else
                                writer.Write($"{val * p.Scale:f1} {p.Units}")
                            End If

                        Case GetType(OnOff_Enum)
                            par = GetInt()
                            Dim val = par And 1         ' only bit0 encodes on/off
                            writer.Write($"{DirectCast(System.Enum.Parse(GetType(OnOff_Enum), val), OnOff_Enum)} ({par})")

                        Case GetType(LSRMode_Enum)
                            par = GetInt()
                            If Not IsValidEnumValue(Of LSRMode_Enum)(par) Then Throw New Exception($"{par} is not a member of enum {GetType(LSRMode_Enum)}")
                            writer.Write($"{DirectCast(System.Enum.Parse(GetType(LSRMode_Enum), par), LSRMode_Enum)} ({par})")

                        Case GetType(Acceleration_Enum)
                            par = GetInt()
                            If Not IsValidEnumValue(Of Acceleration_Enum)(par) Then Throw New Exception($"{par} is not a member of enum {GetType(Acceleration_Enum)}")
                            writer.Write($"{DirectCast(System.Enum.Parse(GetType(Acceleration_Enum), par), Acceleration_Enum)} ({par})")

                        Case GetType(Axis_Enum)
                            par = GetInt()
                            If Not IsValidEnumValue(Of Axis_Enum)(par) Then Throw New Exception($"{par} is not a member of enum {GetType(Axis_Enum)}")
                            writer.Write($"{DirectCast(System.Enum.Parse(GetType(Axis_Enum), par), Axis_Enum)} ({par})")

                        Case GetType(List(Of OnOffSteps))    ' a list of On/Off steps
                            writer.Write($"List of {cmd_len} On/Off steps ")
                            For i1 = 1 To cmd_len    ' one structure for each word
                                OneStep.Steps = GetInt()     ' get 32 bit word
                                writer.Write($" {OneStep.OnSteps * p.Scale:f1}/{OneStep.OffSteps * p.Scale:f1} {p.Units}")
                            Next
                            Exit For ' all parameters have been consumed

                        Case GetType(OnOffSteps)    ' a  On/Off steps
                            OneStep.Steps = GetInt()     ' get 32 bit word
                            writer.Write($" One On/Off step {OneStep.OnSteps * p.Scale:f1}/{OneStep.OffSteps * p.Scale:f1} {p.Units}")

                        Case GetType(List(Of Integer))    ' a list of integer
                            writer.Write($"List of {cmd_len - 1} Power levels ")
                            For i2 = 2 To cmd_len    ' one structure for each word
                                writer.Write($" {GetInt() * p.Scale:f1} {p.Units}")
                            Next
                            Exit For ' all parameters have been consumed

                        Case Else
                            Throw New System.Exception($"{value.Mnemonic}: Unrecognised parameter type of {p.Typ}")
                    End Select
                Next
            Else
                ' UNKNOWN command. Just show parameters
                writer.Write($" Unknown: 0x{cmd:x8} Params {cmd_len}: ")
                For i3 = 1 To cmd_len
                    n = GetInt()
                    writer.Write($" 0x{n:x8}")
                    If n < 0 Or n > 500 Then writer.Write($" ({Float2Double(n)}f)")
                Next
            End If
            writer.WriteLine() : writer.Flush()
            'reader.BaseStream.Seek(cmdBegin + cmd_len * 4 + 4,  SeekOrigin.Begin)       ' move to next command
            If cmd <> MOL_MCBLK Then MCBLKCounter -= (cmd_len + 1)
            Return True         ' more commands follow
        Catch ex As Exception
            MsgBox($"cmd={cmd:x8} @0x{reader.BaseStream.Position:x8}", vbCritical + vbOK, "Error")
            Throw New Exception(ex.Message)
        End Try
    End Function
    ''' <summary>
    ''' Executes a stream of commands, rendering in DXF as we go.
    ''' </summary>
    ''' <param name="StartAddress">The start address of the commands to execute.</param>
    ''' <param name="dxf">The DXF document to render the commands in.</param>
    ''' <param name="DefaultLayer">The default layer to render the commands in.</param>
    Public Sub ExecuteStream(StartAddress As Integer, dxf As DxfDocument, DefaultLayer As Layer)
        Dim AddrTrace As New List(Of Integer)
        Dim layer = DefaultLayer
        ' Set start position for this block
        ' Translate start address to subroutine
        ' Look in start address

        LaserMode = LSRMode_Enum.Off
        Dim addr = StartAddress
        Do
            'AddrTrace.Add(addr)
            addr = ExecuteCmd(addr, dxf, layer)       ' execute command at addr
        Loop Until addr = 0         ' loop until an END statement
    End Sub
    ''' <summary>
    ''' Executes a command at the specified address and returns the address of the next instruction.
    ''' </summary>
    ''' <param name="addr">The address of the command to execute.</param>
    ''' <param name="dxf">The DXF document to render the command in.</param>
    ''' <param name="Layer">The layer to render the command in.</param>
    ''' <returns>The address of the next instruction.</returns>
    Public Function ExecuteCmd(ByVal addr As Integer, dxf As DxfDocument, ByRef Layer As Layer) As Integer
        ' Execute a command at addr. Return addr as start of next instruction
        Dim cmd As Integer, nWords As Integer, Steps As OnOffSteps
        block = GetBlockType(addr)
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

            Case MOL_LASER, MOL_LASER1, MOL_LASER2    ' switch laser
                ' Laser is changing state. Output any pending motion
                If motion.Vertexes.Count > 0 Then
                    ' Display any motion
                    DXF_polyline2d(motion, Layer, 1 / xScale)
                    motion.Vertexes.Clear()
                End If
                i = GetUInt()
                Select Case i
                    Case 0 : LaserMode = LSRMode_Enum.Off
                    Case 1 : LaserMode = LSRMode_Enum.Cut
                    Case 2 : LaserMode = LSRMode_Enum.Engrave
                    Case 3 : LaserMode = LSRMode_Enum.GradeEngrave
                    Case 4 : LaserMode = LSRMode_Enum.Hole
                    Case 5 : LaserMode = LSRMode_Enum.PenCut
                    Case Else
                        Throw New Exception($"Unknown laser mode {i} @0x{addr:x}")
                End Select

            Case MOL_HOLE                ' Burn hole
                Dim duration = GetInt()
                If duration > 0 Then
                    Dim circle = New Circle(New Vector2(position.X / xScale, position.Y / yScale), 0.1)
                    ' Create a hatch entity
                    Dim hatch As New Hatch(HatchPattern.Solid, False)
                    hatch.Pattern.Scale = 1.0
                    ' Add the circle as a boundary to the hatch
                    hatch.BoundaryPaths.Add(New HatchBoundaryPath(New List(Of EntityObject) From {circle}))
                    ' Add the entities to the DXF document
                    dxfBlock.Entities.Add(circle)
                    dxfBlock.Entities.Add(hatch)
                End If

            Case MOL_SCALE      ' also MOL_SPDPWR, MOL_SPDPWRx
                xScale = GetFloat() : yScale = GetFloat() : zScale = GetFloat()        ' x,y,z scale command

            Case MOL_MVREL, MOL_START
                Dim n = GetUInt()           ' always 772
                If n <> 772 Then Throw New Exception($"MOL_MVREL: n is not 772 @0x{addr:x}")
                Dim delta = New IntPoint(GetInt(), GetInt())      ' move relative command
                If delta = ZERO Then Throw New Exception($"MOL_MVREL: delta is (0,0) @0x{addr:x}")

                Select Case block
                    Case BLOCKTYPE.CONFIG
                        ' MVREL in the config block are start positions for each draw block
                        Dim SubBlock = DrawChunks(MVRELInConfig)      ' get chunk for this draw block
                        Dim bookmark = reader.BaseStream.Position         ' remember posn in decode stream
                        Dim subr = GetInt(SubBlock * BLOCK_SIZE + 4)          ' get the subroutine number
                        Dim abs = (MVRELInConfig = 0)       ' first is absolute, other are relative
                        StartPosns.Add(subr, (abs, delta))              ' add subroutine start address
                        reader.BaseStream.Position = bookmark                 ' restore reader position
                        MVRELInConfig += 1

                    Case BLOCKTYPE.TEST, BLOCKTYPE.CUT
                        DXF_line(position, position + delta, Layer)     ' render on this layer ignoring laser commands

                    Case BLOCKTYPE.DRAW
                        If LaserMode <> LSRMode_Enum.Off Then
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
                If LaserMode <> LSRMode_Enum.Off Then MaxLength = Math.Max(delta.Length, MaxLength)
                If block <> BLOCKTYPE.CONFIG Then position += delta       ' update head position

            Case MOL_ENGACD     ' Accelerate distance whilst engraving
                AccelLength = GetInt()

            Case MOL_ENGPWR     ' Power whilst engraving
                EngPower = GetInt() / 100

            Case MOL_ENGSPD
                GetInt()    ' axis
                EngSpeed = GetFloat() / xScale

            Case MOL_ENGMVX   ' Engrave move X  (laser on and off)
                ' ENGMVX is in 3 phases, Accelerate, Engrave, Decelerate
                Dim axis As Integer = GetInt()      ' consume Axis parameter
                If axis <> Axis_Enum.X Then Throw New Exception($"MOL_ENGMVX: Axis value {axis} is not valid @0x{addr:x}")
                Dim posn = position         ' use local copy of position
                Dim TravelDist As Integer = GetInt()       ' the total travel distance for this operation
                If TravelDist = 0 Then Throw New Exception($"MOL_ENGMVX: TravelDist is zero @0x{addr:x}")
                Dim direction As Integer = CInt(Math.Sign(TravelDist))     ' 1=LtoR, -1=RtoL
                ' Move for the initial acceleration
                Dim delta As New IntPoint(AccelLength * direction, 0)
                DXF_line(posn, posn + delta, MoveLayer)
                posn += delta
                ' do On/Off steps
                For Each stp In ENGLSRsteps
                    ' The On portion of the delta
                    If stp.OnSteps > 0 Then
                        delta = New IntPoint(stp.OnSteps * direction, 0)    ' delta in X direction
                        DXF_line(posn, posn + delta, EngraveLayer, PowerSpeedColor(EngPower, EngSpeed))     ' color set to represent engrave shade
                        posn += delta
                    End If
                    ' The Off portion of the delta
                    If stp.OffSteps > 0 Then
                        delta = New IntPoint(stp.OffSteps * direction, 0)    ' delta in X direction
                        DXF_line(posn, posn + delta, MoveLayer)
                        posn += delta
                    End If
                Next
                ' delta for the deceleration
                delta = New IntPoint(AccelLength * direction, 0)
                DXF_line(posn, posn + delta, MoveLayer)
                posn += delta

                'position += New IntPoint(TravelDist * direction, 0)
                position = posn       ' move the position along
                ENGLSRsteps.Clear()     ' clear the steps

            Case MOL_ENGMVY     ' Engrave move Y (laser off)
                Dim axis = GetInt()      ' consume Axis parameter
                If axis <> Axis_Enum.Y Then Throw New Exception($"MOL_ENGMVY: Axis value {axis} is not valid @0x{addr:x}")
                Dim dy = GetInt()
                If dy = 0 Then Throw New Exception($"MOL_ENGMVY: dy is zero @0x{addr:x}")
                Dim delta As New IntPoint(0, dy)    ' move in Y direction
                DXF_line(position, position + delta, MoveLayer)
                position += delta      ' move position along

            Case MOL_ENGLSR     ' engraving movement
                ENGLSRsteps.Clear()
                For i = 1 To nWords
                    Steps.Steps = GetInt()      ' get 2 16 bit values, accessable through OnSteps & OffSteps
                    ENGLSRsteps.Add(Steps)
                Next

            Case MOL_BEGSUB, MOL_BEGSUBa  ' begin subroutine
                Dim n = GetUInt()
                dxfBlock = New Blocks.Block($"Subr {n}")        ' create the block for this subroutine
                CurrentSubr = n
                If n < 100 Then
                    If StartPosns.ContainsKey(n) Then
                        If StartPosns(n).Absolute Then position = StartPosns(n).position Else position += StartPosns(n).position
                        TextBox1.AppendText($"Starting subr {n} with position ({position.X},{position.Y}){vbCrLf}")
                    Else
                        Throw New SystemException($"GOSUB {n} has no start position")
                    End If
                End If
                Select Case block
                    Case BLOCKTYPE.TEST
                        Layer = TestLayer
                    Case BLOCKTYPE.CUT
                        Layer = CutLayer
                    Case BLOCKTYPE.DRAW
                        Layer = EngraveLayer
                    Case Else
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
                dxf.Blocks.Add(dxfBlock)        ' add the block to the dxf file
                Dim insert As New Insert(dxfBlock)      ' create an insert
                dxf.Entities.Add(insert)                ' add the insert to Entities
                dxfBlock = Nothing                      ' deallocate block
                ' Output any buffered polyline
                If motion.Vertexes.Count > 0 Then
                    ' add the motion to the dxf file    
                    DXF_polyline2d(motion, Layer, 1 / xScale)
                    motion.Vertexes.Clear()
                End If
                TextBox1.AppendText($"Position at end of subr {n} is ({position.X},{position.Y}){vbCrLf}")
                If Stack.Count > 0 Then
                    Dim popped = Stack.Pop
                    Layer = popped.item2    'lyr
                    reader.BaseStream.Seek(popped.item1, SeekOrigin.Begin)     ' return to the saved return address
                    Return reader.BaseStream.Position
                Else
                    Throw New System.Exception($"Stack is exhausted - no return address for ENDSUB {n}")
                End If

            Case MOL_GOSUBn, MOL_GOSUBb       ' Call subroutine with parameters
                Dim n = GetInt()     ' get subroutine number
                Dim posn = New IntPoint(GetFloat(), GetFloat())     ' subroutine parameter ignored
                If n < 100 Then
                    Stack.Push((addr:=reader.BaseStream.Position, lyr:=Layer))              ' address of next instruction + current 
                    Return SubrAddrs(n)                     ' jump to start address and continue
                Else
                    Return reader.BaseStream.Position       ' ignore call. Goto next instruction
                End If

            Case MOL_GOSUB ' GOSUB
                Dim n As Integer = GetInt()
                dxfBlock = New Blocks.Block($"Subr {n}")

            Case MOL_ACCELERATION, MOL_SETSPD, MOL_PWRSPD5, MOL_PWRSPD7, MOL_BLWRa, MOL_BLWRb, MOL_X5_FIRST, MOL_MOTION, MOL_ENGSPD1, MOL_ENGPWR1
                ' Nothing to do

            Case Else
                Dim value As MOLcmd = Nothing
                If Commands.TryGetValue(cmd, value) Then Throw New Exception($"Unhandled command {value.Mnemonic} @0x{addr:x}") ' Check for un-handled command
                ' consume all parameters
                For i = 1 To nWords
                    GetInt()    ' consume parameters
                Next
        End Select
        Return param_posn + nWords * 4       ' address of next instruction
    End Function

    Private Function GetBlockType(addr As Integer) As BLOCKTYPE
        ' Translate an address to a block type
        Select Case addr
            Case 0 To BLOCK_SIZE - 1
                Return BLOCKTYPE.HEADER
            Case ConfigChunk * BLOCK_SIZE To TestChunk * BLOCK_SIZE - 1
                Return BLOCKTYPE.CONFIG
            Case TestChunk * BLOCK_SIZE To CutChunk * BLOCK_SIZE - 1
                Return BLOCKTYPE.TEST
            Case CutChunk * BLOCK_SIZE To DrawChunks(0) * BLOCK_SIZE - 1
                Return BLOCKTYPE.CUT
            Case DrawChunks(0) * BLOCK_SIZE To FileSize
                Return BLOCKTYPE.DRAW
            Case Else
                Throw New System.Exception($"Unknown block type at address {addr:x8}")
        End Select
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
        dxfBlock.Entities.Add(line)
    End Sub

    Public Sub DXF_polyline2d(Polyline As Polyline2D, Layer As Layer, Optional ByVal scale As Double = 1.0)
        ' Add a polyline to the specified layer
        Dim ply As Polyline2D = Polyline.Clone
        ply.Layer = Layer
        If scale <> 1.0 Then
            Dim transform As New Matrix3(scale, 0, 0, 0, scale, 0, 0, 0, scale)
            ply.TransformBy(transform, New Vector3(0, 0, 0))
        End If
        dxfBlock.Entities.Add(ply)
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
            Dim value As MOLcmd = Nothing
            If Commands.TryGetValue(cmd, value) Then TextBox1.AppendText($" ({value.Mnemonic})")
            TextBox1.AppendText($"{vbCrLf}")
        End While
    End Sub

    Private Sub WordAccessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WordAccessToolStripMenuItem.Click
        Dim value As Integer, value1() As Byte
        stream = System.IO.File.Open(dlg.filename, FileMode.Open)
        reader = New BinaryReader(stream, System.Text.Encoding.Unicode, False)
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

        Dim mode As String, IntervalStr As String
        If My.Settings.Engrave Then
            mode = "E"
            IntervalStr = $"{CInt(My.Settings.Interval * 10)}"       ' interval in 0.1 mm units
        Else
            mode = "C"
            IntervalStr = ""    ' Inverval is only relevant for engrave
        End If
        ' Make filename for output. Restricted to 8.3 format
        Dim filename = $"TC{mode}{IntervalStr}.MOL"     ' output file name
        If filename.IndexOf("."c) > 7 Then Throw New Exception($"Filename {filename} is invalid. Not 8.3")
        writer = New BinaryWriter(System.IO.File.Open(filename, FileMode.Create), System.Text.Encoding.Unicode, False)
        dxf = New DxfDocument()     ' create empty DXF file

        ' Copy the first 5 blocks of a template file to initialise the new MOL file
        stream = System.IO.File.Open(MOLTemplate, FileMode.Open)           ' file containing template blocks
        reader = New BinaryReader(stream, System.Text.Encoding.Unicode, False)
        GetHeader()
        reader.BaseStream.Seek(0, SeekOrigin.Begin)            ' set reader to start of file
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

        dxfBlock = New Blocks.Block("Test")        ' create the block for this subroutine
        MakeTestBlock(writer, dxf, Outline, TestLayer)
        Dim insert As New Insert(dxfBlock)      ' create an insert
        dxf.Entities.Add(insert)                ' add the insert to Entities
        dxfBlock = Nothing

        dxfBlock = New Blocks.Block("Cut")        ' create the block for this subroutine
        MakeCutBlock(writer, dxf, CutLine, CutLayer)
        insert = New Insert(dxfBlock)      ' create an insert
        dxf.Entities.Add(insert)                ' add the insert to Entities
        dxfBlock = Nothing

        dxfBlock = New Blocks.Block("Engrave")        ' create the block for this subroutine
        MakeEngraveBlock(writer, dxf, GridLine, EngraveLayer, Speeds, Powers)
        insert = New Insert(dxfBlock)      ' create an insert
        dxf.Entities.Add(insert)                ' add the insert to Entities
        dxfBlock = Nothing

        dxfBlock = New Blocks.Block("Text")        ' create the block for this subroutine
        MakeTextBlock(writer, dxf, GridLine, TextLayer, Speeds, Powers)
        insert = New Insert(dxfBlock)      ' create an insert
        dxf.Entities.Add(insert)                ' add the insert to Entities
        dxfBlock = Nothing

        dxf.Save("Test card.dxf")       ' save the DXF file

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
            WriteMOL(MOL_MOTION, {Float.Encode(StartSpeed * xScale), Float.Encode(QuickSpeed * xScale), Float.Encode(WorkAcc * xScale)})
            WriteMOL(MOL_START, {772, StartPosns(i).position.X, StartPosns(i).position.Y})
            WriteMOL(MOL_GOSUBn, {3, i, 0, 0})
        Next
        WriteMOL(MOL_LASER1, {LSRMode_Enum.Off})
        ' add GOSUBn 603
        WriteMOL(MOL_GOSUBn, {3, 603, 0, 0})
        ' end the config block
        EndBlock()        ' pad to end of block with 0

        writer.Close()
        TextBox1.AppendText($"Done - output in {filename}")
    End Sub

    Public Sub MakeTestBlock(writer As BinaryWriter, dxf As DxfDocument, outline As Rect, layer As Layer)
        ' Make the test block
        Const testSpeed = 180         ' speed when tracing TEST block
        position = New IntPoint(0, 0)       ' Layer starts at (0,0)
        WriteMOL(MOL_BEGSUBa, {1}, &H400)    ' begin SUB 1
        WriteMOL(&H80600148, {6, &H25B, &H91D4000, &HF13A300, &H1113A340, &H101D7A80, &HF716C00})   ' Unknown command
        DrawBox(writer, dxf, outline, layer,,, testSpeed)
        WriteMOL(MOL_ENDSUB, {1})
        EndBlock()        ' pad to end of block with 0
    End Sub
    ''' <summary>
    ''' Make the Cut block
    ''' </summary>
    Public Sub MakeCutBlock(writer As BinaryWriter, dxf As DxfDocument, outline As Rect, layer As Layer)
        position = New IntPoint(0, 0)       ' Layer starts at (0,0)
        WriteMOL(MOL_BEGSUBa, {2}, &H600)    ' begin SUB 2
        WriteMOL(&H80600148, {6, &H25B, &H91D4000, &HF13A300, &H1113A340, &H101D7A80, &HF716C00})   ' Unknown command
        WriteMOL(MOL_PWRSPD5, {ClipBoxPower * 100, ClipBoxPower * 100, Float.Encode(ClipBoxSpeed * xScale), Float.Encode(ClipBoxSpeed * xScale), Float.Encode(0.0)})    ' set power & speed
        UseMCBLK = True
        Dim delta = CType(outline.BottomRight, IntPoint)     ' convert to steps
        DXF_line(position, position + delta, CutLayer)
        MoveRelativeSplit(outline.BottomRight, ClipBoxSpeed)  ' Move to start position in two goes
        DrawBox(writer, dxf, outline, layer, False, ClipBoxPower, ClipBoxSpeed)
        FlushMCBLK(False)
        WriteMOL(MOL_ENDSUB, {2})
        EndBlock()        ' pad to end of block with 0
    End Sub

    ''' <summary>
    ''' Creates an engraving block in the DXF document and writes corresponding MOL commands.
    ''' </summary>
    ''' <param name="writer">The BinaryWriter to write the MOL commands to.</param>
    ''' <param name="dxf">The DXF document to add the engraving block to.</param>
    ''' <param name="outline">The outline rectangle for the engraving block.</param>
    ''' <param name="layer">The layer to add the engraving block to.</param>
    ''' <param name="speeds">An array of speeds for the engraving block.</param>
    ''' <param name="powers">An array of powers for the engraving block.</param>
    Public Sub MakeEngraveBlock(writer As BinaryWriter, dxf As DxfDocument, outline As Rect, layer As Layer, speeds() As Integer, powers() As Integer)

        ' Create engraving as SUBR 3. Subroutine engraves a single cell
        Const BlockNumber = 5       ' block number of first subroutine
        Dim StartofBlock = BlockNumber * BLOCK_SIZE   ' address of start of next block. First subr at A00
        PutInt(BlockNumber, &H7C) '  DrawChunks.Add(BlockNumber)    ' was added earlier
        WriteMOL(MOL_BEGSUBa, {3}, StartofBlock)    ' begin SUB 
        WriteMOL(&H1000D46, {&HBB8})         ' unknown command

        Dim cellsize = New Size(outline.Width / 10, outline.Height / 10)    ' there are 10 x 10 cells 

        ' Now create the engraved box for each setting
        ' Power goes left to right across page
        ' Speed goes top to bottom down the page
        ' boxes are drawn top down

        For power = 0 To powers.Length - 1
            For speed = 0 To speeds.Length - 1
                Dim cell = New Rect(New System.Windows.Point(outline.Left + power * cellsize.Width, outline.Top + speed * cellsize.Height), cellsize)  ' one 10x10 cell
                cell = Rect.Inflate(cell, -0.75, -0.75)       ' shrink it a bit to create a margin
                DrawBox(writer, dxf, cell, layer, My.Settings.Engrave, powers(power), speeds(speed))
                WriteMOL(MOL_LASER, {LSRMode_Enum.Off})    ' turn laser off
            Next
        Next
        WriteMOL(MOL_BLWRb, {OnOff_Enum.Off})
        WriteMOL(MOL_BLWRb, {OnOff_Enum.Off})
        WriteMOL(MOL_ENDSUB, {3})    ' end SUB 
        EndBlock()        ' pad to end of block with 0
    End Sub

    ''' <summary>
    ''' Creates a text block in the DXF document and writes corresponding MOL commands.
    ''' </summary>
    ''' <param name="writer">The BinaryWriter to write the MOL commands to.</param>
    ''' <param name="dxf">The DXF document to add the text block to.</param>
    ''' <param name="Outline">The outline rectangle for the text block.</param>
    ''' <param name="layer">The layer to add the text block to.</param>
    ''' <param name="speeds">An array of speeds for the text block.</param>
    ''' <param name="powers">An array of powers for the text block.</param>
    Public Sub MakeTextBlock(writer As BinaryWriter, dxf As DxfDocument, Outline As Rect, layer As Layer, speeds() As Integer, powers() As Integer)
        ' Create Text layer as SUBR 4
        Dim StartofBlock = (writer.BaseStream.Position \ BLOCK_SIZE + 1) * BLOCK_SIZE   ' address of start of next block
        Dim BlockNumber = StartofBlock \ BLOCK_SIZE
        PutInt(BlockNumber, &H7C + 4) : DrawChunks.Add(BlockNumber)
        Dim cellsize = New Size(Outline.Width / 10, Outline.Height / 10)
        WriteMOL(MOL_BEGSUBa, {4}, StartofBlock)    ' begin SUB
        UseMCBLK = True
        DrawText(writer, dxf, My.Settings.Material, System.Windows.TextAlignment.Center, New System.Windows.Point(Outline.Left + (Outline.Width / 2), -5), 3, 0)
        If My.Settings.Engrave Then DrawText(writer, dxf, $"Interval: {My.Settings.Interval} mm", System.Windows.TextAlignment.Center, New System.Windows.Point(Outline.Left + Outline.Width / 2, -9.5), 3, 0)  ' Interval only relevant for engrave
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
        EndBlock()        ' pad to end of block with 0
    End Sub
    ''' <summary>
    ''' Writes a MOL command with optional parameters to the output stream.
    ''' </summary>
    ''' <param name="command">The MOL command to write.</param>
    ''' <param name="Parameters">Optional parameters for the command.</param>
    ''' <param name="posn">Optional position to write the command to. If not specified, writes at the current position.</param>
    ''' <exception cref="System.Exception">Thrown if there is a parameter mismatch or unsupported parameter type.</exception>
    Public Sub WriteMOL(command As Integer, Optional ByVal Parameters() As Object = Nothing, Optional posn As Integer = -1)
        ' Write a MOL command, with parameters to MOL file
        ' Parameters are written in scaled values
        ' Parameters will be written (but not length) if present
        ' writing occurs at current writer position, or "posn" if present
        If UseMCBLK And posn <> -1 Then Throw New System.Exception($"You can't write explicitly to address {posn:x} as an MCBLK in in operation")

        Dim value As MOLcmd = Nothing        ' Check the correct number of parameters
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
                    Case GetType(Float) : PutInt(p.value)
                    Case GetType(OnOffSteps) : PutInt(p)
                    Case GetType(System.Collections.Generic.List(Of OnOffSteps))
                        For Each s In p
                            PutInt(s.steps)
                        Next
                    Case GetType(OnOff_Enum) : PutInt(p)
                    Case GetType(Acceleration_Enum), GetType(LSRMode_Enum), GetType(Axis_Enum)
                        PutInt(p)
                    Case Else
                        Throw New System.Exception($"WRITEMOL: Parameter of unsupported type {p.GetType}")
                End Select
            Next
        End If

        Select Case command
            Case MOL_MVREL, MOL_UNKNOWN07, MOL_UNKNOWN09        ' Count the number of MOL_MVREL, MOL_UNKNOWN07 & MOL_UNKNOWN09 commands
                MVRELCnt += 1
            Case MOL_LASER, MOL_LASER1, MOL_LASER2  ' track mode of laser
                If Parameters(0) Is Nothing Then Throw New Exception($"{value.Mnemonic}: Laser mode not set")
                LaserMode = Parameters(0)
        End Select
        If MCBLK.Count >= MCBLKMax Then
            FlushMCBLK(True)      ' MCBLK limited in size
        End If
    End Sub

    ''' <summary>
    ''' Flushes the contents of the MCBLK (Motion Control Block) buffer to the writer.
    ''' </summary>
    ''' <param name="KeepOpen">If true, keeps the MCBLK buffer open after flushing; otherwise, closes it.</param>
    Public Sub FlushMCBLK(KeepOpen As Boolean)
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
    End Sub

    Public Sub CutLinePrecise(delta As System.Windows.Point, speed As Double, acceleration As Double)
        CutLinePrecise(CType(delta, IntPoint), speed, acceleration)       ' convert mm to steps
    End Sub
    Public Sub CutLinePrecise(delta As IntPoint, speed As Double, accel As Double)
        ' Create a series of move commands from the current position by delta
        ' Each move is a maximum of 0.1mm
        ' The move is done in 3 phases
        ' - Accelerate to speed
        ' - Coast at speed
        ' - Decelerate to stop
        ' The movements are governed by the formula v2 = sqrt(v1^2 + 2as), where v1 is the initial speed, v2 is the final speed, a is the acceleration and s is the distance
        Dim InitialSpeed = CInt(5 * xScale)    ' initial speed in steps/s
        Dim CurrentSpeed As Integer = InitialSpeed
        Dim TargetSpeed = CInt(speed * xScale)          ' maximum speed
        Dim NextSpeed As Integer
        Dim Acceleration = CInt(accel * xScale)   ' acceleration in steps/s
        Dim Increment As Double = 0.1 * xScale
        Dim moves As New List(Of (InitSpeed As Integer, MaxSpeed As Integer, dx As Integer, dy As Integer))
        Dim CurrentPosition As IntPoint = ZERO      ' treat delta as vector, origin (0,0)
        Dim NextPosition As IntPoint            ' position to move to
        Dim TotalDistance As Double = delta.Length       ' total distance to move
        Dim Ratio As Double
        Dim AccelerateDistance As Double       ' distance required to achieve MaxSpeed
        Dim CurrentDistance As Double
        Dim MoveDelta As IntPoint               ' amount to move as a delta
        Dim points As Integer                   ' which point along the line

        ' Calculate the distance it would take to accelerate to speed if we did it in 1 step
        AccelerateDistance = (TargetSpeed ^ 2 - InitialSpeed ^ 2) / (2 * Acceleration)    ' distance to accelerate to target speed
        If AccelerateDistance * 2 > TotalDistance Then Throw New Exception($"Acceleration distance {AccelerateDistance / xScale:f1}mm is greater than half total distance {TotalDistance / xScale:f1}mm. Speed = {speed:f1}mm/s")       ' not enough distance to accelerate

        ' Phase 1 - Accelerate until we reach speed
        While CurrentSpeed < TargetSpeed
            points = moves.Count + 1             ' which point are we calculating
            Ratio = points * Increment / TotalDistance
            NextPosition = New IntPoint(CInt(Ratio * delta.X), CInt(Ratio * delta.Y))
            MoveDelta = NextPosition - CurrentPosition
            NextSpeed = CInt(Math.Sqrt(CurrentSpeed ^ 2 + 2 * Acceleration * Increment))
            NextSpeed = CInt(Math.Min(NextSpeed, TargetSpeed))        ' limit to TargetSpeed
            moves.Add((InitSpeed:=CurrentSpeed, MaxSpeed:=NextSpeed, dx:=MoveDelta.X, dy:=MoveDelta.Y))
            CurrentSpeed = NextSpeed
            CurrentPosition = NextPosition
            CurrentDistance = CurrentPosition.Length        ' distance travelled
            If CurrentSpeed = TargetSpeed Then      ' reached target speed
                Exit While
            End If
        End While
        AccelerateDistance = CurrentDistance    ' the distance we required to accelerate

        ' Phase 2 - coast until we reach time to decelerate
        Do
            points = moves.Count + 1             ' which point are we calculating
            Ratio = points * Increment / TotalDistance
            NextPosition = New IntPoint(CInt(Ratio * delta.X), CInt(Ratio * delta.Y))
            MoveDelta = NextPosition - CurrentPosition
            moves.Add((InitSpeed:=CurrentSpeed, MaxSpeed:=CurrentSpeed, dx:=MoveDelta.X, dy:=MoveDelta.Y))
            CurrentPosition = NextPosition
            CurrentDistance = CurrentPosition.Length       ' distance travelled
        Loop Until CurrentDistance >= TotalDistance - AccelerateDistance

        ' Phase 3 - Decelerate to stop
        While CurrentSpeed > InitialSpeed
            points = moves.Count + 1             ' which point are we calculating
            NextSpeed = CInt(Math.Sqrt(Math.Max(CurrentSpeed ^ 2 - 2 * Acceleration * Increment, 0)))
            Ratio = points * Increment / TotalDistance
            NextPosition = New IntPoint(CInt(Ratio * delta.X), CInt(Ratio * delta.Y))
            MoveDelta = NextPosition - CurrentPosition
            moves.Add((InitSpeed:=CurrentSpeed, MaxSpeed:=NextSpeed, dx:=MoveDelta.X, dy:=MoveDelta.Y))
            CurrentSpeed = NextSpeed
            CurrentPosition = NextPosition
        End While
        position = CurrentPosition      ' update position

        ' post calc check - that total dx,dy matches requested
        'Dim dxtotal As Integer = 0, dytotal As Integer = 0
        'For Each Move As (InitSpeed As Integer, MaxSpeed As Integer, dx As Integer, dy As Integer) In moves
        '    dxtotal += Move.dx
        '    dytotal += Move.dy
        'Next
        'If dxtotal <> delta.X Or dytotal <> delta.Y Then
        '    Throw New System.Exception($"CutLine: Calculated distance {dxtotal},{dytotal} does not match requested {delta.X},{delta.Y}")
        'End If

        ' Execute the moves
        For Each Move As (InitSpeed As Integer, MaxSpeed As Integer, dx As Integer, dy As Integer) In moves
            If Move.MaxSpeed > Move.InitSpeed Then
                WriteMOL(MOL_ACCELERATION, {Acceleration_Enum.Accelerate})      ' we are accelerating
                WriteMOL(MOL_SETSPD, {Float.Encode(Move.InitSpeed), Float.Encode(Move.MaxSpeed), Float.Encode(accel * xScale)})
            ElseIf Move.MaxSpeed < Move.InitSpeed Then
                WriteMOL(MOL_ACCELERATION, {Acceleration_Enum.Decelerate})      ' we are decelerating
                WriteMOL(MOL_SETSPD, {Float.Encode(Move.MaxSpeed), Float.Encode(Move.InitSpeed), Float.Encode(accel * xScale)})
            Else
                WriteMOL(MOL_SETSPD, {Float.Encode(Move.MaxSpeed), Float.Encode(Move.MaxSpeed), Float.Encode(accel * xScale)})
            End If
            WriteMOL(MOL_MVREL, {772, Move.dx, Move.dy})
        Next
    End Sub
    Public Sub MoveRelativeSplit(delta As System.Windows.Point, speed As Double)
        MoveRelativeSplit(CType(delta, IntPoint), speed)       ' convert mm to steps
    End Sub

    ''' <summary>
    ''' Creates a move command from the current position by the specified delta.
    ''' The move is split into multiple segments to handle acceleration and deceleration.
    ''' </summary>
    ''' <param name="delta">The distance to move, specified as an IntPoint.</param>
    ''' <param name="speed">The speed of the movement in mm/s.</param>
    Public Sub MoveRelativeSplit(delta As IntPoint, speed As Double)
        ' Pieces = 2 or 3
        ' Accelerate before first one
        ' Decelerate before last one
        ' There may be a middle piece
        ' Phase 1 is slow, Phase 2 fast, phase 3 slow
        Dim moves As New List(Of IntPoint)

        ' Work out whether we can do a 2 part move, or it's too long and we need 3

        If delta = ZERO Then Return     ' we're already there
        Dim Accel = IIf(LaserMode = LSRMode_Enum.Off, SpaceAcc, WorkAcc)   ' Acceleration differs if laser is on or not
        Dim T = (speed - StartSpeed) / Accel       ' T =(max speed-initial speed)/acceleration time taken to reach QuickSpeed
        Dim S As Integer = CInt(StartSpeed * T + xScale * 0.5 * Accel * T ^ 2)  ' s=ut+0.5t*t  distance (steps) travelled whilst reaching this speed
        ' Calculate 2 or 3 part move
        Dim Dist = delta.Length         ' distance to move

        If S > Dist / 2 Or block = BLOCKTYPE.TEST Or block = BLOCKTYPE.CUT Or LaserMode = LSRMode_Enum.Off Then     ' move in 2 steps
            ' we can get more than halfway if we needed, so 2 pieces enough
            ' only Accelerate till halfway, then decelerate

            Dim delta1 = New IntPoint(delta.X / 2, delta.Y / 2)
            moves.Add(delta1)
            Dim delta2 = delta - delta1     ' We do it this way to avoid rounding errors
            moves.Add(delta2)
        Else
            ' we can't make it halfway whilst accelerating, so will need to coast
            Dim ratio As Double = S / Dist        ' percentage of accelerate/decelerate phases
            Dim delta1 = New IntPoint(delta.X * ratio, delta.Y * ratio)
            Dim delta2 = New IntPoint(delta.X * (1 - 2 * ratio), delta.Y * (1 - 2 * ratio))
            Dim delta3 = delta - delta1 - delta2
            moves.Add(delta1)       ' initial move
            moves.Add(delta2)       ' middle move
            moves.Add(delta3)       ' final move
        End If

        For m = 0 To moves.Count - 1
            If m = 0 Then    ' the first one
                WriteMOL(MOL_ACCELERATION, {Acceleration_Enum.Accelerate})
                WriteMOL(MOL_SETSPD, {Float.Encode(StartSpeed * xScale), Float.Encode(speed * xScale), Float.Encode(Accel * xScale)})
            ElseIf m = moves.Count - 1 Then    ' the last one
                WriteMOL(MOL_ACCELERATION, {Acceleration_Enum.Decelerate})
                WriteMOL(MOL_SETSPD, {Float.Encode(StartSpeed * xScale), Float.Encode(speed * xScale), Float.Encode(Accel * xScale)})
            Else
                WriteMOL(MOL_SETSPD, {Float.Encode(speed * xScale), Float.Encode(speed * xScale), Float.Encode(Accel * xScale)})
            End If
            WriteMOL(MOL_MVREL, {772, moves(m).X, moves(m).Y})
        Next
        position += delta      ' update position
    End Sub

    ''' <summary>
    ''' Draws a box in the DXF document and writes corresponding MOL commands.
    ''' </summary>
    ''' <param name="writer">The BinaryWriter to write the MOL commands to.</param>
    ''' <param name="dxf">The DXF document to add the box to.</param>
    ''' <param name="outline">The outline rectangle for the box.</param>
    ''' <param name="Layer">The layer to add the box to.</param>
    ''' <param name="shaded">If true, the box will be shaded (engraved); otherwise, it will be a simple outline.</param>
    ''' <param name="power">The power setting for the laser (percentage).</param>
    ''' <param name="speed">The speed of the movement in mm/s.</param>
    Public Sub DrawBox(writer As BinaryWriter, dxf As DxfDocument, outline As Rect, Layer As Layer, Optional shaded As Boolean = False, Optional power As Integer = 0, Optional speed As Double = 150)
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
                MoveRelativeSplit(delta, QuickSpeed)        ' Move to start of box (BottomRight) in 2 pieces
            End If
            WriteMOL(MOL_PWRSPD5, {power * 100, power * 100, Float.Encode(5 * xScale), Float.Encode(speed * xScale), Float.Encode(0.0)})    ' set power & speed
            UseMCBLK = True     ' open MCB
            WriteMOL(&H1000B06, {&H201})        ' Unknown magic command
            ' Ordinary box with 4 sides
            block = GetBlockType(writer.BaseStream.Position)    ' get block type   
            If block = BLOCKTYPE.DRAW Or block = BLOCKTYPE.CUT Then WriteMOL(MOL_LASER, {LSRMode_Enum.Cut}) : 
            'If block = BLOCKTYPE.DRAW Then
            '    ' We are cutting, so use precise cut line
            '    CutLinePrecise(New IntPoint(-outline.Width * xScale, 0), speed, WorkAcc)
            '    CutLinePrecise(New IntPoint(0, -outline.Height * yScale), speed, WorkAcc)
            '    CutLinePrecise(New IntPoint(outline.Width * xScale, 0), speed, WorkAcc)
            '    CutLinePrecise(New IntPoint(0, outline.Height * yScale), speed, WorkAcc)
            'Else
            ' Use quick and dirty method to move/cut
            MoveRelativeSplit(New System.Windows.Point(-outline.Width, 0), speed)
            MoveRelativeSplit(New System.Windows.Point(0, -outline.Height), speed)
            MoveRelativeSplit(New System.Windows.Point(outline.Width, 0), speed)
            MoveRelativeSplit(New System.Windows.Point(0, outline.Height), speed)
            'End If
            If LaserMode <> LSRMode_Enum.Off Then WriteMOL(MOL_LASER, {LSRMode_Enum.Off})     ' Turn laser off
            FlushMCBLK(False)       ' close MCB
            WriteMOL(&H1000B06, {&H200})        ' Unknown magic command
            WriteMOL(&H1000B06, {&H200})        ' Unknown magic command
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
            MoveRelativeSplit(delta, speed)
            WriteMOL(MOL_LASER1, {LSRMode_Enum.Engrave})     ' turn laser on engrave mode
            Dim Steps As OnOffSteps                 ' construct steps as on for cellwidth, and off for 0
            Steps.OnSteps = outline.Width * xScale
            Steps.OffSteps = 0
            WriteMOL(MOL_ENGACD, {engacd})     ' define the acceleration start distance
            WriteMOL(&H1000246, {0})            ' unknown
            WriteMOL(MOL_ENGSPD, {Axis_Enum.X, Float.Encode(speed * xScale)})          ' define speed
            WriteMOL(MOL_ENGSPD1, {Axis_Enum.X, Float.Encode(7559.0), Float.Encode(speed * xScale), Float.Encode(881889.0)})          ' define speed
            WriteMOL(&H2010041, {3, &HB6C3000})     ' unknown
            WriteMOL(MOL_PWRSPD5, {power * 100, power * 100, Float.Encode(0.0), Float.Encode(speed * xScale), Float.Encode(0.0)})
            WriteMOL(MOL_ENGPWR, {power * 100})             ' define power
            WriteMOL(MOL_PWRSPD7, {power * 100, power * 100, power * 100, power * 100, Float.Encode(5 * xScale), Float.Encode(speed * xScale), Float.Encode(0.0)})    ' set power & speed
            WriteMOL(MOL_ENGPWR1, {power * 100, 6000})
            WriteMOL(MOL_BLWRb, {OnOff_Enum.On})
            Dim OnOffTotal As Integer = engacd * 2 + Steps.OnSteps + Steps.OffSteps
            Dim OnOffOnly As Integer = Steps.OnSteps + Steps.OffSteps   ' only on/off step distance
            Dim direction As Integer = 1        ' moving L to R
            Dim height As Integer = 0    ' count lines as integer and multiply to get height to avoid rounding errors
            Dim YInc As New IntPoint(0, My.Settings.Interval * yScale)  ' Y increment steps
            While height < outline.Height * xScale     ' do until we reach cell height
                WriteMOL(MOL_ENGLSR, {1, New List(Of OnOffSteps) From {{Steps}}})      ' one full line of on/off
                WriteMOL(MOL_ENGMVX, {Axis_Enum.X, OnOffTotal * direction})
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
                WriteMOL(MOL_ENGMVY, {Axis_Enum.Y, YInc.Y})
                DXF_line(position, position + YInc, EngraveMoveLayer)
                position += YInc
                direction *= -1           ' reverse direction for next line
                height += YInc.Y     ' move up to next line
            End While
            WriteMOL(MOL_BLWRb, {OnOff_Enum.Off})
        End If
    End Sub

    ''' <summary>
    ''' Draws text in the DXF document and writes corresponding MOL commands.
    ''' </summary>
    ''' <param name="writer">The BinaryWriter to write the MOL commands to.</param>
    ''' <param name="dxf">The DXF document to add the text to.</param>
    ''' <param name="text">The text to draw.</param>
    ''' <param name="alignment">The text alignment (Left, Center, Right).</param>
    ''' <param name="origin">The origin point for the text.</param>
    ''' <param name="fontsize">The font size of the text in mm.</param>
    ''' <param name="angle">The rotation angle of the text in degrees.</param>
    Public Sub DrawText(writer As BinaryWriter, dxf As DxfDocument, text As String, alignment As System.Windows.TextAlignment, origin As System.Windows.Point, fontsize As Double, angle As Double)
        ' Write some text at specified size, position and angle

        Const TextSpeed = 40      ' speed for text
        Const TextPower = 40      ' % power for text

        Dim Strokes = DisplayString(text, alignment, origin, fontsize, angle)       ' Convert the text to strokes
        Dim shading = PowerSpeedColor(My.Settings.PowerMax / 2, My.Settings.SpeedMax / 2)      ' use 50% power and speed for text

        ' draw the text
        ' DXF first
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
        UseMCBLK = True
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

            ' Set speed and power for text
            WriteMOL(MOL_PWRSPD5, {TextPower * 100, TextPower * 100, Float.Encode(TextSpeed * xScale), Float.Encode(TextSpeed * xScale), Float.Encode(0.0)})    ' set power & speed
            ' Get to the start position of this stroke
            Dim delta As IntPoint = StartPoint - position
            If delta <> ZERO Then
                DXF_line(position, position + delta, MoveLayer)
                MoveRelativeSplit(delta, QuickSpeed)      ' move in 2 segments
            End If
            WriteMOL(MOL_LASER, {LSRMode_Enum.Cut})
            ' Process all vertexes, except the first
            For p = 1 To stroke.Vertexes.Count - 1
                Dim thispoint = New IntPoint(stroke.Vertexes(p).Position.X, stroke.Vertexes(p).Position.Y)
                delta = thispoint - position
                MoveRelativeSplit(delta, TextSpeed)          ' draw stroke
            Next
            WriteMOL(MOL_LASER, {LSRMode_Enum.Off})
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

        TextBox1.AppendText($"Round trip discrepancy for 0.1 is {Float2Double(Double2Float(0.1))}{vbCrLf}")
        ' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]
        For i = 0.1 To 1 Step 0.1
            Dim n = Double2Float(i)
            TextBox1.AppendText($"{i:f1}  0x{n:x8} Sign={(n >> 23) And 1} exp={(n >> 24) And &HFF - 256} Mantissa={n And &H7FFFFF}{vbCrLf}")
        Next
        ' Test known cases
        TextBox1.AppendText($"Known Equivalents test{vbCrLf}")
        For Each t In KnownEquivalents
            Dim IEEE = t.IEEE
            Dim Leetro = t.Leetro 'And &HFFFFFE00
            Dim toIEEE = Float2Double(Leetro)
            Dim toLeetro = Double2Float(IEEE)
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
            Dim ToInt = Double2Float(t)
            Dim FromInt As Double = Float2Double(ToInt)
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
            GetHeader()         ' Get critical parameters from header
            Dim Size As Long = reader.BaseStream.Length / 4     ' size of the inputfile in 4 byte words
            ' Open a text file to write to
            Dim TextFile As System.IO.StreamWriter
            Dim Basename = System.IO.Path.GetFileNameWithoutExtension(dlg.filename)
            Dim textname = $"{Basename}.txt"
            TextFile = My.Computer.FileSystem.OpenTextFileWriter(textname, False)
            TextFile.AutoFlush = True
            ' DISASSEMBLE everything
            TextFile.WriteLine($"Disassembly of {dlg.filename} on {Now}")
            DisplayHeader(TextFile)
            DecodeStream(TextFile, ConfigChunk * BLOCK_SIZE)
            DecodeStream(TextFile, TestChunk * BLOCK_SIZE)
            DecodeStream(TextFile, CutChunk * BLOCK_SIZE)
            For Each Chunk In DrawChunks
                DecodeStream(TextFile, Chunk * BLOCK_SIZE)
            Next
            ' Display command usage
            TextFile.WriteLine()
            TextFile.WriteLine("Command frequency")
            TextFile.WriteLine()
            For Each cmd In CommandUsage
                Dim value As MOLcmd = Nothing, decode As String
                If Commands.TryGetValue(cmd.Key, value) Then decode = value.Mnemonic Else decode = $"0x{cmd.Key:x8}"
                TextFile.WriteLine($"{decode}  {cmd.Value}")
            Next

            ' Other metrics

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

            position = ZERO
            TextBox1.Clear()
            stream = System.IO.File.Open(dlg.filename, FileMode.Open)
            reader = New BinaryReader(stream, System.Text.Encoding.Unicode, False)
            GetHeader()         ' Get critical parameters from header
            ' Harvest subroutine start addresses
            ' Look at the first instruction of each block. BEGSUBa instruction there has the SUBR number
            Dim blocks As New List(Of Integer) From {
                TestChunk,      ' TEST
                CutChunk       ' CUT
            }
            For Each b In DrawChunks
                blocks.Add(b)
            Next
            For Each block In blocks
                Dim addr = block * BLOCK_SIZE
                Dim n = GetInt(addr)
                If n <> MOL_BEGSUB And n <> MOL_BEGSUBa Then Throw New System.Exception($"Block {block} @ 0x{addr:x} does not start with a BEGSUB or BEGSUBa command")
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
            Stack.Push((addr:=0, lyr:=TestLayer))   ' any old rubbish to keep ENDSUB happy
            position = ZERO      ' this block has an origin of (0,0)
            ExecuteStream(TestChunk * BLOCK_SIZE, dxf, TestLayer)   ' test

            Stack.Push((addr:=0, lyr:=CutLayer))   ' any old rubbish to keep ENDSUB happy
            position = ZERO      ' this block has an origin of (0,0)
            ExecuteStream(CutChunk * BLOCK_SIZE, dxf, CutLayer)   ' cut

            Stack.Push((addr:=0, lyr:=ConfigLayer))   ' any old rubbish to keep ENDSUB happy
            position = ZERO      ' this block has an origin of (0,0)
            ExecuteStream(ConfigChunk * BLOCK_SIZE, dxf, ConfigLayer)   ' config (which calls everything else)

            ' Now execute any that weren't called from config
            For Each s In ChunksWithCode
                If Not s.Value Then
                    Stack.Push((addr:=0, lyr:=DrawLayer))   ' any old rubbish to keep ENDSUB happy
                    position = ZERO      ' this block has an origin of (0,0)
                    ExecuteStream(s.Key * BLOCK_SIZE, dxf, DrawLayer)   ' config (which calls everything else)
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
            TextBox1.AppendText($"maximum vector drawn = {MaxLength / xScale}mm{vbCrLf}")
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

    ''' <summary>
    ''' Initializes all variables and clears any existing data.
    ''' This method is called to reset the state of the application before processing a new file.
    ''' </summary>
    Public Sub Initialise()
        ' Initialize file streams and readers/writers
        stream = Nothing
        reader = Nothing
        writer = Nothing

        ' Initialize header information
        TopRight = New IntPoint(0, 0)
        BottomLeft = New IntPoint(0, 0)
        startposn = New IntPoint(0, 0)
        delta = New IntPoint(0, 0)
        ZERO = New IntPoint(0, 0)
        position = ZERO

        ' Clear command usage statistics
        CommandUsage.Clear()

        ' Initialize steps array and stack
        StepsArray = Nothing
        Stack.Clear()

        ' Initialize drawing layer and MCBLK buffer
        layer = Nothing
        MCBLK.Clear()
        UseMCBLK = False
        MCBLKCount = 0
        MVRELCnt = 0

        ' Clear chunks with code and reset position
        ChunksWithCode.Clear()
        position = ZERO

        ' Initialize laser mode and acceleration length
        LaserMode = LSRMode_Enum.Off
        AccelLength = 0

        ' Initialize subroutine information
        CurrentSubr = 0
        StartPosns.Clear()
        SubrAddrs.Clear()

        ' Initialize chunk addresses
        ConfigChunk = 0
        TestChunk = 0
        CutChunk = 0
        DrawChunks.Clear()

        ' Initialize engraving steps and flags
        ENGLSRsteps.Clear()
        FirstMove = True
        EngPower = 0
        EngSpeed = 0
        MVRELInConfig = 0
        MCBLKCounter = 0

        ' Clear the text box for output
        TextBox1.Clear()

        ' Initialize start positions for subroutines
        StartPosns.Clear()
        StartPosns.Add(1, (True, New IntPoint(0, 0)))
        StartPosns.Add(2, (True, New IntPoint(0, 0)))

        ' Clear chunks with code
        ChunksWithCode.Clear()
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
        'dxfblock.entities.add(New Line(New Vector2(0, 0), New Vector2(0, 10)))
        DXF_line(New System.Windows.Point(0, 0), New System.Windows.Point(0, 10), MoveLayer)
        dxf.Save("LineTest.dxf")
    End Sub

    ''' <summary>
    ''' Handles the click event for the "Convert Leetro to IEEE" menu item.
    ''' Prompts the user to input a Leetro float in hex format, converts it to an IEEE float, and displays the result.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The event data.</param>
    Private Sub ConvertLeetroToIEEEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConvertLeetroToIEEEToolStripMenuItem.Click
        ' Prompt the user to input a Leetro float in hex format
        Dim hex As String = InputBox("Input Leetro float in hex format", "Convert Leetro float to Double")
        Try
            ' Convert the input hex string to a UInt32 value
            Dim value As UInt32 = Convert.ToUInt32(hex, 16)
            ' Convert the Leetro float to an IEEE float
            Dim float = Float2Double(value)
            ' Display the result in a message box
            MsgBox($"The value of 0x{hex} as a Leetro float is {float}", vbInformation + vbOKOnly, "Conversion to float")
        Catch ex As Exception
            ' Display an error message if the input is invalid
            MsgBox($"Invalid input. {ex.Message}", vbExclamation + vbOKOnly, "Error")
        End Try
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
            dxfBlock.Entities.Add(pl)
        Next

        ' Draw a crosshair
        Dim Cross As New System.Windows.Point(300, 300)
        Dim w As New System.Windows.Size(10, 10)
        Dim ln As New Line(New Vector2(Cross.X - w.Width, Cross.Y), New Vector2(Cross.X + w.Width, Cross.Y)) With {.Color = AciColor.Red}
        dxfBlock.Entities.Add(ln)
        ln = New Line(New Vector2(Cross.X, Cross.Y - w.Height), New Vector2(Cross.X, Cross.Y + w.Height)) With {.Color = AciColor.Red}
        dxfBlock.Entities.Add(ln)
        ' 5,2.5;0,2.5,A-1;0,6.5;5,6.5,A-1;5,2.5 = "O"
        ' 1,10;1,-1,A0.181818   = "("

        Dim polyl As New Polyline2D
        polyl.Vertexes.Add(New Polyline2DVertex(New Vector2(1, 10)))
        polyl.Vertexes.Add(New Polyline2DVertex(New Vector2(1, -1), 0.181818))

        st = "String at 0 center"
        ply = DisplayString(st, System.Windows.TextAlignment.Center, Cross, 10, 0)
        For Each pl In ply
            dxfBlock.Entities.Add(pl)
        Next
        st = "String at 90 left"
        ply = DisplayString(st, System.Windows.TextAlignment.Left, Cross, 10, 90)
        For Each pl In ply
            dxfBlock.Entities.Add(pl)
        Next
        st = "String at 45 right"
        ply = DisplayString(st, System.Windows.TextAlignment.Right, Cross, 10, 45)
        For Each pl In ply
            dxfBlock.Entities.Add(pl)
        Next

        st = "Power (%) - 10mm"
        ply = DisplayString(st, System.Windows.TextAlignment.Left, New System.Windows.Point(50, 50), 10, 45)
        For Each pl In ply
            dxfBlock.Entities.Add(pl)
        Next

        st = "This is left justified - 5mm"
        ply = DisplayString(st, System.Windows.TextAlignment.Left, New System.Windows.Point(200, 60), 5, 0)
        For Each pl In ply
            dxfBlock.Entities.Add(pl)
        Next
        st = "This is center justified - 5mm"
        ply = DisplayString(st, System.Windows.TextAlignment.Center, New System.Windows.Point(200, 70), 5, 0)
        For Each pl In ply
            dxfBlock.Entities.Add(pl)
        Next
        st = "This is right justified - 5mm"
        ply = DisplayString(st, System.Windows.TextAlignment.Right, New System.Windows.Point(200, 80), 5, 0)
        For Each pl In ply
            dxfBlock.Entities.Add(pl)
        Next
        st = "This is center justified - 20mm"
        ply = DisplayString(st, System.Windows.TextAlignment.Center, New System.Windows.Point(200, 90), 20, 0)
        For Each pl In ply
            dxfBlock.Entities.Add(pl)
        Next

        dxf.Save("GlyphTest.dxf")
        TextBox1.AppendText("Done")
    End Sub

    ''' <summary>
    ''' Generates a list of Polyline2D objects representing the specified text at the given origin, with specified scale and rotation.
    ''' </summary>
    ''' <param name="text">The string to be rendered.</param>
    ''' <param name="alignment">The text alignment (Left, Center, Right) about the origin.</param>
    ''' <param name="origin">The origin point for the text.</param>
    ''' <param name="fontsize">The font size of the text in mm.</param>
    ''' <param name="rotation">The rotation angle of the text in degrees.</param>
    ''' <returns>A list of Polyline2D objects representing the strokes of the text.</returns>
    Function DisplayString(text As String, alignment As System.Windows.TextAlignment, origin As System.Windows.Point, fontsize As Double, rotation As Double) As List(Of Polyline2D)
        Const RawFontSize = 9     ' the fonts are defined 9 units high
        Const LetterSpacing = 3       ' space between letters
        Dim Width As Double
        Dim Strokes As List(Of Polyline2D)
        Dim utf8Encoding As New UTF8Encoding()
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

    ''' <summary>
    ''' Loads the LibreCAD single stroke font from the specified font file.
    ''' The font data is stored in the FontData dictionary, where each glyph is represented by a list of strokes.
    ''' </summary>
    Friend Sub LoadFonts()
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
        If Not Directory.Exists(FontPath) Then
            MsgBox($"The font path {FontPath} cannot be found. Please use the parameters dialog to set one.", vbAbort + vbOK, "Font folder not found")
            Exit Sub
        Else
            Dim FontFile = My.Settings.FontFile
            If Not File.Exists($"{FontPath}\{FontFile}") Then
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
                                LastVertex = Nothing        ' just to supress a warning
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
                                                MsgBox($"Illegal data on line {LineNumber}: {line}. Font file {FontFile}, glyph {Key},0x{Key:x}", vbAbort + vbOKOnly, "Illegal glyph data")
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
                TextBox1.AppendText($"Font with {FontData.Count} glyphs loaded from {FontPath}\{FontFile}{vbCrLf}")
            End If
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadFonts()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Current.Shutdown()
        End
    End Sub

    ''' <summary>
    ''' Handles the click event for the "Command Spreadsheet" menu item.
    ''' Generates a spreadsheet containing a formatted list of all MOL commands and their parameters.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The event data.</param>
    Private Sub CommandSpreadsheetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CommandSpreadsheetToolStripMenuItem.Click
        ' Set the license key for GemBox.Spreadsheet
        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY")

        ' Create a new Excel workbook and add a worksheet named "Commands"
        Dim workbook As New ExcelFile
        Dim worksheet = workbook.Worksheets.Add("Commands")

        ' Define cell styles for horizontal and vertical alignment
        Dim HorizCenteredStyle = New CellStyle With {
            .HorizontalAlignment = HorizontalAlignmentStyle.Center
        }
        Dim VertCenteredStyle = New CellStyle With {
            .VerticalAlignment = VerticalAlignmentStyle.Center
        }

        ' Add headers to the worksheet
        Dim row = 0
        Dim col = 0
        For Each h In {"Code", "Mnemonic", "Description", "Parameters"}
            worksheet.Cells(row, col).Value = h
            col += 1
        Next

        ' Add second row of headers for parameters
        row = 1
        col = 3
        For Each h In {"#", "Description", "Name", "Type", "Scale", "Units"}
            worksheet.Cells(row, col).Value = h
            col += 1
        Next

        ' Merge cells for the "Parameters" header and apply horizontal centering
        With worksheet.Cells.GetSubrange("D1:I1")
            .Merged = True
            .Style = HorizCenteredStyle
        End With

        ' Apply heading style to the first two rows
        worksheet.Rows("1").Style = workbook.Styles(BuiltInCellStyleName.Heading1)
        worksheet.Rows("2").Style = workbook.Styles(BuiltInCellStyleName.Heading1)

        ' Populate the worksheet with command data
        row = 2
        ' Sort the commands dictionary by the low 24 bits of the command code
        Dim sorted = From item In Commands
                     Order By item.Key And &HFFFFFF
                     Select item
        For Each c In sorted
            ' Populate the rows with command data
            worksheet.Rows(row).Cells("A").Value = $"0x{c.Key:x8}"
            worksheet.Rows(row).Cells("B").Value = c.Value.Mnemonic
            worksheet.Rows(row).Cells("C").Value = c.Value.Description
            Dim pNum = 1
            For Each p In c.Value.Parameters
                worksheet.Rows(row).Cells("D").Value = $"p{pNum}"
                worksheet.Rows(row).Cells("E").Value = p.Description
                worksheet.Rows(row).Cells("F").Value = p.Name
                Dim typ = p.Typ.ToString
                If typ.StartsWith("MOL.") Then typ = typ.Substring(4)    ' remove MOL qualifier
                Select Case typ
                    Case "System.Int32" : typ = "Int32"
                    Case "System.Collections.Generic.List`1[MOL.OnOffSteps]" : typ = "List(Of OnOffSteps)"
                    Case "System.Collections.Generic.List`1[MOL.Int32]" : typ = "List(Of Integer)"
                End Select
                worksheet.Rows(row).Cells("G").Value = typ
                worksheet.Rows(row).Cells("H").Value = p.Scale
                worksheet.Rows(row).Cells("I").Value = p.Units

                pNum += 1
                row += 1
            Next

            ' Merge cells for commands with multiple parameters
            Dim pars = c.Value.Parameters.Count
            If pars > 1 Then
                For col = 0 To 2
                    Dim Mstart = A1(row - pars, col)
                    Dim Mend = A1(row - 1, col)
                    With worksheet.Cells.GetSubrange($"{Mstart}:{Mend}")
                        .Merged = True
                        .Style = VertCenteredStyle
                    End With
                Next
            End If
            row += 1
        Next

        ' Apply number format to the "Scale" column
        worksheet.Columns("H").Style.NumberFormat = NumberFormatBuilder.Number(5)

        ' Autofit column widths
        Dim columnCount = worksheet.CalculateMaxUsedColumns()
        For i As Integer = 0 To columnCount - 1
            worksheet.Columns(i).AutoFit(1, worksheet.Rows(1), worksheet.Rows(worksheet.Rows.Count - 1))
        Next

        ' Create a new worksheet for enums
        Dim EnumSheet As ExcelWorksheet = workbook.Worksheets.Add("Enums")
        EnumSheet.Rows("1").Style = workbook.Styles(BuiltInCellStyleName.Heading1)

        ' Collect custom types (enums) used in command parameters
        Dim CustomTypes As New List(Of Type)
        For Each cmd In Commands
            For Each p In cmd.Value.Parameters
                If p.Typ.BaseType.Name = "Enum" Then
                    If Not CustomTypes.Contains(p.Typ) Then CustomTypes.Add(p.Typ)
                End If
            Next
        Next

        ' Populate the enums worksheet with custom types
        EnumSheet.Cells("A1").Value = "Custom types"
        row = 2
        For Each ct In CustomTypes
            Dim enums As New List(Of String)
            For Each value In [Enum].GetValues(ct)
                Dim name = value.ToString
                Dim valu = CInt(value).ToString
                enums.Add($"{name}={valu}")
            Next
            EnumSheet.Rows(row).Cells("A").Value = ct.Name
            EnumSheet.Rows(row).Cells("B").Value = $"{String.Join(", ", enums)}"
            row += 1
        Next

        ' Autofit column widths for the enums worksheet
        columnCount = EnumSheet.CalculateMaxUsedColumns()
        For i As Integer = 0 To columnCount - 1
            EnumSheet.Columns(i).AutoFit(1, EnumSheet.Rows(1), EnumSheet.Rows(EnumSheet.Rows.Count - 1))
        Next

        ' Save the workbook to a file
        workbook.Save("commands.ods")
        TextBox1.AppendText($"Done{vbCrLf}")
    End Sub

    Shared Function A1(row As Integer, col As Integer) As String
        ' Convert row, col to A1 style reference
        Return Chr(Asc("A") + col) & row + 1
    End Function

    Private Sub TestCutlineToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TestCutlineToolStripMenuItem.Click
        Dim buffer() As Byte
        writer = New BinaryWriter(System.IO.File.Open("TestCutLine.mol", FileMode.Create), System.Text.Encoding.Unicode, False)

        ' Copy the first 5 blocks of a template file to initialise the new MOL file
        Dim ReadStream = System.IO.File.Open(MOLTemplate, FileMode.Open)           ' file containing template blocks
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
        ReadStream.Close()

        writer.BaseStream.Position = 5 * BLOCK_SIZE
        ' Now write the cutline
        position = New IntPoint(100, 200)
        delta = New IntPoint(6000, 8000)
        CutLinePrecise(delta, 50, 500)
        EndBlock()        ' pad to end of block with 0
        writer.Close()
        TextBox1.AppendText("Done")
    End Sub

    ''' <summary>
    ''' Pads the current block to the end with zeros to ensure it is the correct size.
    ''' </summary>
    Private Sub EndBlock()
        ' Get the current position in the stream
        Dim pos = writer.BaseStream.Position
        ' Calculate the number of 4-byte words needed to pad the block to the end
        Dim pad = (BLOCK_SIZE - (pos Mod BLOCK_SIZE)) / 4
        ' Write the padding words
        For i = 1 To pad
            WriteMOL(0)
        Next
    End Sub

    Private Sub FindCommandToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindCommandToolStripMenuItem.Click
        ' Search all .MOL files for specific command
        ' display location and parameters
        Dim cmd_len As Integer, target As Integer, count As Integer = 0, Mnemonic As String = "", chunks As List(Of Integer)

        Dim command As String = InputBox("Input command name or hex value", "Find command in all .MOL files")
        command = UCase(command)
        If command.StartsWith("0X") Then command = command.Remove(0, 2)         ' remove hex prefix
        ' First check if it's a mnemonic
        Dim entry = Commands.FirstOrDefault(Function(x) x.Value.Mnemonic.Equals(command, StringComparison.OrdinalIgnoreCase))
        If entry.Value IsNot Nothing Then
            target = entry.Key
        Else
            Try
                target = Convert.ToInt32(command, 16)
            Catch ex As Exception
                MsgBox($"Invalid input. {ex.Message}", vbExclamation + vbOKOnly, "Error")
                Exit Sub
            End Try
        End If
        Dim value As MOLcmd = Nothing
        If Commands.TryGetValue(target, value) Then Mnemonic = $" - ({value.Mnemonic})"     ' get mnemonic for command (if any)
        TextBox1.Clear()
        TextBox1.AppendText($"Searching for command 0x{target:x8}{Mnemonic}{vbCrLf}")
        Dim workingDirectory = Environment.CurrentDirectory
        Dim projectDirectory = Directory.GetParent(workingDirectory).Parent.Parent.FullName
        Dim molFiles As String() = Directory.GetFiles(projectDirectory, "*.mol", SearchOption.AllDirectories)

        For Each fi In molFiles
            TextBox1.AppendText($"Searching {fi}{vbCrLf}")
            Dim stream = System.IO.File.Open(fi, FileMode.Open)
            reader = New BinaryReader(stream, System.Text.Encoding.Unicode, False)
            GetHeader()
            ' Make list of chunks to be scanned
            chunks = New List(Of Integer) From {ConfigChunk, TestChunk, CutChunk}
            chunks.AddRange(DrawChunks)

            For Each chunk In chunks        ' look in all chunks
                reader.BaseStream.Position = chunk * BLOCK_SIZE         ' position the reader at the start of the chunk
                Dim addr = reader.BaseStream.Position
                Dim cmd = GetInt()
                While cmd <> MOL_END
                    cmd_len = cmd >> 24 And &HFF    ' length of command
                    If cmd_len = &H80 Then
                        cmd_len = GetInt() And &H1FF    ' length of command
                    End If
                    If cmd = target Then
                        TextBox1.AppendText($"Found at 0x{addr:x}: {GetBlockType(addr)} ({cmd_len}) ")
                        For i = 1 To cmd_len : TextBox1.AppendText($" 0x{GetUInt():x8}") : Next
                        TextBox1.AppendText(vbCrLf)
                        count += 1
                    Else
                        If cmd = MOL_MCBLK Then cmd_len = 1             ' process contents of MCBLK
                        For i = 1 To cmd_len : GetInt() : Next ' consume command
                    End If
                    addr = reader.BaseStream.Position
                    cmd = GetInt()
                End While
            Next
            reader.Close()
        Next
        TextBox1.AppendText($"Done. {molFiles.Length} files processed, {count} instances of command found{vbCrLf}")
    End Sub
End Class
