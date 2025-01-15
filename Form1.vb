Public Class Form1
    Const BLOCK_SIZE = 512
    Const STEP_SIZE = 208.33      ' steps/mm
    Dim data() As Byte
    Dim index As Integer = 0      ' index into data
    Dim filesize As Integer
    Private Sub OpenToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem1.Click
        Dim dlg = New OpenFileDialog
        With dlg
            .Filter = "MOL file|*.mol|All files|*.*"
        End With
        Dim result = dlg.ShowDialog()
        If result.OK Then
            data = My.Computer.FileSystem.ReadAllBytes(dlg.FileName)
            DisplayHeader()
            DisplayConfig()
            DisplayTest()
            DisplayCut()
            DisplayDraw()
        End If

    End Sub
    Function getint() As Integer
        ' read 4 byte integer from data(index)
        Dim result As Integer
        result = BitConverter.ToInt32({data(index), data(index + 1), data(index + 2), data(index + 3)})
        index += 4      ' increment to next int
        Return result
    End Function
    Function getint(n As Integer) As Integer
        index = n
        Return getint()
    End Function

    Function getfloat() As Single
        ' read 4 byte float from data
        ' Leetro float format  [eeeeeeee|smmmmmmm|mmmmmmm0|00000000]
        ' TODO: needs to be tested
        Dim b = getint()   ' get the float
        Dim vExp As Integer = b >> 24 And &HFF      ' exponent
        Dim vMan As Integer = b >> 9 And &H7FFF     ' mantissa
        Dim vSgn As Integer = b >> 23 And 1         ' sign
        If vSgn > 0 Then vSgn = -1 Else vSgn = 1
        Return (vMan / 16384.0) * Math.Pow(2, vExp) * vSgn
    End Function
    Sub DisplayHeader()
        ' Display header info
        TextBox1.AppendText($"HEADER &0{vbCrLf}")
        filesize = getint()
        TextBox1.AppendText($"Size of file: &h{filesize:x}{vbCrLf}")
        TextBox1.AppendText($"Version: {data(11)}.{data(10)}.{data(9)}.{data(8)}{vbCrLf}")
        TextBox1.AppendText($"Origin: X {getint(&H18)} ({getint(&H18) / STEP_SIZE:f1}mm), Y {getint(&H1C)} ({getint(&H1C) / STEP_SIZE:f1}mm){vbCrLf}")
        TextBox1.AppendText($"Bottom Left: X {getint(&H20)} ({getint(&H20) / STEP_SIZE:f1}mm), Y {getint(&H24)} ({getint(&H24) / STEP_SIZE:f1}mm){vbCrLf}")
        TextBox1.AppendText($"Config chunk: {getint(&H70)}{vbCrLf}")
        TextBox1.AppendText($"Test chunk: {getint(&H74)}{vbCrLf}")
        TextBox1.AppendText($"Cut chunk: {getint(&H78)}{vbCrLf}")
        TextBox1.AppendText($"Draw chunk: {getint(&H7C)}{vbCrLf}")
        TextBox1.AppendText($"{vbCrLf}")
    End Sub

    Sub DisplayConfig()
        ' Display config block info
        Dim StartOfBlock = getint(&H70) * BLOCK_SIZE
        TextBox1.AppendText($"CONFIG @&{StartOfBlock:x}{vbCrLf}")
        index = StartOfBlock
        For i = 1 To 8
            TextBox1.AppendText($"Item {Index:x}: {getint():x}{vbCrLf}")
        Next
        TextBox1.AppendText($"{vbCrLf}")
    End Sub
    Sub DisplayTest()
        ' Display test block info
        Dim StartOfBlock = getint(&H74) * BLOCK_SIZE
        TextBox1.AppendText($"TEST @&{StartOfBlock:x}{vbCrLf}")
        HexDumpBlock(StartOfBlock)
        TextBox1.AppendText($"{vbCrLf}")
    End Sub

    Sub DisplayCut()
        ' Display cut block info
        Dim StartOfBlock = getint(&H78) * BLOCK_SIZE
        If StartOfBlock > 0 Then
            TextBox1.AppendText($"CUT @&{StartOfBlock:x}{vbCrLf}")
            HexDumpBlock(StartOfBlock)
            TextBox1.AppendText($"{vbCrLf}")
        End If
    End Sub
    Sub DisplayDraw()
        ' Display cut block info
        Dim StartOfBlock = getint(&H7C) * BLOCK_SIZE
        TextBox1.AppendText($"DRAW @&{StartOfBlock:x}{vbCrLf}")
        HexDumpBlock(StartOfBlock)
        Dim done As Boolean = False
        Dim LaserIsOn As Boolean = False
        Dim spdMin As Single = 100
        Dim spdMax As Single = 100
        Dim px As Integer = 0
        Dim py As Integer = 0
        Dim xyScale As Single = 208.33

        index = StartOfBlock
        While Not done
            Dim cmd = getint()
            TextBox1.AppendText($"{index - 4:x} CMD {cmd}{vbCrLf}")
            Select Case cmd
                Case &H80000946 ' begin motion block
                    getint()
                    TextBox1.AppendText($"Begin motion block{vbCrLf}")
                Case &H1000606     ' switch laser
                    LaserIsOn = getint()
                    TextBox1.AppendText($"Switching laser {LaserIsOn}{vbCrLf}")
                Case &H3000301 ' set speeds
                    spdMin = getint() / xyScale
                    spdMax = getint() / xyScale
                    getfloat() ' acceleration
                    TextBox1.AppendText($"Set speed Min={spdMin}, Max={spdMax}{vbCrLf}")
                Case &H3026000     ' move relative
                    Dim axis = getint()
                    Dim dx = getint()
                    Dim dy = getint()
                    If LaserIsOn Then
                        ' Now draw 
                        TextBox1.AppendText($"Moving relative dx={dx}, dy={dy}{vbCrLf}")
                    End If
                    px += dx
                    py += dy
                Case &H1400048, &H0
                    done = True
                Case Else
                    Dim nWords = cmd >> 24
                    If nWords = &H80 Or nWords = -128 Then
                        nWords = getint()
                        TextBox1.AppendText($"Skipping {nWords} words at {index:x}{vbCrLf}")
                        index += 4 * nWords
                    End If
            End Select
        End While
        TextBox1.AppendText($"{vbCrLf}")
    End Sub

    Sub HexDumpBlock(addr As Integer)
        ' dump a block in hex
        Dim done As Boolean = False
        index = addr
        While Not done And index <= filesize
            TextBox1.AppendText($"&h{index:x8}: ")
            Dim cmd As Integer = getint()
            If cmd = 0 Then Exit While
            TextBox1.AppendText($"&h{cmd:x8}")
            Dim nWords = cmd >> 24 And &H7F
            For i = 1 To nWords
                cmd = getint()
                TextBox1.AppendText($" &h{cmd:x8}")
            Next
            TextBox1.AppendText($"{vbCrLf}")
        End While
    End Sub
End Class
