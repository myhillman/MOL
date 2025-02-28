Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Controls

Public Class TestPanel
    Private material As String, interval As Single, passes As Integer, speedMin As Integer, speedMax As Integer, powerMin As Integer, powerMax As Integer
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        ' All parameters are validated. Save them
        My.Settings.Material = tbMaterial.Text
        My.Settings.Interval = tbInterval.Text
        My.Settings.Passes = tbPasses.Text
        My.Settings.SpeedMin = tbSpeedMin.Text
        My.Settings.SpeedMax = tbSpeedMax.Text
        My.Settings.PowerMin = CSng(tbPowerMin.Text)
        My.Settings.PowerMax = CSng(tbPowerMax.Text)
        My.Settings.Engrave = CheckBox1.Checked
        My.Settings.FontPath = tbFontPath.Text
        Dim FontChanged = (My.Settings.FontFile <> cbFontFile.Text)
        My.Settings.FontFile = cbFontFile.Text
        My.Settings.Save()
        If FontChanged Then Form1.LoadFonts()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub TestPanel_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Load form with existing values
        With Me
            .tbMaterial.Text = My.Settings.Material
            .tbInterval.Text = My.Settings.Interval
            .tbInterval.Enabled = My.Settings.Engrave    ' Interval only relevant for engraving
            .tbPasses.Text = My.Settings.Passes
            .tbSpeedMin.Text = My.Settings.SpeedMin
            .tbSpeedMax.Text = My.Settings.SpeedMax
            .tbPowerMin.Text = My.Settings.PowerMin
            .tbPowerMax.Text = My.Settings.PowerMax
            .CheckBox1.Checked = My.Settings.Engrave
            .tbFontPath.Text = My.Settings.FontPath
            ' Fill the list box
            cbFontFile.Items.Clear()
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(My.Settings.FontPath, FileIO.SearchOption.SearchTopLevelOnly, "*.lff")
                Dim filename = Path.GetFileName(foundFile)
                cbFontFile.Items.Add(filename)
            Next
            .cbFontFile.SelectedIndex = cbFontFile.FindStringExact(My.Settings.FontFile)  ' select the current font file
        End With
    End Sub
    Private Sub tbMaterial_Validating(sender As Object, e As CancelEventArgs) Handles tbMaterial.Validating
        If tbMaterial.Text.Length = 0 Then
            e.Cancel = True
            ErrorProvider1.SetError(sender, "Material cannot be blank")
        End If
    End Sub

    Private Sub tbMaterial_Validated(sender As Object, e As EventArgs) Handles tbMaterial.Validated
        ErrorProvider1.SetError(sender, "")     ' remove error message (if any)
    End Sub

    Private Sub tbInterval_Validated(sender As Object, e As EventArgs) Handles tbInterval.Validated
        ErrorProvider1.SetError(sender, "")     ' remove error message (if any)
    End Sub

    Private Sub tbInterval_Validating(sender As Object, e As CancelEventArgs) Handles tbInterval.Validating
        If Not IsNumeric(tbInterval.Text) Then
            e.Cancel = True
            ErrorProvider1.SetError(sender, "Field must be a number")
            Exit Sub
        End If
        Const min = 0.1, max = 1
        If CSng(tbInterval.Text) < min Or CSng(tbInterval.Text) > max Then
            ErrorProvider1.SetError(sender, $"Value must be between {min} and {max}")
            e.Cancel = True
        End If
    End Sub

    Private Sub tbPasses_Validating(sender As Object, e As CancelEventArgs) Handles tbPasses.Validating
        Dim result As Integer
        e.Cancel = Not Int32.TryParse(sender.Text, result)
        If e.Cancel Then
            ErrorProvider1.SetError(sender, "Field must be an integer")
            Exit Sub
        End If
        Const min = 1, max = 5
        If result < min Or result > max Then
            ErrorProvider1.SetError(sender, $"Value must be between {min} and {max}")
            e.Cancel = True
        End If
    End Sub

    Private Sub tbPasses_Validated(sender As Object, e As EventArgs) Handles tbPasses.Validated
        ErrorProvider1.SetError(sender, "")     ' remove error message (if any)
    End Sub

    Private Sub tbSpeedMin_Validating(sender As Object, e As CancelEventArgs) Handles tbSpeedMin.Validating
        Dim result As Integer
        e.Cancel = Not Int32.TryParse(sender.Text, result)
        If e.Cancel Then
            ErrorProvider1.SetError(sender, "Field must be an integer")
            Exit Sub
        End If
        Dim min = 1, max = CInt(sender.Text)
        If result < min Or result > max Then
            ErrorProvider1.SetError(sender, $"Value must be between {min} and {max}")
            e.Cancel = True
        End If
    End Sub

    Private Sub tbSpeedMax_Validating(sender As Object, e As CancelEventArgs) Handles tbSpeedMax.Validating
        Dim result As Integer
        e.Cancel = Not Int32.TryParse(sender.Text, result)
        If e.Cancel Then
            ErrorProvider1.SetError(sender, "Field must be an integer")
            Exit Sub
        End If
        Dim min = CInt(sender.Text), max = 500
        If result < min Or result > max Then
            ErrorProvider1.SetError(sender, $"Value must be between {min} and {max}")
            e.Cancel = True
        End If
    End Sub

    Private Sub tbSpeedMin_Validated(sender As Object, e As EventArgs) Handles tbSpeedMin.Validated
        ErrorProvider1.SetError(sender, "")     ' remove error message (if any)
    End Sub

    Private Sub tbSpeedMax_Validated(sender As Object, e As EventArgs) Handles tbSpeedMax.Validated
        ErrorProvider1.SetError(sender, "")     ' remove error message (if any)
    End Sub

    Private Sub tbPowerMin_Validating(sender As Object, e As CancelEventArgs) Handles tbPowerMin.Validating
        Dim result As Integer
        e.Cancel = Not Int32.TryParse(sender.Text, result)
        If e.Cancel Then
            ErrorProvider1.SetError(sender, "Field must be an integer")
            Exit Sub
        End If
        Dim min = 1, max = CInt(sender.Text)
        If result < min Or result > max Then
            ErrorProvider1.SetError(sender, $"Value must be between {min} and {max}")
            e.Cancel = True
        End If
    End Sub

    Private Sub tbPowerMax_Validating(sender As Object, e As CancelEventArgs) Handles tbPowerMax.Validating
        Dim result As Integer
        e.Cancel = Not Int32.TryParse(sender.Text, result)
        If e.Cancel Then
            ErrorProvider1.SetError(sender, "Field must be an integer")
            Exit Sub
        End If
        Dim min = CInt(sender.Text), max = 100
        If result < min Or result > max Then
            ErrorProvider1.SetError(sender, $"Value must be between {min} and {max}")
            e.Cancel = True
        End If
    End Sub

    Private Sub tbPowerMin_Validated(sender As Object, e As EventArgs) Handles tbPowerMin.Validated
        ErrorProvider1.SetError(sender, "")     ' remove error message (if any)
    End Sub

    Private Sub tbPowerMax_Validated(sender As Object, e As EventArgs) Handles tbPowerMax.Validated
        ErrorProvider1.SetError(sender, "")     ' remove error message (if any)
    End Sub

    Private Sub btnFontFolder_Click(sender As Object, e As EventArgs) Handles btnFontFolder.Click

        If FolderBrowserDialog1.ShowDialog() = vbOK Then
            ' User has selected path
            My.Settings.FontPath = FolderBrowserDialog1.SelectedPath
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        tbInterval.Enabled = Me.CheckBox1.Checked
    End Sub
End Class
