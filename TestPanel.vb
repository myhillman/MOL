Imports System.ComponentModel
Imports System.Windows.Forms

Public Class TestPanel
    Dim material As String, interval As Single, passes As Integer, speedMin As Integer, speedMax As Integer, powerMin As Integer, powerMax As Integer
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        ' All parameters are validated. Save them
        My.Settings.Material = Me.tbMaterial.Text
        My.Settings.Interval = Me.tbInterval.Text
        My.Settings.Passes = Me.tbPasses.Text
        My.Settings.SpeedMin = Me.tbSpeedMin.Text
        My.Settings.SpeedMax = Me.tbSpeedMax.Text
        My.Settings.PowerMin = CSng(Me.tbPowerMin.Text)
        My.Settings.PowerMax = CSng(Me.tbPowerMax.Text)
        My.Settings.Engrave = Me.CheckBox1.Checked
        My.Settings.Save()
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
            .tbPasses.Text = My.Settings.Passes
            .tbSpeedMin.Text = My.Settings.SpeedMin
            .tbSpeedMax.Text = My.Settings.SpeedMax
            .tbPowerMin.Text = My.Settings.PowerMin
            .tbPowerMax.Text = My.Settings.PowerMax
            .CheckBox1.Checked = My.Settings.Engrave
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
        Const min = 1, max = 10
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


End Class
