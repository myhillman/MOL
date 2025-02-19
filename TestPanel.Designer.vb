<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TestPanel
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TestPanel))
        TableLayoutPanel1 = New TableLayoutPanel()
        OK_Button = New Button()
        Cancel_Button = New Button()
        GroupBox1 = New GroupBox()
        tbSpeedMax = New TextBox()
        Label2 = New Label()
        Label1 = New Label()
        tbSpeedMin = New TextBox()
        GroupBox2 = New GroupBox()
        tbPowerMax = New TextBox()
        Label3 = New Label()
        Label4 = New Label()
        tbPowerMin = New TextBox()
        tbPasses = New TextBox()
        tbInterval = New TextBox()
        tbMaterial = New TextBox()
        Label5 = New Label()
        Label6 = New Label()
        Label7 = New Label()
        ErrorProvider1 = New ErrorProvider(components)
        CheckBox1 = New CheckBox()
        Label8 = New Label()
        Label9 = New Label()
        Label10 = New Label()
        tbFontPath = New TextBox()
        FolderBrowserDialog1 = New FolderBrowserDialog()
        btnFontFolder = New Button()
        cbFontFile = New ComboBox()
        TableLayoutPanel1.SuspendLayout()
        GroupBox1.SuspendLayout()
        GroupBox2.SuspendLayout()
        CType(ErrorProvider1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' TableLayoutPanel1
        ' 
        TableLayoutPanel1.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        TableLayoutPanel1.ColumnCount = 2
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50F))
        TableLayoutPanel1.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50F))
        TableLayoutPanel1.Controls.Add(OK_Button, 0, 0)
        TableLayoutPanel1.Controls.Add(Cancel_Button, 1, 0)
        TableLayoutPanel1.Location = New Point(442, 466)
        TableLayoutPanel1.Margin = New Padding(4, 3, 4, 3)
        TableLayoutPanel1.Name = "TableLayoutPanel1"
        TableLayoutPanel1.RowCount = 1
        TableLayoutPanel1.RowStyles.Add(New RowStyle(SizeType.Percent, 50F))
        TableLayoutPanel1.Size = New Size(170, 33)
        TableLayoutPanel1.TabIndex = 0
        ' 
        ' OK_Button
        ' 
        OK_Button.Anchor = AnchorStyles.None
        OK_Button.Location = New Point(4, 3)
        OK_Button.Margin = New Padding(4, 3, 4, 3)
        OK_Button.Name = "OK_Button"
        OK_Button.Size = New Size(77, 27)
        OK_Button.TabIndex = 0
        OK_Button.Text = "OK"
        ' 
        ' Cancel_Button
        ' 
        Cancel_Button.Anchor = AnchorStyles.None
        Cancel_Button.Location = New Point(89, 3)
        Cancel_Button.Margin = New Padding(4, 3, 4, 3)
        Cancel_Button.Name = "Cancel_Button"
        Cancel_Button.Size = New Size(77, 27)
        Cancel_Button.TabIndex = 1
        Cancel_Button.Text = "Cancel"
        ' 
        ' GroupBox1
        ' 
        GroupBox1.Controls.Add(tbSpeedMax)
        GroupBox1.Controls.Add(Label2)
        GroupBox1.Controls.Add(Label1)
        GroupBox1.Controls.Add(tbSpeedMin)
        GroupBox1.Location = New Point(25, 99)
        GroupBox1.Name = "GroupBox1"
        GroupBox1.Size = New Size(154, 87)
        GroupBox1.TabIndex = 1
        GroupBox1.TabStop = False
        GroupBox1.Text = "Speed (mm/s)"
        ' 
        ' tbSpeedMax
        ' 
        tbSpeedMax.Location = New Point(87, 48)
        tbSpeedMax.Name = "tbSpeedMax"
        tbSpeedMax.Size = New Size(52, 23)
        tbSpeedMax.TabIndex = 3
        tbSpeedMax.Tag = "SpeedMax"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New Point(15, 51)
        Label2.Name = "Label2"
        Label2.Size = New Size(29, 15)
        Label2.TabIndex = 2
        Label2.Text = "Max"
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New Point(15, 22)
        Label1.Name = "Label1"
        Label1.Size = New Size(28, 15)
        Label1.TabIndex = 1
        Label1.Text = "Min"
        ' 
        ' tbSpeedMin
        ' 
        tbSpeedMin.Location = New Point(87, 19)
        tbSpeedMin.Name = "tbSpeedMin"
        tbSpeedMin.Size = New Size(52, 23)
        tbSpeedMin.TabIndex = 0
        tbSpeedMin.Tag = "SpeedMin"
        ' 
        ' GroupBox2
        ' 
        GroupBox2.Controls.Add(tbPowerMax)
        GroupBox2.Controls.Add(Label3)
        GroupBox2.Controls.Add(Label4)
        GroupBox2.Controls.Add(tbPowerMin)
        GroupBox2.Location = New Point(25, 203)
        GroupBox2.Name = "GroupBox2"
        GroupBox2.Size = New Size(154, 87)
        GroupBox2.TabIndex = 2
        GroupBox2.TabStop = False
        GroupBox2.Text = "Power (%)"
        ' 
        ' tbPowerMax
        ' 
        tbPowerMax.Location = New Point(87, 51)
        tbPowerMax.Name = "tbPowerMax"
        tbPowerMax.Size = New Size(52, 23)
        tbPowerMax.TabIndex = 3
        tbPowerMax.Tag = "PowerMax"
        ' 
        ' Label3
        ' 
        Label3.AutoSize = True
        Label3.Location = New Point(15, 54)
        Label3.Name = "Label3"
        Label3.Size = New Size(29, 15)
        Label3.TabIndex = 2
        Label3.Text = "Max"
        ' 
        ' Label4
        ' 
        Label4.AutoSize = True
        Label4.Location = New Point(15, 22)
        Label4.Name = "Label4"
        Label4.Size = New Size(28, 15)
        Label4.TabIndex = 1
        Label4.Text = "Min"
        ' 
        ' tbPowerMin
        ' 
        tbPowerMin.Location = New Point(87, 22)
        tbPowerMin.Name = "tbPowerMin"
        tbPowerMin.Size = New Size(52, 23)
        tbPowerMin.TabIndex = 0
        tbPowerMin.Tag = "PowerMin"
        ' 
        ' tbPasses
        ' 
        tbPasses.Location = New Point(112, 70)
        tbPasses.Name = "tbPasses"
        tbPasses.Size = New Size(52, 23)
        tbPasses.TabIndex = 3
        tbPasses.Tag = "Passes"
        ' 
        ' tbInterval
        ' 
        tbInterval.Location = New Point(112, 41)
        tbInterval.Name = "tbInterval"
        tbInterval.Size = New Size(52, 23)
        tbInterval.TabIndex = 4
        tbInterval.Tag = "Interval"
        ' 
        ' tbMaterial
        ' 
        tbMaterial.Location = New Point(112, 12)
        tbMaterial.Name = "tbMaterial"
        tbMaterial.Size = New Size(251, 23)
        tbMaterial.TabIndex = 5
        tbMaterial.Tag = "Material"
        ' 
        ' Label5
        ' 
        Label5.AutoSize = True
        Label5.Location = New Point(25, 18)
        Label5.Name = "Label5"
        Label5.Size = New Size(50, 15)
        Label5.TabIndex = 6
        Label5.Text = "Material"
        ' 
        ' Label6
        ' 
        Label6.AutoSize = True
        Label6.Location = New Point(25, 44)
        Label6.Name = "Label6"
        Label6.Size = New Size(46, 15)
        Label6.TabIndex = 7
        Label6.Text = "Interval"
        ' 
        ' Label7
        ' 
        Label7.AutoSize = True
        Label7.Location = New Point(25, 73)
        Label7.Name = "Label7"
        Label7.Size = New Size(41, 15)
        Label7.TabIndex = 8
        Label7.Text = "Passes"
        ' 
        ' ErrorProvider1
        ' 
        ErrorProvider1.ContainerControl = Me
        ' 
        ' CheckBox1
        ' 
        CheckBox1.AutoSize = True
        CheckBox1.Location = New Point(112, 318)
        CheckBox1.Name = "CheckBox1"
        CheckBox1.Size = New Size(15, 14)
        CheckBox1.TabIndex = 9
        CheckBox1.UseVisualStyleBackColor = True
        ' 
        ' Label8
        ' 
        Label8.AutoSize = True
        Label8.Location = New Point(33, 318)
        Label8.Name = "Label8"
        Label8.Size = New Size(49, 15)
        Label8.TabIndex = 10
        Label8.Text = "Engrave"
        ' 
        ' Label9
        ' 
        Label9.AutoSize = True
        Label9.Location = New Point(33, 355)
        Label9.Name = "Label9"
        Label9.Size = New Size(58, 15)
        Label9.TabIndex = 11
        Label9.Text = "Font path"
        ' 
        ' Label10
        ' 
        Label10.AutoSize = True
        Label10.Location = New Point(33, 389)
        Label10.Name = "Label10"
        Label10.Size = New Size(50, 15)
        Label10.TabIndex = 12
        Label10.Text = "Font file"
        ' 
        ' tbFontPath
        ' 
        tbFontPath.Location = New Point(112, 352)
        tbFontPath.Name = "tbFontPath"
        tbFontPath.Size = New Size(435, 23)
        tbFontPath.TabIndex = 13
        ' 
        ' btnFontFolder
        ' 
        btnFontFolder.Image = CType(resources.GetObject("btnFontFolder.Image"), Image)
        btnFontFolder.Location = New Point(553, 339)
        btnFontFolder.Name = "btnFontFolder"
        btnFontFolder.Size = New Size(55, 47)
        btnFontFolder.TabIndex = 15
        btnFontFolder.Text = "Path"
        btnFontFolder.UseVisualStyleBackColor = True
        ' 
        ' cbFontFile
        ' 
        cbFontFile.FormattingEnabled = True
        cbFontFile.Location = New Point(112, 389)
        cbFontFile.Name = "cbFontFile"
        cbFontFile.Size = New Size(203, 23)
        cbFontFile.TabIndex = 16
        ' 
        ' TestPanel
        ' 
        AcceptButton = OK_Button
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        CancelButton = Cancel_Button
        ClientSize = New Size(626, 513)
        Controls.Add(cbFontFile)
        Controls.Add(btnFontFolder)
        Controls.Add(tbFontPath)
        Controls.Add(Label10)
        Controls.Add(Label9)
        Controls.Add(Label8)
        Controls.Add(CheckBox1)
        Controls.Add(Label7)
        Controls.Add(Label6)
        Controls.Add(Label5)
        Controls.Add(tbMaterial)
        Controls.Add(tbInterval)
        Controls.Add(tbPasses)
        Controls.Add(GroupBox2)
        Controls.Add(GroupBox1)
        Controls.Add(TableLayoutPanel1)
        FormBorderStyle = FormBorderStyle.FixedDialog
        Margin = New Padding(4, 3, 4, 3)
        MaximizeBox = False
        MinimizeBox = False
        Name = "TestPanel"
        ShowInTaskbar = False
        StartPosition = FormStartPosition.CenterParent
        Text = "Test Panel"
        TableLayoutPanel1.ResumeLayout(False)
        GroupBox1.ResumeLayout(False)
        GroupBox1.PerformLayout()
        GroupBox2.ResumeLayout(False)
        GroupBox2.PerformLayout()
        CType(ErrorProvider1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OK_Button As System.Windows.Forms.Button
    Friend WithEvents Cancel_Button As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents tbSpeedMax As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents tbSpeedMin As TextBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents tbPowerMax As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents tbPowerMin As TextBox
    Friend WithEvents tbPasses As TextBox
    Friend WithEvents tbInterval As TextBox
    Friend WithEvents tbMaterial As TextBox
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents ErrorProvider1 As ErrorProvider
    Friend WithEvents Label8 As Label
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Label10 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents tbFontPath As TextBox
    Friend WithEvents btnFontFolder As Button
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents cbFontFile As ComboBox

End Class
