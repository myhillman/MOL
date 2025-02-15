<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        MenuStrip1 = New MenuStrip()
        OpenToolStripMenuItem = New ToolStripMenuItem()
        DisassembleToolStripMenuItem = New ToolStripMenuItem()
        RenderToolStripMenuItem = New ToolStripMenuItem()
        TestCardParametersToolStripMenuItem = New ToolStripMenuItem()
        MakeTestCardToolStripMenuItem = New ToolStripMenuItem()
        DXFMetricsToolStripMenuItem = New ToolStripMenuItem()
        ReconstructMINIMARIOMOLToolStripMenuItem = New ToolStripMenuItem()
        CommandSummaryToolStripMenuItem = New ToolStripMenuItem()
        ExitToolStripMenuItem = New ToolStripMenuItem()
        DebugToolStripMenuItem = New ToolStripMenuItem()
        HEXDUMPONToolStripMenuItem = New ToolStripMenuItem()
        HEXDUMPOFFToolStripMenuItem = New ToolStripMenuItem()
        DEBUGONToolStripMenuItem = New ToolStripMenuItem()
        DEBUGOFFToolStripMenuItem = New ToolStripMenuItem()
        TestToolStripMenuItem = New ToolStripMenuItem()
        WordAccessToolStripMenuItem = New ToolStripMenuItem()
        TestIEEEToLeetroFpConversionToolStripMenuItem = New ToolStripMenuItem()
        TextTestToolStripMenuItem = New ToolStripMenuItem()
        TextBox1 = New TextBox()
        ProgressBar1 = New ProgressBar()
        MenuStrip1.SuspendLayout()
        SuspendLayout()
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.Items.AddRange(New ToolStripItem() {OpenToolStripMenuItem, DebugToolStripMenuItem, TestToolStripMenuItem})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Size = New Size(1047, 24)
        MenuStrip1.TabIndex = 0
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' OpenToolStripMenuItem
        ' 
        OpenToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {DisassembleToolStripMenuItem, RenderToolStripMenuItem, TestCardParametersToolStripMenuItem, MakeTestCardToolStripMenuItem, DXFMetricsToolStripMenuItem, ReconstructMINIMARIOMOLToolStripMenuItem, CommandSummaryToolStripMenuItem, ExitToolStripMenuItem})
        OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        OpenToolStripMenuItem.Size = New Size(37, 20)
        OpenToolStripMenuItem.Text = "File"
        ' 
        ' DisassembleToolStripMenuItem
        ' 
        DisassembleToolStripMenuItem.Name = "DisassembleToolStripMenuItem"
        DisassembleToolStripMenuItem.Size = New Size(233, 22)
        DisassembleToolStripMenuItem.Text = "Disassemble"
        ' 
        ' RenderToolStripMenuItem
        ' 
        RenderToolStripMenuItem.Name = "RenderToolStripMenuItem"
        RenderToolStripMenuItem.Size = New Size(233, 22)
        RenderToolStripMenuItem.Text = "Render"
        ' 
        ' TestCardParametersToolStripMenuItem
        ' 
        TestCardParametersToolStripMenuItem.Name = "TestCardParametersToolStripMenuItem"
        TestCardParametersToolStripMenuItem.Size = New Size(233, 22)
        TestCardParametersToolStripMenuItem.Text = "Test card parameters"
        ' 
        ' MakeTestCardToolStripMenuItem
        ' 
        MakeTestCardToolStripMenuItem.Name = "MakeTestCardToolStripMenuItem"
        MakeTestCardToolStripMenuItem.Size = New Size(233, 22)
        MakeTestCardToolStripMenuItem.Text = "Make Test card"
        ' 
        ' DXFMetricsToolStripMenuItem
        ' 
        DXFMetricsToolStripMenuItem.Name = "DXFMetricsToolStripMenuItem"
        DXFMetricsToolStripMenuItem.Size = New Size(233, 22)
        DXFMetricsToolStripMenuItem.Text = "DXF metrics"
        ' 
        ' ReconstructMINIMARIOMOLToolStripMenuItem
        ' 
        ReconstructMINIMARIOMOLToolStripMenuItem.Name = "ReconstructMINIMARIOMOLToolStripMenuItem"
        ReconstructMINIMARIOMOLToolStripMenuItem.Size = New Size(233, 22)
        ReconstructMINIMARIOMOLToolStripMenuItem.Text = "Reconstruct MINIMARIO.MOL"
        ' 
        ' CommandSummaryToolStripMenuItem
        ' 
        CommandSummaryToolStripMenuItem.Name = "CommandSummaryToolStripMenuItem"
        CommandSummaryToolStripMenuItem.Size = New Size(233, 22)
        CommandSummaryToolStripMenuItem.Text = "Command summary"
        ' 
        ' ExitToolStripMenuItem
        ' 
        ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        ExitToolStripMenuItem.Size = New Size(233, 22)
        ExitToolStripMenuItem.Text = "Exit"
        ' 
        ' DebugToolStripMenuItem
        ' 
        DebugToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {HEXDUMPONToolStripMenuItem, HEXDUMPOFFToolStripMenuItem, DEBUGONToolStripMenuItem, DEBUGOFFToolStripMenuItem})
        DebugToolStripMenuItem.Name = "DebugToolStripMenuItem"
        DebugToolStripMenuItem.Size = New Size(54, 20)
        DebugToolStripMenuItem.Text = "Debug"
        ' 
        ' HEXDUMPONToolStripMenuItem
        ' 
        HEXDUMPONToolStripMenuItem.Name = "HEXDUMPONToolStripMenuItem"
        HEXDUMPONToolStripMenuItem.Size = New Size(154, 22)
        HEXDUMPONToolStripMenuItem.Text = "HEXDUMP ON"
        ' 
        ' HEXDUMPOFFToolStripMenuItem
        ' 
        HEXDUMPOFFToolStripMenuItem.Name = "HEXDUMPOFFToolStripMenuItem"
        HEXDUMPOFFToolStripMenuItem.Size = New Size(154, 22)
        HEXDUMPOFFToolStripMenuItem.Text = "HEXDUMP OFF"
        ' 
        ' DEBUGONToolStripMenuItem
        ' 
        DEBUGONToolStripMenuItem.Name = "DEBUGONToolStripMenuItem"
        DEBUGONToolStripMenuItem.Size = New Size(154, 22)
        DEBUGONToolStripMenuItem.Text = "DEBUG ON"
        ' 
        ' DEBUGOFFToolStripMenuItem
        ' 
        DEBUGOFFToolStripMenuItem.Name = "DEBUGOFFToolStripMenuItem"
        DEBUGOFFToolStripMenuItem.Size = New Size(154, 22)
        DEBUGOFFToolStripMenuItem.Text = "DEBUG OFF"
        ' 
        ' TestToolStripMenuItem
        ' 
        TestToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {WordAccessToolStripMenuItem, TestIEEEToLeetroFpConversionToolStripMenuItem, TextTestToolStripMenuItem})
        TestToolStripMenuItem.Name = "TestToolStripMenuItem"
        TestToolStripMenuItem.Size = New Size(40, 20)
        TestToolStripMenuItem.Text = "Test"
        ' 
        ' WordAccessToolStripMenuItem
        ' 
        WordAccessToolStripMenuItem.Name = "WordAccessToolStripMenuItem"
        WordAccessToolStripMenuItem.Size = New Size(244, 22)
        WordAccessToolStripMenuItem.Text = "Word access"
        ' 
        ' TestIEEEToLeetroFpConversionToolStripMenuItem
        ' 
        TestIEEEToLeetroFpConversionToolStripMenuItem.Name = "TestIEEEToLeetroFpConversionToolStripMenuItem"
        TestIEEEToLeetroFpConversionToolStripMenuItem.Size = New Size(244, 22)
        TestIEEEToLeetroFpConversionToolStripMenuItem.Text = "Test IEEE to Leetro fp conversion"
        ' 
        ' TextTestToolStripMenuItem
        ' 
        TextTestToolStripMenuItem.Name = "TextTestToolStripMenuItem"
        TextTestToolStripMenuItem.Size = New Size(244, 22)
        TextTestToolStripMenuItem.Text = "Text test"
        ' 
        ' TextBox1
        ' 
        TextBox1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        TextBox1.Location = New Point(12, 31)
        TextBox1.Multiline = True
        TextBox1.Name = "TextBox1"
        TextBox1.ScrollBars = ScrollBars.Both
        TextBox1.Size = New Size(1023, 554)
        TextBox1.TabIndex = 1
        ' 
        ' ProgressBar1
        ' 
        ProgressBar1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        ProgressBar1.Location = New Point(12, 591)
        ProgressBar1.Name = "ProgressBar1"
        ProgressBar1.Size = New Size(1023, 21)
        ProgressBar1.TabIndex = 2
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1047, 618)
        Controls.Add(ProgressBar1)
        Controls.Add(TextBox1)
        Controls.Add(MenuStrip1)
        MainMenuStrip = MenuStrip1
        MinimizeBox = False
        Name = "Form1"
        Text = "MOL"
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents OpenToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TestToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents WordAccessToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DebugToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HEXDUMPONToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HEXDUMPOFFToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DEBUGONToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DEBUGOFFToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MakeTestCardToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TestIEEEToLeetroFpConversionToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents TestCardParametersToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TextTestToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DXFMetricsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ReconstructMINIMARIOMOLToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CommandSummaryToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DisassembleToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RenderToolStripMenuItem As ToolStripMenuItem

End Class
