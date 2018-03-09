<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
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
        Me.IDD_ToolsMenuStrip = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrepareIDDInputCollectorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MergeIDDInputsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FormatDeliveredIDDToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AnnotateIDDToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GenerateAnnotationPagesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CompareWorkbooksToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.statusStrip = New System.Windows.Forms.StatusStrip()
        Me.statusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.WordToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AcronymBuilderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IDD_ToolsMenuStrip.SuspendLayout()
        Me.statusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'IDD_ToolsMenuStrip
        '
        Me.IDD_ToolsMenuStrip.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.IDD_ToolsMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.WordToolsToolStripMenuItem})
        Me.IDD_ToolsMenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.IDD_ToolsMenuStrip.Name = "IDD_ToolsMenuStrip"
        Me.IDD_ToolsMenuStrip.Padding = New System.Windows.Forms.Padding(5, 1, 0, 1)
        Me.IDD_ToolsMenuStrip.Size = New System.Drawing.Size(1480, 26)
        Me.IDD_ToolsMenuStrip.TabIndex = 0
        Me.IDD_ToolsMenuStrip.Text = "IDD Tools"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(44, 24)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.Q), System.Windows.Forms.Keys)
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(161, 26)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PrepareIDDInputCollectorToolStripMenuItem, Me.MergeIDDInputsToolStripMenuItem, Me.FormatDeliveredIDDToolStripMenuItem, Me.AnnotateIDDToolStripMenuItem, Me.GenerateAnnotationPagesToolStripMenuItem, Me.CompareWorkbooksToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(94, 24)
        Me.ToolsToolStripMenuItem.Text = "Excel Tools"
        '
        'PrepareIDDInputCollectorToolStripMenuItem
        '
        Me.PrepareIDDInputCollectorToolStripMenuItem.Name = "PrepareIDDInputCollectorToolStripMenuItem"
        Me.PrepareIDDInputCollectorToolStripMenuItem.Size = New System.Drawing.Size(267, 26)
        Me.PrepareIDDInputCollectorToolStripMenuItem.Text = "Prepare IDD Input Collector"
        Me.PrepareIDDInputCollectorToolStripMenuItem.ToolTipText = "Generates a workbook for collecting IDD inputs from stakeholders"
        '
        'MergeIDDInputsToolStripMenuItem
        '
        Me.MergeIDDInputsToolStripMenuItem.Name = "MergeIDDInputsToolStripMenuItem"
        Me.MergeIDDInputsToolStripMenuItem.Size = New System.Drawing.Size(267, 26)
        Me.MergeIDDInputsToolStripMenuItem.Text = "Merge IDD Inputs"
        Me.MergeIDDInputsToolStripMenuItem.ToolTipText = "Merges data in the preparation workbook for capture into the data model"
        '
        'FormatDeliveredIDDToolStripMenuItem
        '
        Me.FormatDeliveredIDDToolStripMenuItem.Name = "FormatDeliveredIDDToolStripMenuItem"
        Me.FormatDeliveredIDDToolStripMenuItem.Size = New System.Drawing.Size(267, 26)
        Me.FormatDeliveredIDDToolStripMenuItem.Text = "Format Delivered IDD"
        Me.FormatDeliveredIDDToolStripMenuItem.ToolTipText = "Reformats formal IDD to be more usable"
        '
        'AnnotateIDDToolStripMenuItem
        '
        Me.AnnotateIDDToolStripMenuItem.Name = "AnnotateIDDToolStripMenuItem"
        Me.AnnotateIDDToolStripMenuItem.Size = New System.Drawing.Size(267, 26)
        Me.AnnotateIDDToolStripMenuItem.Text = "Annotate IDD"
        Me.AnnotateIDDToolStripMenuItem.ToolTipText = "Manages data in the program specific annotation database"
        '
        'GenerateAnnotationPagesToolStripMenuItem
        '
        Me.GenerateAnnotationPagesToolStripMenuItem.Name = "GenerateAnnotationPagesToolStripMenuItem"
        Me.GenerateAnnotationPagesToolStripMenuItem.Size = New System.Drawing.Size(267, 26)
        Me.GenerateAnnotationPagesToolStripMenuItem.Text = "Generate Annotation Pages"
        Me.GenerateAnnotationPagesToolStripMenuItem.ToolTipText = "Generates annotation data from the selected annotation database"
        '
        'CompareWorkbooksToolStripMenuItem
        '
        Me.CompareWorkbooksToolStripMenuItem.Name = "CompareWorkbooksToolStripMenuItem"
        Me.CompareWorkbooksToolStripMenuItem.Size = New System.Drawing.Size(267, 26)
        Me.CompareWorkbooksToolStripMenuItem.Text = "Compare Workbooks"
        '
        'statusStrip
        '
        Me.statusStrip.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.statusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statusLabel})
        Me.statusStrip.Location = New System.Drawing.Point(0, 800)
        Me.statusStrip.Name = "statusStrip"
        Me.statusStrip.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.statusStrip.Size = New System.Drawing.Size(1480, 25)
        Me.statusStrip.TabIndex = 2
        '
        'statusLabel
        '
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(234, 20)
        Me.statusLabel.Text = "Select a tool from the tools menu."
        '
        'WordToolsToolStripMenuItem
        '
        Me.WordToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AcronymBuilderToolStripMenuItem})
        Me.WordToolsToolStripMenuItem.Name = "WordToolsToolStripMenuItem"
        Me.WordToolsToolStripMenuItem.Size = New System.Drawing.Size(96, 24)
        Me.WordToolsToolStripMenuItem.Text = "Word Tools"
        '
        'AcronymBuilderToolStripMenuItem
        '
        Me.AcronymBuilderToolStripMenuItem.Name = "AcronymBuilderToolStripMenuItem"
        Me.AcronymBuilderToolStripMenuItem.Size = New System.Drawing.Size(216, 26)
        Me.AcronymBuilderToolStripMenuItem.Text = "Acronym Builder"
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1480, 825)
        Me.Controls.Add(Me.statusStrip)
        Me.Controls.Add(Me.IDD_ToolsMenuStrip)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.IDD_ToolsMenuStrip
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "MainForm"
        Me.Text = "Office Productivity Tools"
        Me.IDD_ToolsMenuStrip.ResumeLayout(False)
        Me.IDD_ToolsMenuStrip.PerformLayout()
        Me.statusStrip.ResumeLayout(False)
        Me.statusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents IDD_ToolsMenuStrip As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PrepareIDDInputCollectorToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MergeIDDInputsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents FormatDeliveredIDDToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AnnotateIDDToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents GenerateAnnotationPagesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents statusStrip As StatusStrip
    Friend WithEvents statusLabel As ToolStripStatusLabel
    Friend WithEvents CompareWorkbooksToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents WordToolsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AcronymBuilderToolStripMenuItem As ToolStripMenuItem
End Class
