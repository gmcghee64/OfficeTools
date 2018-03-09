<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Main
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
        Me.ListBoxAcronyms = New System.Windows.Forms.ListBox()
        Me.TextBoxDefinition = New System.Windows.Forms.TextBox()
        Me.ButtonOpen = New System.Windows.Forms.Button()
        Me.ButtonApply = New System.Windows.Forms.Button()
        Me.ButtonExit = New System.Windows.Forms.Button()
        Me.StatusStripMain = New System.Windows.Forms.StatusStrip()
        Me.StatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.CheckBoxAcronymTable = New System.Windows.Forms.CheckBox()
        Me.StatusStripMain.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListBoxAcronyms
        '
        Me.ListBoxAcronyms.FormattingEnabled = True
        Me.ListBoxAcronyms.ItemHeight = 16
        Me.ListBoxAcronyms.Location = New System.Drawing.Point(37, 49)
        Me.ListBoxAcronyms.Name = "ListBoxAcronyms"
        Me.ListBoxAcronyms.Size = New System.Drawing.Size(140, 148)
        Me.ListBoxAcronyms.TabIndex = 0
        '
        'TextBoxDefinition
        '
        Me.TextBoxDefinition.Enabled = False
        Me.TextBoxDefinition.Location = New System.Drawing.Point(193, 49)
        Me.TextBoxDefinition.Name = "TextBoxDefinition"
        Me.TextBoxDefinition.Size = New System.Drawing.Size(301, 22)
        Me.TextBoxDefinition.TabIndex = 1
        '
        'ButtonOpen
        '
        Me.ButtonOpen.Location = New System.Drawing.Point(108, 229)
        Me.ButtonOpen.Name = "ButtonOpen"
        Me.ButtonOpen.Size = New System.Drawing.Size(89, 47)
        Me.ButtonOpen.TabIndex = 2
        Me.ButtonOpen.Text = "Open Document"
        Me.ButtonOpen.UseVisualStyleBackColor = True
        '
        'ButtonApply
        '
        Me.ButtonApply.Enabled = False
        Me.ButtonApply.Location = New System.Drawing.Point(223, 229)
        Me.ButtonApply.Name = "ButtonApply"
        Me.ButtonApply.Size = New System.Drawing.Size(89, 47)
        Me.ButtonApply.TabIndex = 3
        Me.ButtonApply.Text = "Apply Acronyms"
        Me.ButtonApply.UseVisualStyleBackColor = True
        '
        'ButtonExit
        '
        Me.ButtonExit.Location = New System.Drawing.Point(338, 229)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(89, 47)
        Me.ButtonExit.TabIndex = 4
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.UseVisualStyleBackColor = True
        '
        'StatusStripMain
        '
        Me.StatusStripMain.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatusLabel})
        Me.StatusStripMain.Location = New System.Drawing.Point(0, 296)
        Me.StatusStripMain.Name = "StatusStripMain"
        Me.StatusStripMain.Size = New System.Drawing.Size(530, 25)
        Me.StatusStripMain.TabIndex = 5
        Me.StatusStripMain.Text = "StatusStrip1"
        '
        'StatusLabel
        '
        Me.StatusLabel.Name = "StatusLabel"
        Me.StatusLabel.Size = New System.Drawing.Size(168, 20)
        Me.StatusLabel.Text = "Open a Word document"
        '
        'CheckBoxAcronymTable
        '
        Me.CheckBoxAcronymTable.AutoSize = True
        Me.CheckBoxAcronymTable.Location = New System.Drawing.Point(210, 122)
        Me.CheckBoxAcronymTable.Name = "CheckBoxAcronymTable"
        Me.CheckBoxAcronymTable.Size = New System.Drawing.Size(284, 21)
        Me.CheckBoxAcronymTable.TabIndex = 6
        Me.CheckBoxAcronymTable.Text = "Insert acronym table at end of document"
        Me.CheckBoxAcronymTable.UseVisualStyleBackColor = True
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(530, 321)
        Me.Controls.Add(Me.CheckBoxAcronymTable)
        Me.Controls.Add(Me.StatusStripMain)
        Me.Controls.Add(Me.ButtonExit)
        Me.Controls.Add(Me.ButtonApply)
        Me.Controls.Add(Me.ButtonOpen)
        Me.Controls.Add(Me.TextBoxDefinition)
        Me.Controls.Add(Me.ListBoxAcronyms)
        Me.Name = "Main"
        Me.Text = "Acronym Definer"
        Me.StatusStripMain.ResumeLayout(False)
        Me.StatusStripMain.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ListBoxAcronyms As ListBox
    Friend WithEvents TextBoxDefinition As TextBox
    Friend WithEvents ButtonOpen As Button
    Friend WithEvents ButtonApply As Button
    Friend WithEvents ButtonExit As Button
    Friend WithEvents StatusStripMain As StatusStrip
    Friend WithEvents StatusLabel As ToolStripStatusLabel
    Friend WithEvents CheckBoxAcronymTable As CheckBox
End Class
