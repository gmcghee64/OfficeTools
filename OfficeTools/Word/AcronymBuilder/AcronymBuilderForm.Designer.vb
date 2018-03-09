<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AcronymBuilderForm
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
        Me.ButtonExit = New System.Windows.Forms.Button()
        Me.ButtonOpenDocument = New System.Windows.Forms.Button()
        Me.ButtonUpdateDocument = New System.Windows.Forms.Button()
        Me.CheckBoxGenerateList = New System.Windows.Forms.CheckBox()
        Me.TextBoxAcronymDefinition = New System.Windows.Forms.TextBox()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.SuspendLayout()
        '
        'ListBoxAcronyms
        '
        Me.ListBoxAcronyms.Enabled = False
        Me.ListBoxAcronyms.FormattingEnabled = True
        Me.ListBoxAcronyms.ItemHeight = 16
        Me.ListBoxAcronyms.Location = New System.Drawing.Point(27, 53)
        Me.ListBoxAcronyms.Name = "ListBoxAcronyms"
        Me.ListBoxAcronyms.Size = New System.Drawing.Size(194, 244)
        Me.ListBoxAcronyms.TabIndex = 0
        '
        'ButtonExit
        '
        Me.ButtonExit.Location = New System.Drawing.Point(32, 338)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(131, 51)
        Me.ButtonExit.TabIndex = 1
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.UseVisualStyleBackColor = True
        '
        'ButtonOpenDocument
        '
        Me.ButtonOpenDocument.Location = New System.Drawing.Point(189, 338)
        Me.ButtonOpenDocument.Name = "ButtonOpenDocument"
        Me.ButtonOpenDocument.Size = New System.Drawing.Size(131, 51)
        Me.ButtonOpenDocument.TabIndex = 2
        Me.ButtonOpenDocument.Text = "Open Document"
        Me.ButtonOpenDocument.UseVisualStyleBackColor = True
        '
        'ButtonUpdateDocument
        '
        Me.ButtonUpdateDocument.Enabled = False
        Me.ButtonUpdateDocument.Location = New System.Drawing.Point(346, 338)
        Me.ButtonUpdateDocument.Name = "ButtonUpdateDocument"
        Me.ButtonUpdateDocument.Size = New System.Drawing.Size(131, 51)
        Me.ButtonUpdateDocument.TabIndex = 3
        Me.ButtonUpdateDocument.Text = "Update Document"
        Me.ButtonUpdateDocument.UseVisualStyleBackColor = True
        '
        'CheckBoxGenerateList
        '
        Me.CheckBoxGenerateList.AutoSize = True
        Me.CheckBoxGenerateList.Enabled = False
        Me.CheckBoxGenerateList.Location = New System.Drawing.Point(266, 205)
        Me.CheckBoxGenerateList.Name = "CheckBoxGenerateList"
        Me.CheckBoxGenerateList.Size = New System.Drawing.Size(212, 21)
        Me.CheckBoxGenerateList.TabIndex = 4
        Me.CheckBoxGenerateList.Text = "Generate Table of Acronyms"
        Me.CheckBoxGenerateList.UseVisualStyleBackColor = True
        '
        'TextBoxAcronymDefinition
        '
        Me.TextBoxAcronymDefinition.Enabled = False
        Me.TextBoxAcronymDefinition.Location = New System.Drawing.Point(266, 53)
        Me.TextBoxAcronymDefinition.Name = "TextBoxAcronymDefinition"
        Me.TextBoxAcronymDefinition.Size = New System.Drawing.Size(212, 22)
        Me.TextBoxAcronymDefinition.TabIndex = 5
        '
        'StatusStrip1
        '
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 428)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(523, 22)
        Me.StatusStrip1.TabIndex = 6
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'AcronymBuilderForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(523, 450)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.TextBoxAcronymDefinition)
        Me.Controls.Add(Me.CheckBoxGenerateList)
        Me.Controls.Add(Me.ButtonUpdateDocument)
        Me.Controls.Add(Me.ButtonOpenDocument)
        Me.Controls.Add(Me.ButtonExit)
        Me.Controls.Add(Me.ListBoxAcronyms)
        Me.Name = "AcronymBuilderForm"
        Me.Text = "Acronym Builder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ListBoxAcronyms As ListBox
    Friend WithEvents ButtonExit As Button
    Friend WithEvents ButtonOpenDocument As Button
    Friend WithEvents ButtonUpdateDocument As Button
    Friend WithEvents CheckBoxGenerateList As CheckBox
    Friend WithEvents TextBoxAcronymDefinition As TextBox
    Friend WithEvents StatusStrip1 As StatusStrip
End Class
