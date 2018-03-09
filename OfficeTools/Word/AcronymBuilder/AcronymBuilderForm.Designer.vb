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
        Me.lstAcronyms = New System.Windows.Forms.ListBox()
        Me.ButtonExit = New System.Windows.Forms.Button()
        Me.ButtonOpenDocument = New System.Windows.Forms.Button()
        Me.ButtonUpdateDocument = New System.Windows.Forms.Button()
        Me.chkGenerateList = New System.Windows.Forms.CheckBox()
        Me.txtAcronymDefinition = New System.Windows.Forms.TextBox()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.SuspendLayout()
        '
        'lstAcronyms
        '
        Me.lstAcronyms.Enabled = False
        Me.lstAcronyms.FormattingEnabled = True
        Me.lstAcronyms.ItemHeight = 16
        Me.lstAcronyms.Location = New System.Drawing.Point(27, 53)
        Me.lstAcronyms.Name = "lstAcronyms"
        Me.lstAcronyms.Size = New System.Drawing.Size(194, 244)
        Me.lstAcronyms.TabIndex = 0
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
        'chkGenerateList
        '
        Me.chkGenerateList.AutoSize = True
        Me.chkGenerateList.Enabled = False
        Me.chkGenerateList.Location = New System.Drawing.Point(266, 205)
        Me.chkGenerateList.Name = "chkGenerateList"
        Me.chkGenerateList.Size = New System.Drawing.Size(212, 21)
        Me.chkGenerateList.TabIndex = 4
        Me.chkGenerateList.Text = "Generate Table of Acronyms"
        Me.chkGenerateList.UseVisualStyleBackColor = True
        '
        'txtAcronymDefinition
        '
        Me.txtAcronymDefinition.Enabled = False
        Me.txtAcronymDefinition.Location = New System.Drawing.Point(266, 53)
        Me.txtAcronymDefinition.Name = "txtAcronymDefinition"
        Me.txtAcronymDefinition.Size = New System.Drawing.Size(212, 22)
        Me.txtAcronymDefinition.TabIndex = 5
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
        Me.Controls.Add(Me.txtAcronymDefinition)
        Me.Controls.Add(Me.chkGenerateList)
        Me.Controls.Add(Me.ButtonUpdateDocument)
        Me.Controls.Add(Me.ButtonOpenDocument)
        Me.Controls.Add(Me.ButtonExit)
        Me.Controls.Add(Me.lstAcronyms)
        Me.Name = "AcronymBuilderForm"
        Me.Text = "Acronym Builder"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lstAcronyms As ListBox
    Friend WithEvents ButtonExit As Button
    Friend WithEvents ButtonOpenDocument As Button
    Friend WithEvents ButtonUpdateDocument As Button
    Friend WithEvents chkGenerateList As CheckBox
    Friend WithEvents txtAcronymDefinition As TextBox
    Friend WithEvents StatusStrip1 As StatusStrip
End Class
