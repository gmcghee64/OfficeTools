<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IDD_ConsolidationForm
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
        Me.cmdOpen = New System.Windows.Forms.Button()
        Me.lstBoxTabs = New System.Windows.Forms.ListBox()
        Me.lstBoxTags = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmdMerge = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'cmdOpen
        '
        Me.cmdOpen.Location = New System.Drawing.Point(69, 300)
        Me.cmdOpen.Name = "cmdOpen"
        Me.cmdOpen.Size = New System.Drawing.Size(179, 74)
        Me.cmdOpen.TabIndex = 0
        Me.cmdOpen.Text = "Open Workbook"
        Me.cmdOpen.UseVisualStyleBackColor = True
        '
        'lstBoxTabs
        '
        Me.lstBoxTabs.Enabled = False
        Me.lstBoxTabs.FormattingEnabled = True
        Me.lstBoxTabs.ItemHeight = 20
        Me.lstBoxTabs.Location = New System.Drawing.Point(52, 57)
        Me.lstBoxTabs.Name = "lstBoxTabs"
        Me.lstBoxTabs.Size = New System.Drawing.Size(420, 204)
        Me.lstBoxTabs.TabIndex = 1
        '
        'lstBoxTags
        '
        Me.lstBoxTags.Enabled = False
        Me.lstBoxTags.FormattingEnabled = True
        Me.lstBoxTags.ItemHeight = 20
        Me.lstBoxTags.Location = New System.Drawing.Point(532, 57)
        Me.lstBoxTags.Name = "lstBoxTags"
        Me.lstBoxTags.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lstBoxTags.Size = New System.Drawing.Size(433, 204)
        Me.lstBoxTags.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(48, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(295, 20)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Select tab into which data will be merged"
        '
        'cmdMerge
        '
        Me.cmdMerge.Enabled = False
        Me.cmdMerge.Location = New System.Drawing.Point(638, 306)
        Me.cmdMerge.Name = "cmdMerge"
        Me.cmdMerge.Size = New System.Drawing.Size(196, 68)
        Me.cmdMerge.TabIndex = 4
        Me.cmdMerge.Text = "Merge Selected Data"
        Me.cmdMerge.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(528, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(215, 20)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Select columns to be merged"
        '
        'IDD_ConsolidationForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1248, 474)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmdMerge)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lstBoxTags)
        Me.Controls.Add(Me.lstBoxTabs)
        Me.Controls.Add(Me.cmdOpen)
        Me.Name = "IDD_ConsolidationForm"
        Me.Text = "IDD Consolidation Form"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdOpen As Button
    Friend WithEvents lstBoxTabs As ListBox
    Friend WithEvents lstBoxTags As ListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents cmdMerge As Button
    Friend WithEvents Label2 As Label
End Class
