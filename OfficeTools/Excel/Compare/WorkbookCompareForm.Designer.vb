<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WorkbookCompare
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
        Me.ButtonOriginalFile = New System.Windows.Forms.Button()
        Me.ButtonNewFile = New System.Windows.Forms.Button()
        Me.ListBoxOriginalFileTabList = New System.Windows.Forms.ListBox()
        Me.ListBoxNewFileTabList = New System.Windows.Forms.ListBox()
        Me.ButtonCompare = New System.Windows.Forms.Button()
        Me.grpKeyValue = New System.Windows.Forms.GroupBox()
        Me.OptionRowByRow = New System.Windows.Forms.RadioButton()
        Me.OptionColumnA = New System.Windows.Forms.RadioButton()
        Me.grpCompareConfiguration = New System.Windows.Forms.GroupBox()
        Me.OptionCompareAll = New System.Windows.Forms.RadioButton()
        Me.OptionCompareSingleSheet = New System.Windows.Forms.RadioButton()
        Me.ButtonReset = New System.Windows.Forms.Button()
        Me.ButtonExit = New System.Windows.Forms.Button()
        Me.grpKeyValue.SuspendLayout()
        Me.grpCompareConfiguration.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdOriginalFile
        '
        Me.ButtonOriginalFile.Location = New System.Drawing.Point(68, 22)
        Me.ButtonOriginalFile.Name = "cmdOriginalFile"
        Me.ButtonOriginalFile.Size = New System.Drawing.Size(95, 45)
        Me.ButtonOriginalFile.TabIndex = 0
        Me.ButtonOriginalFile.Text = "Select Original File"
        Me.ButtonOriginalFile.UseVisualStyleBackColor = True
        '
        'cmdNewFile
        '
        Me.ButtonNewFile.Location = New System.Drawing.Point(272, 22)
        Me.ButtonNewFile.Name = "cmdNewFile"
        Me.ButtonNewFile.Size = New System.Drawing.Size(95, 45)
        Me.ButtonNewFile.TabIndex = 1
        Me.ButtonNewFile.Text = "Select New File"
        Me.ButtonNewFile.UseVisualStyleBackColor = True
        '
        'lstOriginalFileTabList
        '
        Me.ListBoxOriginalFileTabList.Enabled = False
        Me.ListBoxOriginalFileTabList.FormattingEnabled = True
        Me.ListBoxOriginalFileTabList.Location = New System.Drawing.Point(26, 85)
        Me.ListBoxOriginalFileTabList.Name = "lstOriginalFileTabList"
        Me.ListBoxOriginalFileTabList.Size = New System.Drawing.Size(167, 147)
        Me.ListBoxOriginalFileTabList.TabIndex = 2
        '
        'lstNewFileTabList
        '
        Me.ListBoxNewFileTabList.Enabled = False
        Me.ListBoxNewFileTabList.FormattingEnabled = True
        Me.ListBoxNewFileTabList.Location = New System.Drawing.Point(236, 85)
        Me.ListBoxNewFileTabList.Name = "lstNewFileTabList"
        Me.ListBoxNewFileTabList.Size = New System.Drawing.Size(167, 147)
        Me.ListBoxNewFileTabList.TabIndex = 3
        '
        'cmdCompare
        '
        Me.ButtonCompare.Enabled = False
        Me.ButtonCompare.Location = New System.Drawing.Point(168, 244)
        Me.ButtonCompare.Name = "cmdCompare"
        Me.ButtonCompare.Size = New System.Drawing.Size(95, 45)
        Me.ButtonCompare.TabIndex = 4
        Me.ButtonCompare.Text = "Compare Files"
        Me.ButtonCompare.UseVisualStyleBackColor = True
        '
        'grpKeyValue
        '
        Me.grpKeyValue.Controls.Add(Me.OptionRowByRow)
        Me.grpKeyValue.Controls.Add(Me.OptionColumnA)
        Me.grpKeyValue.Location = New System.Drawing.Point(9, 298)
        Me.grpKeyValue.Name = "grpKeyValue"
        Me.grpKeyValue.Size = New System.Drawing.Size(196, 72)
        Me.grpKeyValue.TabIndex = 6
        Me.grpKeyValue.TabStop = False
        Me.grpKeyValue.Text = "Key Value Configuration"
        '
        'optRowByRow
        '
        Me.OptionRowByRow.AutoSize = True
        Me.OptionRowByRow.Location = New System.Drawing.Point(9, 44)
        Me.OptionRowByRow.Name = "optRowByRow"
        Me.OptionRowByRow.Size = New System.Drawing.Size(183, 17)
        Me.OptionRowByRow.TabIndex = 1
        Me.OptionRowByRow.Text = "Perform Row by Row Comparison"
        Me.OptionRowByRow.UseVisualStyleBackColor = True
        '
        'optColumnA
        '
        Me.OptionColumnA.AutoSize = True
        Me.OptionColumnA.Location = New System.Drawing.Point(9, 21)
        Me.OptionColumnA.Name = "optColumnA"
        Me.OptionColumnA.Size = New System.Drawing.Size(157, 17)
        Me.OptionColumnA.TabIndex = 0
        Me.OptionColumnA.Text = "Use Column A as Key Value"
        Me.OptionColumnA.UseVisualStyleBackColor = True
        '
        'grpCompareConfiguration
        '
        Me.grpCompareConfiguration.Controls.Add(Me.OptionCompareAll)
        Me.grpCompareConfiguration.Controls.Add(Me.OptionCompareSingleSheet)
        Me.grpCompareConfiguration.Location = New System.Drawing.Point(233, 298)
        Me.grpCompareConfiguration.Name = "grpCompareConfiguration"
        Me.grpCompareConfiguration.Size = New System.Drawing.Size(196, 72)
        Me.grpCompareConfiguration.TabIndex = 7
        Me.grpCompareConfiguration.TabStop = False
        Me.grpCompareConfiguration.Text = "Comparison Configuration"
        '
        'optCompareAll
        '
        Me.OptionCompareAll.AutoSize = True
        Me.OptionCompareAll.Location = New System.Drawing.Point(11, 44)
        Me.OptionCompareAll.Name = "optCompareAll"
        Me.OptionCompareAll.Size = New System.Drawing.Size(117, 17)
        Me.OptionCompareAll.TabIndex = 1
        Me.OptionCompareAll.Text = "Compare All Sheets"
        Me.OptionCompareAll.UseVisualStyleBackColor = True
        '
        'optCompareSingleSheet
        '
        Me.OptionCompareSingleSheet.AutoSize = True
        Me.OptionCompareSingleSheet.Location = New System.Drawing.Point(11, 21)
        Me.OptionCompareSingleSheet.Name = "optCompareSingleSheet"
        Me.OptionCompareSingleSheet.Size = New System.Drawing.Size(130, 17)
        Me.OptionCompareSingleSheet.TabIndex = 0
        Me.OptionCompareSingleSheet.Text = "Compare Single Sheet"
        Me.OptionCompareSingleSheet.UseVisualStyleBackColor = True
        '
        'cmdReset
        '
        Me.ButtonReset.Location = New System.Drawing.Point(68, 386)
        Me.ButtonReset.Name = "cmdReset"
        Me.ButtonReset.Size = New System.Drawing.Size(95, 45)
        Me.ButtonReset.TabIndex = 8
        Me.ButtonReset.Text = "Reset"
        Me.ButtonReset.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.ButtonExit.Location = New System.Drawing.Point(272, 386)
        Me.ButtonExit.Name = "cmdExit"
        Me.ButtonExit.Size = New System.Drawing.Size(95, 45)
        Me.ButtonExit.TabIndex = 9
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.UseVisualStyleBackColor = True
        '
        'WorkbookCompare
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(436, 467)
        Me.Controls.Add(Me.ButtonExit)
        Me.Controls.Add(Me.ButtonReset)
        Me.Controls.Add(Me.grpCompareConfiguration)
        Me.Controls.Add(Me.grpKeyValue)
        Me.Controls.Add(Me.ButtonCompare)
        Me.Controls.Add(Me.ListBoxNewFileTabList)
        Me.Controls.Add(Me.ListBoxOriginalFileTabList)
        Me.Controls.Add(Me.ButtonNewFile)
        Me.Controls.Add(Me.ButtonOriginalFile)
        Me.Name = "WorkbookCompare"
        Me.Text = "Workbook Compare"
        Me.grpKeyValue.ResumeLayout(False)
        Me.grpKeyValue.PerformLayout()
        Me.grpCompareConfiguration.ResumeLayout(False)
        Me.grpCompareConfiguration.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ButtonOriginalFile As Button
    Friend WithEvents ButtonNewFile As Button
    Friend WithEvents ListBoxOriginalFileTabList As ListBox
    Friend WithEvents ListBoxNewFileTabList As ListBox
    Friend WithEvents ButtonCompare As Button
    Friend WithEvents grpKeyValue As GroupBox
    Friend WithEvents OptionRowByRow As RadioButton
    Friend WithEvents OptionColumnA As RadioButton
    Friend WithEvents grpCompareConfiguration As GroupBox
    Friend WithEvents OptionCompareAll As RadioButton
    Friend WithEvents OptionCompareSingleSheet As RadioButton
    Friend WithEvents ButtonReset As Button
    Friend WithEvents ButtonExit As Button
End Class
