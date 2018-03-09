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
        Me.lstOriginalFileTabList = New System.Windows.Forms.ListBox()
        Me.lstNewFileTabList = New System.Windows.Forms.ListBox()
        Me.cmdCompare = New System.Windows.Forms.Button()
        Me.grpKeyValue = New System.Windows.Forms.GroupBox()
        Me.optRowByRow = New System.Windows.Forms.RadioButton()
        Me.optColumnA = New System.Windows.Forms.RadioButton()
        Me.grpCompareConfiguration = New System.Windows.Forms.GroupBox()
        Me.optCompareAll = New System.Windows.Forms.RadioButton()
        Me.optCompareSingleSheet = New System.Windows.Forms.RadioButton()
        Me.cmdReset = New System.Windows.Forms.Button()
        Me.cmdExit = New System.Windows.Forms.Button()
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
        Me.lstOriginalFileTabList.Enabled = False
        Me.lstOriginalFileTabList.FormattingEnabled = True
        Me.lstOriginalFileTabList.Location = New System.Drawing.Point(26, 85)
        Me.lstOriginalFileTabList.Name = "lstOriginalFileTabList"
        Me.lstOriginalFileTabList.Size = New System.Drawing.Size(167, 147)
        Me.lstOriginalFileTabList.TabIndex = 2
        '
        'lstNewFileTabList
        '
        Me.lstNewFileTabList.Enabled = False
        Me.lstNewFileTabList.FormattingEnabled = True
        Me.lstNewFileTabList.Location = New System.Drawing.Point(236, 85)
        Me.lstNewFileTabList.Name = "lstNewFileTabList"
        Me.lstNewFileTabList.Size = New System.Drawing.Size(167, 147)
        Me.lstNewFileTabList.TabIndex = 3
        '
        'cmdCompare
        '
        Me.cmdCompare.Enabled = False
        Me.cmdCompare.Location = New System.Drawing.Point(168, 244)
        Me.cmdCompare.Name = "cmdCompare"
        Me.cmdCompare.Size = New System.Drawing.Size(95, 45)
        Me.cmdCompare.TabIndex = 4
        Me.cmdCompare.Text = "Compare Files"
        Me.cmdCompare.UseVisualStyleBackColor = True
        '
        'grpKeyValue
        '
        Me.grpKeyValue.Controls.Add(Me.optRowByRow)
        Me.grpKeyValue.Controls.Add(Me.optColumnA)
        Me.grpKeyValue.Location = New System.Drawing.Point(9, 298)
        Me.grpKeyValue.Name = "grpKeyValue"
        Me.grpKeyValue.Size = New System.Drawing.Size(196, 72)
        Me.grpKeyValue.TabIndex = 6
        Me.grpKeyValue.TabStop = False
        Me.grpKeyValue.Text = "Key Value Configuration"
        '
        'optRowByRow
        '
        Me.optRowByRow.AutoSize = True
        Me.optRowByRow.Location = New System.Drawing.Point(9, 44)
        Me.optRowByRow.Name = "optRowByRow"
        Me.optRowByRow.Size = New System.Drawing.Size(183, 17)
        Me.optRowByRow.TabIndex = 1
        Me.optRowByRow.Text = "Perform Row by Row Comparison"
        Me.optRowByRow.UseVisualStyleBackColor = True
        '
        'optColumnA
        '
        Me.optColumnA.AutoSize = True
        Me.optColumnA.Location = New System.Drawing.Point(9, 21)
        Me.optColumnA.Name = "optColumnA"
        Me.optColumnA.Size = New System.Drawing.Size(157, 17)
        Me.optColumnA.TabIndex = 0
        Me.optColumnA.Text = "Use Column A as Key Value"
        Me.optColumnA.UseVisualStyleBackColor = True
        '
        'grpCompareConfiguration
        '
        Me.grpCompareConfiguration.Controls.Add(Me.optCompareAll)
        Me.grpCompareConfiguration.Controls.Add(Me.optCompareSingleSheet)
        Me.grpCompareConfiguration.Location = New System.Drawing.Point(233, 298)
        Me.grpCompareConfiguration.Name = "grpCompareConfiguration"
        Me.grpCompareConfiguration.Size = New System.Drawing.Size(196, 72)
        Me.grpCompareConfiguration.TabIndex = 7
        Me.grpCompareConfiguration.TabStop = False
        Me.grpCompareConfiguration.Text = "Comparison Configuration"
        '
        'optCompareAll
        '
        Me.optCompareAll.AutoSize = True
        Me.optCompareAll.Location = New System.Drawing.Point(11, 44)
        Me.optCompareAll.Name = "optCompareAll"
        Me.optCompareAll.Size = New System.Drawing.Size(117, 17)
        Me.optCompareAll.TabIndex = 1
        Me.optCompareAll.Text = "Compare All Sheets"
        Me.optCompareAll.UseVisualStyleBackColor = True
        '
        'optCompareSingleSheet
        '
        Me.optCompareSingleSheet.AutoSize = True
        Me.optCompareSingleSheet.Location = New System.Drawing.Point(11, 21)
        Me.optCompareSingleSheet.Name = "optCompareSingleSheet"
        Me.optCompareSingleSheet.Size = New System.Drawing.Size(130, 17)
        Me.optCompareSingleSheet.TabIndex = 0
        Me.optCompareSingleSheet.Text = "Compare Single Sheet"
        Me.optCompareSingleSheet.UseVisualStyleBackColor = True
        '
        'cmdReset
        '
        Me.cmdReset.Location = New System.Drawing.Point(68, 386)
        Me.cmdReset.Name = "cmdReset"
        Me.cmdReset.Size = New System.Drawing.Size(95, 45)
        Me.cmdReset.TabIndex = 8
        Me.cmdReset.Text = "Reset"
        Me.cmdReset.UseVisualStyleBackColor = True
        '
        'cmdExit
        '
        Me.cmdExit.Location = New System.Drawing.Point(272, 386)
        Me.cmdExit.Name = "cmdExit"
        Me.cmdExit.Size = New System.Drawing.Size(95, 45)
        Me.cmdExit.TabIndex = 9
        Me.cmdExit.Text = "Exit"
        Me.cmdExit.UseVisualStyleBackColor = True
        '
        'WorkbookCompare
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(436, 467)
        Me.Controls.Add(Me.cmdExit)
        Me.Controls.Add(Me.cmdReset)
        Me.Controls.Add(Me.grpCompareConfiguration)
        Me.Controls.Add(Me.grpKeyValue)
        Me.Controls.Add(Me.cmdCompare)
        Me.Controls.Add(Me.lstNewFileTabList)
        Me.Controls.Add(Me.lstOriginalFileTabList)
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
    Friend WithEvents lstOriginalFileTabList As ListBox
    Friend WithEvents lstNewFileTabList As ListBox
    Friend WithEvents cmdCompare As Button
    Friend WithEvents grpKeyValue As GroupBox
    Friend WithEvents optRowByRow As RadioButton
    Friend WithEvents optColumnA As RadioButton
    Friend WithEvents grpCompareConfiguration As GroupBox
    Friend WithEvents optCompareAll As RadioButton
    Friend WithEvents optCompareSingleSheet As RadioButton
    Friend WithEvents cmdReset As Button
    Friend WithEvents cmdExit As Button
End Class
