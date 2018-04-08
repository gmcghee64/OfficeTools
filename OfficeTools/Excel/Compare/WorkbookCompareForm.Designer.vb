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
        Me.Status = New System.Windows.Forms.StatusStrip()
        Me.grpKeyValue.SuspendLayout()
        Me.grpCompareConfiguration.SuspendLayout()
        Me.SuspendLayout()
        '
        'ButtonOriginalFile
        '
        Me.ButtonOriginalFile.Location = New System.Drawing.Point(91, 27)
        Me.ButtonOriginalFile.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonOriginalFile.Name = "ButtonOriginalFile"
        Me.ButtonOriginalFile.Size = New System.Drawing.Size(127, 55)
        Me.ButtonOriginalFile.TabIndex = 0
        Me.ButtonOriginalFile.Text = "Select Original File"
        Me.ButtonOriginalFile.UseVisualStyleBackColor = True
        '
        'ButtonNewFile
        '
        Me.ButtonNewFile.Location = New System.Drawing.Point(363, 27)
        Me.ButtonNewFile.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonNewFile.Name = "ButtonNewFile"
        Me.ButtonNewFile.Size = New System.Drawing.Size(127, 55)
        Me.ButtonNewFile.TabIndex = 1
        Me.ButtonNewFile.Text = "Select New File"
        Me.ButtonNewFile.UseVisualStyleBackColor = True
        '
        'ListBoxOriginalFileTabList
        '
        Me.ListBoxOriginalFileTabList.Enabled = False
        Me.ListBoxOriginalFileTabList.FormattingEnabled = True
        Me.ListBoxOriginalFileTabList.ItemHeight = 16
        Me.ListBoxOriginalFileTabList.Location = New System.Drawing.Point(35, 105)
        Me.ListBoxOriginalFileTabList.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ListBoxOriginalFileTabList.Name = "ListBoxOriginalFileTabList"
        Me.ListBoxOriginalFileTabList.Size = New System.Drawing.Size(221, 180)
        Me.ListBoxOriginalFileTabList.TabIndex = 2
        '
        'ListBoxNewFileTabList
        '
        Me.ListBoxNewFileTabList.Enabled = False
        Me.ListBoxNewFileTabList.FormattingEnabled = True
        Me.ListBoxNewFileTabList.ItemHeight = 16
        Me.ListBoxNewFileTabList.Location = New System.Drawing.Point(315, 105)
        Me.ListBoxNewFileTabList.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ListBoxNewFileTabList.Name = "ListBoxNewFileTabList"
        Me.ListBoxNewFileTabList.Size = New System.Drawing.Size(221, 180)
        Me.ListBoxNewFileTabList.TabIndex = 3
        '
        'ButtonCompare
        '
        Me.ButtonCompare.Enabled = False
        Me.ButtonCompare.Location = New System.Drawing.Point(224, 300)
        Me.ButtonCompare.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonCompare.Name = "ButtonCompare"
        Me.ButtonCompare.Size = New System.Drawing.Size(127, 55)
        Me.ButtonCompare.TabIndex = 4
        Me.ButtonCompare.Text = "Compare Files"
        Me.ButtonCompare.UseVisualStyleBackColor = True
        '
        'grpKeyValue
        '
        Me.grpKeyValue.Controls.Add(Me.OptionRowByRow)
        Me.grpKeyValue.Controls.Add(Me.OptionColumnA)
        Me.grpKeyValue.Location = New System.Drawing.Point(12, 367)
        Me.grpKeyValue.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpKeyValue.Name = "grpKeyValue"
        Me.grpKeyValue.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpKeyValue.Size = New System.Drawing.Size(261, 89)
        Me.grpKeyValue.TabIndex = 6
        Me.grpKeyValue.TabStop = False
        Me.grpKeyValue.Text = "Key Value Configuration"
        '
        'OptionRowByRow
        '
        Me.OptionRowByRow.AutoSize = True
        Me.OptionRowByRow.Location = New System.Drawing.Point(12, 54)
        Me.OptionRowByRow.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OptionRowByRow.Name = "OptionRowByRow"
        Me.OptionRowByRow.Size = New System.Drawing.Size(239, 21)
        Me.OptionRowByRow.TabIndex = 1
        Me.OptionRowByRow.Text = "Perform Row by Row Comparison"
        Me.OptionRowByRow.UseVisualStyleBackColor = True
        '
        'OptionColumnA
        '
        Me.OptionColumnA.AutoSize = True
        Me.OptionColumnA.Location = New System.Drawing.Point(12, 26)
        Me.OptionColumnA.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OptionColumnA.Name = "OptionColumnA"
        Me.OptionColumnA.Size = New System.Drawing.Size(205, 21)
        Me.OptionColumnA.TabIndex = 0
        Me.OptionColumnA.Text = "Use Column A as Key Value"
        Me.OptionColumnA.UseVisualStyleBackColor = True
        '
        'grpCompareConfiguration
        '
        Me.grpCompareConfiguration.Controls.Add(Me.OptionCompareAll)
        Me.grpCompareConfiguration.Controls.Add(Me.OptionCompareSingleSheet)
        Me.grpCompareConfiguration.Location = New System.Drawing.Point(311, 367)
        Me.grpCompareConfiguration.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpCompareConfiguration.Name = "grpCompareConfiguration"
        Me.grpCompareConfiguration.Padding = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.grpCompareConfiguration.Size = New System.Drawing.Size(261, 89)
        Me.grpCompareConfiguration.TabIndex = 7
        Me.grpCompareConfiguration.TabStop = False
        Me.grpCompareConfiguration.Text = "Comparison Configuration"
        '
        'OptionCompareAll
        '
        Me.OptionCompareAll.AutoSize = True
        Me.OptionCompareAll.Location = New System.Drawing.Point(15, 54)
        Me.OptionCompareAll.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OptionCompareAll.Name = "OptionCompareAll"
        Me.OptionCompareAll.Size = New System.Drawing.Size(153, 21)
        Me.OptionCompareAll.TabIndex = 1
        Me.OptionCompareAll.Text = "Compare All Sheets"
        Me.OptionCompareAll.UseVisualStyleBackColor = True
        '
        'OptionCompareSingleSheet
        '
        Me.OptionCompareSingleSheet.AutoSize = True
        Me.OptionCompareSingleSheet.Location = New System.Drawing.Point(15, 26)
        Me.OptionCompareSingleSheet.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.OptionCompareSingleSheet.Name = "OptionCompareSingleSheet"
        Me.OptionCompareSingleSheet.Size = New System.Drawing.Size(170, 21)
        Me.OptionCompareSingleSheet.TabIndex = 0
        Me.OptionCompareSingleSheet.Text = "Compare Single Sheet"
        Me.OptionCompareSingleSheet.UseVisualStyleBackColor = True
        '
        'ButtonReset
        '
        Me.ButtonReset.Location = New System.Drawing.Point(91, 475)
        Me.ButtonReset.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonReset.Name = "ButtonReset"
        Me.ButtonReset.Size = New System.Drawing.Size(127, 55)
        Me.ButtonReset.TabIndex = 8
        Me.ButtonReset.Text = "Reset"
        Me.ButtonReset.UseVisualStyleBackColor = True
        '
        'ButtonExit
        '
        Me.ButtonExit.Location = New System.Drawing.Point(363, 475)
        Me.ButtonExit.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.ButtonExit.Name = "ButtonExit"
        Me.ButtonExit.Size = New System.Drawing.Size(127, 55)
        Me.ButtonExit.TabIndex = 9
        Me.ButtonExit.Text = "Exit"
        Me.ButtonExit.UseVisualStyleBackColor = True
        '
        'Status
        '
        Me.Status.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.Status.Location = New System.Drawing.Point(0, 553)
        Me.Status.Name = "Status"
        Me.Status.Size = New System.Drawing.Size(581, 22)
        Me.Status.TabIndex = 10
        Me.Status.Text = "StatusStrip1"
        '
        'WorkbookCompare
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(581, 575)
        Me.Controls.Add(Me.Status)
        Me.Controls.Add(Me.ButtonExit)
        Me.Controls.Add(Me.ButtonReset)
        Me.Controls.Add(Me.grpCompareConfiguration)
        Me.Controls.Add(Me.grpKeyValue)
        Me.Controls.Add(Me.ButtonCompare)
        Me.Controls.Add(Me.ListBoxNewFileTabList)
        Me.Controls.Add(Me.ListBoxOriginalFileTabList)
        Me.Controls.Add(Me.ButtonNewFile)
        Me.Controls.Add(Me.ButtonOriginalFile)
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "WorkbookCompare"
        Me.Text = "Workbook Compare"
        Me.grpKeyValue.ResumeLayout(False)
        Me.grpKeyValue.PerformLayout()
        Me.grpCompareConfiguration.ResumeLayout(False)
        Me.grpCompareConfiguration.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

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
    Friend WithEvents Status As StatusStrip
End Class
