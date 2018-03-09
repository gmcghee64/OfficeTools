<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class IDD_PreparationForm
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
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdCreate = New System.Windows.Forms.Button()
        Me.lstSheetForIndex = New System.Windows.Forms.ListBox()
        Me.lstSheetForExport = New System.Windows.Forms.ListBox()
        Me.lblSheetForIndex = New System.Windows.Forms.Label()
        Me.lblExportSheet = New System.Windows.Forms.Label()
        Me.pbPrepProgress = New System.Windows.Forms.ProgressBar()
        Me.cmdSelectWorkbook = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(390, 225)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(148, 54)
        Me.cmdCancel.TabIndex = 0
        Me.cmdCancel.Text = "Exit"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdCreate
        '
        Me.cmdCreate.Enabled = False
        Me.cmdCreate.Location = New System.Drawing.Point(596, 225)
        Me.cmdCreate.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmdCreate.Name = "cmdCreate"
        Me.cmdCreate.Size = New System.Drawing.Size(176, 55)
        Me.cmdCreate.TabIndex = 1
        Me.cmdCreate.Text = "Create IDD"
        Me.cmdCreate.UseVisualStyleBackColor = True
        '
        'lstSheetForIndex
        '
        Me.lstSheetForIndex.FormattingEnabled = True
        Me.lstSheetForIndex.ItemHeight = 20
        Me.lstSheetForIndex.Location = New System.Drawing.Point(26, 54)
        Me.lstSheetForIndex.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstSheetForIndex.Name = "lstSheetForIndex"
        Me.lstSheetForIndex.Size = New System.Drawing.Size(318, 144)
        Me.lstSheetForIndex.TabIndex = 2
        '
        'lstSheetForExport
        '
        Me.lstSheetForExport.FormattingEnabled = True
        Me.lstSheetForExport.ItemHeight = 20
        Me.lstSheetForExport.Location = New System.Drawing.Point(390, 54)
        Me.lstSheetForExport.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.lstSheetForExport.Name = "lstSheetForExport"
        Me.lstSheetForExport.Size = New System.Drawing.Size(382, 144)
        Me.lstSheetForExport.TabIndex = 3
        '
        'lblSheetForIndex
        '
        Me.lblSheetForIndex.AutoSize = True
        Me.lblSheetForIndex.Location = New System.Drawing.Point(26, 29)
        Me.lblSheetForIndex.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSheetForIndex.Name = "lblSheetForIndex"
        Me.lblSheetForIndex.Size = New System.Drawing.Size(316, 20)
        Me.lblSheetForIndex.TabIndex = 4
        Me.lblSheetForIndex.Text = "Select the sheet containing the list of topics"
        '
        'lblExportSheet
        '
        Me.lblExportSheet.AutoSize = True
        Me.lblExportSheet.Location = New System.Drawing.Point(390, 29)
        Me.lblExportSheet.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblExportSheet.Name = "lblExportSheet"
        Me.lblExportSheet.Size = New System.Drawing.Size(383, 20)
        Me.lblExportSheet.TabIndex = 5
        Me.lblExportSheet.Text = "Select the sheet containing the data for the IDD tabs"
        '
        'pbPrepProgress
        '
        Me.pbPrepProgress.Location = New System.Drawing.Point(396, 306)
        Me.pbPrepProgress.Name = "pbPrepProgress"
        Me.pbPrepProgress.Size = New System.Drawing.Size(375, 15)
        Me.pbPrepProgress.TabIndex = 6
        Me.pbPrepProgress.Visible = False
        '
        'cmdSelectWorkbook
        '
        Me.cmdSelectWorkbook.Location = New System.Drawing.Point(100, 225)
        Me.cmdSelectWorkbook.Name = "cmdSelectWorkbook"
        Me.cmdSelectWorkbook.Size = New System.Drawing.Size(154, 54)
        Me.cmdSelectWorkbook.TabIndex = 7
        Me.cmdSelectWorkbook.Text = "Select Template Workbook"
        Me.cmdSelectWorkbook.UseVisualStyleBackColor = True
        '
        'IDD_PreparationForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(801, 377)
        Me.Controls.Add(Me.cmdSelectWorkbook)
        Me.Controls.Add(Me.pbPrepProgress)
        Me.Controls.Add(Me.lblExportSheet)
        Me.Controls.Add(Me.lblSheetForIndex)
        Me.Controls.Add(Me.lstSheetForExport)
        Me.Controls.Add(Me.lstSheetForIndex)
        Me.Controls.Add(Me.cmdCreate)
        Me.Controls.Add(Me.cmdCancel)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.Name = "IDD_PreparationForm"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.Text = "Prepare IDD Collection Workbook"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents cmdCancel As Windows.Forms.Button
    Friend WithEvents cmdCreate As Windows.Forms.Button
    Friend WithEvents lstSheetForIndex As Windows.Forms.ListBox
    Friend WithEvents lstSheetForExport As Windows.Forms.ListBox
    Friend WithEvents lblSheetForIndex As Windows.Forms.Label
    Friend WithEvents lblExportSheet As Windows.Forms.Label
    Friend WithEvents pbPrepProgress As Windows.Forms.ProgressBar
    Friend WithEvents cmdSelectWorkbook As Windows.Forms.Button
End Class
