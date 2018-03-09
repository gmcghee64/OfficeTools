Public Class IDD_PreparationForm
    Private myConfiguration As New IDD_PreparationConfiguration
    Private isConfiguredForTest As Boolean = True
    Private myStatusBar As New System.Windows.Forms.StatusBar
    Private myStatusBarPanel As New StatusBarPanel()

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        myStatusBarPanel.BorderStyle = StatusBarPanelBorderStyle.Sunken
        myStatusBarPanel.AutoSize = StatusBarPanelAutoSize.Spring
        myStatusBar.ShowPanels = True
        myStatusBar.Panels.Add(myStatusBarPanel)
        Me.Controls.Add(myStatusBar)

    End Sub

    Protected Overrides Sub Finalize()
        If Not IsNothing(myConfiguration) Then
            myConfiguration = Nothing
        End If
        MyBase.Finalize()
    End Sub

    ''' <summary>
    ''' This operation opens the Excel workbook containing a sheet with the index of topics to include in the preparation workbook and
    ''' a sheet with the data model classes and attributes with logical topic cross reference information in the last column.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub SelectWorkbook_Click(sender As Object, e As EventArgs) Handles cmdSelectWorkbook.Click
        Dim myExcelTools As New ExcelHandler
        myConfiguration.ExcelHandler = myExcelTools

        Me.lstSheetForExport.Items.Clear()
        Me.lstSheetForIndex.Items.Clear()
        Me.pbPrepProgress.Value = 0
        Me.pbPrepProgress.Visible = False
        Me.cmdCreate.Enabled = False
        If myExcelTools.OpenWorkbook() = True Then
            For Each item In myExcelTools.SheetsInWorkbook
                lstSheetForExport.Items.Add(item.Name)
                lstSheetForIndex.Items.Add(item.Name)
            Next
        End If
    End Sub

    Private Sub Cancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Hide()
    End Sub

    Private Sub SheetForIndex_SelectedIndexChange(sender As Object, e As EventArgs) Handles lstSheetForIndex.SelectedIndexChanged
        Try
            myConfiguration.IndexWorksheet = lstSheetForIndex.SelectedItem.ToString
            If IsConfigurationComplete() = True Then
                cmdCreate.Enabled = True
            Else
                cmdCreate.Enabled = False
            End If
        Catch
        End Try
    End Sub

    Private Sub SheetForExport_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstSheetForExport.SelectedIndexChanged
        Try
            myConfiguration.ExportWorksheet = lstSheetForExport.SelectedItem.ToString
            If IsConfigurationComplete() = True Then
                cmdCreate.Enabled = True
            Else
                cmdCreate.Enabled = False
            End If
        Catch
        End Try
    End Sub

    Private Function IsConfigurationComplete() As Boolean
        Try
            If myConfiguration.IndexWorksheet <> "" And myConfiguration.ExportWorksheet <> "" And myConfiguration.IndexWorksheet <> myConfiguration.ExportWorksheet Then
                IsConfigurationComplete = True
            Else
                IsConfigurationComplete = False
            End If
        Catch
            IsConfigurationComplete = False
        End Try
    End Function

    Private Sub Create_Click(sender As Object, e As EventArgs) Handles cmdCreate.Click
        Dim myGenerateWB As New GeneratePreparationWorkbook
        Me.cmdCreate.Enabled = False
        pbPrepProgress.Visible = True
        myGenerateWB.Generate(myConfiguration, myStatusBar, pbPrepProgress)
        myGenerateWB = Nothing
        Me.lstSheetForExport.Items.Clear()
        Me.lstSheetForIndex.Items.Clear()
        Me.Refresh()
    End Sub

End Class