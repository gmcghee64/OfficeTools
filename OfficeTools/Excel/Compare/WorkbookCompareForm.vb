Public Class WorkbookCompare
    Private myConfiguration As New CompareWorkbooks

    Protected Overrides Sub Finalize()
        myConfiguration = Nothing
        MyBase.Finalize()
    End Sub

    Public Sub New()
        InitializeComponent()
        With myConfiguration
            OptionColumnA.Checked = .KeyValueColumnA
            OptionCompareSingleSheet.Checked = .CompareSingleSheet
            .ExcelHandler = New ExcelHandler
        End With
    End Sub

    Private Sub ButtonOriginalFile_Click(sender As Object, e As EventArgs) Handles ButtonOriginalFile.Click
        Try
            With myConfiguration
                .ExcelHandler.OpenWorkbook()
                .OriginalWorkbook = .ExcelHandler.ActiveWorkbook
                ListBoxOriginalFileTabList.Items.Clear()
                For Each sheet In .OriginalWorkbook.Sheets
                    ListBoxOriginalFileTabList.Items.Add(sheet.name)
                Next
            End With
            ConfigureFormSelections()
        Catch ex As Exception
            'Workbook not opened, clear configuration
            Debug.Print(ex.ToString)
            myConfiguration.OriginalWorkbook = Nothing
            Me.ListBoxOriginalFileTabList.Items.Clear()
            ConfigureFormSelections()
        End Try
    End Sub

    Private Sub ButtonNewFile_Click(sender As Object, e As EventArgs) Handles ButtonNewFile.Click
        Try
            With myConfiguration
                .ExcelHandler.OpenWorkbook()
                .NewWorkbook = .ExcelHandler.ActiveWorkbook
                ListBoxNewFileTabList.Items.Clear()
                For Each sheet In .NewWorkbook.Sheets
                    ListBoxNewFileTabList.Items.Add(sheet.name)
                Next
            End With
            ConfigureFormSelections()
        Catch ex As Exception
            'Workbook not opened, clear configuration
            Debug.Print(ex.ToString)
            myConfiguration.NewWorkbook = Nothing
            Me.ListBoxNewFileTabList.Items.Clear()
            ConfigureFormSelections()
        End Try
    End Sub

    Private Sub ButtonExit_Click(sender As Object, e As EventArgs) Handles ButtonExit.Click
        Me.Hide()
    End Sub

    Private Sub ButtonCompare_Click(sender As Object, e As EventArgs) Handles ButtonCompare.Click
        'TODO:  Implement comparison
    End Sub

    Private Sub OptionColumnA_CheckedChanged(sender As Object, e As EventArgs) Handles OptionColumnA.CheckedChanged
        myConfiguration.KeyValueColumnA = OptionColumnA.Checked
    End Sub

    Private Sub OptionRowByRow_CheckedChanged(sender As Object, e As EventArgs) Handles OptionRowByRow.CheckedChanged
        myConfiguration.KeyValueColumnA = Not (OptionRowByRow.Checked)
    End Sub

    Private Sub OptionCompareSingleSheet_CheckedChanged(sender As Object, e As EventArgs) Handles OptionCompareSingleSheet.CheckedChanged
        myConfiguration.CompareSingleSheet = OptionCompareSingleSheet.Checked
        ConfigureFormSelections()
    End Sub

    Private Sub OptionCompareAll_CheckedChanged(sender As Object, e As EventArgs) Handles OptionCompareAll.CheckedChanged
        myConfiguration.CompareSingleSheet = Not (OptionCompareAll.Checked)
        ConfigureFormSelections()
    End Sub

    Private Sub ButtonReset_Click(sender As Object, e As EventArgs) Handles ButtonReset.Click
        With myConfiguration
            Try
                .ExcelHandler.ExcelApp.Workbooks.Close()
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
            .NewWorkbook = Nothing
            .OriginalWorkbook = Nothing
        End With
        ListBoxNewFileTabList.Items.Clear()
        ListBoxOriginalFileTabList.Items.Clear()
    End Sub

    Private Sub ListBoxOriginalFileTabList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxOriginalFileTabList.SelectedIndexChanged
        Try
            'TODO:  Add selected sheet to myConfiguration
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Private Sub ListBoxNewFileTabList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxNewFileTabList.SelectedIndexChanged
        Try
            'TODO:  Add selected sheet to myConfiguration
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Private Sub ConfigureFormSelections()
        With myConfiguration
            If Not IsNothing(.NewWorkbook) And Not IsNothing(.OriginalWorkbook) Then
                If .CompareSingleSheet = True Then
                    ListBoxNewFileTabList.Enabled = True
                    ListBoxOriginalFileTabList.Enabled = True
                Else
                    ListBoxNewFileTabList.Enabled = False
                    ListBoxOriginalFileTabList.Enabled = False
                End If
            Else
                ListBoxNewFileTabList.Enabled = False
                ListBoxOriginalFileTabList.Enabled = False
                ButtonCompare.Enabled = False
            End If
        End With
    End Sub
End Class