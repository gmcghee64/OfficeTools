Public Class WorkbookCompare
    Private myConfiguration As New CompareWorkbooks

    Protected Overrides Sub Finalize()
        myConfiguration = Nothing
        MyBase.Finalize()
    End Sub

    Public Sub New()
        InitializeComponent()
        With myConfiguration
            optColumnA.Checked = .KeyValueColumnA
            optCompareSingleSheet.Checked = .CompareSingleSheet
            .ExcelHandler = New ExcelHandler
        End With
    End Sub

    Private Sub ButtonOriginalFile_Click(sender As Object, e As EventArgs) Handles ButtonOriginalFile.Click
        Try
            With myConfiguration
                .ExcelHandler.OpenWorkbook()
                .OriginalWorkbook = .ExcelHandler.ActiveWorkbook
                lstOriginalFileTabList.Items.Clear()
                For Each sheet In .OriginalWorkbook.Sheets
                    lstOriginalFileTabList.Items.Add(sheet.name)
                Next
            End With
            ConfigureFormSelections()
        Catch ex As Exception
            'Workbook not opened, clear configuration
            Debug.Print(ex.ToString)
            myConfiguration.OriginalWorkbook = Nothing
            Me.lstOriginalFileTabList.Items.Clear()
            ConfigureFormSelections()
        End Try
    End Sub

    Private Sub ButtonNewFile_Click(sender As Object, e As EventArgs) Handles ButtonNewFile.Click
        Try
            With myConfiguration
                .ExcelHandler.OpenWorkbook()
                .NewWorkbook = .ExcelHandler.ActiveWorkbook
                lstNewFileTabList.Items.Clear()
                For Each sheet In .NewWorkbook.Sheets
                    lstNewFileTabList.Items.Add(sheet.name)
                Next
            End With
            ConfigureFormSelections()
        Catch ex As Exception
            'Workbook not opened, clear configuration
            Debug.Print(ex.ToString)
            myConfiguration.NewWorkbook = Nothing
            Me.lstNewFileTabList.Items.Clear()
            ConfigureFormSelections()
        End Try
    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        Me.Hide()
    End Sub

    Private Sub cmdCompare_Click(sender As Object, e As EventArgs) Handles cmdCompare.Click
        'TODO:  Implement comparison
    End Sub

    Private Sub optColumnA_CheckedChanged(sender As Object, e As EventArgs) Handles optColumnA.CheckedChanged
        myConfiguration.KeyValueColumnA = optColumnA.Checked
    End Sub

    Private Sub optRowByRow_CheckedChanged(sender As Object, e As EventArgs) Handles optRowByRow.CheckedChanged
        myConfiguration.KeyValueColumnA = Not (optRowByRow.Checked)
    End Sub

    Private Sub optCompareSingleSheet_CheckedChanged(sender As Object, e As EventArgs) Handles optCompareSingleSheet.CheckedChanged
        myConfiguration.CompareSingleSheet = optCompareSingleSheet.Checked
        ConfigureFormSelections()
    End Sub

    Private Sub optCompareAll_CheckedChanged(sender As Object, e As EventArgs) Handles optCompareAll.CheckedChanged
        myConfiguration.CompareSingleSheet = Not (optCompareAll.Checked)
        ConfigureFormSelections()
    End Sub

    Private Sub cmdReset_Click(sender As Object, e As EventArgs) Handles cmdReset.Click
        With myConfiguration
            Try
                .ExcelHandler.ExcelApp.Workbooks.Close()
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
            .NewWorkbook = Nothing
            .OriginalWorkbook = Nothing
        End With
        lstNewFileTabList.Items.Clear()
        lstOriginalFileTabList.Items.Clear()
    End Sub

    Private Sub lstOriginalFileTabList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstOriginalFileTabList.SelectedIndexChanged
        Try
            'TODO:  Add selected sheet to myConfiguration
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Private Sub lstNewFileTabList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstNewFileTabList.SelectedIndexChanged
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
                    lstNewFileTabList.Enabled = True
                    lstOriginalFileTabList.Enabled = True
                Else
                    lstNewFileTabList.Enabled = False
                    lstOriginalFileTabList.Enabled = False
                End If
            Else
                lstNewFileTabList.Enabled = False
                lstOriginalFileTabList.Enabled = False
                cmdCompare.Enabled = False
            End If
        End With
    End Sub
End Class