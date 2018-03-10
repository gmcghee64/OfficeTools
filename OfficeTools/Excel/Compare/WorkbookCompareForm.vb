Public Class WorkbookCompare
    Private myComparison As New CompareWorkbooks

    Protected Overrides Sub Finalize()
        myComparison = Nothing
        MyBase.Finalize()
    End Sub

    Public Sub New()
        InitializeComponent()
        With myComparison
            OptionColumnA.Checked = .KeyValueColumnA
            OptionCompareSingleSheet.Checked = .CompareSingleSheet
            .ExcelHandler = New ExcelHandler
        End With
    End Sub

    ''' <summary>
    ''' This operation prompts the user to select the workbook as it existed prior to making any changes.  Upon selecting
    ''' the workbook, the listbox will be updated to list all of the tabs in the workbook.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonOriginalFile_Click(sender As Object, e As EventArgs) Handles ButtonOriginalFile.Click
        Try
            With myComparison
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
            myComparison.OriginalWorkbook = Nothing
            Me.ListBoxOriginalFileTabList.Items.Clear()
            ConfigureFormSelections()
        End Try
    End Sub

    ''' <summary>
    ''' This operation prompts the user to open the workbook that contains the changes.  Upon opening the workbook, the
    ''' listbox will be updated to include all of the tabs in the workbook.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonNewFile_Click(sender As Object, e As EventArgs) Handles ButtonNewFile.Click
        Try
            With myComparison
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
            myComparison.NewWorkbook = Nothing
            Me.ListBoxNewFileTabList.Items.Clear()
            ConfigureFormSelections()
        End Try
    End Sub

    ''' <summary>
    ''' This operation exits the compare workbooks application.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonExit_Click(sender As Object, e As EventArgs) Handles ButtonExit.Click
        Me.Hide()
    End Sub

    ''' <summary>
    ''' This operation compares the two workbooks.  The new workbook is saved as a new file with a tab that lists all
    ''' of the changes.  The list contains hypertext links to each item that is different.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonCompare_Click(sender As Object, e As EventArgs) Handles ButtonCompare.Click
        'TODO:  Implement comparison
    End Sub

    ''' <summary>
    ''' This operation updates the comparison configuration object to indicate how the comparison should be performed.  Options
    ''' are to use column A as a key-value or to perform a row-by-row comparison.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OptionColumnA_CheckedChanged(sender As Object, e As EventArgs) Handles OptionColumnA.CheckedChanged
        myComparison.KeyValueColumnA = OptionColumnA.Checked
    End Sub

    ''' <summary>
    ''' This operation updates the comparison configuration object to indicate how the comparison should be performed.  Options
    ''' are to use column A as a key-value or to perform a row-by-row comparison.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OptionRowByRow_CheckedChanged(sender As Object, e As EventArgs) Handles OptionRowByRow.CheckedChanged
        myComparison.KeyValueColumnA = Not (OptionRowByRow.Checked)
    End Sub

    ''' <summary>
    ''' This operation updates the comparison configuration object to indicate if a single (selected) sheet or
    ''' the entire workbook should be compared.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OptionCompareSingleSheet_CheckedChanged(sender As Object, e As EventArgs) Handles OptionCompareSingleSheet.CheckedChanged
        myComparison.CompareSingleSheet = OptionCompareSingleSheet.Checked
        ConfigureFormSelections()
    End Sub

    ''' <summary>
    ''' This operation updates the comparison configuration object to indicate if a single (selected) sheet or
    ''' the entire workbook should be compared.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub OptionCompareAll_CheckedChanged(sender As Object, e As EventArgs) Handles OptionCompareAll.CheckedChanged
        myComparison.CompareSingleSheet = Not (OptionCompareAll.Checked)
        ConfigureFormSelections()
    End Sub

    ''' <summary>
    ''' This operation closes any open workbooks and resets the form to its original state.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonReset_Click(sender As Object, e As EventArgs) Handles ButtonReset.Click
        With myComparison
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

    ''' <summary>
    ''' This operation updates the comparison configuration to identify the sheet in the original file that is to be compared.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ListBoxOriginalFileTabList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxOriginalFileTabList.SelectedIndexChanged
        Try
            'TODO:  Add selected sheet to myConfiguration
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' This operation updates the comparison configuration to identify the sheet in the updated file that is to be compared.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ListBoxNewFileTabList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxNewFileTabList.SelectedIndexChanged
        Try
            'TODO:  Add selected sheet to myConfiguration
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' This operation enables and disables form controls based on the options selected by the user.
    ''' </summary>
    Private Sub ConfigureFormSelections()
        With myComparison
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