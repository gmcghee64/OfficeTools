Imports Microsoft.Office.Interop.Excel

Public Class IDD_ConsolidationForm
    Private myConfiguration As New IDD_ConsolidationConfiguration

    Protected Overrides Sub Finalize()
        If Not IsNothing(myConfiguration) Then
            myConfiguration = Nothing
        End If
        MyBase.Finalize()
    End Sub

    Private Sub cmdOpen_Click(sender As Object, e As EventArgs) Handles cmdOpen.Click
        Dim myExcelHandler As New ExcelHandler
        myConfiguration.ExcelHandler = myExcelHandler
        lstBoxTabs.Items.Clear()
        lstBoxTags.Items.Clear()
        lstBoxTabs.Enabled = False
        lstBoxTags.Enabled = False

        If myExcelHandler.OpenWorkbook() = True Then
            For Each item In myExcelHandler.SheetsInWorkbook
                lstBoxTabs.Items.Add(item.Name)
                lstBoxTabs.Enabled = True
            Next
        End If
    End Sub

    Private Sub lstBoxTabs_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstBoxTabs.SelectedIndexChanged

        lstBoxTags.Items.Clear()
        lstBoxTags.Enabled = False
        Try
            Dim myExcelApp = myConfiguration.ExcelHandler.ExcelApp
            myConfiguration.ExportWorksheet = lstBoxTabs.SelectedItem.ToString
            Dim myConsolidationSheet As Worksheet = CType(myExcelApp.ActiveWorkbook.Sheets(myConfiguration.ExportWorksheet), Worksheet)
            Dim myHeaderRow As Range = myConsolidationSheet.ListObjects(1).HeaderRowRange
            For Each item In myHeaderRow
                Select Case item.Text
                    Case "GUID", "Full Name", "Fully Qualified Namespace", "Metaclass", "AttributeType"
                        'Skip this item
                    Case Else
                        lstBoxTags.Items.Add(item.Text)
                        lstBoxTags.Enabled = True
                End Select
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub lstBoxTags_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstBoxTags.SelectedIndexChanged
        Try
            Dim mySelectedTags As New Collection
            For Each item In lstBoxTags.SelectedItems
                mySelectedTags.Add(item.ToString, item.ToString)
            Next
            If lstBoxTags.SelectedItems.Count > 0 Then
                cmdMerge.Enabled = True
                myConfiguration.TagsToMerge = mySelectedTags
            Else
                cmdMerge.Enabled = False
                myConfiguration.TagsToMerge = Nothing
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdMerge_Click(sender As Object, e As EventArgs) Handles cmdMerge.Click
        Dim myConsolidateData As New ConsolidatePreparationWorkbook
        myConsolidateData.GenerateConsoldatedWorkbook(myConfiguration)
    End Sub
End Class