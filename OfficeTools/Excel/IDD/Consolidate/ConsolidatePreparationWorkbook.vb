Imports Microsoft.Office.Interop.Excel
''' <summary>
''' This class merges the data in the preparation workbook with the data contained in the model
''' </summary>
Public Class ConsolidatePreparationWorkbook

    Public Sub New()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Public Overrides Function ToString() As String
        Return MyBase.ToString()
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        Return MyBase.Equals(obj)
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return MyBase.GetHashCode()
    End Function

    Public Sub GenerateConsoldatedWorkbook(ByRef Configuration As IDD_ConsolidationConfiguration)
        Dim myExcelApp As Application = Configuration.ExcelHandler.ExcelApp.Application
        Dim myConsolidationSheet As Worksheet = CType(Configuration.ExcelHandler, Worksheet)
        Dim myTagCollection As Collection = Configuration.TagsToMerge
        Dim mySheetsInWorkbook As Sheets = Configuration.ExcelHandler.SheetsInWorkbook

        'TODO:  Build GUID collection from consolidation sheet
        Dim myConsolidationTable As ListObject = myConsolidationSheet.ListObjects(1)
    End Sub


    'TODO:  The following block of code is from the VBA implementation and needs to be adapted
    '    Private Sub ConsolidateTags(strWorksheetName As String)
    '    On Error GoTo ErrorHandler
    '    Dim strTopicName As String
    '    Dim tblForTopic, tblForImport As ListObject
    '    Dim collTags As New Collection
    '    Dim intNumberOfAttributes, ctr, intImportRow As Integer
    '    Dim TagValue As GUID_Container
    '    Dim rngGUID, rngTagValue, rngSummary As Range

    '        Worksheets(strWorksheetName).Activate
    '        strTopicName = ActiveSheet.Range("B2").Value
    '    Set tblForTopic = ActiveSheet.ListObjects(strTopicName)
    '    tblForTopic.AutoFilter.ShowAllData
    'Gather the tags and GUIDs
    '    Set rngGUID = tblForTopic.DataBodyRange
    '    intNumberOfAttributes = rngGUID.Rows.Count
    '    For ctr = 1 To intNumberOfAttributes
    '    Set TagValue = New GUID_Container
    '    With TagValue
    '    .GUID = rngGUID.Cells(ctr, 1).Value
    '    .TagValue = rngGUID.Cells(ctr, intTagColumnForImport).Value
    '    End With
    '            collTags.Add Item:=TagValue, Key:=TagValue.GUID
    '    Set TagValue = Nothing
    '    Next
    'Add the tags to the main spreadsheet
    '    Set rngGUID = Nothing
    '    Worksheets(strExportedDataSheetName).Activate
    '    Set tblForImport = ActiveSheet.ListObjects(1)
    '    tblForImport.AutoFilter.ShowAllData
    '    If intLastColumnInExportSheet = 0 Then
    '            intLastColumnInExportSheet = tblForImport.Range.Columns.Count + 1
    '    End If
    '    For ctr = 1 To intNumberOfAttributes
    '    Set TagValue = New GUID_Container
    '        TagValue.GUID = collTags.Item(ctr).GUID
    '            TagValue.TagValue = collTags.Item(ctr).TagValue
    '   Set rngGUID = tblForImport.Range.Find(TagValue.GUID)
    '       intImportRow = rngGUID.Row - intFirstDataRow + 2
    '   Set rngTagValue = tblForImport.Range.Cells(intImportRow, intTagColumnForImport)
    '   Set rngSummary = tblForImport.Range.Cells(intImportRow, intLastColumnInExportSheet)
    '  If rngTagValue.Value = "" Then
    '              rngTagValue.Value = TagValue.TagValue
    '              rngSummary.Value = rngSummary.Value & strTopicName & " (" & TagValue.TagValue & ")" & Chr(10)
    '  ElseIf rngTagValue.Value = TagValue.TagValue Then
    '              rngSummary.Value = rngSummary.Value & strTopicName & " (" & TagValue.TagValue & ")" & Chr(10)
    '  ElseIf rngTagValue.Value <> TagValue.TagValue And rngTagValue.Value <> "CONFLICT" Then
    '              'First conflict detected
    '              rngTagValue.Value = "CONFLICT"
    '              rngSummary.Value = rngSummary.Value & strTopicName & " (" & TagValue.TagValue & ")" & Chr(10)
    '  Else
    '              'After detecting first conflict, simply log the contributing topics and their values
    '              rngSummary.Value = rngSummary.Value & strTopicName & " (" & TagValue.TagValue & ")" & Chr(10)
    '  End If
    '    Set TagValue = Nothing
    '    Next

    '    Exit Sub'
    'ErrorHandler:
    '        Cleanup("ConsolidateTags")
    '    Exit Sub
    '    End Sub

End Class
