Imports Microsoft.Office.Interop.Excel

''' <summary>
''' This class contains operations that are used to generate the IDD preparation workboook.
''' </summary>
Public Class GeneratePreparationWorkbook

    Public Sub Generate(ByRef Configuration As IDD_PreparationConfiguration, ByRef statusBar As System.Windows.Forms.StatusBar, ByRef progress As System.Windows.Forms.ProgressBar)
        Dim myExcelApp As Application
        Dim myWorkbook As Workbook
        Dim myIndexSheet, myExportSheet, myMessageTopicSheet As Worksheet
        Dim myIndexTableData, myExportTableData, myMessageTopicRow As Range
        Dim myLogicalTopic As LogicalTopic
        Dim myLogicalTopicCollection As New LogicalTopicCollection
        Dim messageRowCounter, rowsInIndexTable, rowsInExportTable, loopCounter, lastColumnInExport As Integer
        Dim startDateTime As DateTime = Now()

        'myExcelApp = Configuration.ExcelApplication
        myExcelApp = Configuration.ExcelHandler.ExcelApp
        myWorkbook = CType(myExcelApp.ActiveWorkbook, Workbook)
        myIndexSheet = CType(myExcelApp.ActiveWorkbook.Sheets(Configuration.IndexWorksheet), Worksheet)
        myExportSheet = CType(myExcelApp.ActiveWorkbook.Sheets(Configuration.ExportWorksheet), Worksheet)

        'Set up ranges for import and export
        myIndexTableData = myIndexSheet.ListObjects(1).DataBodyRange
        rowsInIndexTable = myIndexTableData.Rows.Count
        myExportTableData = myExportSheet.ListObjects(1).DataBodyRange
        rowsInExportTable = myExportTableData.Rows.Count
        progress.Maximum = (2 * rowsInIndexTable) + rowsInExportTable
        progress.Step = 1

        'Build collection of topics from index
        For loopCounter = 1 To rowsInIndexTable
            Dim myIndexCellValue As Object = CType(myIndexTableData.Cells(loopCounter, 1), Range).Value
            myLogicalTopic = New LogicalTopic With {
                .Topic = myIndexCellValue
            }
            myLogicalTopicCollection.TopicCollection.Add(myIndexCellValue, myLogicalTopic)
            'Update status and progress bar
            statusBar.Panels.Item(0).Text = "Creating collection:  adding " & loopCounter & " of " & rowsInIndexTable & " topics."
            statusBar.Refresh()
            progress.PerformStep()
            progress.Refresh()
            myLogicalTopic = Nothing
        Next

        'Use the last column in export to identify records to add to LogicalTopicIndex entries
        Dim myTopicCellValue As String
        Dim topicStringArray As String()
        Dim myExportTableRange As Range
        lastColumnInExport = myExportTableData.Columns.Count
        For loopCounter = 1 To rowsInExportTable
            statusBar.Panels.Item(0).Text = "Updating collection:  adding row " & loopCounter & " of " & rowsInExportTable & " to logical topics."
            statusBar.Refresh()
            myTopicCellValue = CType(myExportTableData.Cells(loopCounter, lastColumnInExport), Range).Value
            topicStringArray = Split(myTopicCellValue, vbLf)
            myExportTableRange = myExportTableData.Range(myExportTableData.Cells(loopCounter, 1), myExportTableData.Cells(loopCounter, lastColumnInExport - 1))
            For Each item In topicStringArray
                Try
                    myLogicalTopic = myLogicalTopicCollection.TopicCollection.Item(item)
                    myLogicalTopic.DataRange.Add(myExportTableRange)
                    myLogicalTopicCollection.TopicCollection.Remove(item)
                    myLogicalTopicCollection.TopicCollection.Add(item, myLogicalTopic)
                Catch
                    'Topic isn't in collection
                End Try
            Next
            progress.PerformStep()
            progress.Refresh()
        Next

        'Loop through logical topic collection to create sheets and link to index
        Dim myHeaderRow As String = myExportSheet.ListObjects(1).HeaderRowRange.Row.ToString
        Dim myHeaderFirstColumn As String = ColumnIndexToColumnLetter(myExportSheet.ListObjects(1).HeaderRowRange.Column)
        Dim myHeaderLastColumn As String = ColumnIndexToColumnLetter(myExportSheet.ListObjects(1).HeaderRowRange.Columns.Count - 1)
        myMessageTopicRow = myExportSheet.Range(myHeaderFirstColumn & myHeaderRow & ":" & myHeaderLastColumn & myHeaderRow)
        For Each item As Range In myIndexTableData
            'Add new sheet to end
            statusBar.Panels.Item(0).Text = "Creating sheet for " & item.Value & " logical topic."
            statusBar.Refresh()
            myExcelApp.Worksheets.Add(After:=myExcelApp.Sheets(myExcelApp.Sheets.Count))
            myLogicalTopic = myLogicalTopicCollection.TopicCollection.Item(item.Value)
            myMessageTopicSheet = myExcelApp.ActiveSheet
            With myMessageTopicSheet
                'Rename sheet to topic name
                .Name = CalculateTabName(item.Value, myExcelApp.ActiveWorkbook)
                'Add "Back to Index" hyperlink in A1
                item.Hyperlinks.Add(Anchor:= .Range(Cell1:="A1"), Address:="", SubAddress:=myIndexSheet.Name & "!A1", TextToDisplay:="Back to Index")
                'Add hyperlink on index table
                myIndexSheet.Hyperlinks.Add(Anchor:=item, Address:="", SubAddress:=myMessageTopicSheet.Name & "!A1", TextToDisplay:=item.Value)
                'Add "Topic Name" in B1
                .Range("A2").Value = "Message Topic"
                'Insert topic name in B2
                .Range("B2").Value = item.Value
                'Insert table
                Dim myMessageTable As ListObject
                myMessageTable = myMessageTopicSheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
                                                    myMessageTopicSheet.Range(myMessageTopicSheet.Cells(3, 1),
                                                    myMessageTopicSheet.Cells(3, myMessageTopicRow.Count - 1)), , XlYesNoGuess.xlYes)
                myMessageTopicRow.Copy()
                .Range("A3").PasteSpecial(XlPasteType.xlPasteColumnWidths)
                .Range("A3").PasteSpecial(XlPasteType.xlPasteValues)
                'Loop through topic and paste data range beginning with D1.
                myMessageTable.Name = item.Value
                messageRowCounter = 4
                For Each entry As Range In myLogicalTopic.DataRange
                    entry.Copy()
                    .Range("A" & messageRowCounter).PasteSpecial(XlPasteType.xlPasteAll)
                    messageRowCounter = messageRowCounter + 1
                Next
                myMessageTopicSheet.Range("A1").Select()
            End With
            progress.PerformStep()
            progress.Refresh()
        Next
        progress.Value = progress.Maximum
        progress.Refresh()
        myWorkbook.SaveAs(Left(myWorkbook.Name, Len(myWorkbook.Name) - 5) & "_" & Now().ToString("yyyyMMdd_hh_mm") & ".xlsx")
        myWorkbook.Close()
        Dim elapsedTime As Integer = (Now().Hour * 3600 + Now().Minute * 60 + Now().Second) - (startDateTime.Hour * 3600 + startDateTime.Minute * 60 + startDateTime.Second)
        statusBar.Panels.Item(0).Text = "Complete:  Workbook created and saved.  Elapsed time = " & elapsedTime & " seconds."
        statusBar.Refresh()
        myExcelApp.Quit()
    End Sub

    Private Function CalculateTabName(ByVal strInputTopicName As String, ByRef myWorkbook As Workbook) As String
        Dim isSheetDuplicate As Boolean
        Dim nameCtr, topicNameLength As Integer
        Dim strTrimmedTopicName As String
        Dim tmpWorksheet As Worksheet

        nameCtr = 0
        topicNameLength = Len(strInputTopicName)

        Select Case topicNameLength
            Case 1 To 28
                CalculateTabName = strInputTopicName
                Exit Function
            Case Else
                isSheetDuplicate = True
                strTrimmedTopicName = Left(strInputTopicName, 29)
                Do Until isSheetDuplicate = False
                    For Each item In myWorkbook.Sheets
                        Try
                            tmpWorksheet = myWorkbook.Sheets.Item(strTrimmedTopicName)
                            isSheetDuplicate = True
                            nameCtr = nameCtr + 1
                            strTrimmedTopicName = strTrimmedTopicName & nameCtr
                        Catch ex As Exception
                            isSheetDuplicate = False
                        End Try
                    Next
                Loop
                CalculateTabName = strTrimmedTopicName
                Exit Function
        End Select

    End Function

    Private Function ColumnIndexToColumnLetter(colIndex As Integer) As String
        Dim div As Integer = colIndex
        Dim colLetter As String = String.Empty
        Dim modnum As Integer = 0

        While div > 0
            modnum = (div - 1) Mod 26
            colLetter = Chr(65 + modnum) & colLetter
            div = CInt((div - modnum) \ 26)
        End While

        Return colLetter
    End Function

End Class
