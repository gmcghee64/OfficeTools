﻿Imports Microsoft.Office.Interop.Excel
Imports OfficeTools.ComparisonTypes

''' <summary>
''' This class either performs a complete workbook comparison or comparison of selected sheets in two workbooks.
''' The comparison is initiated by executing the PerformComparison operation.  Prior to initiating this operation,
''' the class properties should be populated with the information needed to perform the desired comparison.
''' </summary>
Public Class CompareWorkbooks
    Private _ExcelHandler As ExcelHandler
    Private _OriginalWorkbook As Workbook
    Private _NewWorkbook As Workbook
    Private _KeyValueColumnA As Boolean = True
    Private _CompareSingleSheet As Boolean = True
    Private _OriginalSheet As Worksheet
    Private _NewSheet As Worksheet

    ''' <summary>
    ''' The ExcelHandler property contains the reference to the instance of the Excel application.
    ''' </summary>
    ''' <returns></returns>
    Public Property ExcelHandler As ExcelHandler
        Get
            Return _ExcelHandler
        End Get
        Set(value As ExcelHandler)
            _ExcelHandler = value
        End Set
    End Property

    ''' <summary>
    ''' The OriginalWorkbook property contains the reference to the Workbook as it existed prior to making
    ''' any changes.
    ''' </summary>
    ''' <returns></returns>
    Public Property OriginalWorkbook As Workbook
        Get
            Return _OriginalWorkbook
        End Get
        Set(value As Workbook)
            _OriginalWorkbook = value
        End Set
    End Property

    ''' <summary>
    ''' The NewWorkbook property contains the reference to the Workbook that contains the changes.
    ''' </summary>
    ''' <returns></returns>
    Public Property NewWorkbook As Workbook
        Get
            Return _NewWorkbook
        End Get
        Set(value As Workbook)
            _NewWorkbook = value
        End Set
    End Property

    ''' <summary>
    ''' The KeyValueColumnA property indicates how to perform the comparison.  When set to true, 
    ''' column A will be used as the key value.  Otherwise, a row-by-row comparison will be
    ''' performed.
    ''' </summary>
    ''' <returns></returns>
    Public Property KeyValueColumnA As Boolean
        Get
            Return _KeyValueColumnA
        End Get
        Set(value As Boolean)
            _KeyValueColumnA = value
        End Set
    End Property

    ''' <summary>
    ''' The CompareSingleSheet property indicates if the the entire workbook will be compared.  When
    ''' set to true, the sheets identified by the OriginalSheet and NewSheet properties will be compared.
    ''' Otherwise, all sheets in the workbook will be compared.
    ''' </summary>
    ''' <returns></returns>
    Public Property CompareSingleSheet As Boolean
        Get
            Return _CompareSingleSheet
        End Get
        Set(value As Boolean)
            _CompareSingleSheet = value
        End Set
    End Property

    ''' <summary>
    ''' The OriginalSheet property identifies the worksheet in the original workbook to be compared.
    ''' This property only has meaning when the CompareSingleSheet property is set to true.
    ''' </summary>
    ''' <returns></returns>
    Public Property OriginalSheet As Worksheet
        Get
            Return _OriginalSheet
        End Get
        Set(value As Worksheet)
            _OriginalSheet = value
        End Set
    End Property

    ''' <summary>
    ''' The NewSheet property identifies the worksheet in the original workbook to be compared.
    ''' This property only has meaning when the CompareSingleSheet property is set to true.
    ''' </summary>
    ''' <returns></returns>
    Public Property NewSheet As Worksheet
        Get
            Return _NewSheet
        End Get
        Set(value As Worksheet)
            _NewSheet = value
        End Set
    End Property

    ''' <summary>
    ''' The PerformComparison operation either compares the OriginalSheet to the NewSheet or
    ''' the OriginalWorkbook to the NewWorkbook.  The NewWorkbook is saved as a new
    ''' workbook having the same name as the NewWorkbook appended with a date/time stamp.
    ''' A new sheet is included in the workbook with the comparison results.
    ''' </summary>
    ''' <param name="statusPanel"></param>
    ''' <returns>If the comparison is successful, true is returned.  Otherwise, false is returned.</returns>
    Public Function PerformComparison(statusPanel As StatusStrip) As Boolean
        'TODO:  Implement statusPanel update processing.
        If CheckConfiguration() = True Then
            Dim myResultsCollection As New Collection
            Select Case CompareSingleSheet
                Case True 'Compare OriginalSheet to NewSheet
                    Dim myResults As New SheetComparisonResults
                    If CompareSheet(OriginalSheet, NewSheet, myResults) = True Then
                        myResultsCollection.Add(myResults)
                        GenerateReport(myResultsCollection)
                        Return True
                    Else
                        'An error must have occurred.  Return false.
                        Return False
                    End If
                    Exit Select
                Case False 'Compare OriginalWorkbook to NewWorkbook
                    Dim myResults As New SheetComparisonResults
                    For Each sheet In NewWorkbook.Sheets
                        Try
                            Dim tmpSheetName As String = sheet.name
                            If CompareSheet(OriginalWorkbook.Sheets(tmpSheetName), sheet, myResults) = True Then
                                myResultsCollection.Add(myResults)
                            Else
                                'An error must have occurred.  Return false.
                                Return False
                            End If
                        Catch ex As Exception
                            Debug.Print(ex.ToString)
                            'Sheet must not be in original workbook
                            myResults.SheetName = sheet.name
                            myResults.Result = SheetResultType.SHEET_NEW
                            myResultsCollection.Add(myResults)
                        End Try
                    Next
                    GenerateReport(myResultsCollection)
                    Return True
                    Exit Select
                Case Else
                    Return False
                    Exit Select
            End Select
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' The CheckConfiguration operation verifies that all of the necessary properties have been defined
    ''' prior to executing the comparison.
    ''' </summary>
    ''' <returns>If all properties have been properly configured, true is returned.  Othwerwise, false is returned.</returns>
    Private Function CheckConfiguration() As Boolean
        If Not IsNothing(ExcelHandler) And Not IsNothing(OriginalWorkbook) And
                Not IsNothing(NewWorkbook) Then
            If CompareSingleSheet = True Then
                If Not IsNothing(NewSheet) And Not IsNothing(OriginalSheet) Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return True
            End If
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' This operation compares two worksheets.
    ''' </summary>
    ''' <param name="originalSheet"></param>
    ''' <param name="newSheet"></param>
    ''' <param name="result"></param>
    ''' <returns>This operation returns true if the comparison is successful.  Otherwise, false is returned.  When true
    ''' is returned, the detailed results will be in the result parameter.</returns>
    Private Function CompareSheet(originalSheet As Worksheet, newSheet As Worksheet, result As SheetComparisonResults) As Boolean
        'TODO:  Compare workwheet
        result.Result = SheetResultType.SHEET_IDENTICAL ' Assume sheets are identical
        Try
            result.SheetName = newSheet.Name
            If originalSheet Is newSheet Then
                result.Result = SheetResultType.SHEET_IDENTICAL
                Return True
            Else
                'Sheets are different.  Find and record the different cells.
                Select Case KeyValueColumnA
                    Case True
                        'Use column A as key value
                        'Compare each row in original sheet
                        Dim keyValueColumnOriginal As Range = originalSheet.Range("A1:A" & originalSheet.UsedRange.Rows.Count)
                        Dim keyValueColumnNew As Range = newSheet.Range("A1:A" & newSheet.UsedRange.Rows.Count)
                        For myKeyValueCounter = 1 To originalSheet.UsedRange.Rows.Count
                            Dim myKeyValue As String = originalSheet.Range("A" & myKeyValueCounter).Text
                            'Locate key value row in newSheet
                            Dim myOriginalRow As Range = originalSheet.Range(originalSheet.Cells(myKeyValueCounter, 1),
                                                                                originalSheet.Cells(myKeyValueCounter, originalSheet.UsedRange.Columns.Count))
                            Try
                                Dim myKeyValueRow As Double = ExcelHandler.ExcelApp.WorksheetFunction.Match(myKeyValue, keyValueColumnNew, 0)
                                'Row was found.  Compare the rows.
                                Dim myNewRow As Range = newSheet.Range(newSheet.Cells(myKeyValueCounter, 1),
                                                                       newSheet.Cells(myKeyValueCounter, newSheet.UsedRange.Columns.Count))
                                If Not (myOriginalRow Is myNewRow) Then
                                    'Rows are different.  Find the modified cells and add them to the collection.
                                    result.Result = SheetResultType.SHEET_DIFFERENT
                                    For myCellCounter = 2 To myOriginalRow.Columns.Count
                                        Dim myOriginalCell As Range = originalSheet.Cells(myKeyValueCounter, myCellCounter)
                                        Dim myNewCell As Range = newSheet.Cells(myKeyValueCounter, myCellCounter)
                                        If Not (myOriginalCell Is myNewCell) Then
                                            'Cells are different.  Add to collection.
                                            Dim myCellComparisonResult As New CellComparisonResults
                                            With myCellComparisonResult
                                                .Result = CellResultType.CELL_MODIFIED
                                                .Cell = myNewCell
                                            End With
                                            result.CellList.Add(myCellComparisonResult)
                                        End If
                                    Next
                                    If myNewRow.Columns.Count > myOriginalRow.Columns.Count Then
                                        For myCellCounter = myOriginalRow.ColumnsCount + 1 To myNewRow.columnsCount
                                            'Add new cells to the collection
                                            Dim myCellComparisonResult As New CellComparisonResults
                                            Dim myNewCell As Range = newSheet.Cells(myKeyValueCounter, myCellCounter)
                                            With myCellComparisonResult
                                                .Result = CellResultType.CELL_ADDED
                                                .Cell = myNewCell
                                            End With
                                        Next
                                    End If
                                End If
                            Catch ex As Exception
                                'Row was deleted in newSheet.  Save row in collection as deleted.
                                result.Result = SheetResultType.SHEET_DIFFERENT
                                For Each item In myOriginalRow
                                    Dim myCellComparisonResult As New CellComparisonResults
                                    With myCellComparisonResult
                                        .Result = CellResultType.CELL_DELETED
                                        .Cell = item
                                    End With
                                    result.CellList.Add(myCellComparisonResult)
                                Next
                            End Try
                        Next
                        Return True
                    Case Else
                        'Perform row-by-row comparison
                        'TODO:  Implement
                        Return True
                End Select
            End If
        Catch ex As Exception
            'The new sheet isn't in the original workbook.
            result.Result = SheetResultType.SHEET_NEW
            Debug.Print("Exception in CompareWorkbooks.CompareSheet.  " & ex.ToString)
            Return True
        End Try
    End Function
    Private Sub GenerateReport(results As Collection)

    End Sub
End Class
