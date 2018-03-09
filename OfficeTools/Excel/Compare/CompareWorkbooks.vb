Imports Microsoft.Office.Interop.Excel
Imports OfficeTools

Public Class CompareWorkbooks
    Private _ExcelHandler As ExcelHandler
    Private _OriginalWorkbook As Workbook
    Private _NewWorkbook As Workbook
    Private _KeyValueColumnA As Boolean = True
    Private _CompareSingleSheet As Boolean = True
    Private _OriginalSheetToCompare As Worksheet
    Private _NewSheetToCompare As Worksheet

    ''' <summary>
    ''' The ExcelHandler holds the instance of the Excel application.
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

    Public Property OriginalWorkbook As Workbook
        Get
            Return _OriginalWorkbook
        End Get
        Set(value As Workbook)
            _OriginalWorkbook = value
        End Set
    End Property

    Public Property NewWorkbook As Workbook
        Get
            Return _NewWorkbook
        End Get
        Set(value As Workbook)
            _NewWorkbook = value
        End Set
    End Property

    Public Property KeyValueColumnA As Boolean
        Get
            Return _KeyValueColumnA
        End Get
        Set(value As Boolean)
            _KeyValueColumnA = value
        End Set
    End Property

    Public Property CompareSingleSheet As Boolean
        Get
            Return _CompareSingleSheet
        End Get
        Set(value As Boolean)
            _CompareSingleSheet = value
        End Set
    End Property

    Public Property OriginalSheetToCompare As Worksheet
        Get
            Return _OriginalSheetToCompare
        End Get
        Set(value As Worksheet)
            _OriginalSheetToCompare = value
        End Set
    End Property

    Public Property NewSheetToCompare As Worksheet
        Get
            Return _NewSheetToCompare
        End Get
        Set(value As Worksheet)
            _NewSheetToCompare = value
        End Set
    End Property

    Public Function PerformComparison(statusPanel As StatusBar) As Boolean
        If CheckConfiguration() = True Then
            Dim myResultsCollection As New Collection
            Select Case CompareSingleSheet
                Case True
                    Dim myResults As New SheetComparisonResults
                    If CompareSheet(OriginalSheetToCompare, NewSheetToCompare, myResults) = True Then
                        myResultsCollection.Add(myResults)
                        'TODO:  Generate change report in new workbook and save it
                        Return True
                    Else
                        'An error must have occurred.  Return false.
                        Return False
                    End If
                    Exit Select
                Case False
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
                            'TODO:  Need to record result as new sheet.  How to modify enum property?
                            myResultsCollection.Add(myResults)
                        End Try
                    Next
                    'TODO:  Generate change report in new workbook and save it
                    Exit Select
                Case Else
                    Return False
                    Exit Select
            End Select
        Else
            Return False
        End If
    End Function

    Private Function CheckConfiguration() As Boolean
        If Not IsNothing(ExcelHandler) And Not IsNothing(OriginalWorkbook) And
                Not IsNothing(NewWorkbook) Then
            If CompareSingleSheet = True Then
                If Not IsNothing(NewSheetToCompare) And Not IsNothing(OriginalSheetToCompare) Then
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

    Private Function CompareSheet(originalSheet As Worksheet, newSheet As Worksheet, result As SheetComparisonResults) As Boolean
        'TODO:  Compare workwheet
    End Function
End Class
