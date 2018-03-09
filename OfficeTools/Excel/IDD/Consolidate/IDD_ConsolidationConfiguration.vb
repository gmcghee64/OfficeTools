Imports OfficeTools

Public Class IDD_ConsolidationConfiguration
    Private _ExportWorksheet As String
    Private _ExcelHandler As ExcelHandler
    Private _TagsToMerge As Collection

    ''' <summary>
    ''' This property identifies the worksheet on which data will be consolidated
    ''' </summary>
    ''' <returns></returns>
    Public Property ExportWorksheet As String
        Get
            Return _ExportWorksheet
        End Get
        Set(value As String)
            _ExportWorksheet = value
        End Set
    End Property

    ''' <summary>
    ''' This property identifies the excel application and helper functions
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
    ''' This property is the collection of column header names containing data to be merged
    ''' </summary>
    ''' <returns></returns>
    Public Property TagsToMerge As Collection
        Get
            Return _TagsToMerge
        End Get
        Set(value As Collection)
            _TagsToMerge = value
        End Set
    End Property
End Class
