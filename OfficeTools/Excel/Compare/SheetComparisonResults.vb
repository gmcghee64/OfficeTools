Imports Microsoft.Office.Interop.Excel
Imports OfficeTools
Imports OfficeTools.ComparisonTypes

Public Class SheetComparisonResults
    Private _SheetName As String
    Private _CellList As New Collection
    Private _Result As ResultType

    Public Property SheetName As String
        Get
            Return _SheetName
        End Get
        Set(value As String)
            _SheetName = value
        End Set
    End Property

    Public Property CellList As Collection
        Get
            Return _CellList
        End Get
        Set(value As Collection)
            _CellList = value
        End Set
    End Property

    Private Property Result As ResultType
        Get
            Return _Result
        End Get
        Set(value As ResultType)
            _Result = value
        End Set
    End Property


End Class

