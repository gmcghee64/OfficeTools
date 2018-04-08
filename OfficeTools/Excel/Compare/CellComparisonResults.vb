Imports Microsoft.Office.Interop.Excel
Imports OfficeTools
Imports OfficeTools.ComparisonTypes
''' <summary>
''' This class contains the results of comparing two cells
''' </summary>
Public Class CellComparisonResults
    Private _Cell As Range
    Private _Result As CellResultType

    ''' <summary>
    ''' The Cell property contains a copy of the cell that was modified or added to the sheet.
    ''' </summary>
    ''' <returns></returns>
    Public Property Cell As Range
        Get
            Return _Cell
        End Get
        Set(value As Range)
            _Cell = value
        End Set
    End Property

    ''' <summary>
    ''' The Result property indicates whether or not the Cell was added or modified.
    ''' </summary>
    ''' <returns></returns>
    Public Property Result As CellResultType
        Get
            Return _Result
        End Get
        Set(value As CellResultType)
            _Result = value
        End Set
    End Property
End Class
