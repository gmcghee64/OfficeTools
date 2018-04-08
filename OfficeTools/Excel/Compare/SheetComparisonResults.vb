Imports OfficeTools.ComparisonTypes

''' <summary>
''' This class contains the results from comparing two worksheets.
''' </summary>
Public Class SheetComparisonResults
    Private _SheetName As String
    Private _CellList As New Collection
    Private _Result As SheetResultType

    ''' <summary>
    ''' The SheetName property contains the name of the sheet that was evaluated.
    ''' </summary>
    ''' <returns></returns>
    Public Property SheetName As String
        Get
            Return _SheetName
        End Get
        Set(value As String)
            _SheetName = value
        End Set
    End Property

    ''' <summary>
    ''' The CellList property contains the collection of cells that are different.  This collection is populated
    ''' with CellComparisonResults objects.  This collection will be empty if the sheets are identical or if the
    ''' sheet is new.
    ''' </summary>
    ''' <returns></returns>
    Public Property CellList As Collection
        Get
            Return _CellList
        End Get
        Set(value As Collection)
            _CellList = value
        End Set
    End Property

    ''' <summary>
    ''' The Result property contains the result of the sheet comparison.
    ''' </summary>
    ''' <returns></returns>
    Public Property Result As SheetResultType
        Get
            Return _Result
        End Get
        Set(value As SheetResultType)
            _Result = value
        End Set
    End Property


End Class

