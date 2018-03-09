Imports OfficeTools

Public Class IDD_PreparationConfiguration
    Private _IndexWorksheet As String
    Private _ExportWorksheet As String
    Private _ExcelHandler As ExcelHandler

    ''' <summary>
    ''' This property is used to identify the sheet that contains the list of logical topics to be processed.
    ''' </summary>
    ''' <returns></returns>
    Property IndexWorksheet As String
        Get
            Return _IndexWorksheet
        End Get
        Set
            _IndexWorksheet = Value
        End Set
    End Property

    ''' <summary>
    ''' This property is used to identify the sheet that contains the classes, attributes, and logical topic cross reference from
    ''' the data model.
    ''' </summary>
    ''' <returns></returns>
    Property ExportWorksheet As String
        Get
            Return _ExportWorksheet
        End Get
        Set
            _ExportWorksheet = Value
        End Set
    End Property

    Public Property ExcelHandler As ExcelHandler
        Get
            Return _ExcelHandler
        End Get
        Set(value As ExcelHandler)
            _ExcelHandler = value
        End Set
    End Property

    Protected Overrides Sub Finalize()
    End Sub
End Class
