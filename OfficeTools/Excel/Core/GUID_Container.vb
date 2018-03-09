Public Class GUID_Container
    Dim _GUID As String
    Dim _CellValue As String

    Public Property GUID As String
        Get
            Return _GUID
        End Get
        Set(value As String)
            _GUID = value
        End Set
    End Property

    Public Property CellValue As String
        Get
            Return _CellValue
        End Get
        Set(value As String)
            _CellValue = value
        End Set
    End Property
End Class
