Imports OfficeTools
Public Class AcronymBuilderConfiguration
    Private _WordHelper As WordHandler
    Private _AcronymList As Collection

    Public Property WordHelper As WordHandler
        Get
            Return _WordHelper
        End Get
        Set(value As WordHandler)
            _WordHelper = value
        End Set
    End Property

    Public Property AcronymList As Collection
        Get
            Return _AcronymList
        End Get
        Set(value As Collection)
            _AcronymList = value
        End Set
    End Property
End Class
