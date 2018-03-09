Imports OfficeTools
Public Class AcronymBuilderConfiguration
    ''' <summary>
    ''' The WordHelper property is a reference to the instance of the Word application.  It also contains commonly
    ''' used functions.
    ''' </summary>
    Private _WordHelper As WordHandler
    ''' <summary>
    ''' The AcronymList property is used to store the list of acronym objects.  It should be populated with objects
    ''' defined by the AcronymDetails class.
    ''' </summary>
    Private _AcronymList As Collection
    ''' <summary>
    ''' The GenerateAcronymListPage property indicates whether or not a table of acronyms will be appended to the end
    ''' of the document.
    ''' </summary>
    Private _GenerateAcronymListPage As Boolean

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

    Public Property GenerateAcronymListPage As Boolean
        Get
            Return _GenerateAcronymListPage
        End Get
        Set(value As Boolean)
            _GenerateAcronymListPage = value
        End Set
    End Property
End Class
