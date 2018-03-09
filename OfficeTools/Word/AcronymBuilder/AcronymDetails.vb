Imports Microsoft.Office.Interop.Word
''' <summary>
''' This class defines properties for an acronym appearing in a word document.
''' </summary>
Public Class AcronymDetails
    Private _Acronym As String
    Private _Definition As String
    Private _FirstOccurrence As Bookmark

    ''' <summary>
    ''' The Acronym property is a string that is found in the analyzed document that is in all caps.
    ''' </summary>
    ''' <returns></returns>
    Public Property Acronym As String
        Get
            Return _Acronym
        End Get
        Set(value As String)
            _Acronym = value
        End Set
    End Property

    ''' <summary>
    ''' The Definition property is the string that provides the acronym definition.
    ''' </summary>
    ''' <returns></returns>
    Public Property Definition As String
        Get
            Return _Definition
        End Get
        Set(value As String)
            _Definition = value
        End Set
    End Property

    ''' <summary>
    ''' The FirstOccurrence property is a bookmark in the document that marks the first location at which the acronym
    ''' appears.
    ''' </summary>
    ''' <returns></returns>
    Public Property FirstOccurrence As Bookmark
        Get
            Return _FirstOccurrence
        End Get
        Set(value As Bookmark)
            _FirstOccurrence = value
        End Set
    End Property
End Class
