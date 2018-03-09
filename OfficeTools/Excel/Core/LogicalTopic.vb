''' <summary>
''' This class is used to define the ranges that contains data for a given logical topic.
''' </summary>
Public Class LogicalTopic
    Private _Topic As String
    Private _DataRange As New Collection

    ''' <summary>
    ''' This property contains the name of the logical topic.
    ''' </summary>
    ''' <returns>Returns the topic name.
    ''' </returns>
    Property Topic As String
        Get
            Return _Topic
        End Get
        Set
            _Topic = Value
        End Set
    End Property

    ''' <summary>
    ''' This property contains a collection of the ranges containing data associated with the logical topic.
    ''' </summary>
    ''' <returns></returns>
    Property DataRange As Collection
        Get
            Return _DataRange
        End Get
        Set
            _DataRange = Value
        End Set
    End Property

End Class
