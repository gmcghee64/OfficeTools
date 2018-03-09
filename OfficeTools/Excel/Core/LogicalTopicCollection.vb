''' <summary>
''' This class is used to maintain a sorted collection of LogicalTopic classes
''' </summary>
Public Class LogicalTopicCollection
    Private _TopicCollection As New Collections.SortedList

    ''' <summary>
    ''' This property defiunes a sorted collection of logical topics.  It should be populated with
    ''' LogicalTopic objects.
    ''' </summary>
    ''' <returns></returns>
    Property TopicCollection As Collections.SortedList
        Get
            Return _TopicCollection
        End Get
        Set
            _TopicCollection = Value
        End Set
    End Property

End Class
