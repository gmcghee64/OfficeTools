Imports Microsoft.Office.Interop.Word

Public Class WordHandler
    ''' <summary>
    '''This class creates an instance of an excel application and provides common functions for processing excel files. 
    ''' </summary>
    Private _myWordApp As New Application

    Public Sub New()
    End Sub

    ''' <summary>
    ''' This is a handle to the excel application instance.  When using the ExcelHandler class, users should retrieve
    ''' this application instance to execute excel functions to manipulate the workbook.  Prior to manipulating the 
    ''' workbook, users will need to execute the OpenWorkbook method to open a workbook.
    ''' </summary>
    ''' <returns>
    ''' The function returns an instance of the active application.
    ''' </returns>
    Public ReadOnly Property WordApp As Application
        Get
            Return _myWordApp
        End Get
    End Property


    Protected Overrides Sub Finalize()
        WordApp.Quit()
        MyBase.Finalize()
    End Sub

    Public Overrides Function ToString() As String
        Return MyBase.ToString()
    End Function

    Public Overrides Function Equals(obj As Object) As Boolean
        Return MyBase.Equals(obj)
    End Function

    Public Overrides Function GetHashCode() As Integer
        Return MyBase.GetHashCode()
    End Function

    ''' <summary>
    ''' This function is used to open a word document for processing.
    ''' </summary>
    ''' <returns>
    ''' If the document is successfully opened, the function returns a value of true.  Otherwise, a value
    ''' of false is returned.
    ''' </returns>
    Public Function OpenDocument() As Boolean
        Try
            With WordApp
                .Visible = True
                .Documents.Open(.GetOpenFilename(FileFilter:="Word Files (*.doc; *.docx), _
                                                                *.doc;*.docx")))
                .Visible = False
                OpenDocument = True
            End With
        Catch ex As Exception
            OpenDocument = False
            WordApp.Visible = False
        End Try
    End Function
End Class
