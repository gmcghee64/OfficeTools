Imports Microsoft.Office.Interop.Excel

Public Class ExcelHandler
    ''' <summary>
    '''This class creates an instance of an excel application and provides common functions for processing excel files. 
    ''' </summary>
    Private _myExcelApp As New Application
    Private _mySheets As Sheets
    Private _myActiveWorkbook As Workbook

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
    Public ReadOnly Property ExcelApp As Application
        Get
            Return _myExcelApp
        End Get
    End Property

    ''' <summary>
    ''' This is the collection of worksheets in the active workbook
    ''' </summary>
    ''' <returns></returns>
    Public Property SheetsInWorkbook As Sheets
        Get
            Return _mySheets
        End Get
        Set(value As Sheets)
            _mySheets = value
        End Set
    End Property

    Public Property ActiveWorkbook As Workbook
        Get
            Return _myActiveWorkbook
        End Get
        Set(value As Workbook)
            _myActiveWorkbook = value
        End Set
    End Property

    Protected Overrides Sub Finalize()
        Try
            ExcelApp.Quit()
        Catch ex As Exception
            Debug.Print("Exception in ExcelHandler.Finalize.  " & ex.ToString)
        End Try
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
    ''' This function is used to open a workbook for processing.
    ''' </summary>
    ''' <returns>
    ''' If the workbook is successfully opened, the function returns a value of true.  Otherwise, a value
    ''' of false is returned.
    ''' </returns>
    Public Function OpenWorkbook() As Boolean
        Try
            With ExcelApp
                .Visible = True
                .Workbooks.Open(.GetOpenFilename(FileFilter:="Excel Files (*.xls; *.xlsx), _
                                                                *.xls;*.xlsx"))
                .Visible = False
                OpenWorkbook = True
                SheetsInWorkbook = .Sheets
                ActiveWorkbook = .ActiveWorkbook
            End With
        Catch ex As Exception
            OpenWorkbook = False
            ExcelApp.Visible = False
        End Try
    End Function
End Class
