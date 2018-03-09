Imports OfficeTools

Public Class AcronymBuilderForm
    ''' <summary>
    ''' The myAcronymBuilder property contains the configuration settings for the acronym builder session.
    ''' </summary>
    Private myAcronymBuilder As New AcronymBuilderConfiguration
    ''' <summary>
    ''' The myWordApp property is the Word application instance in the configuration settings.
    ''' </summary>
    Private myWordApp = myAcronymBuilder.WordHelper

    Private Sub ButtonExit_Click(sender As Object, e As EventArgs) Handles ButtonExit.Click
        Try
            myAcronymBuilder = Nothing
            myWordApp = Nothing
        Catch ex As Exception
            Debug.Print("Exception in AcronymBuilderForm.ButtonExit_Click.  " & ex.ToString)
        End Try
    End Sub

    Private Sub ListBoxAcronyms_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBoxAcronyms.SelectedIndexChanged
        'TODO:  Implement this function
    End Sub

    Private Sub CheckBoxGenerateList_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBoxGenerateList.CheckedChanged
        myAcronymBuilder.GenerateAcronymListPage = Me.CheckBoxGenerateList.Enabled
    End Sub

    Private Sub ButtonOpenDocument_Click(sender As Object, e As EventArgs) Handles ButtonOpenDocument.Click
        Try
            myAcronymBuilder.WordHelper.OpenDocument()
            'TODO:  Scan the document and build the acronym collection
        Catch ex As Exception
            Debug.Print("Exception in AcronymBuilderForm.ButtonOpenDocument_Click.  " & ex.ToString)
        End Try
    End Sub

    Private Sub ButtonUpdateDocument_Click(sender As Object, e As EventArgs) Handles ButtonUpdateDocument.Click
        'TODO:  Implement this function
    End Sub

    Private Sub TextBoxAcronymDefinition_ModifiedChanged(sender As Object, e As EventArgs) Handles TextBoxAcronymDefinition.ModifiedChanged
        'TODO:  Implement this function
    End Sub
End Class