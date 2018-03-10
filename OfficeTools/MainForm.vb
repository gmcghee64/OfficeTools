''' <summary>
''' The MainForm is the primary menu for accessing the office productivity tools.
''' </summary>
Public Class MainForm
    ''' <summary>
    ''' This operation opens the Prepare IDD input application.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub PrepareIDDInputCollectorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrepareIDDInputCollectorToolStripMenuItem.Click
        'IDD_PreparationForm.Show()
        Dim myInputCollector As New IDD_PreparationForm With {
            .MdiParent = Me
        }
        myInputCollector.Show()
    End Sub

    ''' <summary>
    ''' This operation terminates the office productivity tools application.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    ''' <summary>
    ''' This operation opens the Merge IDD Inputs application.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub MergeIDDInputsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MergeIDDInputsToolStripMenuItem.Click
        Dim myConsolidator As New IDD_ConsolidationForm With {
            .MdiParent = Me
        }
        myConsolidator.Show()
    End Sub

    ''' <summary>
    ''' This operation opens the Format Delivered IDD application.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub FormatDeliveredIDDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FormatDeliveredIDDToolStripMenuItem.Click
        Dim myFormatter = New FormatIDD
        myFormatter = Nothing
    End Sub

    ''' <summary>
    ''' This operation opens the Compare Workbooks application.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub CompareWorkbooksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompareWorkbooksToolStripMenuItem.Click
        Dim myComparer As New WorkbookCompare With {
            .MdiParent = Me
        }
        myComparer.Show()
    End Sub

    ''' <summary>
    ''' This operation opens the Acronym Builder application.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AcronymBuilderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AcronymBuilderToolStripMenuItem.Click
        Dim myAcronymBuilder = New AcronymBuilderForm With {
            .MdiParent = Me
        }
        myAcronymBuilder.Show()
    End Sub

    ''' <summary>
    ''' This operation opens the Annotate IDD application.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub AnnotateIDDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnnotateIDDToolStripMenuItem.Click
        'TODO:  Implement this function
        MsgBox("Not implemented")
    End Sub

    ''' <summary>
    ''' This operation opens the Generate IDD Annotation pages application.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub GenerateAnnotationPagesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateAnnotationPagesToolStripMenuItem.Click
        'TODO:  Implement this function
        MsgBox("Not implemented")
    End Sub
End Class
