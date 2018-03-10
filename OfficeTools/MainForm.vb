Public Class MainForm

    Private Sub PrepareIDDInputCollectorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrepareIDDInputCollectorToolStripMenuItem.Click
        'IDD_PreparationForm.Show()
        Dim myInputCollector As New IDD_PreparationForm With {
            .MdiParent = Me
        }
        myInputCollector.Show()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub MergeIDDInputsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MergeIDDInputsToolStripMenuItem.Click
        Dim myConsolidator As New IDD_ConsolidationForm With {
            .MdiParent = Me
        }
        myConsolidator.Show()
    End Sub

    Private Sub FormatDeliveredIDDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FormatDeliveredIDDToolStripMenuItem.Click
        Dim myFormatter = New FormatIDD
        myFormatter = Nothing
    End Sub

    Private Sub CompareWorkbooksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompareWorkbooksToolStripMenuItem.Click
        Dim myComparer As New WorkbookCompare With {
            .MdiParent = Me
        }
        myComparer.Show()
    End Sub

    Private Sub AcronymBuilderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AcronymBuilderToolStripMenuItem.Click
        Dim myAcronymBuilder = New AcronymBuilderForm With {
            .MdiParent = Me
        }
        myAcronymBuilder.Show()
    End Sub

    Private Sub AnnotateIDDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnnotateIDDToolStripMenuItem.Click
        'TODO:  Implement this function
        MsgBox("Not implemented")
    End Sub

    Private Sub GenerateAnnotationPagesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateAnnotationPagesToolStripMenuItem.Click
        'TODO:  Implement this function
        MsgBox("Not implemented")
    End Sub
End Class
