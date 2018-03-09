Public Class MainForm

    Private Sub PrepareIDDInputCollectorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrepareIDDInputCollectorToolStripMenuItem.Click
        'IDD_PreparationForm.Show()
        Dim myInputCollector As New IDD_PreparationForm
        myInputCollector.MdiParent = Me
        myInputCollector.Show()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub MergeIDDInputsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MergeIDDInputsToolStripMenuItem.Click
        Dim myConsolidator As New IDD_ConsolidationForm
        myConsolidator.MdiParent = Me
        myConsolidator.Show()
    End Sub

    Private Sub FormatDeliveredIDDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FormatDeliveredIDDToolStripMenuItem.Click
        Dim myFormatter = New FormatIDD
        myFormatter = Nothing
    End Sub

    Private Sub CompareWorkbooksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompareWorkbooksToolStripMenuItem.Click
        Dim myComparer As New WorkbookCompare
        myComparer.MdiParent = Me
        myComparer.Show()
    End Sub
End Class
