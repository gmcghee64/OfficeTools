Imports Microsoft.Office.Interop.Excel

Public Class FormatIDD
    Dim myExcelHandler As New ExcelHandler

    Public Sub New()
        Try
            If myExcelHandler.OpenWorkbook = True Then
                With myExcelHandler.ExcelApp.ActiveWorkbook
                    For Each item In .Sheets
                        Dim myWorksheet As Worksheet = item
                        MainForm.statusLabel.Text = "Format IDD:  Formatting sheet " & myWorksheet.Index & " of " & .Worksheets.Count
                        If myWorksheet.Range("B3").Text = "MESSAGE NAME" Then
                            myWorksheet.Range("A7").Value = "Applicability"
                            Dim myWorksheetLastRow As Integer = myExcelHandler.ExcelApp.WorksheetFunction.CountA(myWorksheet.Range("B:B")) + 2
                            Dim myTableRange As Range = myWorksheet.Range("A7:I" & myWorksheetLastRow)
                            Dim myTable As ListObject = myWorksheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange, myTableRange,, XlYesNoGuess.xlYes)
                            myTable.Name = myWorksheet.Name
                        End If
                    Next
                    .SaveAs(Left(.Name, Len(.Name) - 5) & "_" & Now().ToString("yyyyMMdd_hh_mm") & ".xlsx")
                    .Close()
                    MainForm.statusLabel.Text = "Format IDD:  IDD formatting complete."
                End With
            End If
        Catch ex As Exception
            MainForm.statusLabel.Text = "FormatIDD:  Unhandled exception" & ex.ToString
            Debug.Print("Exception in FormatIDD.New() - " & ex.ToString)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
