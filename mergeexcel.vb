Imports System.IO
Imports System.Linq
Imports OfficeOpenXml


Public Class Mergeexcel
    Shared Sub Mergetest()
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim folder = "T:\dokumente\main"
        Dim outPath = Path.Combine(folder, "MERGED.xlsx")

        Dim files = Directory.GetFiles(folder) _
            .Where(Function(f)
                       Dim ext = Path.GetExtension(f).ToLowerInvariant()
                       Return ext = ".xlsx"
                   End Function) _
            .OrderBy(Function(f) f) _
            .ToArray()

        If files.Length = 0 Then
            Console.WriteLine("Keine .xlsx-Dateien gefunden.")
            Return
        End If

        Using outPkg As New ExcelPackage()
            Dim outWs = outPkg.Workbook.Worksheets.Add("Merged")
            Dim nextRow As Integer = 1

            For Each file In files
                Using inPkg As New ExcelPackage(New FileInfo(file))
                    Dim inWs = inPkg.Workbook.Worksheets.FirstOrDefault()
                    If inWs Is Nothing Then Continue For

                    Dim endRow As Integer = 0
                    Dim endCol As Integer = 0
                    If inWs.Dimension IsNot Nothing Then
                        endRow = inWs.Dimension.End.Row
                        endCol = inWs.Dimension.End.Column
                    End If

                    If endRow = 0 OrElse endCol = 0 Then Continue For

                    For r As Integer = 1 To endRow
                        For c As Integer = 1 To endCol
                            outWs.Cells(nextRow, c).Value = inWs.Cells(r, c).Value
                        Next
                        nextRow += 1
                    Next
                End Using
            Next

            If File.Exists(outPath) Then File.Delete(outPath)
            outPkg.SaveAs(New FileInfo(outPath))
        End Using

        Console.WriteLine("Fertig: " & outPath)
    End Sub
End Class


