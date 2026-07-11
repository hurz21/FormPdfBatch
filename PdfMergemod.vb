Imports iText.Kernel.Pdf
Imports iText.Kernel.Utils


Imports iText.Kernel.Pdf.Canvas
Imports iText.Kernel.Utils.PdfMerger



Module PdfMerge


    Public Sub MergePdfs(outputPath As String, inputPaths As IEnumerable(Of String))
        Using targetWriter As New PdfWriter(outputPath)
            Using targetDoc As New PdfDocument(targetWriter)
                Dim merger As New PdfMerger(targetDoc)

                For Each p In inputPaths
                    If String.IsNullOrWhiteSpace(p) OrElse Not IO.File.Exists(p) Then
                        Throw New IO.FileNotFoundException("PDF nicht gefunden.", p)
                    End If

                    Using srcReader As New PdfReader(p)
                        Using srcDoc As New PdfDocument(srcReader)
                            merger.Merge(srcDoc, 1, srcDoc.GetNumberOfPages())
                        End Using
                    End Using
                Next
            End Using
        End Using
    End Sub

    Sub testmerge()
        Dim inputs As String() = {
            "C:\Users\Feinen_j\Documents\itext\a.pdf",
            "C:\Users\Feinen_j\Documents\itext\b.pdf",
            "C:\Users\Feinen_j\Documents\itext\c.pdf"
        }

        Dim outFile As String = "C:\Users\Feinen_j\Documents\itext\merged.pdf"
        MergePdfs(outFile, inputs)
    End Sub

End Module


