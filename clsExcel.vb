Imports Microsoft.Office.Interop.Excel
Public Class clsExcel
    Public Shared Function konvOne(xlsfile As String, xlsXfile As String) As Boolean
        Dim objxls As New Application
        Dim retvalue As Boolean = False
        Dim app As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Dim Workbook As Microsoft.Office.Interop.Excel.Workbook = app.Workbooks.Open(xlsfile)
        Try
            app.DisplayAlerts = False
            Workbook.SaveAs(xlsXfile, 51)
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
        Workbook.Close()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook)
        app = Nothing
        Workbook = Nothing
        GC.Collect() '// force final cleanup!
        Return True
    End Function
End Class
