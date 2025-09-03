
'Imports System.Drawing
'Imports System.Drawing.Imaging
'Imports System.Windows
Imports Microsoft.Office.Interop.Word


Public Class clsWordTest

    Private Shared Function textExtraktionTeil1(strall As String) As String
        Try
            strall = strall.Replace(vbCrLf, " "c)
            strall = strall.Replace(vbLf, " "c)
            strall = strall.Replace(vbTab, " "c)
            strall = strall.Replace(vbCr, " "c)
            strall = strall.Replace(vbVerticalTab, " "c)
            'strAll = strAll.Replace(vbNullString, " "c)
            strall = strall.Replace(vbNewLine, " "c)
            strall = strall.Replace(vbFormFeed, " "c)
            strall = strall.Replace(vbCr, " "c)
            strall = strall.Replace(ChrW(7), " "c)
            strall = strall.Replace(ChrW(21), " "c)
            strall = strall.Replace("kreis offenbach", " "c)
            strall = strall.Replace("Kreis Offenbach", " "c)
            strall = strall.Replace("FD Umwelt", " "c)
            strall = strall.Replace("·", " "c)
            strall = strall.Replace(". ", " "c)
            strall = strall.Replace(", ", " "c)
            strall = strall.Replace(": ", " "c)
            strall = strall.Replace("? ", " "c)
            strall = strall.Replace("! ", " "c)
            strall = strall.Replace("", " "c)
            For i = 1 To 100
                strall = strall.Replace("  ", " "c)
            Next
            strall = strall.Replace("  ", " "c)
            Return strall
        Catch ex As Exception
            Return ""
        End Try
    End Function
    'Public Shared Function konvOneDoc2pdf(infile As String, Pdffile As String) As Boolean
    '    Dim objWord As New Application
    '    Dim objDoc As Document
    '    Dim YourSourcePath, pdfpath As Object
    '    Dim retvalue As Boolean = False
    '    YourSourcePath = CType(infile, Object)
    '    pdfpath = CType(Pdffile, Object)
    '    ' YourSourcePath = "C:\muell\34463\16070850_174865.docx"
    '    If YourSourcePath = "C:\muell\34463\16070850_174865.docx" Then
    '        Debug.Print("")
    '    End If
    '    Try
    '        objDoc = objWord.Documents.OpenNoRepairDialog(YourSourcePath)
    '        If objDoc.HasPassword Then
    '            objDoc.Close()
    '            retvalue = True
    '            Return retvalue
    '        End If
    '        objDoc.SaveAs2(pdfpath, 17)
    '        If objDoc IsNot Nothing Then
    '            objDoc.Close()
    '            retvalue = True
    '        Else
    '            retvalue = True
    '        End If

    '        Return retvalue
    '    Catch ex As Exception
    '        Console.WriteLine(("fehler: " & ex.Message.ToString()))
    '        Return False
    '    Finally

    '        objWord.Application.Quit(False)
    '        objDoc = Nothing
    '        objWord = Nothing
    '        ReleaseComObject(objWord)
    '        ReleaseComObject(objDoc)

    '    End Try
    'End Function
    'Public Shared Function konvOneDoc2Docx(infile As String, Pdffile As String) As Boolean
    '    Dim objWord As New Application
    '    Dim objDoc As Document
    '    Dim YourSourcePath, pdfpath As Object
    '    Dim retvalue As Boolean = False
    '    YourSourcePath = CType(infile, Object)
    '    pdfpath = CType(Pdffile, Object)
    '    ' YourSourcePath = "C:\muell\34463\16070850_174865.docx"
    '    If YourSourcePath = "C:\muell\34463\16070850_174865.docx" Then
    '        Debug.Print("")
    '    End If
    '    Try
    '        objDoc = objWord.Documents.OpenNoRepairDialog(YourSourcePath)
    '        objDoc.SaveAs2(pdfpath, 16) '16 'objDoc.wWdSaveFormat.WordProcessingML
    '        If objDoc IsNot Nothing Then
    '            objDoc.Close()
    '            retvalue = True
    '        Else
    '            retvalue = True
    '        End If

    '        Return retvalue
    '    Catch ex As Exception
    '        Console.WriteLine(("fehler: " & ex.Message.ToString()))
    '        Return False
    '    Finally

    '        objWord.Application.Quit(False)
    '        objDoc = Nothing
    '        objWord = Nothing
    '        ReleaseComObject(objWord)
    '        ReleaseComObject(objDoc)

    '    End Try
    'End Function
    'Private Function GetEncoder(ByVal format As ImageFormat) As ImageCodecInfo

    '    Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageDecoders()

    '    Dim codec As ImageCodecInfo
    '    For Each codec In codecs
    '        If codec.FormatID = format.Guid Then
    '            Return codec
    '        End If
    '    Next codec
    '    Return Nothing

    'End Function
    Public Shared Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub
End Class
