Imports System.Data

Module modThumbnail


    Public Sub datenholden(ByRef vid As Integer, ByRef relativpfad As String, ByRef dateinameext As String, ByRef typ As String, ByRef newsavemode As Boolean, ByRef dokumentid As Integer, ByVal drr As DataRow,
                           ByRef datum As Date, ByRef init5ial As String)
        vid = CInt(drr.Item("bearbeiterID"))
        dokumentid = CInt(drr.Item("dokumentid"))
        relativpfad = CStr(drr.Item("relativpfad"))
        dateinameext = CStr(drr.Item("dateinameext"))
        newsavemode = CBool(drr.Item("newsavemode"))
        init5ial = CStr(drr.Item("initial_"))
        datum = CDate(drr.Item("checkindatum"))
        typ = CStr(drr.Item("typ"))
    End Sub
    Public Function GetInputfileWordFullPath(ByVal innDir As String, ByVal relativpfad As String, ByVal dokumentid As Integer) As String
        Dim inputfile As String
        inputfile = innDir & IO.Path.Combine(relativpfad, CType(dokumentid, String))
        inputfile = inputfile.Replace("/", "\")
        '  inputfile = Chr(34) & inputfile & Chr(34)
        Return inputfile
    End Function
    Public Function GetInputfile1WordFullPath(ByVal innDir As String, ByVal relativpfad As String, ByVal dateinameext As String) As String
        Dim inputfile As String
        inputfile = innDir & IO.Path.Combine(relativpfad, dateinameext)
        inputfile = inputfile.Replace("/", "\")
        '   inputfile = Chr(34) & inputfile & Chr(34)
        Return inputfile
    End Function

    Public Function auscheckenword(ByVal inputfile As String, ByVal checkoutfile As String, sw As IO.StreamWriter, vid As String, dokid As String) As Boolean
        Dim fo As New IO.FileInfo(checkoutfile)
        If fo.Exists Then
            Return True
        End If
        Dim fe As IO.FileInfo
        Try
            fe = New IO.FileInfo(inputfile.Replace(Chr(34), ""))
            If Not fe.Exists Then
                l("INPUTFILE EXISTIERT NICHT: " & inputfile)
                sw.WriteLine("INPUTFILE EXISTIERT NICHT: " & vid & "/" & dokid & inputfile)
                Return False
            End If
            fe.CopyTo(checkoutfile, True)
            fe = Nothing
            Return True
        Catch ex As Exception
            l("fehler in auschekcen: " & ex.ToString)
            Return False
        End Try
    End Function
    Public Sub deleteCheckoutfileWord(ByVal checkoutfile As String)
        Dim fi As IO.FileInfo
        Try
            fi = New IO.FileInfo(checkoutfile)
            fi.Delete()

        Catch ex As Exception
            'l("fehler in delete: " & checkoutfile & Environment.NewLine & ex.ToString)
        End Try

    End Sub


    Public Sub alteProizessekillen()
        Try
            For Each Process In System.Diagnostics.Process.GetProcessesByName("WINWORD")
                Process.Kill()
            Next
        Catch ex As Exception

        End Try
    End Sub

    Public Function getCheckoutfileWord(inputfile As String, checkoutRoot As String, dokumentid As Integer,
                                        vid As Integer, prenom As String) As String
        Dim outfile As String
        outfile = checkoutRoot & IO.Path.Combine(vid.ToString, prenom & CType(dokumentid, String))
        outfile = outfile.Replace("/", "\")

        Return outfile
    End Function
    Public Function getOutfileExcel(inputfile As String, checkoutRoot As String, dokumentid As Integer,
                                        vid As Integer, prenom As String) As String
        Dim outfile As String
        outfile = checkoutRoot & IO.Path.Combine(vid.ToString, prenom & CType(dokumentid, String))
        outfile = outfile.Replace("/", "\")

        Return outfile
    End Function

    Public Function wordJob1(inputfile As String, outfile As String) As Boolean
        Dim aaa As New clsWordTest
        Dim resilt As Boolean = aaa.konvOneDoc2pdf(inputfile, outfile)
        aaa = Nothing
        GC.Collect()
        GC.WaitForFullGCComplete()
        Return resilt
    End Function
    Public Function GetOutfileWORD(ByVal vid As Integer, ByVal outDir As String, ByVal dokumentid As Integer) As String
        Dim outfile As String
        outfile = outDir & IO.Path.Combine(vid.ToString, CType(dokumentid, String)) & ".jpg"
        outfile = outfile.Replace("/", "\")
        'outfile = Chr(34) & outfile & Chr(34)
        Return outfile
    End Function
    Public Function GetOutfileEXCEL(ByVal vid As Integer, ByVal outDir As String, ByVal dokumentid As Integer) As String
        Dim outfile As String
        outfile = outDir & IO.Path.Combine(vid.ToString, CType(dokumentid, String)) & ".xlsx"
        outfile = outfile.Replace("/", "\")
        'outfile = Chr(34) & outfile & Chr(34)
        Return outfile
    End Function
End Module
