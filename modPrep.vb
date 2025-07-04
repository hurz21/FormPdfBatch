Module modPrep
    Public Function getVid() As String
        Dim vid As String
        Dim cmd As String
        Dim a As String()
        Try
            cmd = Environment.CommandLine

            If cmd.ToLower.Contains("vshost") Then
                vid = "9609"
            Else
                a = cmd.Split(" "c)
                cmd = a(1).Trim
                vid = cmd
                If Not IsNumeric(vid) Then
                    Return "fehler not numeric"
                End If
            End If
            Return vid
        Catch ex As Exception
            Return "fehler"
        End Try
    End Function



    Public Sub DbMetaDatenDokumentHolen(ByRef vid As String, ByRef relativpfad As String, ByRef dateinameext As String,
                           ByRef typ As String, ByRef newsavemode As Boolean, ByRef dokumentid As String,
                           ByVal drr As DataRow, ByRef datumDB As Date, ByRef istRevisionssicher As Boolean,
ByRef initial As String, ByRef eid As Integer, ByRef beschreibung As String, ByRef eingang As Date, ByRef fullfilename As String)
        Try
            vid = CStr(drr.Item("Bearbeitungsart"))
            dokumentid = CStr(drr.Item("dokumentid"))
            eid = CStr(drr.Item("eid"))
            relativpfad = CStr(drr.Item("relativpfad"))
            dateinameext = CStr(drr.Item("dateinameext"))
            newsavemode = CBool(drr.Item("newsavemode"))
            datumDB = CDate(drr.Item("checkindatum"))
            eingang = CDate(drr.Item("checkindatum"))
            initial = CStr(drr.Item("initial_"))
            istRevisionssicher = CBool(drr.Item("revisionssicher"))
            beschreibung = CStr(drr.Item("Beschreibung"))
            fullfilename = CStr(drr.Item("tooltip"))
            typ = CStr(drr.Item("typ"))
        Catch ex As Exception
            l("fehler in DbMetaDatenDokumentHolen:" & vid & ex.ToString)
        End Try
    End Sub
    Public Function GetInputfilename(ByVal innDir As String, ByVal relativpfad As String, ByVal dokumentid As Integer) As String
        Dim inputfile As String
        inputfile = innDir & IO.Path.Combine(relativpfad, CType(dokumentid, String))
        inputfile = inputfile.Replace("/", "\")
        '  inputfile = Chr(34) & inputfile & Chr(34)
        Return inputfile
    End Function
    Public Function GetInputfile1Name(ByVal innDir As String, ByVal relativpfad As String, ByVal dateinameext As String) As String
        Dim inputfile As String
        inputfile = innDir & IO.Path.Combine(relativpfad, dateinameext)
        inputfile = inputfile.Replace("/", "\")
        ' inputfile = Chr(34) & inputfile & Chr(34)
        Return inputfile
    End Function

    Public Function auschekcen(ByVal inputfile As String, ByVal checkoutfile As String) As Boolean
        Dim fe As IO.FileInfo
        Try
            l("auschekcen: " & inputfile & "  " & checkoutfile)
            fe = New IO.FileInfo(inputfile.Replace(Chr(34), ""))

            fe.CopyTo(checkoutfile, True)
            fe = Nothing
            Return True
        Catch ex As Exception
            l("fehler in auschekcen:  " & ex.ToString)
            Return False
        End Try
    End Function
    Public Sub deleteCheckoutfile(ByVal checkoutfile As String)
        Dim fi As IO.FileInfo
        Try
            fi = New IO.FileInfo(checkoutfile)
            fi.Delete()

        Catch ex As Exception
            'l("fehler in delete: " & ex.ToString)
        End Try

    End Sub

    Public Function GetOutfileName(ByVal vid As Integer, ByVal outDir As String, ByVal dokumentid As Integer, endung As String) As String
        Dim outfile As String
        outfile = outDir & IO.Path.Combine(vid.ToString, CType(dokumentid, String)) & endung
        outfile = outfile.Replace("/", "\")
        'outfile = Chr(34) & outfile & Chr(34)
        Return outfile
    End Function

    Public Function getCheckoutfile(inputfile As String, checkoutRoot As String, dokumentid As Integer, vid As Integer) As String
        Dim outfile As String
        outfile = checkoutRoot & IO.Path.Combine(vid.ToString, CType(dokumentid, String))
        outfile = outfile.Replace("/", "\")
        Return outfile
    End Function
    Sub l(t As String)
        nachricht(t)
    End Sub

    Sub nachricht(t As String)
        '  Form1.sw.WriteLine(t)
        '  Console.WriteLine(t)
        My.Application.Log.WriteEntry(t)
    End Sub
End Module
