Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices

Public Class Form1
    Public wordVorlagen As New Microsoft.Office.Interop.Word.Application 'habe hier new ergänzt ????
    Public docVorlagen As New Microsoft.Office.Interop.Word.Document

    Private immerUeberschreiben As Boolean
    Private nichtUeberschreibenAusserWennNeuer As Boolean
    Property dt As New Data.DataTable
    Private count As Integer = 0

    Public Shared inndir, outdir, vid, checkoutRoot As String
    Public Shared nichtUeberschreiben As Boolean = True
    Public Shared sw As IO.StreamWriter
    Public Shared swfehlt As IO.StreamWriter
    Public batchmode As Boolean = False
    Private icntREADONLYentfernt As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Minimized
        protokoll()
        ' fullpathdokumenteErzeugen()
        'PDFumwandeln()
        'DOCXumwandeln(2113, False)
    End Sub
    Public Sub protokoll()
        With My.Application.Log.DefaultFileLogWriter
#If DEBUG Then
            .CustomLocation = "L:\system\batch\margit\" & ""
#Else
            .CustomLocation = "L:\system\batch\margit\" & ""
#End If

            .BaseFileName = "form2pdf" & "_" & Environment.UserName
            .AutoFlush = True
            .Append = False
        End With
        ' zeitStart = Now
        l("protokoll now: " & Now)
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            PDFumwandeln()
        Catch ex As Exception
            Debug.Print("")
        End Try


    End Sub
    Private Shared Function RevSicherdokumentDatenHolen(sql As String) As DataTable

        Dim dt As New DataTable
        Try

            'MsgBox(Sql)
            dt = getDT(sql)
            'MsgBox(dt.Rows.Count)
            'l("nach getDT")
            Return dt
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Private Shared Function PDFdokumentDatenHolen() As DataTable
        Dim Sql As String
        Dim dt As New DataTable
        Try
            Sql = "SELECT * FROM dokumente where   dokumentid<2000000 and dokumentid>0  " &
                  " and ( typ='pdf') order by dokumentid desc "
            'MsgBox(Sql)
            dt = getDT(Sql)
            'MsgBox(dt.Rows.Count)
            'l("nach getDT")
            Return dt
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Private Shared Function alleDokumentDatenHolen(sql As String) As DataTable

        Dim dt As New DataTable
        Try

            'MsgBox(Sql)
            dt = getDT(sql)
            'MsgBox(dt.Rows.Count)
            'l("nach getDT")
            Return dt
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Private Shared Function alleDokumentDatenHolenohnemb() As DataTable
        Dim Sql As String
        Dim dt As New DataTable
        Try
            Sql = "SELECT * FROM dokumente where   dokumentid<2000000 and dokumentid>0  " &
                  "  and mb =0 " &
                  "  order by dokumentid desc "
            'MsgBox(Sql)
            dt = getDT(Sql)
            'MsgBox(dt.Rows.Count)
            'l("nach getDT")
            Return dt
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Friend Shared Function getFileSize4Length(mySize As Double) As String
        Dim result As String = ""
        Try
            l(" MOD getFileSize4Length anfang")
            Select Case mySize
                Case 0 To 1023
                    Return mySize & " Bytes"
                Case 1024 To 1048575
                    Return Format(mySize / 1024, "###0.00") & " KB"
                Case 1048576 To 1043741824
                    Return Format(mySize / 1024 ^ 2, "###0.00") & " MB"
                Case Is > 1043741824
                    Return Format(mySize / 1024 ^ 3, "###0.00") & " GB"
            End Select
            Return "0 bytes"
            l(" MOD getFileSize4Length ende")

        Catch ex As Exception
            l("Fehler in getFileSize4Length: " & ex.ToString())
            Return result
        End Try
    End Function

    'Private Shared Function convertPDF(ByVal filename As String, outfilename As String) As Boolean
    '    'http://www.codeproject.com/Articles/37637/View-PDF-files-in-C-using-the-Xpdf-and-muPDF-libra
    '    Dim _pdfdoc = New PDFLibNet.PDFWrapper
    '    Dim pic As PictureBox = New PictureBox
    '    Dim backbuffer As Bitmap
    '    Try
    '        _pdfdoc.LoadPDF(filename)
    '        _pdfdoc.CurrentPage = 1
    '        pic.Width = 800
    '        pic.Height = 1024
    '        _pdfdoc.FitToWidth(pic.Handle)
    '        pic.Height = _pdfdoc.PageHeight
    '        _pdfdoc.RenderPage(pic.Handle)
    '        backbuffer = New Bitmap(_pdfdoc.PageWidth, _pdfdoc.PageHeight)
    '        Using g As Graphics = Graphics.FromImage(backbuffer)
    '            _pdfdoc.RenderPage(g.GetHdc)
    '            g.ReleaseHdc()
    '        End Using
    '        pic.Image = backbuffer
    '        'filename = getJPGfilename(filename, outdir, vid)
    '        _pdfdoc.ExportJpg(outfilename, 1, 1, 150, 5, -1)
    '        ' _pdfdoc.ExportText(filename, 1, 1, True, True)
    '        Return True
    '    Catch ex As Exception
    '        Return False
    '    Finally
    '        backbuffer.Dispose()
    '        backbuffer = Nothing
    '        pic.Dispose()
    '        pic = Nothing
    '        _pdfdoc.Dispose()
    '        _pdfdoc = Nothing
    '        GC.Collect()
    '        GC.WaitForFullGCComplete()
    '    End Try
    'End Function


    Private Shared Function getJPGfilename(filename As String, outdir As String, vid As String) As String
        Dim fi As New IO.FileInfo(filename)
        Dim outfile As String = outdir & vid & "\" & fi.Name.Replace(".pdf", ".jpg")
        fi = Nothing
        Return outfile
    End Function

    Private Sub DokExistsMain()
        Dim DT As DataTable
        Dim logfile As String = "\\file-paradigma\paradigma\test\thumbnails\dokuFehlt_" & Format(Now, "ddhhmmss") & ".txt"
        'logfile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\paradigma\muell\thumbnailer.log"
        sw = New IO.StreamWriter(logfile)
        sw.AutoFlush = True
        sw.WriteLine(Now)
        If vid = "fehler" Then End
        Dim Sql As String
        Sql = "SELECT * FROM dokumente where   dokumentid<2000000 and dokumentid>0  " &
                  "  order by dokumentid desc "
        DT = alleDokumentDatenHolen(Sql)
        l("vor prüfung")
        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        Dim ic As Integer = 0
        Dim eid As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim initial As String
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenHolen(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid)
                '   l(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")
                '   sw.WriteLine(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")
                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If

                If istRevisionssicher Then
                    If CheckBox1.Checked Then
                        inputFileReadonlysetzen(inputfile)
                    End If
                Else
                    If inputFileReadonlyEntfernen(inputfile) Then
                        icntREADONLYentfernt += 1
                        l("icntREADONLYentfernt: " & inputfile)
                    End If

                End If
                Dim fo As New IO.FileInfo(inputfile)
                TextBox3.Text = igesamt & " von " & DT.Rows.Count
                Application.DoEvents()
                If fo.Exists Then
                    'l("exists")
                    'inputFileReadonlyEntfernen(inputfile)
                    Continue For
                Else
                    ic += 1
                    l("dokument fehlt: " & ic.ToString & Environment.NewLine & " " &
                                     inputfile & Environment.NewLine &
                                     vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")")


                    TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                      inputfile & Environment.NewLine &
                      vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                      TextBox2.Text
                    Application.DoEvents()
                End If

            Catch ex As Exception
                l("fehler1: " & ex.ToString)
            End Try
        Next
        Debug.Print(icntREADONLYentfernt)
        Process.Start(logfile)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        DokExistsMain()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DOCXumwandeln(2113, False)
    End Sub



    'Private Sub PDFSverarbeiten(outdir As String, vid As String, dt As DataTable)
    '    Dim ic As Integer = 0
    '    Dim relativpfad As String = "", dateinameext As String = "", typ As String, batchfile As String
    '    Dim newsavemode As Boolean
    '    Dim dbdatum As Date
    '    Dim inputfile, checkoutfile, outfile As String
    '    Dim dokumentid As String
    '    For Each drr As DataRow In dt.Rows
    '        Try
    '            ic += 1
    '            DbMetaDatenHolen(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum)
    '            l(vid & " " & dokumentid.ToString & " " & ic & " (" & dt.Rows.Count & ")")
    '            TextBox1.Text = TextBox1.Text & vid & " " & dokumentid.ToString & " " & ic & " (" & dt.Rows.Count & ")"
    '            If newsavemode Then
    '                inputfile = GetInputfile(inndir, relativpfad, CInt(dokumentid))
    '            Else
    '                inputfile = GetInputfile1(inndir, relativpfad, dateinameext)
    '            End If
    '            outfile = modPrep.GetOutfile(CInt(vid), outdir, CInt(dokumentid), ".jpg")
    '            Dim fi As New IO.FileInfo(outfile.Replace(Chr(34), ""))

    '            If fi.Exists Then
    '                l("exists")
    '                Continue For
    '            End If
    '            If Not IO.Directory.Exists(outdir & vid.ToString) Then
    '                IO.Directory.CreateDirectory(outdir & vid.ToString)
    '            End If

    '        Catch ex As Exception
    '            l("fehler1: " & ex.ToString)
    '        End Try
    '        Try
    '            convertPDF(inputfile.Replace(Chr(34), ""), outfile)
    '        Catch ex As Exception
    '            l("fehler2: " & ex.ToString)
    '        End Try
    '    Next
    'End Sub
    Private Sub PDFumwandeln()
        Dim DT As DataTable

        l("PDFumwandeln ")
        'Dim logfile As String = "\\file-paradigma\paradigma\test\thumbnails\PDFlog" & Format(Now, "ddhhmmss") & ".txt"

        Dim dateifehlt As String = "\\file-paradigma\paradigma\test\thumbnails\dateifehlt_pdf" & Environment.UserName & ".txt"
        swfehlt = New IO.StreamWriter(dateifehlt)
        swfehlt.AutoFlush = True
        swfehlt.WriteLine(Now)



        Dim logfile As String = "\\file-paradigma\paradigma\test\thumbnails\PDFlog_" & Environment.UserName & ".txt"
        sw = New IO.StreamWriter(logfile)
        sw.AutoFlush = True
        sw.WriteLine(Now)
        'outdir = "L:\cache\paradigma\thumbnails\"
        outdir = "\\file-paradigma\paradigma\test\thumbnails\"

        checkoutRoot = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\paradigma\muell\"
        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        '  vid = modPrep.getVid()

        sw.WriteLine(vid)
        If vid = "fehler" Then End
        DT = PDFdokumentDatenHolen()
        'teil1 = pdf -----------------------------------------------

        l("vor pdfverarbeiten")
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim eid As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim initial As String

        l("PDFumwandeln 2 ")
        l("PDFumwandeln 2 ")
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenHolen(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid)
                l(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")
                sw.WriteLine(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")


                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If
                outfile = modPrep.GetOutfileName(CInt(vid), outdir, CInt(dokumentid), ".jpg")
                Dim fo As New IO.FileInfo(outfile.Replace(Chr(34), ""))
                TextBox3.Text = igesamt & " von " & DT.Rows.Count
                Application.DoEvents()
                If fo.Exists Then
                    ' l("exists")
                    Continue For
                End If


                TextBox1.Text = ic.ToString & " / " & dateinameext & Environment.NewLine & " " &
                inputfile & Environment.NewLine &
                vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")"
                Application.DoEvents()


                If Not IO.Directory.Exists(outdir & vid.ToString) Then
                    IO.Directory.CreateDirectory(outdir & vid.ToString)
                End If
            Catch ex As Exception
                l("fehler1: " & ex.ToString)
            End Try
            Try
                sw.WriteLine(inputfile)
                Application.DoEvents()
                If dokumentid = "60091" Then
                    'Continue For
                    Debug.Print("")
                End If
                'If dokumentid = "77828" Then Continue For
                'If dokumentid = "80043" Then Continue For
                'If dokumentid = "80071" Then Continue For
                Dim fi As New IO.FileInfo(inputfile.Replace(Chr(34), ""))
                If Not fi.Exists Then
                    swfehlt.WriteLine(vid & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                    Continue For
                Else

                End If

                If convertPDF2(inputfile, outfile) Then
                    l("erfolg")
                    ic += 1
                    TextBox1.Text = ic.ToString & " / " & dateinameext & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")"
                    Application.DoEvents()
                Else
                    l("erfolglos " & ic.ToString & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")")
                    TextBox2.Text = ic.ToString & " / " & dateinameext & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                        TextBox2.Text
                    Application.DoEvents()
                End If
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       inputfile & Environment.NewLine &
                       vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()

        Next
        If batchmode = True Then

        End If
        swfehlt.Close()
        l("logfile  " & logfile)
        Process.Start(logfile)
    End Sub



    Private Function convertPDF2(inputfile As String, outputfile As String) As Boolean
        ' Acrobat objects
        Dim pdfDoc As Acrobat.CAcroPDDoc
        Dim pdfPage As Acrobat.CAcroPDPage
        Dim pdfRect As Acrobat.CAcroRect
        Dim pdfRectTemp As Object



        Dim pdfInputPath As String
        Dim pngOutputPath As String
        Dim pageCount As Integer
        Dim ret As Integer

        Try
            ' Could skip if thumbnail already exists in output path
            ''Dim fi As New FileInfo(inputFile)
            ''If Not fi.Exists() Then
            ''
            ''End If
            pdfDoc = CreateObject("AcroExch.PDDoc")

            ' Open the document
            ret = pdfDoc.Open(inputfile)

            If ret = False Then
                Return False
            End If

            ' Get the number of pages
            pageCount = pdfDoc.GetNumPages()

            ' Get the first page
            pdfPage = pdfDoc.AcquirePage(0)

            ' Get the size of the page
            ' This is really strange bug/documentation problem
            ' The PDFRect you get back from GetSize has properties
            ' x and y, but the PDFRect you have to supply CopyToClipboard
            ' has left, right, top, bottom
            pdfRectTemp = pdfPage.GetSize

            ' Create PDFRect to hold dimensions of the page
            pdfRect = CreateObject("AcroExch.Rect")

            pdfRect.Left = 0
            pdfRect.right = pdfRectTemp.x
            pdfRect.Top = 0
            pdfRect.bottom = pdfRectTemp.y

            ' Render to clipboard, scaled by 100 percent (ie. original size)
            ' Even though we want a smaller image, better for us to scale in .NET
            ' than Acrobat as it would greek out small text
            ' see http://www.adobe.com/support/techdocs/1dd72.htm

            Call pdfPage.CopyToClipboard(pdfRect, 0, 0, 100)

            Dim clipboardData As IDataObject = Clipboard.GetDataObject()

            If (clipboardData.GetDataPresent(DataFormats.Bitmap)) Then

                Dim pdfBitmap As Bitmap = clipboardData.GetData(DataFormats.Bitmap)

                ' Size of generated thumbnail in pixels
                Dim thumbnailWidth As Integer = 600
                Dim thumbnailHeight As Integer = 900

                Dim templateFile As String

                ' Switch between portrait and landscape
                If (pdfRectTemp.x < pdfRectTemp.y) Then
                    thumbnailWidth = 600
                    thumbnailHeight = 900
                    thumbnailWidth = 3700
                    thumbnailHeight = 5800
                    thumbnailWidth = pdfRectTemp.x * 1
                    thumbnailHeight = pdfRectTemp.y * 1
                Else
                    thumbnailWidth = 900
                    thumbnailHeight = 600
                    thumbnailWidth = 5800
                    thumbnailHeight = 3700
                    thumbnailWidth = pdfRectTemp.y * 1
                    thumbnailHeight = pdfRectTemp.x * 1
                End If


                ' Load the template graphic
                'Dim templateBitmap As Bitmap = New Bitmap(templateFile)
                'Dim templateImage As Image = Image.FromFile(templateFile)
                Dim myImageCodecInfo As ImageCodecInfo
                Dim myEncoder As Imaging.Encoder
                Dim myEncoderParameter As Imaging.EncoderParameter
                Dim myEncoderParameters As Imaging.EncoderParameters

                ' Render to small image using the bitmap class
                Dim pdfImage As Image = pdfBitmap.GetThumbnailImage(thumbnailWidth,
                                                                    thumbnailHeight,
                                                                    Nothing, Nothing)

                ' Create new blank bitmap (+ 7 for template border)
                Dim thumbnailBitmap As Bitmap = New Bitmap(thumbnailWidth + 7,
                                                           thumbnailHeight + 7,
                                                           Imaging.PixelFormat.Format16bppRgb565
                                                           )
                'Format32bppArgb,Format24bppRgb,Format16bppRgb565
                ' To overlayout the template with the image, we need to set the transparency
                ' http://www.sellsbrothers.com/writing/default.aspx?content=dotnetimagerecoloring.htm
                '  templateBitmap.MakeTransparent()

                Dim thumbnailGraphics As Graphics = Graphics.FromImage(thumbnailBitmap)

                ' Draw rendered pdf image to new blank bitmap
                thumbnailGraphics.DrawImage(pdfImage, 2, 2, thumbnailWidth, thumbnailHeight)

                ' Draw template outline over the bitmap (pdf with show through the transparent area)
                '  thumbnailGraphics.DrawImage(templateImage, 0, 0)

                myImageCodecInfo = GetEncoderInfo(ImageFormat.Jpeg)

                ' Create an Encoder object based on the GUID
                ' for the Quality parameter category.
                myEncoder = Encoder.Quality

                ' Create an EncoderParameters object.
                ' An EncoderParameters object has an array of EncoderParameter
                ' objects. In this case, there is only one
                ' EncoderParameter object in the array.
                myEncoderParameters = New EncoderParameters(1)

                '            Dim ep As Imaging.EncoderParameters = New Imaging.EncoderParameters
                'ep.Param(0) = New System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, komprimierung)

                ' Save the bitmap as a JPEG file with quality level 25.
                'myEncoderParameter = New EncoderParameter(myEncoder, CType(5L, Int32))
                'myEncoderParameters.Param(0) = myEncoderParameter
                myEncoderParameters.Param(0) = New EncoderParameter(myEncoder, 15)
                '    myBitmap.Save("Shapes025.jpg", myImageCodecInfo, myEncoderParameters)


                ' Save as .png file
                thumbnailBitmap.Save(outputfile, myImageCodecInfo, myEncoderParameters)

                Console.WriteLine("Generated thumbnail... {0}", outputfile)
                thumbnailGraphics.Dispose()


            End If
            Return True


        Catch ex As Exception
            Console.WriteLine("fehler in convert2pdf: " & ex.ToString)

            Return False

        Finally


            pdfDoc.Close()
            If pdfPage IsNot Nothing Then Marshal.ReleaseComObject(pdfPage)
            If pdfRect IsNot Nothing Then Marshal.ReleaseComObject(pdfRect)
            If pdfDoc IsNot Nothing Then Marshal.ReleaseComObject(pdfDoc)
        End Try
    End Function
    Function GetEncoderInfo(ByVal format As ImageFormat) As ImageCodecInfo
        Dim j As Integer
        Dim encoders() As ImageCodecInfo
        encoders = ImageCodecInfo.GetImageEncoders()

        j = 0
        While j < encoders.Length
            If encoders(j).FormatID = format.Guid Then
                Return encoders(j)
            End If
            j += 1
        End While
        Return Nothing

    End Function 'GetEncoderInfo

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Friend Sub DOCXumwandeln(vid As Integer, isDebugmode As Boolean)
        '    If Not IsNumeric(vid) Then Exit Sub
        Dim inputfile, outfileJPG, parameter As String
        Dim innDir, outDir, checkoutfile, pdffile As String
        'Dim checkoutRoot As String = "C:\muell\" 
        parameter = " /1 1"
        Dim dateifehlt As String = "\\file-paradigma\paradigma\test\thumbnails\dateifehlt_doc2" & Environment.UserName & ".txt"
        swfehlt = New IO.StreamWriter(dateifehlt)
        swfehlt.AutoFlush = True
        swfehlt.WriteLine(Now)


        Dim checkoutRoot As String = "C:\muell\"
        innDir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv" '"\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        outDir = "\\file-paradigma\paradigma\test\thumbnails\"
        'outDir = "c:\muell\"
        l("DOCXumwandeln         ")
        l(innDir)
        l(outDir)
        If isDebugmode Then
            '  outDir = "l:\cache\paradigma\thumbnails\"
        End If


        '  DBfestlegen()
        l("in getallin")
        Dim Sql As String
        Dim oben, unten As String
        ' oben = "200000" : unten = "142568" ' muss am schluss nachgeholt werden
        oben = "2000000" : unten = "0"
        '  oben = "139340" : unten = "0"
        'oben = "134133" : unten = "0"
        'oben = "40777" : unten = "0"

        Sql = "SELECT * FROM dokumente where   dokumentid > " & unten & "  and dokumentid < " & oben & "  " &
              "and ( typ='docx' or  typ='doc'  or  typ='rtf' )  " &
              "order by dokumentid desc"
        'Sql = "SELECT * FROM dokumente where   vid=9609 " &
        '      "and ( typ='docx' or  typ='doc' or  typ='rtf')  " &
        '      "order by dokumentid desc"
        l(Sql)

        immerUeberschreiben = True
        dt = getDT(Sql)

        l("nach getDT")
        Dim relativpfad As String = "", dateinameext As String = "", typ As String, logfile As String
        Dim newsavemode As Boolean

        Dim dokumentid As Integer = 0
        'l("nach 1: " & outDir & vid.ToString)
        IO.Directory.CreateDirectory(outDir & vid.ToString)
        'l("nach 2")
        logfile = outDir & "\tnmaker_" & Environment.UserName & "wordgen.txt"
        'l("nach 3")
        Dim ic As Integer = 0
        Dim ierfolg As Integer = 0
        Dim soll As Integer = 0
        Dim dbdatum As Date
        Dim initial As String
        typ = "1"
        Using sw As New IO.StreamWriter(logfile)
            sw.AutoFlush = True
            Application.DoEvents()
            For Each drr As DataRow In dt.Rows
                TextBox3.Text = ic & " von " & dt.Rows.Count


                ic += 1
                TextBox2.Text = " " & ierfolg & "   konvertierungen von word nach jpg erfolgreich von " & soll & Environment.NewLine
                Application.DoEvents()
                datenholden(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, initial)


                If vid < 1000 Then
                    Continue For
                End If
                If dokumentid = 174865 Then
                    Debug.Print("")

                End If
                Console.Write(vid & "/" & dokumentid & "----")
                If newsavemode Then
                    inputfile = GetInputfileWordFullPath(innDir, relativpfad, dokumentid)
                Else
                    inputfile = GetInputfile1WordFullPath(innDir, relativpfad, dateinameext)
                End If

                Dim fi As New IO.FileInfo(inputfile.Replace(Chr(34), ""))
                If Not fi.Exists Then
                    'swfehlt.WriteLine(vid & "," & dokumentid & ", " & dbdatum & "," & initial & ", " & inputfile & "")
                    swfehlt.WriteLine(vid & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                    Continue For
                Else
                    TextBox1.Text = dateinameext & " ist dran  " & Environment.NewLine
                    Application.DoEvents()
                End If

                'inputFileReadonlyEntfernen(inputfile)
                checkoutfile = getCheckoutfileWord(inputfile, checkoutRoot, dokumentid, vid, Format(Now, "yyMMddmm_"))

                'checkoutfile = checkoutfile.Replace("\muell\", "\muell\AA_")


                pdffile = checkoutfile & ".pdf"
                checkoutfile = checkoutfile & "." & typ

                outfileJPG = GetOutfileWORD(vid, outDir, dokumentid)
                Dim fo As New IO.FileInfo(outfileJPG)
                fi = New IO.FileInfo(inputfile)
                If fo.Exists Then
                    If fo.LastWriteTime > fi.LastWriteTime Then
                        'keine änderung
                        Continue For
                    Else

                    End If

                End If
                soll += 1
                IO.Directory.CreateDirectory(checkoutRoot & vid.ToString)

                If Not auscheckenword(inputfile, checkoutfile, sw, CType(vid, String), CType(dokumentid, String)) Then
                    l("-- " & dateinameext)
                    'sw.WriteLine("-- " & dateinameext)
                    Continue For
                Else
                    'l(" ")
                    'sw.WriteLine("-- " & dateinameext)
                    inputFileReadonlyEntfernen(checkoutfile)
                End If
                TextBox1.Text = checkoutfile & " checkout erfolgreich   " & Environment.NewLine
                Application.DoEvents()
                If clsWordTest.konvOneDoc2pdf(checkoutfile, pdffile) Then
                    TextBox1.Text = TextBox1.Text & " " & pdffile & " pdf erfolgreich" & Environment.NewLine
                    Application.DoEvents()
                    If convertPDF2(pdffile, outfileJPG) Then
                        l("erfolg")
                        ic += 1
                        TextBox1.Text = TextBox1.Text & " / " & dateinameext & " " & "jpg erfolgreich: " & ic.ToString & Environment.NewLine & " " &
                            outfileJPG & Environment.NewLine &
                            vid & "/" & dokumentid & " " & ic & "(" & dt.Rows.Count.ToString & ")"
                        Application.DoEvents()
                        ierfolg += 1
                    Else
                        l("pdf2jpg erfolglos " & ic.ToString & Environment.NewLine & " " &
                            outfileJPG & Environment.NewLine &
                            vid & "/" & dokumentid & " " & ic & "(" & dt.Rows.Count.ToString & ")")
                        TextBox2.Text = TextBox1.Text & " " & "jpg nicht erfolgreich: " & ic.ToString & Environment.NewLine & " " &
                            inputfile & Environment.NewLine &
                            vid & "/" & dokumentid & " " & ic & "(" & dt.Rows.Count.ToString & ")" & Environment.NewLine &
                            TextBox2.Text
                        sw.WriteLine("fehlerin convertPDF2: " & vid & "/" & dokumentid & " " & outfileJPG & " " & inputfile)
                        Application.DoEvents()
                    End If
                Else
                    l("word2pdf erfolglos " & ic.ToString & Environment.NewLine & " " &
                        outfileJPG & Environment.NewLine &
                        vid & "/" & dokumentid & " " & ic & "(" & dt.Rows.Count.ToString & ")")
                    sw.WriteLine("fehlerin word2pdf: " & vid & "/" & dokumentid & " " & outfileJPG & " " & inputfile)
                    Application.DoEvents()
                End If




                GC.Collect()
                GC.WaitForFullGCComplete()
                IO.Directory.CreateDirectory(outDir & vid.ToString)

                deleteCheckoutfileWord(checkoutfile)
                deleteCheckoutfileWord(pdffile)

                Threading.Thread.Sleep(1000)
                Try
                    IO.Directory.Delete(checkoutRoot & "\" & vid)
                Catch ex As Exception

                End Try

            Next
        End Using
        swfehlt.Close()
        Try
            Process.Start(logfile)
        Catch ex As Exception
            l("fehler : " & ex.ToString)
        End Try
    End Sub



    Public Function inputFileReadonlyEntfernen(inputfile As String) As Boolean
        Dim retval As Boolean = False
        Try
            Dim fi As New IO.FileInfo(inputfile)
            If CBool(fi.Attributes And IO.FileAttributes.ReadOnly) Then
                ' Datei ist schreibgeschützt
                ' Jetzt Schreibschutz-Attribut entfernen
                '  fi.Attributes = fi.Attributes Xor IO.FileAttributes.ReadOnly
                fi.IsReadOnly = False
                retval = True
            End If
            fi = Nothing
            Return retval
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Sub inputFileReadonlysetzen(inputfile As String)
        Try
            Dim fi As New IO.FileInfo(inputfile)
            If CBool(fi.Attributes And Not IO.FileAttributes.ReadOnly) Then
                ' Datei ist nicht schreibgeschützt
                ' Jetzt Schreibschutz-Attribut setzen
                fi.IsReadOnly = True
                ' fi.Attributes = fi.Attributes Or IO.FileAttributes.ReadOnly
                fi = Nothing
            End If
        Catch ex As Exception
            nachricht("inputFileReadonlysetzen " & inputfile & " / " & ex.ToString)
        End Try
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Xls2XlsxKonv(9609, True)
    End Sub
    Friend Sub Xls2XlsxKonv(vid As Integer, isDebugmode As Boolean)
        '    If Not IsNumeric(vid) Then Exit Sub
        Dim inputfile, outfile, parameter As String
        Dim innDir, outDir, checkoutfile, pdffile As String
        'Dim checkoutRoot As String = "C:\muell\"

        Dim dateifehlt As String = "\\file-paradigma\paradigma\test\thumbnails\dateifehltDoc_" & Environment.UserName & ".txt"
        swfehlt = New IO.StreamWriter(dateifehlt)
        swfehlt.AutoFlush = True
        swfehlt.WriteLine(Now)

        parameter = " /1 1"
        Dim checkoutRoot As String = "C:\muell\"
        innDir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv" '"\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        outDir = "\\file-paradigma\paradigma\test\thumbnails\"
        'outDir = "c:\muell\"
        If isDebugmode Then
            '  outDir = "l:\cache\paradigma\thumbnails\"
        End If


        '  DBfestlegen()
        l("in getallin")
        Dim Sql As String
        Dim oben, unten As String
        ' oben = "200000" : unten = "142568" ' muss am schluss nachgeholt werden
        oben = "2000000" : unten = "0"
        '  oben = "139340" : unten = "0"
        'oben = "134133" : unten = "0"
        'oben = "40777" : unten = "0"

        Sql = "SELECT * FROM dokumente where   dokumentid > " & unten & "  and dokumentid < " & oben & "  " &
              "and (  lower(dateinameext) like '%.xls' )  " &
              "order by dokumentid desc"
        'Sql = "SELECT * FROM dokumente where   vid=9609 " &
        '      "and ( typ='docx' or  typ='doc' or  typ='rtf')  " &
        '      "order by dokumentid desc"


        immerUeberschreiben = True
        dt = getDT(Sql)

        l("nach getDT")
        Dim relativpfad As String = "", dateinameext As String = "", typ As String, logfile As String
        Dim newsavemode As Boolean

        Dim dokumentid As Integer = 0
        'l("nach 1: " & outDir & vid.ToString)
        IO.Directory.CreateDirectory(outDir & vid.ToString)
        'l("nach 2")
        logfile = outDir & "\tnmaker_" & "wordgen.txt"
        'l("nach 3")
        Dim ic As Integer = 0
        Dim ierfolg As Integer = 0
        Dim soll As Integer = 0
        Dim dbdatum As Date
        Dim initial As String
        typ = "1"
        Using sw As New IO.StreamWriter(logfile)
            sw.AutoFlush = True
            Application.DoEvents()

            For Each drr As DataRow In dt.Rows
                TextBox3.Text = ic & " von " & dt.Rows.Count

                ic += 1
                TextBox2.Text = " " & ierfolg & "   konvertierungen von word nach jpg erfolgreich von " & soll & Environment.NewLine
                Application.DoEvents()
                datenholden(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, initial)
                If dokumentid = 174865 Then
                    Debug.Print("")

                End If
                Console.Write(vid & "/" & dokumentid & "----")
                'sw.WriteLine(vid & "/" & dokumentid & "----")
                If newsavemode Then
                    inputfile = GetInputfileWordFullPath(innDir, relativpfad, dokumentid)
                Else
                    inputfile = GetInputfile1WordFullPath(innDir, relativpfad, dateinameext)
                End If

                'inputFileReadonlyEntfernen(inputfile)
                checkoutfile = getCheckoutfileWord(inputfile, checkoutRoot, dokumentid, vid, Format(Now, "yyMMddmm_"))

                'checkoutfile = checkoutfile.Replace("\muell\", "\muell\AA_")


                pdffile = checkoutfile & ".pdf"
                checkoutfile = checkoutfile & "." & typ

                outfile = GetOutfileEXCEL(vid, checkoutRoot, dokumentid)
                Dim fo As New IO.FileInfo(outfile)
                Dim fi As New IO.FileInfo(inputfile)
                If fo.Exists Then
                    If fo.LastWriteTime > fi.LastWriteTime Then
                        'keine änderung
                        Continue For
                    Else

                    End If

                End If
                soll += 1
                IO.Directory.CreateDirectory(checkoutRoot & vid.ToString)


                fi = New IO.FileInfo(inputfile.Replace(Chr(34), ""))
                If Not fi.Exists Then
                    'swfehlt.WriteLine(vid & "," & dokumentid & ", " & dbdatum & "," & initial & ", " & inputfile & "")
                    swfehlt.WriteLine(vid & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                    Continue For
                Else

                End If

                If Not auscheckenword(inputfile, checkoutfile, sw, CType(vid, String), CType(dokumentid, String)) Then
                    l("-- " & dateinameext)
                    'sw.WriteLine("-- " & dateinameext)
                    Continue For
                Else
                    'l(" ")
                    'sw.WriteLine("-- " & dateinameext)
                    inputFileReadonlyEntfernen(checkoutfile)
                End If
                TextBox1.Text = checkoutfile & " checkout erfolgreich   " & Environment.NewLine
                Application.DoEvents()
                If clsExcel.konvOne(checkoutfile, outfile) Then
                    TextBox1.Text = TextBox1.Text & " " & pdffile & " xls erfolgreich" & Environment.NewLine
                    Application.DoEvents()
                    'altearchivdatei umbenennen
                    'neuedatei im archiv speichern
                    'in db: typ anpassen
                    'in db: typ in dateinameext anpassen
                    If newsavemode Then
                        If altearchivdatei_umbenennen(newsavemode, inputfile) Then
                            If neuedatei_im_archiv_speichern(inputfile, outfile) Then
                                If db_eintragExcelAendern(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr) Then
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End Using
        Try
            Process.Start(logfile)
        Catch ex As Exception
            l("fehler : " & ex.ToString)
        End Try
    End Sub



    Private Function neuedatei_im_archiv_speichern(inputfile As String, outfile As String) As Boolean
        Try
            FileSystem.FileCopy(outfile, inputfile)
            Return True
        Catch ex As Exception
            l("fehelr in neuedatei_im_archiv_speichern " & ex.ToString)
            Return False
        End Try

    End Function

    Private Function altearchivdatei_umbenennen(newsavemode As Boolean, inputfile As String) As Boolean
        Dim neuername As String
        Try
            neuername = inputfile & "_xls"
            FileSystem.Rename(inputfile, neuername)
            Return True
        Catch ex As Exception
            l("fehelr in altearchivdatei_umbenennen" & ex.ToString)
            Return False
        End Try
    End Function
    Private Function db_eintragExcelAendern(vid As Integer, relativpfad As String, dateinameext As String, typ As String, newsavemode As Boolean, dokumentid As Integer, drr As DataRow) As Boolean
        'modOracle.setExcelAttribute2(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr)
    End Function
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'doc2docx
        Dim quellverz, zielverzeichnis As String
        quellverz = "O:\UMWELT\B\Vordruck_paradigma\"
        zielverzeichnis = "O:\UMWELT-PARADIGMA\Vordruck_paradigmaNEU\"

        quellverz = "O:\UMWELT\B\Vordrucke\"
        zielverzeichnis = "O:\UMWELT-PARADIGMA\Vordrucke\"


        doc2docxKonv(quellverz.ToLower, zielverzeichnis.ToLower)
    End Sub

    Private Sub doc2docxKonv(quellverz As String, zielverzeichnis As String)
        dirSearch(quellverz, zielverzeichnis)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        bplantn()
    End Sub

    Private Sub bplantn()
        Dim DT As DataTable


        Dim logfile As String = "\\file-paradigma\paradigma\test\thumbnails\PDFlog" & Format(Now, "ddhhmmss") & ".txt"
        'logfile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\paradigma\muell\thumbnailer.log"
        sw = New IO.StreamWriter(logfile)
        sw.AutoFlush = True
        sw.WriteLine(Now)
        'outdir = "L:\cache\paradigma\thumbnails\"
        outdir = "\\file-paradigma\paradigma\test\thumbnails\"

        checkoutRoot = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\paradigma\muell\"
        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        '  vid = modPrep.getVid()

        sw.WriteLine(vid)
        If vid = "fehler" Then End
        DT = PDFdokumentDatenHolen()
        'teil1 = pdf -----------------------------------------------

        l("vor pdfverarbeiten")
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim initial As String
        Dim eid As Integer = 0
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                If vid = 8930 Then
                    Debug.Print("")
                End If
                igesamt += 1
                DbMetaDatenHolen(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid)
                l(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")

                sw.WriteLine(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")
                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If
                outfile = modPrep.GetOutfileName(CInt(vid), outdir, CInt(dokumentid), ".jpg")
                Dim fo As New IO.FileInfo(outfile.Replace(Chr(34), ""))
                TextBox3.Text = igesamt & " von " & DT.Rows.Count
                Application.DoEvents()
                If fo.Exists Then
                    l("exists")
                    Continue For
                End If
                If Not IO.Directory.Exists(outdir & vid.ToString) Then
                    IO.Directory.CreateDirectory(outdir & vid.ToString)
                End If
            Catch ex As Exception
                l("fehler1: " & ex.ToString)
            End Try
            Try
                sw.WriteLine(inputfile)
                Application.DoEvents()
                If dokumentid = "60091" Then
                    'Continue For
                    Debug.Print("")
                End If
                'If dokumentid = "77828" Then Continue For
                'If dokumentid = "80043" Then Continue For
                'If dokumentid = "80071" Then Continue For
                If convertPDF2(inputfile, outfile) Then
                    l("erfolg")
                    ic += 1
                    TextBox1.Text = ic.ToString & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")"
                    Application.DoEvents()
                Else
                    l("erfolglos " & ic.ToString & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")")
                    TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                        TextBox2.Text
                    Application.DoEvents()
                End If
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       inputfile & Environment.NewLine &
                       vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()

        Next
        If batchmode = True Then
            End
        End If
        Process.Start(logfile)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            PDFumwandeln()
            DOCXumwandeln(2113, False)
        Catch ex As Exception
            Debug.Print("")
        End Try
    End Sub

    Public Sub dirSearch(strDir As String, zieldir As String)
        Dim zielname As String = ""
        Try
            For Each strDirectory As String In IO.Directory.GetDirectories(strDir)
                ' mach etwas....
                For Each strFile As String In IO.Directory.GetFiles(strDirectory, "*.doc*")
                    Debug.Print(strFile)
                    'dateiImZielVerzLoeschen
                    If dateiImZielVerzLoeschen(strFile, strDir, zieldir, zielname) Then
                        If zielname.EndsWith(".doc") Then
                            zielname = zielname.Replace(".doc", ".docx")
                        End If

                        If worddateiAlsDocxSpeichern(strFile, strDir, zieldir, zielname) Then
                        Else
                            Debug.Print("problem1 mit " & strFile)
                        End If
                    Else
                        Debug.Print("problem2 mit " & strFile)
                    End If
                Next
                dirSearch(strDirectory, zieldir)
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub
    Public Sub dirSearchVorlagen(strDir As String)
        Dim zielname As String = ""
        Dim cntBookmark As Integer
        Dim fi As IO.FileInfo
        Try
            For Each strDirectory As String In IO.Directory.GetDirectories(strDir)
                'If Not strDirectory.Contains("allgemein") Then Continue For
                For Each strFile As String In IO.Directory.GetFiles(strDirectory, "*.docx")
                    TextBox1.Text = ""
                    TextBox2.Text = TextBox2.Text & " " & strFile & Environment.NewLine
                    count += 1
                    Application.DoEvents()
                    If strFile.Contains("~$") Then Continue For
                    Debug.Print(strFile)
                    fi = New IO.FileInfo(strFile)
                    If fi.LastWriteTime.Day = Now.Day And fi.LastWriteTime.Month = Now.Month And fi.LastWriteTime.Year = Now.Year Then
                        Continue For
                    End If
                    cntBookmark = TM_ernteBookmarksAusVorlagenDoc(strFile)
                    TextBox3.Text = count.ToString & ", " & cntBookmark
                    '  If cntBookmark > 0 Then loescheBookmarktsAusDocX(strFile)
                Next
                dirSearchVorlagen(strDirectory)
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Function loescheBookmarktsAusDocX(vorlageFullname As String) As Integer
        nachricht("cropBookmarksList ---------------------- ")
        Dim obj As Object
        Try
            Dim int As Integer
            nachricht("cropBookmarksList vor öffnen ")
            obj = vorlageFullname
            docVorlagen = wordVorlagen.Documents.OpenNoRepairDialog(obj)
            docVorlagen.Activate()
            nachricht("cropBookmarksList nach activate - vor schleife")
            nachricht("cropBookmarksList anzahl textmarken: " & docVorlagen.Bookmarks.Count)
            TextBox1.Text = TextBox1.Text & " " & "loeschen " & Environment.NewLine
            'ReDim bookmarkArray(.Bookmarks.Count - 1)
            For int = 1 To docVorlagen.Bookmarks.Count
                'bookmarkArray(int - 1) = .Bookmarks(int).Name
                nachricht("Textmarke gefunden: " & docVorlagen.Bookmarks(int).Name)
                TextBox1.Text = TextBox1.Text & " " & "löschen " & docVorlagen.Bookmarks(int).Name
                DeleteBookmark(docVorlagen.Bookmarks(int).Name, "#" & docVorlagen.Bookmarks(int).Name & "#", docVorlagen)
            Next
            Return docVorlagen.Bookmarks.Count
        Catch ex As Exception
            nachricht("cropBookmarksList: " & ex.ToString)
            If docVorlagen IsNot Nothing Then
                docVorlagen.Close()
                docVorlagen = Nothing
            End If
            'wordVorla
            'wordVorlagen.Application.Quit()
            'wordVorlagen = Nothing
            Return -1
        Finally
            If docVorlagen IsNot Nothing Then
                docVorlagen.Close()
                docVorlagen = Nothing
            End If
            'wordVorlagen.Application.Quit()
            'wordVorlagen = Nothing
            'ReleaseComObj(word)
            'ReleaseComObj(doc)
            ' Die Speichert freigeben
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            'GC.WaitForPendingFinalizers() 
        End Try
    End Function

    Public Function TM_ernteBookmarksAusVorlagenDoc(vorlageFullname As String) As Integer 'liefert leere bookmarks
        nachricht("cropBookmarksList ---------------------- ")
        Dim obj As Object
        Try
            Dim int As Integer
            nachricht("cropBookmarksList vor öffnen ")
            obj = vorlageFullname
            docVorlagen = wordVorlagen.Documents.OpenNoRepairDialog(obj)
            docVorlagen.Activate()
            TextBox1.Text = TextBox1.Text & " " & "geöffnet" & docVorlagen.Bookmarks.Count & Environment.NewLine
            nachricht("cropBookmarksList nach activate - vor schleife " & docVorlagen.Bookmarks.Count)
            nachricht("cropBookmarksList anzahl textmarken: " & docVorlagen.Bookmarks.Count)

            For int = 1 To docVorlagen.Bookmarks.Count
                'bookmarkArray(int - 1) = .Bookmarks(int).Name
                nachricht("Textmarke gefunden: " & docVorlagen.Bookmarks(int).Name)
                TextBox1.Text = TextBox1.Text & " " & "change" & docVorlagen.Bookmarks(int).Name
                changeAndDeleteBookmark(docVorlagen.Bookmarks(int).Name, "#" & docVorlagen.Bookmarks(int).Name & "#", docVorlagen)
                TextBox1.Text = TextBox1.Text & " " & "change fertig" & Environment.NewLine
            Next
            Return docVorlagen.Bookmarks.Count
        Catch ex As Exception
            nachricht("cropBookmarksList: " & ex.ToString)
            If docVorlagen IsNot Nothing Then
                docVorlagen.Close()
                docVorlagen = Nothing
            End If
            'wordVorla
            'wordVorlagen.Application.Quit()
            'wordVorlagen = Nothing
            Return -1
        Finally
            If docVorlagen IsNot Nothing Then
                docVorlagen.Close()
                docVorlagen = Nothing
            End If
            'wordVorlagen.Application.Quit()
            'wordVorlagen = Nothing
            'ReleaseComObj(word)
            'ReleaseComObj(doc)
            ' Die Speichert freigeben
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
            'GC.Collect()
            'GC.WaitForPendingFinalizers() 
        End Try
    End Function
    Private Shared Function changeAndDeleteBookmark(ByVal textmarke As String, ByVal textm_value As String, ByVal doc As Microsoft.Office.Interop.Word.Document) As Integer
        Try
            '   nachricht("In changeBookmark------------------")
            Dim test = textm_value.Trim.Replace("""", "")

            If test = "0" Then
                Return 0
            End If
            If doc.Range.Bookmarks.Exists(textmarke) Then
                doc.Bookmarks().Item(textmarke).Range.Text = textm_value
                doc.Bookmarks().Item(textmarke).Delete()
                Return 1
            Else
                Return 0
            End If
        Catch ex As Exception
            nachricht(String.Format("Fehler in changeBookmark:{0}{1}", vbCrLf, ex))
            nachricht("Fehler bei: " & textmarke & "_" & textm_value)
            Return -1
        End Try
    End Function
    Private Shared Function DeleteBookmark(ByVal textmarke As String, ByVal textm_value As String, ByVal doc As Microsoft.Office.Interop.Word.Document) As Integer
        Try
            Dim test = textm_value.Trim.Replace("""", "")
            If test = "0" Then
                Return 0
            End If
            If doc.Range.Bookmarks.Exists(textmarke) Then
                doc.Bookmarks().Item(textmarke).Delete()
                Return 1
            Else
                '  nachricht("Warnung:changeBookmark: Textmarke nicht vorhanden: " & textmarke)
                Return 0
            End If
        Catch ex As Exception
            nachricht(String.Format("Fehler in changeBookmark:{0}{1}", vbCrLf, ex))
            nachricht("Fehler bei: " & textmarke & "_" & textm_value)
            Return -1
        End Try
    End Function
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim quellverz As String
        quellverz = "O:\UMWELT\B\Vordruck_paradigma_hashtag"
        quellverz = "C:\3\Vordruck_paradigma_hashtag"

        dirSearchVorlagen(quellverz)
        'dirSearch(quellverz.ToLower, zielverzeichnis.ToLower)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        '    If Not IsNumeric(vid) Then Exit Sub
        Dim inputfile, outfile, parameter As String
        Dim innDir, outDir, checkoutfile, pdffile As String
        parameter = " /1 1"
        Dim checkoutRoot As String = "C:\muell\"
        innDir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv" '"\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        outDir = "\\file-paradigma\paradigma\test\thumbnailsOOOO\"
        '  DBfestlegen()
        l("in getallin")
        Dim Sql As String
        Dim oben, unten As String
        Dim relativpfad As String = "", dateinameext As String = "", typ As String, logfile As String
        Dim newsavemode As Boolean

        Dim dokumentid As Integer = 0
        oben = "2000000" : unten = "0"
        Sql = "SELECT *   FROM [Paradigma].[dbo].[DOKUMENTE] where lower(typ)='docx' or lower(typ)='doc' or lower(typ)='pdf' "
        immerUeberschreiben = True
        dt = getDT(Sql)

        l("nach getDT")

        'l("nach 1: " & outDir & vid.ToString)
        IO.Directory.CreateDirectory(outDir)
        'l("nach 2")
        logfile = outDir & "\" & "dok.txt"
        'l("nach 3")
        Dim ic As Integer = 0
        Dim ierfolg As Integer = 0
        Dim soll As Integer = 0
        typ = "1"
        Dim sizeSumme As Long
        Dim dbdatum As Date
        Dim initial As Long
        Dim fi As IO.FileInfo
        Using sw As New IO.StreamWriter(logfile)
            sw.AutoFlush = True
            Application.DoEvents()

            For Each drr As DataRow In dt.Rows
                TextBox3.Text = ic & " von " & dt.Rows.Count

                ic += 1
                TextBox2.Text = " " & ierfolg & "   konvertierungen von word nach jpg erfolgreich von " & soll & Environment.NewLine
                Application.DoEvents()
                datenholden(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, initial)
                If dokumentid = 174865 Then
                    Debug.Print("")

                End If
                Console.Write(vid & "/" & dokumentid & "----")
                'sw.WriteLine(vid & "/" & dokumentid & "----")
                If newsavemode Then
                    inputfile = GetInputfileWordFullPath(innDir, relativpfad, dokumentid)
                Else
                    inputfile = GetInputfile1WordFullPath(innDir, relativpfad, dateinameext)
                End If
                fi = New IO.FileInfo(inputfile)
                If Not fi.Exists Then
                    Continue For
                End If
                sizeSumme += fi.Length
                TextBox1.Text = TextBox1.Text & ic & " " & inputfile & " " & fi.Length & Environment.NewLine
                '   sw.WriteLine(TextBox1.Text)
                Application.DoEvents()
                'If ic = 1000 Then Exit For

            Next
        End Using
        MsgBox("anzahl:" & Environment.NewLine &
               dt.Rows.Count & ", sizesumme: " & getFileSize4Length(sizeSumme) & Environment.NewLine &
                ", im schnitt: " & getFileSize4Length(sizeSumme / dt.Rows.Count))
        sw.WriteLine("typ:" & typ & ", anzahl:" & Environment.NewLine &
               dt.Rows.Count & ", sizesumme: " & getFileSize4Length(sizeSumme) & Environment.NewLine &
                ", im schnitt: " & getFileSize4Length(sizeSumme / dt.Rows.Count))
        Try
            Process.Start(logfile)
        Catch ex As Exception
            l("fehler : " & ex.ToString)
        End Try
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim summe As String = ""
        Dim summe2 As String = ""
        Dim kuerzel As String = ""
        Dim inputfile, outfile, parameter As String
        Dim innDir, outDir, checkoutfile, pdffile As String
        Dim bearbeiterDT As DataTable
        parameter = " /1 1"
        Dim checkoutRoot As String = "C:\muell\"
        innDir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv" '"\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        outDir = "\\file-paradigma\paradigma\test\thumbnailsOOOO\"
        '  DBfestlegen()
        l("in getallin")
        Dim Sql As String
        Dim oben, unten As String
        Dim bearbeiter As String = "", dateinameext As String = "", typ As String, logfile As String
        Dim newsavemode As Boolean

        Dim initial_ As String = 0
        oben = "2000000" : unten = "0"
        Sql = "SELECT *  FROM [Paradigma].[dbo].[t05]   "
        immerUeberschreiben = True
        dt = getDT(Sql)


        'Sql = "SELECT bearbeiter  FROM [Paradigma].[dbo].[t41] where bearbeiterid=0 "
        'immerUeberschreiben = True
        'dt = getDT(Sql)

        l("nach getDT")

        'l("nach 1: " & outDir & vid.ToString)
        IO.Directory.CreateDirectory(outDir)
        'l("nach 2")
        logfile = outDir & "\" & "bearbeiterid.txt"
        'l("nach 3")
        Dim ic As Integer = 0
        Dim ierfolg As Integer = 0
        Dim bearbeiterid As Integer = 0
        typ = "1"
        Dim sizeSumme As Long
        Dim fi As IO.FileInfo
        Using sw As New IO.StreamWriter(logfile)
            sw.AutoFlush = True
            Application.DoEvents()

            For Each drr As DataRow In dt.Rows
                TextBox3.Text = ic & " von " & dt.Rows.Count

                ic += 1
                'TextBox2.Text = " " & ierfolg & "   konvertierungen von word nach jpg erfolgreich von " & soll & Environment.NewLine
                Application.DoEvents()
                initial_ = (drr.Item("initial_"))
                'If initial_.ToLower <> "kosh" Then Continue For
                bearbeiterid = CInt(drr.Item("bearbeiterid"))
                kuerzel = CStr(drr.Item("kuerzel1"))

                Sql = "update t41 set bearbeiterid=" & bearbeiterid &
                    " where lower(bearbeiter)='" & initial_.ToLower & "' and bearbeiterid<>" & bearbeiterid & ";" & Environment.NewLine
                summe += Sql
                Sql = "update t41 set bearbeiterid=" & bearbeiterid &
                    " where lower(bearbeiter)='" & kuerzel.ToLower & "' and bearbeiterid<>" & bearbeiterid & ";" & Environment.NewLine
                summe2 += Sql

                Console.Write(vid & "/" & initial_ & "----")
                'sw.WriteLine(vid & "/" & dokumentid & "----")

                'sizeSumme += fi.Length
                'TextBox1.Text = TextBox1.Text & ic & " " & inputfile & " " & fi.Length & Environment.NewLine
                '   sw.WriteLine(TextBox1.Text)
                Application.DoEvents()
                'If ic = 1000 Then Exit For

            Next
            TextBox1.Text = summe
            TextBox2.Text = summe2
            sw.WriteLine(summe)
            sw.WriteLine(summe2)
        End Using
        'MsgBox("anzahl:" & Environment.NewLine &
        '       dt.Rows.Count & ", sizesumme: " & getFileSize4Length(sizeSumme) & Environment.NewLine &
        '        ", im schnitt: " & getFileSize4Length(sizeSumme / dt.Rows.Count))
        'sw.WriteLine("typ:" & typ & ", anzahl:" & Environment.NewLine &
        '       dt.Rows.Count & ", sizesumme: " & getFileSize4Length(sizeSumme) & Environment.NewLine &
        '        ", im schnitt: " & getFileSize4Length(sizeSumme / dt.Rows.Count))
        Try
            Process.Start(logfile)
        Catch ex As Exception
            l("fehler : " & ex.ToString)
        End Try
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dim summe As String = ""
        Dim summe2 As String = ""
        Dim kuerzel As String = ""
        Dim inputfile, outfile, parameter As String
        Dim innDir, outDir, checkoutfile, pdffile As String
        Dim bearbeiterDT As DataTable
        parameter = " /1 1"
        Dim checkoutRoot As String = "C:\muell\"
        innDir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv" '"\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        outDir = "\\file-paradigma\paradigma\test\thumbnailsOOOO\"
        '  DBfestlegen()
        l("in getallin")
        Dim Sql As String
        Dim oben, unten As String
        Dim bearbeiter As String = "", dateinameext As String = "", typ As String, logfile As String
        Dim newsavemode As Boolean

        Dim initial_ As String = 0
        oben = "2000000" : unten = "0"
        Sql = "SELECT *  FROM [Paradigma].[dbo].[t05]   "
        immerUeberschreiben = True
        bearbeiterDT = getDT(Sql)


        Sql = "SELECT vorgangsid,weitereBearb  FROM [Paradigma].[dbo].[t41] where len(weitereBearb)>0 "
        immerUeberschreiben = True
        dt = getDT(Sql)

        l("nach getDT")

        'l("nach 1: " & outDir & vid.ToString)
        IO.Directory.CreateDirectory(outDir)
        'l("nach 2")
        logfile = outDir & "\" & "bearbeiterid.sql"
        'l("nach 3")
        Dim ic As Integer = 0
        Dim ierfolg As Integer = 0
        Dim bearbeiterid As Integer = 0
        typ = "1"
        Dim sizeSumme As Long
        Dim weitere As String
        Dim fi As IO.FileInfo
        Dim b() As String
        Dim bids() As Integer
        Using sw As New IO.StreamWriter(logfile)
            sw.AutoFlush = True
            Application.DoEvents()

            For Each drr As DataRow In dt.Rows
                TextBox3.Text = ic & " von " & dt.Rows.Count

                ic += 1
                'TextBox2.Text = " " & ierfolg & "   konvertierungen von word nach jpg erfolgreich von " & soll & Environment.NewLine
                Application.DoEvents()
                weitere = (drr.Item("weitereBearb")).tolower
                vid = (drr.Item("vorgangsid"))
                b = weitere.Split(New Char() {";"c},
                        StringSplitOptions.RemoveEmptyEntries)
                ReDim bids(b.Length - 1)
                For i = 0 To b.Length - 1
                    bids(i) = getBearbeiterID(b(i), bearbeiterDT)
                    sw.WriteLine(getinsertSql(bids(i), vid))
                    Console.Write(vid & "/" & initial_ & "----")
                Next






                'sw.WriteLine(vid & "/" & dokumentid & "----")

                'sizeSumme += fi.Length
                'TextBox1.Text = TextBox1.Text & ic & " " & inputfile & " " & fi.Length & Environment.NewLine
                '   sw.WriteLine(TextBox1.Text)
                Application.DoEvents()
                'If ic = 1000 Then Exit For

            Next
            'TextBox1.Text = summe
            'sw.WriteLine(summe)
        End Using
        'MsgBox("anzahl:" & Environment.NewLine &
        '       dt.Rows.Count & ", sizesumme: " & getFileSize4Length(sizeSumme) & Environment.NewLine &
        '        ", im schnitt: " & getFileSize4Length(sizeSumme / dt.Rows.Count))
        'sw.WriteLine("typ:" & typ & ", anzahl:" & Environment.NewLine &
        '       dt.Rows.Count & ", sizesumme: " & getFileSize4Length(sizeSumme) & Environment.NewLine &
        '        ", im schnitt: " & getFileSize4Length(sizeSumme / dt.Rows.Count))
        Try
            Process.Start(logfile)
        Catch ex As Exception
            l("fehler : " & ex.ToString)
        End Try
    End Sub

    Private Function getinsertSql(bearbeiterid As Integer, vid As String) As String
        Dim Sql = "INSERT INTO [dbo].[t47] ([VORGANGSID],[BEARBEITERID]) VALUES (" & vid & "," & bearbeiterid & ");" & Environment.NewLine

        Return Sql

    End Function

    Private Function getBearbeiterID(initial As String, bearbeiterDT As DataTable) As Integer
        For i = 0 To bearbeiterDT.Rows.Count - 1
            If CStr(bearbeiterDT.Rows(i).Item("INITIAL_")).ToLower = initial Then
                Return bearbeiterDT.Rows(i).Item("BEARBEITERID")
            End If
        Next
        Return 0
    End Function

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        'alle dokus auf vorhandensein prüfen
        Dim DT As DataTable
        l("PDFumwandeln ")
        'Dim logfile As String = "\\file-paradigma\paradigma\test\thumbnails\PDFlog" & Format(Now, "ddhhmmss") & ".txt"

        Dim dateifehlt As String = "\\file-paradigma\paradigma\test\thumbnails\dateifehlt_alle1" & Environment.UserName & ".txt"
        swfehlt = New IO.StreamWriter(dateifehlt)
        swfehlt.AutoFlush = True
        swfehlt.WriteLine(Now)

        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        '  vid = modPrep.getVid()


        If vid = "fehler" Then End
        Dim Sql As String
        Sql = "SELECT * FROM dokumente where   dokumentid<2000000 and dokumentid>0  " &
                  "  order by dokumentid desc "
        DT = alleDokumentDatenHolen(Sql)
        'teil1 = pdf -----------------------------------------------

        l("vor pdfverarbeiten")
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim initial As String
        Dim eid As Integer = 0
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()

        l("PDFumwandeln 2 ")
        l("PDFumwandeln 2 ")
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenHolen(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid)
                l(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")

                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If
                'clsBlob.dokufull_speichern(dokumentid, myoracle, inputfile)
                TextBox3.Text = igesamt & " von " & DT.Rows.Count
                Application.DoEvents()
                Dim fi As New IO.FileInfo(inputfile.Replace(Chr(34), ""))
                If Not fi.Exists Then
                    swfehlt.WriteLine(vid & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                    Continue For
                Else

                End If
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       inputfile & Environment.NewLine &
                       vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()
        Next
        If batchmode = True Then

        End If
        swfehlt.Close()
        l("dateifehlt  " & dateifehlt)
        Process.Start(dateifehlt)
    End Sub



    Private Function dateiImZielVerzLoeschen(strFile As String, quellstrDir As String, zieldir As String, ByRef zielname As String) As Boolean
        Dim qfile As New IO.FileInfo(strFile)
        Dim zielunterverzeichnis As String
        zielname = qfile.Name
        zielunterverzeichnis = qfile.DirectoryName
        zielunterverzeichnis = zielunterverzeichnis.Replace(quellstrDir, zieldir)
        zielname = zielunterverzeichnis & "\" & zielname
        Dim zfile As New IO.FileInfo(zielname)
        Try
            If zfile.Exists Then
                zfile.Delete()
            End If
            Return True
        Catch ex As Exception
            Debug.Print("fehler beim löschen: " & zielname)
            Return False
        End Try
    End Function
    Private Function worddateiAlsDocxSpeichern(strFile As String, quellDir As String, zieldir As String, zielname As String) As Boolean
        clsWordTest.konvOneDoc2Docx(strFile, zielname)
        Return True
    End Function



    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        IO.Directory.SetCurrentDirectory("L:\system\batch\margit")
        'MessageBox.Show("You are in the Form.Shown event.")
        If Environment.CommandLine.ToLower.Contains("batchmode=true") Then
            Application.DoEvents()
            batchmode = True
            PDFumwandeln()
            Button7.Text = "jetzt DOCXs"
            DOCXumwandeln(2113, False)
            BackColor = Color.Aquamarine
            '  MessageBox.Show("fertig")
            End
        End If
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        'revisionssichere Dokumente zusätzlich nach BLOB sichern
        Dim DT As DataTable
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim initial As String
        Dim eid As Integer
        l("revisionssicher ")
        Dim logfile As String = "C:\tempout\blob\in_" & Environment.UserName & ".txt"
        sw = New IO.StreamWriter(logfile)
        sw.AutoFlush = True
        sw.WriteLine(Now)
        'outdir = "L:\cache\paradigma\thumbnails\"
        outdir = "C:\tempout\blob\"

        checkoutRoot = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\paradigma\muell\"
        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        '  vid = modPrep.getVid()

        sw.WriteLine(vid)
        If vid = "fehler" Then End
        Dim Sql As String
        Sql = "SELECT * FROM dokumente where   dokumentid<2000000 and dokumentid>0  " &
                  " and (revisionssicher=1) order by dokumentid desc "
        Sql = "SELECT * FROM dokumente " &
            " LEFT JOIN t08 " &
            " ON dokumente.DOKUMENTID = t08.DOKID " &
            " where  (dokid is null) and (revisionssicher=1) and lower(typ)<>'jpg' and lower(typ)<>'png'"
        DT = RevSicherdokumentDatenHolen(Sql)
        'teil1 = pdf -----------------------------------------------
        Dim igesamt As Integer = 0
        Dim ic As Integer = 0

        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()

        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                If igesamt > 500 Then
                    Debug.Print("top")
                End If
                DbMetaDatenHolen(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid)
                l(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")
                sw.WriteLine(vid & " did: " & CStr(dokumentid) & " " & ic & " (count: " & DT.Rows.Count & ")")
                If istFoto(dateinameext) Then
                    Continue For
                End If
                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If
                Dim blobid As Long
                blobid = clsBlob.db_speichern(inputfile, dokumentid, myoracle, eid, vid)

            Catch ex As Exception
                l("fehler1: " & ex.ToString)
            End Try
        Next
    End Sub

    Shared Function getMSSQLCon() As SqlClient.SqlConnection
        Dim myoracle As SqlClient.SqlConnection
        Dim host, datenbank, schema, tabelle, dbuser, dbpw, dbport As String
        host = "kh-w-sql02" : schema = "Paradigma" : dbuser = "sgis" : dbpw = " WinterErschranzt.74"
        Dim conbuil As New SqlClient.SqlConnectionStringBuilder
        Dim v = "Data Source=" & host & ";User ID=" & dbuser & ";Password=" & dbpw & ";" +
                "Initial Catalog=" & schema & ";"

        myoracle = New SqlClient.SqlConnection(v)
        Return myoracle
    End Function

    Private Shared Function istFoto(dateinameext As String) As Boolean
        Return dateinameext.ToLower.EndsWith(".jpg") Or
                            dateinameext.ToLower.EndsWith(".jpeg") Or
                            dateinameext.ToLower.EndsWith(".png") Or
                            dateinameext.ToLower.EndsWith(".tif") Or
                            dateinameext.ToLower.EndsWith(".tiff")
    End Function

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Dim DT As DataTable
        Dim logfile As String = "\\file-paradigma\paradigma\test\thumbnailsOOOO\dokufilesize_" & Format(Now, "ddhhmmss") & ".txt"
        'logfile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\paradigma\muell\thumbnailer.log"
        sw = New IO.StreamWriter(logfile)
        sw.AutoFlush = True
        'sw.WriteLine(Now)
        If vid = "fehler" Then End
        DT = alleDokumentDatenHolenohnemb()
        'DT = alleDokumentDatenHolen()
        l("vor prüfung")
        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        Dim ic As Integer = 0
        Dim eid As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim str As String

        Dim initial As String
        '  Using sw As New IO.StreamWriter(logfile)
        Dim Sql As String
        getvidAufnahemdatum(DT, Sql)

        Dim altdatum As Date = Today
        Dim altvid As Integer = 0
        Dim days As Long
        'For i = 50000 To 1 Step -1
        '    If CStr(drr.Item("vid")) Then
        'Next
        For Each drr As DataRow In DT.Rows
            If CDate(drr.Item("aufnahme")) > altdatum Then
                Debug.Print("zualt " & drr.Item("vorgangsid") & " " & drr.Item("aufnahme"))
                ' Determine the number of days between the two dates.
                days = DateDiff(DateInterval.Day, drr.Item("aufnahme"), altdatum)
                If days > 1 Then
                    Debug.Print("zualt " & drr.Item("vorgangsid") & " " & drr.Item("aufnahme"))
                End If
            Else
                Debug.Print("istok " & drr.Item("vorgangsid") & " " & drr.Item("aufnahme"))

            End If
            altdatum = drr.Item("aufnahme")
            altvid = drr.Item("vorgangsid")
        Next
    End Sub

    Private Shared Sub getvidAufnahemdatum(ByRef DT As DataTable, ByRef Sql As String)
        Try
            Sql = "SELECT  [VORGANGSID]    ,[aufnahme]" &
                "  FROM [Paradigma].[dbo].[t41]" &
                "  order by VORGANGSID desc "
            DT = getDT(Sql)
            l("nach getDT")

        Catch ex As Exception

        End Try
    End Sub



    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        DokFileSize()
    End Sub
    Private Sub DokFileSize()
        Dim DT As DataTable
        Dim logfile As String = "\\file-paradigma\paradigma\test\thumbnailsOOOO\dokufilesize_" & Format(Now, "ddhhmmss") & ".txt"
        'logfile = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\paradigma\muell\thumbnailer.log"
        sw = New IO.StreamWriter(logfile)
        sw.AutoFlush = True
        'sw.WriteLine(Now)
        If vid = "fehler" Then End
        DT = alleDokumentDatenHolenohnemb()
        'DT = alleDokumentDatenHolen()
        l("vor prüfung")
        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        Dim ic As Integer = 0
        Dim eid As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim str As String

        Dim initial As String
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenHolen(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid)
                '   l(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")
                '   sw.WriteLine(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")
                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If
                If istRevisionssicher Then
                    If CheckBox1.Checked Then
                        inputFileReadonlysetzen(inputfile)
                    End If
                Else
                    If inputFileReadonlyEntfernen(inputfile) Then
                        icntREADONLYentfernt += 1
                        l("icntREADONLYentfernt: " & inputfile)
                    End If
                End If
                Dim fo As New IO.FileInfo(inputfile)
                TextBox3.Text = igesamt & " von " & DT.Rows.Count
                Application.DoEvents()
                If fo.Exists Then
                    str = GetFileSizeInMB(fo.FullName)
                    If str = "0" Then
                        'Continue For
                        str = "0,00001"
                    End If
                    sw.WriteLine("update dokumente set MB=" & str.Replace(",", ".") &
                         " where dokumentid=" & dokumentid & ";")
                Else
                    Continue For
                    ic += 1
                    l("dokument fehlt: " & ic.ToString & Environment.NewLine & " " &
                                     inputfile & Environment.NewLine &
                                     vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")")
                    TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                      inputfile & Environment.NewLine &
                      vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                      TextBox2.Text
                    Application.DoEvents()
                End If
            Catch ex As Exception
                l("fehler1: " & ex.ToString)
            End Try
        Next
        sw.Close()
        Debug.Print(icntREADONLYentfernt)
        Process.Start(logfile)
    End Sub
    Public Function GetFileSizeInMB(ByVal path As String) As Double
        Dim myFile As IO.FileInfo
        Dim mySize As Single
        Try
            myFile = New IO.FileInfo(path)
            If Not myFile.Exists Then
                mySize = 0
                Return 0
            Else
                mySize = myFile.Length
                Return Format(mySize / 1024 ^ 2, "###0.000") ' & " MB"
            End If
            'Select Case mySize 
            'Case 0 To 1023
            '    Return mySize & " Bytes"
            'Case 1024 To 1048575
            '    Return Format(mySize / 1024, "###0.00") & " KB"
            'Case 1048576 To 1043741824
            'Case Is > 1043741824
            '    Return Format(mySize / 1024 ^ 3, "###0.00") & " GB"
            'End Select
            myFile = Nothing
            Return "0 bytes"
        Catch ex As Exception
            Return "0 bytes"
        End Try
    End Function

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        fullpathdokumenteErzeugen()
    End Sub

    Private Sub fullpathdokumenteErzeugen()
        Dim dateifehlt As String = "\\file-paradigma\paradigma\test\thumbnails\auffueller" & Environment.UserName & ".txt"
        dateifehlt = "L:\system\batch\margit\auffueller" & Environment.UserName & ".txt"
        swfehlt = New IO.StreamWriter(dateifehlt)
        swfehlt.AutoFlush = True
        swfehlt.WriteLine(Now)
        ' S1020dokumenteMitFullpathTabelleErstellen("DOKUFULLNAME", swfehlt) 'referenzfälleNeuZuweisen
        swfehlt.WriteLine("wechsel")
        dokumenteMitFullpathTabelleErstellen(swfehlt)

        swfehlt.Close()
        l("fertig  " & dateifehlt)
        Process.Start(dateifehlt)
    End Sub

    Private Sub S1020dokumenteMitFullpathTabelleErstellen(zieltabelle As String, swfehlt As IO.StreamWriter)
        'alle dokus auf vorhandensein prüfen
        Dim DT As DataTable
        l("PDFumwandeln ")
        'Dim logfile As String = "\\file-paradigma\paradigma\test\thumbnails\PDFlog" & Format(Now, "ddhhmmss") & ".txt"


        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        '  vid = modPrep.getVid()

        swfehlt.WriteLine("Teil1 referenz  Dokumente ausschreiben ---------------------")
        If vid = "fehler" Then End

        Dim Sql As String

        ' 'alle vorgänge mit referenzfällen
        ' proVorgang:  referenzverwandte zum vorgang
        ' proVorgang:  alle referenzdokus zu einem vorgang
        Dim alleVorgaengeMitReferenzen As DataTable
        Dim tempReferenzVorgaenge As DataTable
        Dim tempREfDokumente As DataTable
        Sql = "SELECT  [VORGANGSID]" &
                 " FROM [Paradigma].[dbo].[t44]" &
                 " where FREMDVORGANGSID in" &
                  "(" &
                 " SELECT   VORGANGSID" &
                 " FROM [Paradigma].[dbo].[VORGANG_T43] a, DOKUMENTE b" &
                 " where a.SACHGEBIETNR='1020'" &
                 " and a.VORGANGSID=b.VID" &
                 " )"
        alleVorgaengeMitReferenzen = alleDokumentDatenHolen(Sql)
        ' proVorgang:  referenzverwandte zum vorgang
        ' proVorgang:  alle referenzdokus zu einem vorgang

        'Sql = "SELECT * FROM dokumentefull where   dokumentid<2000000 and dokumentid>0  and fullname is null " &
        '          "  order by dokumentid desc "
        'DT = alleDokumentDatenHolen(Sql)
        'teil1 = pdf -----------------------------------------------

        l("vor pdfverarbeiten")
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim initial As String
        Dim eid As Integer = 0
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        '

        Dim aktVID As Integer = 0
        Dim fremdvorgangsid As Integer = 0
        Dim fremddokumentid As Integer = 0

        swfehlt.WriteLine("Teil1 alleVorgaengeMitReferenzen.Rows: " & alleVorgaengeMitReferenzen.Rows.Count.ToString)
        'Dim max As Integer
        'max = alleVorgaengeMitReferenzen.Rows.Count
        'max = 1000
        'swfehlt.WriteLine("Teit masx:" & max)
        Dim idok As Integer = 0
        For Each drr As DataRow In alleVorgaengeMitReferenzen.Rows
            Try
                igesamt += 1
                aktVID = CStr(drr.Item("VORGANGSID"))
                Sql = "  SELECT   FREMDVORGANGSID  FROM [Paradigma].[dbo].t44 a" &
                     " where     VORGANGSID= " & aktVID & "" &
                     " and FREMDVORGANGSID in (" &
                    "	 SELECT  VORGANGSID" &
                    "	  FROM [Paradigma].[dbo].[VORGANG_T43] b" &
                    "	  where  b.SACHGEBIETNR='1020' " &
                     " )"
                tempReferenzVorgaenge = alleDokumentDatenHolen(Sql)
                For Each fremdv As DataRow In tempReferenzVorgaenge.Rows
                    Try
                        '   igesamt += 1
                        fremdvorgangsid = CStr(fremdv.Item("FREMDVORGANGSID"))
                        Debug.Print(igesamt & ", " &
                                    alleVorgaengeMitReferenzen.Rows.Count.ToString & "/" &
                                    CStr(fremdv.Item("FREMDVORGANGSID")))

                        Sql = " Select distinct  b.*  " &
                                "   FROM [Paradigma].[dbo].[VORGANG_T43] a, DOKUMENTE b" &
                                "   where a.SACHGEBIETNR='1020'" &
                                "   and b.vid=" & fremdvorgangsid & " "
                        tempREfDokumente = alleDokumentDatenHolen(Sql)
                        For Each fremddokus As DataRow In tempREfDokumente.Rows
                            Try
                                fremddokumentid = CStr(fremddokus.Item("DOKUMENTID"))
                                Debug.Print(CStr(fremddokumentid))


                                DbMetaDatenHolen(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, fremddokus, dbdatum, istRevisionssicher, initial, eid)
                                vid = aktVID
                                'l(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")

                                If newsavemode Then
                                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                                Else
                                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                                End If

                                Application.DoEvents()
                                Dim fi As New IO.FileInfo(inputfile.Replace(Chr(34), ""))
                                If Not fi.Exists Then
                                    swfehlt.WriteLine(vid & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                                    Continue For
                                Else
                                    If clsBlob.dokufull_speichern(dokumentid, myoracle, inputfile, vid, zieltabelle) <> 0 Then
                                        MsgBox("Fehler")
                                    Else

                                    End If
                                End If
                                idok += 1
                                swfehlt.WriteLine(idok & " eingefügt/ref")
                            Catch ex3 As Exception
                                Debug.Print(ex3.ToString)
                            End Try
                            l(igesamt & " (" & alleVorgaengeMitReferenzen.Rows.Count & ") " & " aktvid: " & aktVID & " docid" & CStr(dokumentid) & " ")
                            Debug.Print(igesamt & " (" & alleVorgaengeMitReferenzen.Rows.Count & ") " & " aktvid: " & aktVID & " docid" & CStr(dokumentid) & " ")
                        Next
                    Catch ex2 As Exception
                        Debug.Print(ex2.ToString)
                    End Try
                Next
            Catch ex As Exception

                Debug.Print(ex.ToString)

            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()

        Next
        l("PDFumwandeln 2 ")

        If batchmode = True Then

        End If
        swfehlt.WriteLine(idok & " Teil1 fertig  --------------------- " & igesamt)
    End Sub
    Private Sub dokumenteMitFullpathTabelleErstellen(swfehlt As IO.StreamWriter)

        Dim DT As DataTable
        Dim idok As Integer = 0
        l("PDFumwandeln ")
        swfehlt.WriteLine("Teil2 normale Dokumente ausschreiben ---------------------")

        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        If vid = "fehler" Then End
        Dim Sql As String
        Sql = "SELECT * FROM dokumentefull where   dokumentid<20000000 and dokumentid>0  and fullname is null " &
                  "  order by dokumentid desc "
        Sql = "SELECT * FROM dokumente where   dokumentid<20000000 and dokumentid>0  and (tooltip ='' or    tooltip is null) " &
                  "  order by dokumentid desc "
        DT = alleDokumentDatenHolen(Sql)
        'teil1 = pdf -----------------------------------------------

        l("vor pdfverarbeiten")
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim initial As String
        Dim eid As Integer = 0
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        myoracle.Open()
        l("PDFumwandeln 2 ")
        l("PDFumwandeln 2 ")
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenHolen(vid, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid)
                l(vid & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")

                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If

                TextBox3.Text = igesamt & " von " & DT.Rows.Count
                Application.DoEvents()
                Dim fi As New IO.FileInfo(inputfile.Replace(Chr(34), ""))
                If Not fi.Exists Then
                    swfehlt.WriteLine(vid & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                    Continue For
                Else
                    'If clsBlob.dokufull_speichern(dokumentid, myoracle, inputfile, vid, zieltabelle) <> 0 Then
                    '    MsgBox("Fehler")
                    'Else

                    'End If
                    If clsBlob.saveDokumenteTooltip(dokumentid, myoracle, inputfile, vid) <> 0 Then
                        'MsgBox("Fehler")
                    Else
                        'MsgBox("Fehler")
                    End If
                    idok += 1
                    swfehlt.WriteLine(idok & " eingefügt/norm")
                End If
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       inputfile & Environment.NewLine &
                       vid & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()
        Next
        If batchmode = True Then

        End If
        swfehlt.WriteLine(idok & "Teil2 fertig  --------------------- " & igesamt)
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click

        Dim dateifehlt As String = "\\file-paradigma\paradigma\test\thumbnails\referenzdokus" & Environment.UserName & ".txt"
        dateifehlt = "L:\system\batch\margit\referenzdokus" & Environment.UserName & ".txt"
        swfehlt = New IO.StreamWriter(dateifehlt)
        swfehlt.AutoFlush = True
        swfehlt.WriteLine(Now)
        MsgBox("Bitte zuerst die Tabelle DOKUFULLNAME  löschen oder leeren  delete   FROM [Paradigma].[dbo].[DOKUFULLNAME]")
        S1020dokumenteMitFullpathTabelleErstellen("DOKUFULLNAME", swfehlt) 'referenzfälleNeuZuweisen
        swfehlt.WriteLine("feddich")
        'dokumenteMitFullpathTabelleErstellen(swfehlt) 
        swfehlt.Close()
        l("fertig  " & dateifehlt)
        Process.Start(dateifehlt)
    End Sub

    'Private Sub Button16_Click_1(sender As Object, e As EventArgs) Handles Button16.Click

    'End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        ' BLOB als Datei speichern
        Dim DT As DataTable
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim initial As String
        Dim eid As Integer
        Dim ausCheckDokumentid = 414080
        Dim res = InputBox("DOKID: ", "Bitte die DOkID des gewünschten Dokumentes hier angeben!", 414080)
        ausCheckDokumentid = res
        l("revisionssicher ")
        Dim logfile As String = "C:\tempout\blob\Blobout_" & Environment.UserName & ".txt"
        sw = New IO.StreamWriter(logfile)
        sw.AutoFlush = True
        sw.WriteLine(Now)
        'outdir = "L:\cache\paradigma\thumbnails\"
        outdir = "C:\tempout\blob"

        checkoutRoot = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\paradigma\muell\"
        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        '  vid = modPrep.getVid()

        If vid = "fehler" Then End
        Dim Sql As String
        Sql = "SELECT * FROM dokumente where   dokumentid<2000000 and dokumentid>0  " &
                  " and (revisionssicher=1) order by dokumentid desc "
        Sql = "SELECT * FROM dokumente " &
            " LEFT JOIN t08 " &
            " ON dokumente.DOKUMENTID = t08.DOKID " &
            " where  dokid =" & ausCheckDokumentid
        DT = RevSicherdokumentDatenHolen(Sql)
        dateinameext = CStr(DT.Rows(0).Item("dateinameext"))
        'teil1 = pdf -----------------------------------------------
        Dim igesamt As Integer = 0
        Dim ic As Integer = 0
        '##########################
        Dim myoracle = getSQLConnection()
        sw.WriteLine(ausCheckDokumentid)
        outfile = outdir & "\" & dateinameext '"\ausCheckDokumentid.pdf"
        Dim ausgecheckt As Boolean = checkoutNachDatei(ausCheckDokumentid, outfile, myoracle)

        Process.Start(outdir)

    End Sub

    Private Shared Function getSQLConnection() As SqlClient.SqlConnection
        Dim myoracle As SqlClient.SqlConnection
        Dim host, datenbank, schema, tabelle, dbuser, dbpw, dbport As String
        host = "kh-w-sql02" : schema = "Paradigma" : dbuser = "sgis" : dbpw = " WinterErschranzt.74"
        'Dim conbuil As New SqlClient.SqlConnectionStringBuilder
        Dim v = "Data Source=" & host & ";User ID=" & dbuser & ";Password=" & dbpw & ";" +
                "Initial Catalog=" & schema & ";"

        myoracle = New SqlClient.SqlConnection(v)
        Return myoracle
    End Function

    Shared Function checkoutNachDatei(dokmetaDokid As Integer, dateiname As String, myoracle As SqlClient.SqlConnection) As Boolean
        Try
            sw.WriteLine("checkoutNachDatei---------------------- anfang")
            If clsBlob.ausBLOBdbholen(dateiname, dokmetaDokid, myoracle) Then
                Return True
            Else
                Return False
            End If
            sw.WriteLine("checkoutNachDatei---------------------- ende")
        Catch ex As System.Exception
            sw.WriteLine("Fehler in checkoutNachDatei: " & ex.ToString())
            Return False
        End Try
    End Function
End Class
