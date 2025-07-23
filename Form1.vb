Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
'Imports Acrobat
'Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Word

Public Class Form1
    Public wordVorlagen As New Microsoft.Office.Interop.Word.Application 'habe hier new ergänzt ????
    Public docVorlagen As New Microsoft.Office.Interop.Word.Document

    Private immerUeberschreiben As Boolean
    Private nichtUeberschreibenAusserWennNeuer As Boolean
    Property dt As New Data.DataTable
    Private count As Integer = 0

    Public Shared inndir, outdir, Bearbeitungsart, checkoutRoot As String
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
            .CustomLocation = "D:\probaug_Ausgabe\logs\" & ""
#Else
            .CustomLocation = "D:\probaug_Ausgabe\logs\" & ""
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
            Sql = "SELECT * FROM dokumente where   ort<2000000 and ort>0  " &
                  " and ( Vorhaben='pdf') order by ort desc "
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
            Sql = "SELECT * FROM dokumente where   ort<2000000 and ort>0  " &
                  "  and mb =0 " &
                  "  order by ort desc "
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
    '        'filename = getJPGfilename(filename, outdir, Bearbeitungsart)
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
        If Bearbeitungsart = "fehler" Then End
        Dim Sql As String
        Sql = "SELECT * FROM dokumente where   ort<2000000 and ort>0  " &
                  "  order by ort desc "
        DT = alleDokumentDatenHolen(Sql)
        l("vor prüfung")
        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        Dim ic As Integer = 0
        Dim eid As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim beschreibung As String
        Dim fullfilename As String
        Dim dbdatum As Date
        Dim eingang As Date
        Dim initial As String
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenDokumentHolen(Bearbeitungsart, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid,
                                 beschreibung, eingang, fullfilename)
                '   l(Bearbeitungsart & " " & CStr(ort) & " " & ic & " (" & DT.Rows.Count & ")")
                '   sw.WriteLine(Bearbeitungsart & " " & CStr(ort) & " " & ic & " (" & DT.Rows.Count & ")")
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
                    'inputFileReadonlyEntfernen(Vorhabensmerkmal)
                    Continue For
                Else
                    ic += 1
                    l("dokument fehlt: " & ic.ToString & Environment.NewLine & " " &
                                     inputfile & Environment.NewLine &
                                     Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")")


                    TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                      inputfile & Environment.NewLine &
                      Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
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



    'Private Sub PDFSverarbeiten(outdir As String, Bearbeitungsart As String, dt As DataTable)
    '    Dim ic As Integer = 0
    '    Dim sachgebiet As String = "", Verfahrensart As String = "", Vorhaben As String, batchfile As String
    '    Dim newsavemode As Boolean
    '    Dim dbdatum As Date
    '    Dim Vorhabensmerkmal, checkoutfile, bearbeiter As String
    '    Dim ort As String
    '    For Each drr As DataRow In dt.Rows
    '        Try
    '            ic += 1
    '            DbMetaDatenDokumentHolen(Bearbeitungsart, sachgebiet, Verfahrensart, Vorhaben, newsavemode, ort, drr, dbdatum)
    '            l(Bearbeitungsart & " " & ort.ToString & " " & ic & " (" & dt.Rows.Count & ")")
    '            TextBox1.Text = TextBox1.Text & Bearbeitungsart & " " & ort.ToString & " " & ic & " (" & dt.Rows.Count & ")"
    '            If newsavemode Then
    '                Vorhabensmerkmal = GetInputfile(inndir, sachgebiet, CInt(ort))
    '            Else
    '                Vorhabensmerkmal = GetInputfile1(inndir, sachgebiet, Verfahrensart)
    '            End If
    '            bearbeiter = modPrep.GetOutfile(CInt(Bearbeitungsart), outdir, CInt(ort), ".jpg")
    '            Dim fi As New IO.FileInfo(bearbeiter.Replace(Chr(34), ""))

    '            If fi.Exists Then
    '                l("exists")
    '                Continue For
    '            End If
    '            If Not IO.Directory.Exists(outdir & Bearbeitungsart.ToString) Then
    '                IO.Directory.CreateDirectory(outdir & Bearbeitungsart.ToString)
    '            End If

    '        Catch ex As Exception
    '            l("fehler1: " & ex.ToString)
    '        End Try
    '        Try
    '            convertPDF(Vorhabensmerkmal.Replace(Chr(34), ""), bearbeiter)
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
        '  Bearbeitungsart = modPrep.getVid()

        sw.WriteLine(Bearbeitungsart)
        If Bearbeitungsart = "fehler" Then End
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
        Dim eingang As Date
        Dim initial As String
        Dim beschreibung As String
        Dim fullfilename As String

        l("PDFumwandeln 2 ")
        l("PDFumwandeln 2 ")
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenDokumentHolen(Bearbeitungsart, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid, beschreibung, eingang, fullfilename)
                l(Bearbeitungsart & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")
                sw.WriteLine(Bearbeitungsart & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")


                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If
                outfile = modPrep.GetOutfileName(CInt(Bearbeitungsart), outdir, CInt(dokumentid), ".jpg")
                Dim fo As New IO.FileInfo(outfile.Replace(Chr(34), ""))
                TextBox3.Text = igesamt & " von " & DT.Rows.Count
                Application.DoEvents()
                If fo.Exists Then
                    ' l("exists")
                    Continue For
                End If


                TextBox1.Text = ic.ToString & " / " & dateinameext & Environment.NewLine & " " &
                inputfile & Environment.NewLine &
                Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")"
                Application.DoEvents()


                If Not IO.Directory.Exists(outdir & Bearbeitungsart.ToString) Then
                    IO.Directory.CreateDirectory(outdir & Bearbeitungsart.ToString)
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
                'If ort = "77828" Then Continue For
                'If ort = "80043" Then Continue For
                'If ort = "80071" Then Continue For
                Dim fi As New IO.FileInfo(inputfile.Replace(Chr(34), ""))
                If Not fi.Exists Then
                    swfehlt.WriteLine(Bearbeitungsart & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                    Continue For
                Else

                End If

                If convertPDF2(inputfile, outfile) Then
                    l("erfolg")
                    ic += 1
                    TextBox1.Text = ic.ToString & " / " & dateinameext & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")"
                    Application.DoEvents()
                Else
                    l("erfolglos " & ic.ToString & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")")
                    TextBox2.Text = ic.ToString & " / " & dateinameext & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                        TextBox2.Text
                    Application.DoEvents()
                End If
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       inputfile & Environment.NewLine &
                       Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
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
        '    If Not IsNumeric(Bearbeitungsart) Then Exit Sub
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

        Sql = "SELECT * FROM dokumente where   ort > " & unten & "  and ort < " & oben & "  " &
              "and ( Vorhaben='docx' or  Vorhaben='doc'  or  Vorhaben='rtf' )  " &
              "order by ort desc"
        'Sql = "SELECT * FROM dokumente where   Bearbeitungsart=9609 " &
        '      "and ( Vorhaben='docx' or  Vorhaben='doc' or  Vorhaben='rtf')  " &
        '      "order by ort desc"
        l(Sql)

        immerUeberschreiben = True
        dt = getDT(Sql)

        l("nach getDT")
        Dim relativpfad As String = "", dateinameext As String = "", typ As String, logfile As String
        Dim newsavemode As Boolean

        Dim dokumentid As Integer = 0
        'l("nach 1: " & outDir & Bearbeitungsart.ToString)
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
                    'swfehlt.WriteLine(Bearbeitungsart & "," & ort & ", " & dbdatum & "," & initial & ", " & Vorhabensmerkmal & "")
                    swfehlt.WriteLine(vid & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                    Continue For
                Else
                    TextBox1.Text = dateinameext & " ist dran  " & Environment.NewLine
                    Application.DoEvents()
                End If

                'inputFileReadonlyEntfernen(Vorhabensmerkmal)
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
                    'sw.WriteLine("-- " & Verfahrensart)
                    Continue For
                Else
                    'l(" ")
                    'sw.WriteLine("-- " & Verfahrensart)
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
        '    If Not IsNumeric(Bearbeitungsart) Then Exit Sub
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

        Sql = "SELECT * FROM dokumente where   ort > " & unten & "  and ort < " & oben & "  " &
              "and (  lower(Verfahrensart) like '%.xls' )  " &
              "order by ort desc"
        'Sql = "SELECT * FROM dokumente where   Bearbeitungsart=9609 " &
        '      "and ( Vorhaben='docx' or  Vorhaben='doc' or  Vorhaben='rtf')  " &
        '      "order by ort desc"


        immerUeberschreiben = True
        dt = getDT(Sql)

        l("nach getDT")
        Dim relativpfad As String = "", dateinameext As String = "", typ As String, logfile As String
        Dim newsavemode As Boolean

        Dim dokumentid As Integer = 0
        'l("nach 1: " & outDir & Bearbeitungsart.ToString)
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
                'sw.WriteLine(Bearbeitungsart & "/" & ort & "----")
                If newsavemode Then
                    inputfile = GetInputfileWordFullPath(innDir, relativpfad, dokumentid)
                Else
                    inputfile = GetInputfile1WordFullPath(innDir, relativpfad, dateinameext)
                End If

                'inputFileReadonlyEntfernen(Vorhabensmerkmal)
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
                    'swfehlt.WriteLine(Bearbeitungsart & "," & ort & ", " & dbdatum & "," & initial & ", " & Vorhabensmerkmal & "")
                    swfehlt.WriteLine(vid & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                    Continue For
                Else

                End If

                If Not auscheckenword(inputfile, checkoutfile, sw, CType(vid, String), CType(dokumentid, String)) Then
                    l("-- " & dateinameext)
                    'sw.WriteLine("-- " & Verfahrensart)
                    Continue For
                Else
                    'l(" ")
                    'sw.WriteLine("-- " & Verfahrensart)
                    inputFileReadonlyEntfernen(checkoutfile)
                End If
                TextBox1.Text = checkoutfile & " checkout erfolgreich   " & Environment.NewLine
                Application.DoEvents()
                If clsExcel.konvOne(checkoutfile, outfile) Then
                    TextBox1.Text = TextBox1.Text & " " & pdffile & " xls erfolgreich" & Environment.NewLine
                    Application.DoEvents()
                    'altearchivdatei umbenennen
                    'neuedatei im archiv speichern
                    'in db: Vorhaben anpassen
                    'in db: Vorhaben in Verfahrensart anpassen
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
        'modOracle.setExcelAttribute2(Bearbeitungsart, sachgebiet, Verfahrensart, Vorhaben, newsavemode, ort, drr)
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
        '  Bearbeitungsart = modPrep.getVid()

        sw.WriteLine(Bearbeitungsart)
        If Bearbeitungsart = "fehler" Then End
        DT = PDFdokumentDatenHolen()
        'teil1 = pdf -----------------------------------------------

        l("vor pdfverarbeiten")
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim eingang As Date
        Dim initial As String
        Dim beschreibung As String
        Dim fullfilename As String
        Dim eid As Integer = 0
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                If Bearbeitungsart = 8930 Then
                    Debug.Print("")
                End If
                igesamt += 1
                DbMetaDatenDokumentHolen(Bearbeitungsart, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid,
                                 beschreibung, eingang, fullfilename)
                l(Bearbeitungsart & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")

                sw.WriteLine(Bearbeitungsart & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")
                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If
                outfile = modPrep.GetOutfileName(CInt(Bearbeitungsart), outdir, CInt(dokumentid), ".jpg")
                Dim fo As New IO.FileInfo(outfile.Replace(Chr(34), ""))
                TextBox3.Text = igesamt & " von " & DT.Rows.Count
                Application.DoEvents()
                If fo.Exists Then
                    l("exists")
                    Continue For
                End If
                If Not IO.Directory.Exists(outdir & Bearbeitungsart.ToString) Then
                    IO.Directory.CreateDirectory(outdir & Bearbeitungsart.ToString)
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
                'If ort = "77828" Then Continue For
                'If ort = "80043" Then Continue For
                'If ort = "80071" Then Continue For
                If convertPDF2(inputfile, outfile) Then
                    l("erfolg")
                    ic += 1
                    TextBox1.Text = ic.ToString & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")"
                    Application.DoEvents()
                Else
                    l("erfolglos " & ic.ToString & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")")
                    TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                        inputfile & Environment.NewLine &
                        Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                        TextBox2.Text
                    Application.DoEvents()
                End If
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       inputfile & Environment.NewLine &
                       Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
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
        '    If Not IsNumeric(Bearbeitungsart) Then Exit Sub
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
        Sql = "SELECT *   FROM [Paradigma].[dbo].[DOKUMENTE] where lower(Vorhaben)='docx' or lower(Vorhaben)='doc' or lower(Vorhaben)='pdf' "
        immerUeberschreiben = True
        dt = getDT(Sql)

        l("nach getDT")

        'l("nach 1: " & outDir & Bearbeitungsart.ToString)
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
                datenholden(Bearbeitungsart, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, initial)
                If dokumentid = 174865 Then
                    Debug.Print("")

                End If
                Console.Write(Bearbeitungsart & "/" & dokumentid & "----")
                'sw.WriteLine(Bearbeitungsart & "/" & ort & "----")
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
        sw.WriteLine("Vorhaben:" & typ & ", anzahl:" & Environment.NewLine &
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

        'l("nach 1: " & outDir & Bearbeitungsart.ToString)
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

                Console.Write(Form1.Bearbeitungsart & "/" & initial_ & "----")
                'sw.WriteLine(Bearbeitungsart & "/" & ort & "----")

                'sizeSumme += fi.Length
                'TextBox1.Text = TextBox1.Text & ic & " " & Vorhabensmerkmal & " " & fi.Length & Environment.NewLine
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
        'sw.WriteLine("Vorhaben:" & Vorhaben & ", anzahl:" & Environment.NewLine &
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

        'l("nach 1: " & outDir & Bearbeitungsart.ToString)
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
                Form1.Bearbeitungsart = (drr.Item("vorgangsid"))
                b = weitere.Split(New Char() {";"c},
                        StringSplitOptions.RemoveEmptyEntries)
                ReDim bids(b.Length - 1)
                For i = 0 To b.Length - 1
                    bids(i) = getBearbeiterID(b(i), bearbeiterDT)
                    sw.WriteLine(getinsertSql(bids(i), Form1.Bearbeitungsart))
                    Console.Write(Form1.Bearbeitungsart & "/" & initial_ & "----")
                Next






                'sw.WriteLine(Bearbeitungsart & "/" & ort & "----")

                'sizeSumme += fi.Length
                'TextBox1.Text = TextBox1.Text & ic & " " & Vorhabensmerkmal & " " & fi.Length & Environment.NewLine
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
        'sw.WriteLine("Vorhaben:" & Vorhaben & ", anzahl:" & Environment.NewLine &
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
        '  Bearbeitungsart = modPrep.getVid()


        If Bearbeitungsart = "fehler" Then End
        Dim Sql As String
        Sql = "SELECT * FROM dokumente where   ort<2000000 and ort>0  " &
                  "  order by ort desc "
        DT = alleDokumentDatenHolen(Sql)
        'teil1 = pdf -----------------------------------------------

        l("vor pdfverarbeiten")
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim eingang As Date
        Dim initial As String
        Dim fullfilename As String
        Dim beschreibung As String
        Dim eid As Integer = 0
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()

        l("PDFumwandeln 2 ")
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenDokumentHolen(Bearbeitungsart, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid,
                                 beschreibung, eingang, fullfilename)
                l(Bearbeitungsart & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")

                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If
                'clsBlob.dokufull_speichern(ort, myoracle, Vorhabensmerkmal)
                TextBox3.Text = igesamt & " von " & DT.Rows.Count
                Application.DoEvents()
                Dim fi As New IO.FileInfo(inputfile.Replace(Chr(34), ""))
                If Not fi.Exists Then
                    swfehlt.WriteLine(Bearbeitungsart & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                    Continue For
                Else

                End If
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       inputfile & Environment.NewLine &
                       Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
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
        Dim fullfilename As String
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
        '  Bearbeitungsart = modPrep.getVid()

        sw.WriteLine(Bearbeitungsart)
        If Bearbeitungsart = "fehler" Then End
        Dim Sql As String
        Sql = "SELECT * FROM dokumente where   ort<2000000 and ort>0  " &
                  " and (revisionssicher=1) order by ort desc "
        Sql = "SELECT * FROM dokumente " &
            " LEFT JOIN t08 " &
            " ON dokumente.DOKUMENTID = t08.DOKID " &
            " where  (dokid is null) and (revisionssicher=1) and lower(Vorhaben)<>'jpg' and lower(Vorhaben)<>'png'"
        DT = RevSicherdokumentDatenHolen(Sql)
        'teil1 = pdf -----------------------------------------------
        Dim igesamt As Integer = 0
        Dim ic As Integer = 0

        Dim myoracle As SqlClient.SqlConnection
        Dim beschreibung As String
        myoracle = getMSSQLCon()
        Dim eingang As Date
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                If igesamt > 500 Then
                    Debug.Print("top")
                End If
                DbMetaDatenDokumentHolen(Bearbeitungsart, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid,
                                 beschreibung, eingang, fullfilename)
                l(Bearbeitungsart & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")
                sw.WriteLine(Bearbeitungsart & " did: " & CStr(dokumentid) & " " & ic & " (count: " & DT.Rows.Count & ")")
                If istFoto(dateinameext) Then
                    Continue For
                End If
                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If
                Dim blobid As Long
                blobid = clsBlob.db_speichern(inputfile, dokumentid, myoracle, eid, Bearbeitungsart)

            Catch ex As Exception
                l("fehler1: " & ex.ToString)
            End Try
        Next
    End Sub

    Shared Function getMSSQLCon() As SqlClient.SqlConnection
        Dim myoracle As SqlClient.SqlConnection
        Dim host, datenbank, schema, tabelle, dbuser, dbpw, dbport As String
        host = "kh-w-sql02" : schema = "Paradigma" : dbuser = "sgis" : dbpw = "Grunt8-Cornhusk-Reporter"
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
        If Bearbeitungsart = "fehler" Then End
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
        '    If CStr(drr.Item("Bearbeitungsart")) Then
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
        If Bearbeitungsart = "fehler" Then End
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
        Dim beschreibung As String
        Dim eingang As Date
        Dim str As String
        Dim fullfilename As String
        Dim initial As String
        '  Using sw As New IO.StreamWriter(logfile)
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenDokumentHolen(Bearbeitungsart, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid,
                                 beschreibung, eingang, fullfilename)
                '   l(Bearbeitungsart & " " & CStr(ort) & " " & ic & " (" & DT.Rows.Count & ")")
                '   sw.WriteLine(Bearbeitungsart & " " & CStr(ort) & " " & ic & " (" & DT.Rows.Count & ")")
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
                         " where ort=" & dokumentid & ";")
                Else
                    Continue For
                    ic += 1
                    l("dokument fehlt: " & ic.ToString & Environment.NewLine & " " &
                                     inputfile & Environment.NewLine &
                                     Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")")
                    TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                      inputfile & Environment.NewLine &
                      Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
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
        '  Bearbeitungsart = modPrep.getVid()

        swfehlt.WriteLine("Teil1 referenz  Dokumente ausschreiben ---------------------")
        If Bearbeitungsart = "fehler" Then End

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

        'Sql = "SELECT * FROM dokumentefull where   ort<2000000 and ort>0  and fullname is null " &
        '          "  order by ort desc "
        'DT = alleDokumentDatenHolen(Sql)
        'teil1 = pdf -----------------------------------------------

        l("vor pdfverarbeiten")
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim eingang As Date
        Dim initial As String
        Dim eid As Integer = 0
        Dim fullfilename As String
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        '
        Dim beschreibung As String
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
                                "   and b.Bearbeitungsart=" & fremdvorgangsid & " "
                        tempREfDokumente = alleDokumentDatenHolen(Sql)
                        For Each fremddokus As DataRow In tempREfDokumente.Rows
                            Try
                                fremddokumentid = CStr(fremddokus.Item("DOKUMENTID"))
                                Debug.Print(CStr(fremddokumentid))


                                DbMetaDatenDokumentHolen(Bearbeitungsart, relativpfad, dateinameext, typ, newsavemode, dokumentid, fremddokus, dbdatum, istRevisionssicher, initial, eid, beschreibung,
                                                eingang, fullfilename)
                                Bearbeitungsart = aktVID
                                'l(Bearbeitungsart & " " & CStr(ort) & " " & ic & " (" & DT.Rows.Count & ")")

                                If newsavemode Then
                                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                                Else
                                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                                End If

                                Application.DoEvents()
                                Dim fi As New IO.FileInfo(inputfile.Replace(Chr(34), ""))
                                If Not fi.Exists Then
                                    swfehlt.WriteLine(Bearbeitungsart & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext & ", " & inputfile & "")
                                    Continue For
                                Else
                                    If clsBlob.dokufull_speichern(dokumentid, myoracle, inputfile, Bearbeitungsart, zieltabelle) <> 0 Then
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
        If Bearbeitungsart = "fehler" Then End
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
        Dim beschreibung As String
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)
        Dim eingang As Date
        Dim initial As String
        Dim fullfilename As String
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
                DbMetaDatenDokumentHolen(Bearbeitungsart, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher, initial, eid,
                                 beschreibung, eingang, fullfilename)
                l(Bearbeitungsart & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")

                If newsavemode Then
                    inputfile = GetInputfilename(inndir, relativpfad, CInt(dokumentid))
                Else
                    inputfile = GetInputfile1Name(inndir, relativpfad, dateinameext)
                End If

                TextBox3.Text = igesamt & " von " & DT.Rows.Count
                Application.DoEvents()
                Dim fi As New IO.FileInfo(inputfile.Replace(Chr(34), ""))
                If Not fi.Exists Then
                    swfehlt.WriteLine("FEhlt: " & Bearbeitungsart & "," & dokumentid & ", " & dbdatum & "," & initial & "," & dateinameext) '& ", " & Vorhabensmerkmal & "")
                    Continue For
                Else
                    'If clsBlob.dokufull_speichern(ort, myoracle, Vorhabensmerkmal, Bearbeitungsart, zieltabelle) <> 0 Then
                    '    MsgBox("Fehler")
                    'Else

                    'End If
                    If clsBlob.saveDokumenteTooltip(dokumentid, myoracle, inputfile) <> 0 Then
                        'MsgBox("Fehler")
                    Else
                        'MsgBox("Fehler")
                    End If
                    idok += 1
                    If idok > maxobj Then Exit For
                    swfehlt.WriteLine(idok & " eingefügt/norm")
                End If
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       inputfile & Environment.NewLine &
                       Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
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

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Dim puFehler As String = "\\file-paradigma\paradigma\test\thumbnails\PU_ausgabeDoku" & Environment.UserName & ".txt"
        Dim puAusgabe As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\" & "dokumente" & ".csv"
        Dim puAusgabeStream As New IO.StreamWriter(puAusgabe)
        '   dateifehlt = "L:\system\batch\margit\auffueller" & Environment.UserName & ".txt"
        swfehlt = New IO.StreamWriter(puFehler)
        swfehlt.AutoFlush = True
        '  swfehlt.WriteLine(Now)
        ' S1020dokumenteMitFullpathTabelleErstellen("DOKUFULLNAME", swfehlt) 'referenzfälleNeuZuweisen
        ' swfehlt.WriteLine("wechsel")
        ' dokumenteMitFullpathTabelleErstellen(swfehlt)
        Dim Sql As String
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)
        Sql = "SELECT * FROM [Paradigma].[dbo].[probaug_dokumente_vorgang]  order by dokumentid desc "
        TextBox1.Text = puAusgabe
        TextBox2.Text = Sql
        'MsgBox("max. objekte für test: " & maxobj)
        writeDokumentePU(puFehler, puAusgabeStream, Sql, maxobj)
        puAusgabeStream.Close()
        puAusgabeStream.Dispose()
        'Process.Start(puAusgabe)
        '######
        puAusgabe = "D:\probaug_Ausgabe\" & "dokumente_referenz" & ".csv"
        puAusgabeStream = New IO.StreamWriter(puAusgabe)
        Sql = "SELECT * FROM [Paradigma].[dbo].[probaug_dokumente_referenz]  order by ort desc "
        TextBox1.Text = TextBox1.Text & Environment.NewLine & puAusgabe
        TextBox2.Text = TextBox2.Text & Environment.NewLine & Sql
        writeDokumentePU(puFehler, puAusgabeStream, Sql, 500)
        swfehlt.Close()
        l("fertig  " & puFehler)

        'Process.Start(puAusgabe)
    End Sub

    Private Function setMaxObj(maxobj As Integer) As Integer
        If (TextBox4.Text) = 0 Or TextBox4.Text = "" Then
            maxobj = 10000000
        Else
            If IsNumeric(TextBox4.Text) Then
                maxobj = CInt(TextBox4.Text)
            End If
        End If

        Return maxobj
    End Function

    Private Sub writeDokumentePU(puFehler As String, puAusgabeStream As IO.StreamWriter, sql As String, maxobj As Integer)
        '####
        Dim DT As DataTable
        Dim idok As Integer = 0
        puAusgabeStream.AutoFlush = True
        'ausgabeAntragsteller.WriteLine(Now)
        l("PDFumwandeln ")
        swfehlt.WriteLine("Teil2 normale Dokumente ausschreiben ---------------------")

        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        If Bearbeitungsart = "fehler" Then End

        DT = alleDokumentDatenHolen(sql)
        l("vor csvverarbeiten")
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim relativpfad, dateinameext, typ, dokumentid, inputfile, outfile, bearbeiterid As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim initial As String
        Dim beschreibung As String
        Dim eingang As Date
        Dim eid As Integer = 0
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        myoracle.Open()
        Dim zeile As New Text.StringBuilder
        'Dim block As New Text.StringBuilder 
        'Dim blockMAX As Int16 = 50
        'Dim iblock As Int16 = 0
        Dim fullfilename As String
        Dim t As String = ";"
        l("PDFumwandeln 2 ")
        l("PDFumwandeln 2 ")
        '  Using sw As New IO.StreamWriter(logfile)
        bearbeiterid = 0
        'kopfzeile
        zeile.Append("az" & t) 'Az
        zeile.Append("jahr" & t) 'jahr
        zeile.Append("datum" & t) 'datum
        zeile.Append("oberbegriff" & t) 'oberbegriff Protokolle
        zeile.Append((cleanString("bezeichnung")) & t) 'bezeichnung beschreibung
        zeile.Append(("pfad") & t) 'pfad
        zeile.Append("ordner" & t) 'ordner im mediencenter
        zeile.Append("revisionssicher" & t) ' 
        zeile.Append("bearbeiterid" & t) ' 
        csvzeileSpeichern(zeile.ToString, puAusgabeStream)
        zeile.Clear()

        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenDokumentHolen(Bearbeitungsart, relativpfad, dateinameext, typ, newsavemode, dokumentid, drr, dbdatum, istRevisionssicher,
                                         initial, eid,
                                 beschreibung, eingang, fullfilename)
                l(Bearbeitungsart & " " & CStr(dokumentid) & " " & ic & " (" & DT.Rows.Count & ")")

                'If newsavemode Then
                '    Vorhabensmerkmal = GetInputfilename(inndir, sachgebiet, CInt(ort))
                'Else
                '    Vorhabensmerkmal = GetInputfile1Name(inndir, sachgebiet, Verfahrensart)
                'End If
                If fullfilename = String.Empty Then Continue For
                TextBox3.Text = igesamt & " von " & DT.Rows.Count & "   [maxobj4test: " & maxobj & " ]"
                Application.DoEvents()
                'zeilebilden
                zeile.Append(Bearbeitungsart & t) 'Az
                zeile.Append(eingang.ToString("yyyy") & t) 'jahr
                zeile.Append(dbdatum.ToString("yyyyMMdd") & t) 'datum
                zeile.Append(typ & t) 'oberbegriff Protokolle
                zeile.Append((cleanString(beschreibung)) & t) 'bezeichnung beschreibung
                zeile.Append((fullfilename) & t) 'pfad
                zeile.Append(cleanString(dateinameext).Trim & t) 'ordner im mediencenter
                zeile.Append(CInt(istRevisionssicher) & t) 'ordner im mediencenter
                zeile.Append(CInt(bearbeiterid) & t) 'ordner im mediencenter

                'If iblock < blockMAX Then
                '    block.AppendLine(zeileAntragsteller.ToString)
                '    zeileAntragsteller.Clear()
                '    iblock += 1
                'Else
                If csvzeileSpeichern(zeile.ToString, puAusgabeStream) Then
                    'iblock = 0
                    'block.Clear()
                    zeile.Clear()
                Else
                    Debug.Print("oooo")
                End If


                'zeileAntragsteller.Clear()
                idok += 1
                If idok > maxobj Then Exit For
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       inputfile & Environment.NewLine &
                       Bearbeitungsart & "/" & dokumentid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()
        Next
        csvzeileSpeichern(zeile.ToString, puAusgabeStream)
        zeile.Clear()

        swfehlt.WriteLine(idok & "Teil2 fertig  -------" & Now.ToString & "-------------- " & igesamt)

        '####
        'swfehlt.Close()
        l("fertig  " & puFehler)
        ' Process.Start(puFehler)
    End Sub

    Private Function cleanString(Text As String) As String
        Try
            If Text Is Nothing Then Text = " "
            Text = Text.Replace(Chr(34), " ")
            Text = Text.Replace(";", "_")
            Text = Text.Replace(vbCrLf, "")
            Text = clsString.noWhiteSpace(Text, " ")
            Return Text.Trim
        Catch ex As Exception
            Return "clean_error"
        End Try
    End Function

    Private Function csvzeileSpeichern(zeile As String, puausgabestream As IO.StreamWriter) As Boolean
        Try
            puausgabestream.WriteLine(zeile)
            Return True
        Catch ex As Exception
            Debug.Print("error " & ex.ToString)
            Return False
        End Try
    End Function
    ' Funktion zum Escapen von CSV-Feldern
    Function CSVFormat(text As String) As String
        If text Is Nothing Then text = ""
        If text.Contains(";") OrElse text.Contains("""") OrElse text.Contains(vbCrLf) Then
            ' Doppelte Anführungszeichen verdoppeln
            text = text.Replace("""", """""")

            ' In doppelte Anführungszeichen einschließen
            text = $"""{text}"""
        End If
        Return text
    End Function

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        'probaugstammdaten
        Dim puFehler As String = "\\file-paradigma\paradigma\test\thumbnails\PU_Stammdaten" & Environment.UserName & ".txt"
        Dim puAusgabe As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\" & "Stammdaten" & ".csv"
        Dim puAusgabeStream As New IO.StreamWriter(puAusgabe)
        swfehlt = New IO.StreamWriter(puFehler)
        swfehlt.AutoFlush = True
        Dim Sql As String
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)
        Sql = "SELECT * FROM [Paradigma].[dbo].[stammdaten_tutti]  order by vorgangsid desc "
        TextBox1.Text = puAusgabe
        TextBox2.Text = Sql
        'writeDokumentePU(puFehler, ausgabeAntragsteller, Sql, maxobj)
        writeStammdatenPU(puFehler, puAusgabeStream, Sql, maxobj)
        puAusgabeStream.Close()
        puAusgabeStream.Dispose()
        System.Diagnostics.Process.Start("explorer", "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung")
    End Sub

    Private Sub writeStammdatenPU(puFehler As String, puAusgabeStream As IO.StreamWriter, sql As String, maxobj As Integer)
        Dim DT As DataTable
        Dim idok As Integer = 0
        puAusgabeStream.AutoFlush = True
        swfehlt.WriteLine("writeStammdatenPU---")
        DT = alleDokumentDatenHolen(sql)

        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim sachgebiet, Verfahrensart, Vorhaben, Bezeichnung, Vorhabensmerkmal, Notiz, sachbearbeiter As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum, hauptaktenjahr As Date
        Dim Hauptaktenzeichen As String
        Dim beschreibung As String
        Dim eingang, antrag, vollstaendig, bescheid, abgeschlossen As Date
        Dim aktenstandort As String
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        myoracle.Open()
        Dim zeile As New Text.StringBuilder
        Dim fullfilename As String
        Dim t As String = ";"
        Dim geschlossen As String = "0"
        l("stammdaten")

        'kopfzeile
        zeile.Append("az" & t) '                           (VID)
        zeile.Append("jahr" & t) '                       (datum)
        zeile.Append("Fachschale" & t) '                  PROUMWELT
        zeile.Append("ort" & t) '                 (sgtext + / + paragraf + / + vorgangsgegenstand + ) Überwachung einer Kleinkläranlage
        zeile.Append("Eingangsdatum" & t) '              (datum)
        zeile.Append("Antragsdatum" & t) '               (aufnahme)
        zeile.Append("Datum Vollständigkeit geprüft" & t) '(letztebearbeitung)
        zeile.Append("Bescheid Datum" & t) '
        zeile.Append("Datum Abgeschlossen am " & t) '   (letztebearbeitung falls erledigt=1)
        zeile.Append("Kürzel des Aktenstandorts" & t) ' (storaumnr)	3.b.11 oder  mu
        zeile.Append("Sachgebiet" & t) '                (sachgebietnr)
        zeile.Append("Verfahrensart" & t) '
        zeile.Append("Vorhaben" & t) '
        zeile.Append("Vorhabensmerkmal" & t) '
        zeile.Append("Zust. Sachbearbeiter" & t) '      (bearbeiter) schu
        zeile.Append("Objektnummer" & t) '
        zeile.Append("Hauptaktenzeichen" & t) ' (Hauptaktenzeichen) 
        zeile.Append("Hauptaktenjahr" & t) '
        zeile.Append("Notiz" & t) '             (az2 +  altaz + internenr + beschreibung)
        'zeileAntragsteller.Append("geschlossen" & t) '  


        csvzeileSpeichern(zeile.ToString, puAusgabeStream) : zeile.Clear()

        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1

                Bearbeitungsart = CStr(clsDBtools.fieldvalue(drr.Item("VORGANGSID")))
                eingang = CDate(clsDBtools.fieldvalueDate(drr.Item("datum")))
                Bezeichnung = cleanString(makeStammBezeichnung(CStr(clsDBtools.fieldvalue(drr.Item("SACHGEBIETSTEXT"))), CStr(clsDBtools.fieldvalue(drr.Item("PARAGRAF"))),
                                                               CStr(clsDBtools.fieldvalue(drr.Item("VORGANGSGEGENSTAND")))))
                eingang = CDate(clsDBtools.fieldvalueDate(drr.Item("datum")))
                antrag = CDate(clsDBtools.fieldvalueDate(drr.Item("aufnahme")))
                vollstaendig = CDate(clsDBtools.fieldvalueDate(drr.Item("LETZTEBEARBEITUNG")))
                bescheid = CDate(clsDBtools.fieldvalueDate(drr.Item("aufnahme")))
                abgeschlossen = makeAbgeschlossen(CDate(clsDBtools.fieldvalueDate(drr.Item("LETZTEBEARBEITUNG"))),
                                                  CStr(clsDBtools.fieldvalue(drr.Item("erledigt")))) 'falls erledigt
                aktenstandort = CStr(clsDBtools.fieldvalue(drr.Item("STORAUMNR")))
                sachgebiet = cleanString(makeSachgebiet(CStr(clsDBtools.fieldvalue(drr.Item("SACHGEBIETNR"))), CStr(clsDBtools.fieldvalue(drr.Item("SACHGEBIETSTEXT"))), 1))
                Verfahrensart = cleanString(makeSachgebiet(CStr(clsDBtools.fieldvalue(drr.Item("SACHGEBIETNR"))), CStr(clsDBtools.fieldvalue(drr.Item("SACHGEBIETSTEXT"))), 2))
                Vorhaben = cleanString(makeSachgebiet(CStr(clsDBtools.fieldvalue(drr.Item("SACHGEBIETNR"))), CStr(clsDBtools.fieldvalue(drr.Item("SACHGEBIETSTEXT"))), 3))
                Vorhabensmerkmal = cleanString(makeSachgebiet(CStr(clsDBtools.fieldvalue(drr.Item("SACHGEBIETNR"))), CStr(clsDBtools.fieldvalue(drr.Item("SACHGEBIETSTEXT"))), 4))
                sachbearbeiter = CStr(clsDBtools.fieldvalue(drr.Item("bearbeiter")))
                Hauptaktenzeichen = CStr(clsDBtools.fieldvalue(drr.Item("probaugaz")))
                hauptaktenjahr = CDate(clsDBtools.fieldvalueDate(drr.Item("datum")))
                geschlossen = CStr(clsDBtools.fieldvalue(drr.Item("erledigt")))

                Notiz = cleanString(makeStammNotiz(CStr(clsDBtools.fieldvalue(drr.Item("az2"))),
                                                   CStr(clsDBtools.fieldvalue(drr.Item("altaz"))),
                                                   CStr(clsDBtools.fieldvalue(drr.Item("internenr"))),
                                                   CStr(clsDBtools.fieldvalue(drr.Item("beschreibung")))))


                TextBox3.Text = igesamt & " von " & DT.Rows.Count & "   [maxobj4test: " & maxobj & " ]"
                Application.DoEvents()
                'zeilebilden
                zeile.Append(Bearbeitungsart & t) 'Az
                zeile.Append(eingang.ToString("yyyy") & t) 'jahr
                zeile.Append("PROUMWELT" & t) '
                zeile.Append(Bezeichnung & t) ' 
                zeile.Append(eingang.ToString("yyyyMMdd") & t) 'datum
                zeile.Append(antrag.ToString("yyyyMMdd") & t) 'datum
                zeile.Append(vollstaendig.ToString("yyyyMMdd") & t) 'datum
                zeile.Append(bescheid.ToString("yyyyMMdd") & t) 'datum
                zeile.Append(abgeschlossen.ToString("yyyyMMdd") & t) 'datum
                zeile.Append(aktenstandort & t) ' 
                zeile.Append(sachgebiet & t) ' 
                zeile.Append(Verfahrensart & t) ' 
                zeile.Append(Vorhaben & t) ' 
                zeile.Append(Vorhabensmerkmal & t) ' 
                zeile.Append(sachbearbeiter & t) ' 
                zeile.Append("" & t) ' 
                zeile.Append((cleanString(Hauptaktenzeichen)) & t) '   
                zeile.Append(hauptaktenjahr.ToString("yyyyMMdd") & t) 'datum 
                zeile.Append(cleanString(Notiz) & t) ' 
                'zeileAntragsteller.Append(geschlossen)
                'If iblock < blockMAX Then
                '    block.AppendLine(zeileAntragsteller.ToString)
                '    zeileAntragsteller.Clear()
                '    iblock += 1
                'Else
                If csvzeileSpeichern(zeile.ToString, puAusgabeStream) Then
                    'iblock = 0
                    'block.Clear()
                    zeile.Clear()
                Else
                    Debug.Print("oooo")
                End If


                'zeileAntragsteller.Clear()
                idok += 1
                If idok > maxobj Then Exit For
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       Vorhabensmerkmal & Environment.NewLine &
                       Bearbeitungsart & "/" & Bezeichnung & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()
        Next
        'csvzeileSpeichern(zeileAntragsteller.ToString, ausgabeAntragsteller)
        zeile.Clear()

        swfehlt.WriteLine(idok & "Teil2 fertig  -------" & Now.ToString & "-------------- " & igesamt)

        '####
        'swfehlt.Close()
        l("fertig  " & puFehler)
        ' Process.Start(puFehler)
    End Sub

    Private Function makeAbgeschlossen([date] As Date, erledigt As String) As Date
        Try
            If erledigt = 1 Then
                Return [date]
            Else
                Return Nothing
            End If
        Catch ex As Exception
            l("fehler  " & ex.ToString)
            Return "error"
        End Try
    End Function

    Private Function makeSachgebiet(sachgebietnr As String, sachgebiettext As String, modus As Integer) As String
        Dim t As Char = ","
        Dim a() As String
        Try
            sachgebietnr = Trim(sachgebietnr)
            sachgebiettext = Trim(sachgebiettext)
            sachgebiettext = sachgebiettext.Replace(sachgebietnr, "")
            If sachgebietnr.Contains("-") Then '
                'a = sachgebietnr.Split("-"c)
                If modus = "1" Then
                    If sachgebietnr.Substring(0, 1) Then
                        Return "4-Wasser und Bodenschutz"
                    End If
                Else
                    Return sachgebietnr & " - " & sachgebiettext
                End If
            End If
            If sachgebietnr.Count < 4 Then
                Return sachgebietnr & " - " & sachgebiettext
            End If
            If sachgebietnr.Count = 4 Then
                'normalfall
                If modus = "1" Then
                    Select Case sachgebietnr.Substring(0, 1)
                        Case "1"
                            Return "1-FD Umwelt Allgemein"
                        Case "2"
                            Return "2-Grafische Datenverarbeitung"
                        Case "3"
                            Return "3-Naturschutz"
                        Case "4"
                            Return "4-Wasser und Bodenschutz"
                        Case "5"
                            Return "5-Immissionsschutz"
                        Case "6"
                            Return "6-Umweltschutz"
                        Case "7"
                            Return "7-Abfallwirtschaft"
                        Case Else
                            Return sachgebietnr
                    End Select
                    Return sachgebietnr & " - " & sachgebiettext
                End If
                If modus = "2" Then
                    Return sachgebietnr.Substring(1, 3) & " - " & sachgebiettext
                End If
                If modus = "3" Then
                    Return sachgebietnr.Substring(2, 2) & " - " & sachgebiettext
                End If
                If modus = "4" Then
                    Return sachgebietnr.Substring(3, 1) & " - " & sachgebiettext
                End If
            End If
            l("oooo")
        Catch ex As Exception
            l("fehler  " & ex.ToString)
            Return "error"
        End Try
    End Function

    Private Function makeStammNotiz(az As String, altaz As String, interne As String, beschreibung As String) As String
        Dim t As Char = ","
        Dim ret As String = ""
        Try
            az = az.Trim
            altaz = altaz.Trim
            interne = interne.Trim
            beschreibung = beschreibung.Trim
            If altaz.Length < 1 Then
                altaz = ""
            Else
                altaz = ", AltAz: " & altaz
            End If
            If interne.Length < 1 Then
                interne = ""
            Else
                interne = ", IntNr: " & interne
            End If
            Return az & altaz & interne & t & " " & beschreibung

        Catch ex As Exception
            l("fehler  " & ex.ToString)
            Return "error"
        End Try
    End Function

    Private Function makeStammBezeichnung(sgtext As String, paragraf As String, vgGegenstand As String) As String
        Dim t As Char = ","
        Dim ret As String = ""
        Try
            sgtext = Trim(sgtext)
            paragraf = Trim(paragraf)
            vgGegenstand = Trim(vgGegenstand)
            If paragraf.Length < 1 Then
                paragraf = ""
            Else
                paragraf = t & " $: " & paragraf
            End If
            ret = sgtext & paragraf & t & " " & vgGegenstand
            Return ret
        Catch ex As Exception
            l("fehler  " & ex.ToString)
            Return "error"
        End Try
    End Function

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        'adresse
        Dim puFehler As String = "\\file-paradigma\paradigma\test\thumbnails\PU_LageAdresse" & Environment.UserName & ".txt"
        Dim puAusgabe As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\" & "PU_LageAdresse" & ".csv"
        Dim puAusgabeStream As New IO.StreamWriter(puAusgabe)
        swfehlt = New IO.StreamWriter(puFehler)
        swfehlt.AutoFlush = True
        Dim Sql As String
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)
        Sql = "SELECT  *  FROM [Paradigma].[dbo].[PA_SEKID2VID] c inner join [Paradigma].[dbo].PARAADRESSE a  " &
               "     ON c.SEKID = a.ID  " &
               "     order by VORGANGSID desc "
        Sql = "select * from  [Paradigma].[dbo].[PA_mitRH]      order by VORGANGSID desc "
        Sql = "select * from Pa_mitRH as a, raumbezug as r " &
                "where a.id=r.sekid  and typ=1	"



        TextBox1.Text = puAusgabe
        TextBox2.Text = Sql
        writeAdresseausgabePU(puFehler, puAusgabeStream, Sql, maxobj)
        puAusgabeStream.Close()
        puAusgabeStream.Dispose()
        swfehlt.Dispose()
        System.Diagnostics.Process.Start("explorer", "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung")
    End Sub

    Private Sub writeAdresseausgabePU(puFehler As String, puAusgabeStream As IO.StreamWriter, sql As String, maxobj As Integer)
        Dim DT As DataTable
        Dim idok As Integer = 0
        puAusgabeStream.AutoFlush = True
        swfehlt.WriteLine("writeStammdatenPU---")
        DT = alleDokumentDatenHolen(sql)

        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim ortsteil, strasse, gemeindenr, ort, nummer, freitext, funktion, strcode, ost, nord, abstrakt, fs As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum, hauptaktenjahr As Date
        Dim spalte1 As String
        Dim veraltet As String
        Dim eingang, antrag, vollstaendig, bescheid, abgeschlossen As Date
        Dim aktenstandort As String
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        myoracle.Open()
        Dim zeile As New Text.StringBuilder
        Dim fullfilename As String
        Dim t As String = ";"
        Dim geschlossen As String = "0"
        l("stammdaten")

        'kopfzeile
        zeile.Append("az" & t) '       
        zeile.Append("jahr" & t) '     
        zeile.Append("Ort" & t) '      
        zeile.Append("Ortsteil" & t) ' 
        zeile.Append("Strasse" & t) '  
        zeile.Append("nr" & t) '       
        zeile.Append("veraltet" & t) '
        zeile.Append("spalte1" & t) '
        zeile.Append("gemeindenr" & t)
        zeile.Append("strassencode" & t) ' 
        zeile.Append("rechtswert" & t)
        zeile.Append("hochwert" & t) '
        zeile.Append("initial_1" & t) '
        zeile.Append("abteilung" & t) '
        zeile.Append("abstract" & t) 'telefon
        zeile.Append("initial_2" & t) 'fs
        csvzeileSpeichern(zeile.ToString, puAusgabeStream) : zeile.Clear()
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                Bearbeitungsart = CStr(clsDBtools.fieldvalue(drr.Item("VORGANGSID")))
                eingang = CDate(clsDBtools.fieldvalueDate(drr.Item("datum")))
                ort = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("gemeindetext"))))
                ortsteil = ""
                strasse = (clsDBtools.fieldvalue(drr.Item("strassenname")))
                nummer = CStr(clsDBtools.fieldvalue(drr.Item("hausnrkombi")))
                veraltet = ""
                spalte1 = ""
                gemeindenr = CStr(clsDBtools.fieldvalue(drr.Item("gemeindenr")))
                strcode = CStr(clsDBtools.fieldvalue(drr.Item("strcode")))
                ost = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("rechts"))))
                nord = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("hoch"))))
                funktion = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("titel"))))
                freitext = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("abteilung"))))
                abstrakt = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("abstract"))))
                fs = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("fs"))))

                TextBox3.Text = igesamt & " von " & DT.Rows.Count & "   [maxobj4test: " & maxobj & " ]"
                Application.DoEvents()
                'zeilebilden
                zeile.Append(Bearbeitungsart & t) 'Az
                zeile.Append(eingang.ToString("yyyy") & t) 'jahr 
                zeile.Append(ort & t) ' 
                zeile.Append(ortsteil & t) ' 
                zeile.Append(strasse & t) ' 
                zeile.Append(nummer & t) ' 
                zeile.Append(veraltet & t) ' 
                zeile.Append(spalte1 & t) '    
                zeile.Append(gemeindenr & t) '    
                zeile.Append(strcode & t) '    
                zeile.Append(ost & t) ' 
                zeile.Append(nord & t) '  
                zeile.Append(funktion & t) ' 
                zeile.Append(freitext & t) ' 
                zeile.Append(abstrakt & t) '     
                zeile.Append(fs & t) ' 

                If csvzeileSpeichern(zeile.ToString, puAusgabeStream) Then

                    zeile.Clear()
                Else
                    Debug.Print("oooo")
                End If
                idok += 1
                If idok > maxobj Then Exit For
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                          Environment.NewLine &
                       Bearbeitungsart & "/" & ort & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()
        Next
        'csvzeileSpeichern(zeileAntragsteller.ToString, ausgabeAntragsteller)
        zeile.Clear()

        swfehlt.WriteLine(idok & "Teil2 fertig  -------" & Now.ToString & "-------------- " & igesamt)

        '####
        'swfehlt.Close()
        l("fertig  " & puFehler)
        ' Process.Start(puFehler)
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        'kataster 
        Dim puFehler As String = "\\file-paradigma\paradigma\test\thumbnails\PU_kataster" & Environment.UserName & ".txt"
        Dim puAusgabe As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\" & "PU_kataster" & ".csv"
        Dim puAusgabeStream As New IO.StreamWriter(puAusgabe)
        swfehlt = New IO.StreamWriter(puFehler)
        swfehlt.AutoFlush = True
        Dim Sql As String
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)

        Sql = "select * from PF_mitRH as a, raumbezug as r  where a.id=r.sekid and typ=2      order by Bearbeitungsart desc  "
        TextBox1.Text = puAusgabe
        TextBox2.Text = Sql
        'writeAdresseausgabePU(puFehler, ausgabeAntragsteller, Sql, maxobj)
        writeKatasterausgabePU(puFehler, puAusgabeStream, Sql, maxobj)
        puAusgabeStream.Close()
        puAusgabeStream.Dispose()
        System.Diagnostics.Process.Start("explorer", "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung")
        swfehlt.Dispose()
    End Sub

    Private Sub writeKatasterausgabePU(puFehler As String, puAusgabeStream As IO.StreamWriter, sql As String, maxobj As Integer)
        Dim DT As DataTable
        Dim idok As Integer = 0
        puAusgabeStream.AutoFlush = True
        swfehlt.WriteLine("writeKatasterausgabePU---")
        DT = alleDokumentDatenHolen(sql)

        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim gemarkung, flur, flurstueck, ost, nord, freitext, funktion, gemcode, abstrakt, FS, flaecheqm As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum, hauptaktenjahr As Date
        Dim spalte1 As String
        Dim veraltet As String
        Dim eingang, antrag, vollstaendig, bescheid, abgeschlossen As Date
        Dim aktenstandort As String
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        myoracle.Open()
        Dim zeile As New Text.StringBuilder
        Dim fullfilename As String
        Dim t As String = ";"
        Dim geschlossen As String = "0"
        l("writeKatasterausgabePU")

        'kopfzeile
        zeile.Append("az" & t) '     Bearbeitungsart   
        zeile.Append("jahr" & t) '     datum
        zeile.Append("Gemarkung" & t) '  gemarkungstext    
        zeile.Append("nachname" & t) '  nachname
        zeile.Append("vorname" & t) '  znkombi
        zeile.Append("ostwert" & t) '   rechts    
        zeile.Append("nordwert" & t) ' hoch
        zeile.Append("veraltet" & t) '
        zeile.Append("spalte1" & t)
        zeile.Append("spalte2" & t) ' 
        zeile.Append("initial_1" & t) 'titel
        zeile.Append("abteilung" & t) 'abteilung
        zeile.Append("abstract" & t) 'telefon
        zeile.Append("gemarkungscode" & t) 'fax
        zeile.Append("initial_2" & t) 'fs
        zeile.Append("flaeche" & t) 'flaecheqm
        'zeileAntragsteller.Append("Funktion" & t) ' 
        'zeileAntragsteller.Append("Freitext" & t) ' 
        csvzeileSpeichern(zeile.ToString, puAusgabeStream) : zeile.Clear()
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                Bearbeitungsart = CStr(clsDBtools.fieldvalue(drr.Item("VORGANGSID")))
                eingang = CDate(clsDBtools.fieldvalueDate(drr.Item("datum")))
                gemarkung = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("gemarkungstext"))))
                flur = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("nachname"))))
                flurstueck = cleanString((clsDBtools.fieldvalue(drr.Item("znkombi"))))
                ost = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("rechts"))))
                nord = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("hoch"))))

                spalte1 = ""
                spalte1 = ""
                funktion = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("titel"))))
                freitext = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("abteilung"))))
                abstrakt = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("abstract"))))
                gemcode = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("fax"))))
                FS = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("fs"))))
                flaecheqm = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("flaecheqm"))))
                'abteilung = CDate(clsDBtools.fieldvalueDate(drr.Item("abteilung")))
                'initial_1 = CDate(clsDBtools.fieldvalueDate(drr.Item("initial_1"))) 
                TextBox3.Text = igesamt & " von " & DT.Rows.Count & "   [maxobj4test: " & maxobj & " ]"
                Application.DoEvents()
                'zeilebilden
                zeile.Append(Bearbeitungsart & t) 'Az
                zeile.Append(eingang.ToString("yyyy") & t) 'jahr 
                zeile.Append(gemarkung & t) ' 
                zeile.Append(flur & t) ' 
                zeile.Append(flurstueck & t) ' 
                zeile.Append(ost & t) ' 
                zeile.Append(nord & t) ' 
                zeile.Append(spalte1 & t) ' veraltet
                zeile.Append(spalte1 & t) ' 
                zeile.Append(spalte1 & t) ' 
                zeile.Append(funktion & t) ' 
                zeile.Append(freitext & t) ' 
                zeile.Append(abstrakt & t) '    
                zeile.Append(gemcode & t) '    
                zeile.Append(FS & t) '    
                zeile.Append(flaecheqm & t) '    

                If csvzeileSpeichern(zeile.ToString, puAusgabeStream) Then

                    zeile.Clear()
                Else
                    Debug.Print("oooo")
                End If
                idok += 1
                If idok > maxobj Then Exit For
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                          Environment.NewLine &
                       Bearbeitungsart & "/" & Bearbeitungsart & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()
        Next
        'csvzeileSpeichern(zeileAntragsteller.ToString, ausgabeAntragsteller)
        zeile.Clear()

        swfehlt.WriteLine(idok & "Teil2 fertig  -------" & Now.ToString & "-------------- " & igesamt)

        '####
        'swfehlt.Close()
        l("fertig  " & puFehler)
        ' Process.Start(puFehler)

    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        'beteiligte 2_stakeholder 
        Dim puFehler As String = "\\file-paradigma\paradigma\test\thumbnails\PU_beteiligte" & Environment.UserName & ".txt"
        Dim puAusgabe As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\" & "beteiligte_2_stakeholder" & ".csv"
        Dim puAusgabeStream As New IO.StreamWriter(puAusgabe)
        swfehlt = New IO.StreamWriter(puFehler)
        swfehlt.AutoFlush = True
        Dim Sql As String
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)
        Sql = "select * FROM [Paradigma].[dbo].[STAKEHOLDER]    order by rolle desc  "
        TextBox1.Text = puAusgabe
        TextBox2.Text = Sql
        'writeKatasterausgabePU(puFehler, puAusgabeStream, Sql, maxobj)
        'writeAntragstellerausgabePU()
        writeStakeholderAusgabePU(puFehler, puAusgabeStream, Sql, maxobj)
        puAusgabeStream.Close()
        puAusgabeStream.Dispose()
        System.Diagnostics.Process.Start("explorer", "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung")
        swfehlt.Dispose()
    End Sub
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        'wiedervorlagen
        Dim puFehler As String = "\\file-paradigma\paradigma\test\thumbnails\PU_wiedervorlagen" & Environment.UserName & ".txt"
        Dim puAusgabe As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\" & "PU_wiedervorlagen" & ".csv"
        Dim puAusgabeStream As New IO.StreamWriter(puAusgabe)
        swfehlt = New IO.StreamWriter(puFehler)
        swfehlt.AutoFlush = True
        Dim Sql As String
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)
        Sql = "select DATUM,TODO,w.BEMERKUNG,w.TS,w.VORGANGSID,w.ERLEDIGTAM,w.ERLEDIGT,w.BEARBEITER,w.WARTENAUF,w.BEARBEITERID," &
                    "t.EINGANG as teingang " &
              "From [Paradigma].[dbo].[WIEDERVORLAGE_T45] as w, stammdaten_tutti As t " &
              "where w.VORGANGSID=t.VORGANGSID   order by w.VORGANGSID desc"
        TextBox1.Text = puAusgabe
        TextBox2.Text = Sql
        'writeStakeholderAusgabePU(puFehler, puAusgabeStream, Sql, maxobj)
        writeWiedervorlageAusgabePU(puFehler, puAusgabeStream, Sql, maxobj)
        puAusgabeStream.Close()
        puAusgabeStream.Dispose()
        System.Diagnostics.Process.Start("explorer", "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung")
        swfehlt.Dispose()
    End Sub
    Private Sub writeWiedervorlageAusgabePU(puFehler As String, puAusgabeStream As IO.StreamWriter, sql As String, maxobj As Integer)
        Dim DT As DataTable
        Dim idok As Integer = 0
        puAusgabeStream.AutoFlush = True
        swfehlt.WriteLine("Wiedervorlage---")
        DT = alleDokumentDatenHolen(sql)

        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim gemarkung, flur, flurstueck, ost, info, funktion, gemcode, bearbeiterid, FS, Betreff As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum, hauptaktenjahr As Date
        Dim spalte1 As String
        Dim vid As String
        Dim datum, antrag, vollstaendig, eingang As Date
        Dim sachbearbeiter As String
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        myoracle.Open()
        Dim zeile As New Text.StringBuilder
        Dim erledigtam As Date
        Dim erledigt As Boolean
        Dim t As String = ";"
        Dim geschlossen As String = "0"
        l("Wiedervorlage")

        'kopfzeile
        zeile.Append("az" & t) '     Bearbeitungsart   
        zeile.Append("jahr" & t) '     datum
        zeile.Append("Bearbeitungsart" & t) '  gemarkungstext    
        zeile.Append("datum" & t) '  nachname
        zeile.Append("sachbearbeiter" & t) '  znkombi
        zeile.Append("Betreff" & t) '   rechts    
        zeile.Append("info" & t) ' hoch

        zeile.Append("bearbeiterid" & t) '  
        zeile.Append("erledigtam" & t) '  
        zeile.Append("erledigt" & t) '  
        'zeile.Append("angelegtam" & t) '  
        csvzeileSpeichern(zeile.ToString, puAusgabeStream) : zeile.Clear()
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                vid = CStr(clsDBtools.fieldvalue(drr.Item("vorgangsid")))
                eingang = cleanString(CStr(clsDBtools.fieldvalueDate(drr.Item("teingang"))))
                Bearbeitungsart = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("TODO"))))
                datum = CDate(clsDBtools.fieldvalueDate(drr.Item("datum")))  '==FÄLLIG
                sachbearbeiter = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("bearbeiter"))))
                Betreff = "Wartenauf: " & cleanString(CStr(clsDBtools.fieldvalue(drr.Item("wartenauf"))))
                info = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Bemerkung"))))

                bearbeiterid = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("bearbeiterid"))))
                erledigtam = cleanString(CStr(clsDBtools.fieldvalueDate(drr.Item("erledigtam"))))
                erledigt = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("erledigt"))))
                'angelegtam = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("weingang"))))

                TextBox3.Text = igesamt & " von " & DT.Rows.Count & "   [maxobj4test: " & maxobj & " ]"
                Application.DoEvents()
                'zeilebilden
                zeile.Append(vid & t) 'Az
                zeile.Append(eingang.ToString("yyyy") & t) 'jahr 
                zeile.Append(Bearbeitungsart & t) 'Az 
                zeile.Append(datum & t) ' 
                zeile.Append(sachbearbeiter & t) ' 
                zeile.Append(Betreff & t) ' 
                zeile.Append(info & t) ' 


                zeile.Append(bearbeiterid & t) ' 
                zeile.Append(erledigtam.ToString("yyyy") & t) 'jahr 
                zeile.Append(erledigt & t) ' 
                'zeile.Append(angelegtam & t) ' veraltet


                If csvzeileSpeichern(zeile.ToString, puAusgabeStream) Then
                    zeile.Clear()
                Else
                    Debug.Print("oooo")
                End If
                idok += 1
                If idok > maxobj Then Exit For
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                          Environment.NewLine &
                       Bearbeitungsart & "/" & Bearbeitungsart & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()
        Next
        'csvzeileSpeichern(zeileAntragsteller.ToString, ausgabeAntragsteller)
        zeile.Clear()

        swfehlt.WriteLine(idok & "Teil2 fertig  -------" & Now.ToString & "-------------- " & igesamt)

        '####
        'swfehlt.Close()
        l("fertig  " & puFehler)
        ' Process.Start(puFehler)
    End Sub

    Private Sub writeStakeholderAusgabePU(puFehler As String, puAusgabeStream As IO.StreamWriter, sql As String, maxobj As Integer)
        Dim DT As DataTable
        Dim idok As Integer = 0
        swfehlt.WriteLine("writeStakeholderAusgabePU---")
        DT = alleDokumentDatenHolen(sql)
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim eingang As Date
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        myoracle.Open()
        Dim zeileBeteiligte As New Text.StringBuilder
        Dim t As String = ";"
        Dim geschlossen As String = "0"
        l("writeKatasterausgabePU")
        Dim perstemp As New person
        Dim perscoll As New List(Of person)
        'kopfzeile
        bildeKopfZeileStakeholder(zeileBeteiligte, t)
        csvzeileSpeichern(zeileBeteiligte.ToString, puAusgabeStream) : zeileBeteiligte.Clear()
        Dim beteiligter As New person
        Try
            beteiligter = New person
            igesamt += 1
            TextBox3.Text = igesamt & " von " & DT.Rows.Count & "   [maxobj4test: " & maxobj & " ]" : Application.DoEvents()
            Bearbeitungsart = 0 'CStr(clsDBtools.fieldvalue(drr.Item("VORGANGSID")))
            eingang = CDate("1911-01-01") ' ""CStr(clsDBtools.fieldvalueDate(drr.Item("datum")))'""
            perscoll = getAllStakeholders(perstemp, DT)

            For Each perso As person In perscoll
                zeileBeteiligte = bildeZeilePerson(eingang, t, perso)
                If csvzeileSpeichern(zeileBeteiligte.ToString, puAusgabeStream) Then
                    zeileBeteiligte.Clear()
                End If
                igesamt += 1
                TextBox3.Text = igesamt & " von " & DT.Rows.Count & "   [maxobj4test: " & maxobj & " ]" : Application.DoEvents()
            Next
        Catch ex As Exception
            l("ddd" & ex.ToString)
        End Try
        'Next
    End Sub

    Private Function getAllStakeholders(perstemp As person, dt As DataTable) As List(Of person)
        Dim perlist As New List(Of person)
        Dim per As New person
        Try

            For Each drr As DataRow In dt.Rows
                per = New person
                per.Rolle = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("rolle"))))
                per.Name = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("nachname"))))
                per.Vorname = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("vorname"))))
                per.Bemerkung = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("bemerkung"))))
                per.Namenszusatz = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Namenszusatz"))))
                per.Anrede = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Anrede"))))
                per.Kontakt.Anschrift.Gemeindename = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Gemeindename"))))
                per.Kontakt.Anschrift.Strasse = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Strasse"))))
                per.Kontakt.Anschrift.Hausnr = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Hausnr"))))
                per.Kontakt.Anschrift.Postfach = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Postfach"))))
                per.Kontakt.Anschrift.PLZ = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("plz"))))
                per.Kontakt.Org.Name = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgname"))))
                per.Kontakt.Org.Zusatz = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgzusatz"))))
                per.Kontakt.Org.Typ1 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgtyp1"))))
                per.Kontakt.Org.Typ2 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgtyp2"))))
                per.Kontakt.Org.Eigentuemer = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgEigentuemer"))))
                per.Quelle = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("quelle"))))
                per.Kontakt.Org.Bemerkung = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgBemerkung"))))
                per.Kontakt.GesellFunktion = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("GesellFunktion"))))
                per.Kontakt.elektr.Telefon1 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("fftelefon1"))))
                per.Kontakt.elektr.Telefon2 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("fftelefon2"))))
                per.Kontakt.elektr.Fax1 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("FFFax1"))))
                per.Kontakt.elektr.Fax2 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("FFFax2"))))
                per.Kontakt.elektr.MobilFon = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("FFMOBILFON"))))
                per.Kontakt.elektr.Email = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("FFemail"))))
                per.Kontakt.elektr.Homepage = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("FFhomepage"))))

                per.Kassenkonto = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Kassenkonto"))))
                per.VERTRETENDURCH = "" 'cleanString(CStr(clsDBtools.fieldvalue(drr.Item("VERTRETENDURCH"))))

                'für die sepa daten muss man über die personenid holen

                per.Kontakt.BankkontoID = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("personenid"))))
                perlist.Add(per)
            Next
            Return perlist
        Catch ex As Exception
            'swfehlt.Close()
            l("fertig  " & ex.ToString)
            Return Nothing
        End Try
    End Function
    Private Sub bildeKopfZeileWiedervorlage(zeileBeteiligte As System.Text.StringBuilder, t As String)
        zeileBeteiligte.Append("az" & t) '     Bearbeitungsart   
        zeileBeteiligte.Append("jahr" & t) '     datum
        zeileBeteiligte.Append("Obergruppe" & t) '  rolle    
        zeileBeteiligte.Append("Anrede" & t) '   
        zeileBeteiligte.Append("Firma" & t) '  znkombi
        zeileBeteiligte.Append("Titel" & t) '   rechts    
        zeileBeteiligte.Append("Vorname" & t) ' hoch
        zeileBeteiligte.Append("nachname" & t) '
        zeileBeteiligte.Append("strasse" & t)
        zeileBeteiligte.Append("hausnr" & t) ' 
        zeileBeteiligte.Append("hausnr bis" & t) 'titel
        zeileBeteiligte.Append("land" & t) 'abteilung
        zeileBeteiligte.Append("plz" & t) 'telefon
        zeileBeteiligte.Append("ort" & t) 'fax
        zeileBeteiligte.Append("adreszusatzzeile1" & t) 'fs
        zeileBeteiligte.Append("adreszusatzzeile2" & t) 'fs
        zeileBeteiligte.Append("telefon" & t) 'fs
        zeileBeteiligte.Append("fax" & t) 'fs
        zeileBeteiligte.Append("mobil" & t) 'fs
        zeileBeteiligte.Append("email" & t) 'fs
        zeileBeteiligte.Append("de-mail" & t) 'fs
        zeileBeteiligte.Append("web" & t) 'fs
        zeileBeteiligte.Append("zeichen" & t) 'fs
        zeileBeteiligte.Append("personennummer" & t) 'fs
        zeileBeteiligte.Append("spalte1" & t)
        zeileBeteiligte.Append("KASSENKONTO" & t)
    End Sub
    Private Sub bildeKopfZeileStakeholder(zeileBeteiligte As System.Text.StringBuilder, t As String)
        zeileBeteiligte.Append("az" & t) '     Bearbeitungsart   
        zeileBeteiligte.Append("jahr" & t) '     datum
        zeileBeteiligte.Append("Obergruppe" & t) '  rolle    
        zeileBeteiligte.Append("Anrede" & t) '   
        zeileBeteiligte.Append("Firma" & t) '  znkombi
        zeileBeteiligte.Append("Titel" & t) '   rechts    
        zeileBeteiligte.Append("Vorname" & t) ' hoch
        zeileBeteiligte.Append("nachname" & t) '
        zeileBeteiligte.Append("strasse" & t)
        zeileBeteiligte.Append("hausnr" & t) ' 
        zeileBeteiligte.Append("hausnr bis" & t) 'titel
        zeileBeteiligte.Append("land" & t) 'abteilung
        zeileBeteiligte.Append("plz" & t) 'telefon
        zeileBeteiligte.Append("ort" & t) 'fax
        zeileBeteiligte.Append("adreszusatzzeile1" & t) 'fs
        zeileBeteiligte.Append("adreszusatzzeile2" & t) 'fs
        zeileBeteiligte.Append("telefon" & t) 'fs
        zeileBeteiligte.Append("fax" & t) 'fs
        zeileBeteiligte.Append("mobil" & t) 'fs
        zeileBeteiligte.Append("email" & t) 'fs
        zeileBeteiligte.Append("de-mail" & t) 'fs
        zeileBeteiligte.Append("web" & t) 'fs
        zeileBeteiligte.Append("zeichen" & t) 'fs
        zeileBeteiligte.Append("personennummer" & t) 'fs
        zeileBeteiligte.Append("spalte1" & t)
        zeileBeteiligte.Append("KASSENKONTO" & t)
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        'antragsteller
        Dim puFehler As String = "\\file-paradigma\paradigma\test\thumbnails\PU_antragsteller" & Environment.UserName & ".txt"
        Dim puAusgabe As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\" & "PU_antragsteller" & ".csv"
        Dim pubeteiligte As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\" & "PU_beteiligte_1" & ".csv"
        Dim ausgabeAntragsteller As New IO.StreamWriter(puAusgabe)
        Dim ausgabeBeteiligte As New IO.StreamWriter(pubeteiligte)
        swfehlt = New IO.StreamWriter(puFehler)
        swfehlt.AutoFlush = True
        Dim Sql As String
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)

        Sql = "select vorgangsid,datum from stammdaten_tutti     order by vorgangsid desc   "
        TextBox1.Text = puAusgabe
        TextBox2.Text = Sql
        writeAntragstellerausgabePU(puFehler, ausgabeAntragsteller, ausgabeBeteiligte, Sql, maxobj)
        ausgabeAntragsteller.Close()
        ausgabeAntragsteller.Dispose()
        ausgabeBeteiligte.Close()
        ausgabeBeteiligte.Dispose()
        swfehlt.Dispose()
        System.Diagnostics.Process.Start("explorer", "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung")
    End Sub

    Private Sub writeAntragstellerausgabePU(puFehler As String, ausgabeAntragsteller As IO.StreamWriter, ausgabeBeteiligte As IO.StreamWriter, sql As String, maxobj As Integer)
        Dim DT As DataTable
        Dim idok As Integer = 0
        ausgabeAntragsteller.AutoFlush = True
        swfehlt.WriteLine("writeKatasterausgabePU---")
        DT = alleDokumentDatenHolen(sql)

        Dim ic As Integer = 0
        Dim igesamt As Integer = 0

        Dim eingang As Date
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        myoracle.Open()
        Dim zeileAntragsteller As New Text.StringBuilder
        Dim zeileBeteiligte As New Text.StringBuilder
        Dim t As String = ";"
        Dim geschlossen As String = "0"
        l("writeKatasterausgabePU")
        Dim perstemp As New person
        Dim perscoll As New List(Of person)
        'kopfzeile
        bildeKopfZeile(zeileAntragsteller, t)
        bildeKopfZeile(zeileBeteiligte, t)
        csvzeileSpeichern(zeileAntragsteller.ToString, ausgabeAntragsteller) : zeileAntragsteller.Clear()
        csvzeileSpeichern(zeileBeteiligte.ToString, ausgabeBeteiligte) : zeileBeteiligte.Clear()
        Dim antragsteller As New person
        Dim beteiligter As New person
        For Each drr As DataRow In DT.Rows  'alle vorgänge
            Try
                antragsteller = New person
                igesamt += 1
                TextBox3.Text = igesamt & " von " & DT.Rows.Count & "   [maxobj4test: " & maxobj & " ]" : Application.DoEvents()
                Bearbeitungsart = CStr(clsDBtools.fieldvalue(drr.Item("VORGANGSID")))
                eingang = CStr(clsDBtools.fieldvalueDate(drr.Item("datum")))
                perscoll = getAllBeteiligte4vorgang(perstemp, Bearbeitungsart)
                If hatAntragsteller(perscoll) Then
                    antragsteller = getAntragsteller(perscoll)
                    If antragsteller Is Nothing Then Exit For
                    zeileAntragsteller = bildeZeilePerson(eingang, t, antragsteller)
                    'zeile Nach antragsteller ausschreiben
                    csvzeileSpeichern(zeileAntragsteller.ToString, ausgabeAntragsteller)
                    zeileAntragsteller.Clear()
                Else
                    'getErstenEintragOrDummy(persocoll)
                    antragsteller = getErstenEintrag(perscoll)
                    If antragsteller Is Nothing Then
                        antragsteller = New person
                        antragsteller.Name = "dummy"
                        antragsteller.Vorname = "dummy"
                        antragsteller.Rolle = "dummy"
                    End If
                    zeileAntragsteller = bildeZeilePerson(eingang, t, antragsteller)
                    'zeile Nach antragsteller ausschreiben
                    csvzeileSpeichern(zeileAntragsteller.ToString, ausgabeAntragsteller)
                    zeileAntragsteller.Clear()
                End If
                For Each perso As person In perscoll
                    zeileBeteiligte = bildeZeilePerson(eingang, t, perso)
                    If csvzeileSpeichern(zeileBeteiligte.ToString, ausgabeBeteiligte) Then
                        zeileBeteiligte.Clear()
                    End If
                Next
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                          Environment.NewLine &
                       Bearbeitungsart & "/" & Bearbeitungsart & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()
            idok += 1
            If idok > maxobj Then Exit For
        Next
        'csvzeileSpeichern(zeileAntragsteller.ToString, ausgabeAntragsteller)
        zeileAntragsteller.Clear()

        swfehlt.WriteLine(idok & "Teil2 fertig  -------" & Now.ToString & "-------------- " & igesamt)

        '####
        'swfehlt.Close()
        l("fertig  " & puFehler)
        ' Process.Start(puFehler)
    End Sub

    Private Function getErstenEintrag(perscoll As List(Of person)) As person
        Try
            For Each perso As person In perscoll
                If Not (perso.Rolle.ToLower.Contains("ntragsteller") Or perso.Rolle.ToLower.Contains("eschwerde")) Then
                    Return perso
                End If
            Next
            Return Nothing
        Catch ex As Exception
            l("fertig  " & ex.ToString)
            Return Nothing
        End Try
    End Function

    Private Function getAntragsteller(perscoll As List(Of person)) As person
        Try
            For Each perso As person In perscoll
                If perso.Rolle.ToLower.Contains("ntragsteller") Or perso.Rolle.ToLower.Contains("eschwerde") Then
                    Return perso
                End If
            Next
            Return Nothing
        Catch ex As Exception
            l("fertig  " & ex.ToString)
            Return Nothing
        End Try
    End Function

    Private Shared Sub bildeKopfZeile(zeileAntragsteller As System.Text.StringBuilder, t As String)
        zeileAntragsteller.Append("az" & t) '     Bearbeitungsart   
        zeileAntragsteller.Append("jahr" & t) '     datum
        zeileAntragsteller.Append("Obergruppe" & t) '  rolle    
        zeileAntragsteller.Append("Anrede" & t) '   
        zeileAntragsteller.Append("Firma" & t) '  znkombi
        zeileAntragsteller.Append("Titel" & t) '   rechts    
        zeileAntragsteller.Append("Vorname" & t) ' hoch
        zeileAntragsteller.Append("nachname" & t) '
        zeileAntragsteller.Append("strasse" & t)
        zeileAntragsteller.Append("hausnr" & t) ' 
        zeileAntragsteller.Append("hausnr bis" & t) 'titel
        zeileAntragsteller.Append("land" & t) 'abteilung
        zeileAntragsteller.Append("plz" & t) 'telefon
        zeileAntragsteller.Append("ort" & t) 'fax
        zeileAntragsteller.Append("adreszusatzzeile1" & t) 'fs
        zeileAntragsteller.Append("adreszusatzzeile2" & t) 'fs
        zeileAntragsteller.Append("telefon" & t) 'fs
        zeileAntragsteller.Append("fax" & t) 'fs
        zeileAntragsteller.Append("mobil" & t) 'fs
        zeileAntragsteller.Append("email" & t) 'fs
        zeileAntragsteller.Append("de-mail" & t) 'fs
        zeileAntragsteller.Append("web" & t) 'fs
        zeileAntragsteller.Append("zeichen" & t) 'fs
        zeileAntragsteller.Append("personennummer" & t) 'fs
        zeileAntragsteller.Append("spalte1" & t)
        zeileAntragsteller.Append("KASSENKONTO" & t)
    End Sub

    Private Function hatAntragsteller(perscoll As List(Of person)) As Boolean
        Try
            For Each perso As person In perscoll
                If perso.Rolle.ToLower.Contains("ntragsteller") Or perso.Rolle.ToLower.Contains("eschwerde") Then
                    Return True
                End If
            Next
            Return False
        Catch ex As Exception
            l("fertig  " & ex.ToString)
            Return False
        End Try
    End Function

    Private Shared Function bildeZeilePerson(eingang As Date, t As String, perso As person) As Text.StringBuilder
        Dim zeileAntragsteller As New Text.StringBuilder

        zeileAntragsteller.Clear()
        zeileAntragsteller.Append(Bearbeitungsart & t) 'Az
        zeileAntragsteller.Append(eingang.ToString("yyyy") & t) 'jahr 
        zeileAntragsteller.Append(perso.Rolle & t) ' 
        zeileAntragsteller.Append(perso.Anrede & t) ' 
        zeileAntragsteller.Append(bildeFirma(perso) & t) ' 
        zeileAntragsteller.Append(perso.Namenszusatz & t) ' 
        zeileAntragsteller.Append(perso.Vorname & t) ' 
        zeileAntragsteller.Append(perso.Name & t) ' 
        zeileAntragsteller.Append(perso.Kontakt.Anschrift.Strasse & t) ' 
        zeileAntragsteller.Append(perso.Kontakt.Anschrift.Hausnr & t) ' hausnr
        zeileAntragsteller.Append(perso.Kontakt.Anschrift.Hausnr & t) ' hausnr bis
        zeileAntragsteller.Append("Deutschland" & t) ' 
        zeileAntragsteller.Append(perso.Kontakt.Anschrift.PLZ & t) ' 
        zeileAntragsteller.Append(perso.Kontakt.Anschrift.Gemeindename & t) ' 
        zeileAntragsteller.Append(perso.Bemerkung & t) ' zusatz1
        zeileAntragsteller.Append("" & perso.Kontakt.Anschrift.Postfach & t) ' zusatz2
        zeileAntragsteller.Append((perso.Kontakt.elektr.Telefon1 & ", " & perso.Kontakt.elektr.Telefon2).Replace(", ", "") & t) ' 
        zeileAntragsteller.Append((perso.Kontakt.elektr.Fax1 & ", " & perso.Kontakt.elektr.Fax2).Replace(", ", "") & t) ' 
        zeileAntragsteller.Append(perso.Kontakt.elektr.MobilFon & t) ' 
        zeileAntragsteller.Append(perso.Kontakt.elektr.Email & t) ' 
        zeileAntragsteller.Append("" & t) ' de-mail
        zeileAntragsteller.Append(perso.Kontakt.elektr.Homepage & t) ' 
        zeileAntragsteller.Append(perso.Kassenkonto & t) ' ZEICHEN IHAH
        zeileAntragsteller.Append(perso.PersonenID & t) ' 
        zeileAntragsteller.Append("" & t) ' spalte1
        zeileAntragsteller.Append(perso.Kassenkonto & t) '  
        Return zeileAntragsteller
    End Function

    Private Shared Function bildeFirma(perso As person) As String
        Return (perso.Kontakt.Org.Name & ", " & perso.Kontakt.Org.Zusatz & ", " & perso.Kontakt.Org.Bemerkung & ", " & perso.Kontakt.GesellFunktion).Replace(", , , ", "")
    End Function

    Private Function getAllBeteiligte4vorgang(perstemp As person, vid As String) As List(Of person)
        Dim sql As String = "select * from beteiligte_t6 where vorgangsid=" & vid
        Dim per As New person
        Dim perlist As New List(Of person)
        Try
            'myoracle = getMSSQLCon()
            'myoracle.Open()
            dt = alleDokumentDatenHolen(sql)
            For Each drr As DataRow In dt.Rows
                per = New person
                per.Rolle = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("rolle"))))
                per.Name = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("nachname"))))
                per.Vorname = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("vorname"))))
                per.Bemerkung = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("bemerkung"))))
                per.Namenszusatz = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Namenszusatz"))))
                per.Anrede = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Anrede"))))
                per.Kontakt.Anschrift.Gemeindename = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Gemeindename"))))
                per.Kontakt.Anschrift.Strasse = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Strasse"))))
                per.Kontakt.Anschrift.Hausnr = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Hausnr"))))
                per.Kontakt.Anschrift.Postfach = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Postfach"))))
                per.Kontakt.Anschrift.PLZ = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("plz"))))
                per.Kontakt.Org.Name = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgname"))))
                per.Kontakt.Org.Zusatz = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgzusatz"))))
                per.Kontakt.Org.Typ1 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgtyp1"))))
                per.Kontakt.Org.Typ2 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgtyp2"))))
                per.Kontakt.Org.Eigentuemer = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgEigentuemer"))))
                per.Quelle = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("quelle"))))
                per.Kontakt.Org.Bemerkung = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("orgBemerkung"))))
                per.Kontakt.GesellFunktion = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("GesellFunktion"))))
                per.Kontakt.elektr.Telefon1 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("fftelefon1"))))
                per.Kontakt.elektr.Telefon2 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("fftelefon2"))))
                per.Kontakt.elektr.Fax1 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("FFFax1"))))
                per.Kontakt.elektr.Fax2 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("FFFax2"))))
                per.Kontakt.elektr.MobilFon = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("FFMOBILFON"))))
                per.Kontakt.elektr.Email = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("FFemail"))))
                per.Kontakt.elektr.Homepage = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("FFhomepage"))))

                per.Kassenkonto = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Kassenkonto"))))
                per.VERTRETENDURCH = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("VERTRETENDURCH"))))

                'für die sepa daten muss man über die personenid holen

                per.Kontakt.BankkontoID = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("personenid"))))
                perlist.Add(per)
            Next
            Return perlist
        Catch ex As Exception
            'swfehlt.Close()
            l("fertig  " & ex.ToString)
            Return Nothing
        End Try
    End Function

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click

    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        'sachbearbeiter 
        Dim puFehler As String = "\\file-paradigma\paradigma\test\thumbnails\PU_sachbearbeiter" & Environment.UserName & ".txt"
        Dim puAusgabe As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\" & "sachbearbeiter" & ".csv"
        Dim puAusgabeStream As New IO.StreamWriter(puAusgabe)
        swfehlt = New IO.StreamWriter(puFehler)
        swfehlt.AutoFlush = True
        Dim Sql As String
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)

        Sql = "select *   FROM [Paradigma].[dbo].[BEARBEITER_T5] where aktiv=1      order by abteilung desc  "
        TextBox1.Text = puAusgabe
        TextBox2.Text = Sql
        writeSachbearbeiterPU(puFehler, puAusgabeStream, Sql, maxobj)
        puAusgabeStream.Close()
        puAusgabeStream.Dispose()
        System.Diagnostics.Process.Start("explorer", "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung")
        swfehlt.Dispose()
    End Sub

    Private Sub writeSachbearbeiterPU(puFehler As String, puAusgabeStream As IO.StreamWriter, sql As String, maxobj As Integer)
        Dim DT As DataTable
        Dim idok As Integer = 0
        puAusgabeStream.AutoFlush = True
        swfehlt.WriteLine("writeKatasterausgabePU---")
        DT = alleDokumentDatenHolen(sql)

        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim Username, nachname, vorname, rang, raum, abteilung, initial_1, fax, telefon, initial_2, flaecheqm As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum, hauptaktenjahr As Date
        Dim spalte1 As String
        Dim veraltet As String
        Dim eingang, antrag, vollstaendig, bescheid, abgeschlossen As Date
        Dim aktenstandort As String
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        myoracle.Open()
        Dim zeile As New Text.StringBuilder
        Dim fullfilename As String
        Dim t As String = ";"
        Dim geschlossen As String = "0"
        l("writeKatasterausgabePU")
        'kopfzeile
        zeile.Append("Bearbeitungsart" & t) '     Bearbeitungsart   
        zeile.Append("Username" & t) '     datum
        zeile.Append("nachname" & t) '  gemarkungstext    
        zeile.Append("vorname" & t) '  nachname
        zeile.Append("rang" & t) '  znkombi
        zeile.Append("raum" & t) '   rechts    
        zeile.Append("initial_1" & t) ' hoch
        zeile.Append("abteilung" & t) '
        zeile.Append("telefon" & t)
        zeile.Append("fax" & t) '  
        zeile.Append("initial_2" & t) 'abteilung
        zeile.Append("titel" & t) 'telefon
        zeile.Append("email" & t) 'fax
        zeile.Append("rolle" & t) 'fs
        zeile.Append("sachgebiet" & t) ' 
        csvzeileSpeichern(zeile.ToString, puAusgabeStream) : zeile.Clear()
        Dim titel, email, rolle, sachgebiet As String
        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                Bearbeitungsart = CStr(clsDBtools.fieldvalue(drr.Item("Bearbeitungsart")))
                Username = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("Username"))))
                nachname = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("nachname"))))
                vorname = cleanString((clsDBtools.fieldvalue(drr.Item("vorname"))))
                rang = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("rang"))))
                raum = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("rites"))))

                initial_1 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("initial_"))))
                abteilung = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("abteilung"))))
                telefon = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("telefon"))))
                fax = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("fax"))))
                initial_2 = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("kuerzel1"))))
                titel = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("kuerzel1"))))
                email = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("email"))))
                rolle = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("rolle"))))
                sachgebiet = cleanString(CStr(clsDBtools.fieldvalue(drr.Item("EXPANDHEADERINSACHGEBIET"))))

                TextBox3.Text = igesamt & " von " & DT.Rows.Count & "   [maxobj4test: " & maxobj & " ]"
                Application.DoEvents()
                'zeilebilden
                zeile.Append(Bearbeitungsart & t) 'Az 
                zeile.Append(Username & t) ' 
                zeile.Append(nachname & t) ' 
                zeile.Append(vorname & t) ' 
                zeile.Append(rang & t) ' 
                zeile.Append(raum & t) '   
                zeile.Append(initial_1 & t) ' 
                zeile.Append(abteilung & t) ' 
                zeile.Append(telefon & t) '    
                zeile.Append(fax & t) '    
                zeile.Append(initial_2 & t) '    
                zeile.Append(titel & t) '    
                zeile.Append(email & t) '    
                zeile.Append(rolle & t) '    
                zeile.Append(sachgebiet & t) '    

                If csvzeileSpeichern(zeile.ToString, puAusgabeStream) Then

                    zeile.Clear()
                Else
                    Debug.Print("oooo")
                End If
                idok += 1
                If idok > maxobj Then Exit For
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                          Environment.NewLine &
                       Bearbeitungsart & "/" & Bearbeitungsart & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()
        Next
        'csvzeileSpeichern(zeileAntragsteller.ToString, ausgabeAntragsteller)
        zeile.Clear()

        swfehlt.WriteLine(idok & "Teil2 fertig  -------" & Now.ToString & "-------------- " & igesamt)

        '####
        'swfehlt.Close()
        l("fertig  " & puFehler)
        ' Process.Start(puFehler)
    End Sub



    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        Dim puFehler As String = "\\file-paradigma\paradigma\test\thumbnails\PU_ausgabeEreignisse" & Environment.UserName & ".txt"
        Dim puAusgabe As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\" & "dokumente_ereignisse" & ".csv"
        Dim puAusgabeStream As New IO.StreamWriter(puAusgabe)
        '   dateifehlt = "L:\system\batch\margit\auffueller" & Environment.UserName & ".txt"
        swfehlt = New IO.StreamWriter(puFehler)
        swfehlt.AutoFlush = True
        '  swfehlt.WriteLine(Now)
        ' S1020dokumenteMitFullpathTabelleErstellen("DOKUFULLNAME", swfehlt) 'referenzfälleNeuZuweisen
        ' swfehlt.WriteLine("wechsel")
        ' dokumenteMitFullpathTabelleErstellen(swfehlt)
        Dim Sql As String
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)
        'Sql = "SELECT *  FROM [Paradigma].[dbo].[EREIGNIS_T16]  where not( art like '%email%' or art like '%wiederv%')  order by id desc "
        Sql = "    Select   * FROM [Paradigma].[dbo].[EREIGNIS_T16]    e, " &
                "  [Paradigma].[dbo].[stammdaten_tutti] s " &
                " where Not (art Like '%email%' or art like '%wiederv%') " &
                " And e.VORGANGSID = s.VORGANGSID " &
                " order by e.id desc "
        Sql = "    Select   * FROM [Paradigma].[dbo].[EREIGNIS_T16]    e, " &
                "  [Paradigma].[dbo].[stammdaten_tutti] s " &
                " where  " &
                "   e.VORGANGSID = s.VORGANGSID " &
                "  and    not( art like '%email%' or art like '%wiederv%' or  (NOTIZ) is  null  ) " &
                " order by    e.VORGANGSID desc,  e.DATUM desc"

        'Sql = "  SELECT e.id,e.BESCHREIBUNG,datum,art,richtung,notiz,typnr,e.VORGANGSID,s.EINGANG,dateinameext " &
        '        " FROM [Paradigma].[dbo].EREIGNIS_und_DOK e,   " &
        '        "   [Paradigma].[dbo].[stammdaten_tutti] s  " &
        '        "   Where   " &
        '        "     e.VORGANGSID = s.VORGANGSID  " &
        '        "  and    not( art like '%email%' or art like '%wiederv%' or  (NOTIZ) is  null  ) " &
        '        "  order by VORGANGSID desc,datum desc "


        Dim relativpfad As String = "O:\UMWELT\B\GISDatenEkom\proumweltaufbereitung\ereignisse\"


        TextBox1.Text = puAusgabe
        TextBox2.Text = Sql
        'MsgBox("max. objekte für test: " & maxobj)
        writeEreignissePU(puFehler, puAusgabeStream, Sql, maxobj, relativpfad)
        puAusgabeStream.Close()
        puAusgabeStream.Dispose()
        'Process.Start(puAusgabe)
        '######
        'puAusgabe = "D:\probaug_Ausgabe\" & "ereignisse" & ".csv"
        'puAusgabeStream = New IO.StreamWriter(puAusgabe)
        ''     Sql = "SELECT * FROM [Paradigma].[dbo].[probaug_dokumente_referenz]  order by ort desc "
        'TextBox1.Text = TextBox1.Text & Environment.NewLine & puAusgabe
        'TextBox2.Text = TextBox2.Text & Environment.NewLine & Sql
        'writeDokumentePU(puFehler, puAusgabeStream, Sql, 500)
        swfehlt.Close()
        l("fertig  " & puFehler)
        ' SELECT *  FROM [Paradigma].[dbo].[EREIGNIS_T16]  where not( art like '%email%' or art like '%wiederv%')
    End Sub

    Private Sub writeEreignissePU(puFehler As String, puAusgabeStream As IO.StreamWriter, sql As String, maxobj As Integer, relativpfad As String)
        '####
        Dim DT As DataTable
        Dim idok As Integer = 0
        puAusgabeStream.AutoFlush = True
        'ausgabeAntragsteller.WriteLine(Now)
        l("PDFumwandeln ")


        inndir = "\\file-paradigma\paradigma\test\paradigmaArchiv\backup\archiv"
        If Bearbeitungsart = "fehler" Then End

        DT = alleDokumentDatenHolen(sql)
        l("vor csvverarbeiten")
        Dim ic As Integer = 0
        Dim igesamt As Integer = 0
        Dim dateinameext, art, richtung, inputfile, outfile, typnr, outstring, summary As String
        Dim newsavemode As Boolean
        Dim istRevisionssicher As Boolean
        Dim dbdatum As Date
        Dim notiz As String
        Dim beschreibung As String
        Dim eingang As Date
        Dim eid, vid As String
        Dim myoracle As SqlClient.SqlConnection
        myoracle = getMSSQLCon()
        myoracle.Open()
        Dim zeile As New Text.StringBuilder
        'Dim block As New Text.StringBuilder 
        'Dim blockMAX As Int16 = 50
        'Dim iblock As Int16 = 0
        Dim fullfilename As String
        Dim t As String = ";"
        l("PDFumwandeln 2 ")
        l("PDFumwandeln 2 ")
        '  Using sw As New IO.StreamWriter(logfile)

        'kopfzeile
        zeile.Append("az" & t) 'Az
        zeile.Append("jahr" & t) 'jahr
        zeile.Append("datum" & t) 'datum
        zeile.Append("oberbegriff" & t) 'oberbegriff Protokolle
        zeile.Append((cleanString("bezeichnung")) & t) 'bezeichnung beschreibung
        zeile.Append(("pfad") & t) 'pfad
        zeile.Append("ordner" & t) 'ordner im mediencenter
        zeile.Append("revisionssicher" & t) ' 
        zeile.Append("bearbeiterid" & t) ' 
        csvzeileSpeichern(zeile.ToString, puAusgabeStream)
        zeile.Clear()

        For Each drr As DataRow In DT.Rows
            Try
                igesamt += 1
                DbMetaDatenEreignisHolen(art, richtung, notiz, typnr, dbdatum, eingang, vid,
                                           eid, beschreibung, drr)
                If Trim(notiz) = String.Empty Then
                    Continue For
                End If
                l(eid & " " & CStr(art) & " " & ic)
                outfile = dbdatum.ToString("yyyyMMdd_hhmmss") & "_" & cleanString(art) & "_" & cleanString(richtung) & ".txt"
                outfile = relativpfad & vid & "\" & eid & "\" & outfile
                IO.Directory.CreateDirectory(relativpfad & vid & "\" & eid)
                outstring = erzeugeEreignisString(beschreibung, richtung, art, dbdatum, vid, eid, typnr, notiz)
                If schreibeEreignisdatei(outfile, outstring) Then

                Else
                    l("fehler beim erzeugeEreignisString ")
                End If

                TextBox3.Text = igesamt & " von " & DT.Rows.Count & "   [maxobj4test: " & maxobj & " ]"
                Application.DoEvents()
                'zeilebilden
                zeile.Append(vid & t) 'Az
                zeile.Append(eingang.ToString("yyyy") & t) 'jahr
                zeile.Append(dbdatum.ToString("yyyyMMdd_hhmmss") & t) 'datum
                zeile.Append(art & t) 'oberbegriff Protokolle
                zeile.Append((cleanString(beschreibung)) & t) 'bezeichnung beschreibung
                zeile.Append((outfile) & t) 'pfad
                zeile.Append("" & t) 'ordner im mediencenter
                zeile.Append("" & t) 'ordner im mediencenter
                zeile.Append("" & t) 'ordner im mediencenter

                'If iblock < blockMAX Then
                '    block.AppendLine(zeileAntragsteller.ToString)
                '    zeileAntragsteller.Clear()
                '    iblock += 1
                'Else
                If csvzeileSpeichern(zeile.ToString, puAusgabeStream) Then
                    'iblock = 0
                    'block.Clear()
                    zeile.Clear()
                Else
                    Debug.Print("oooo")
                End If


                'zeileAntragsteller.Clear()
                idok += 1
                If idok > maxobj Then Exit For
            Catch ex As Exception
                l("fehler2: " & ex.ToString)
                TextBox2.Text = ic.ToString & Environment.NewLine & " " &
                       inputfile & Environment.NewLine &
                       vid & "/" & eid & " " & igesamt & "(" & DT.Rows.Count.ToString & ")" & Environment.NewLine &
                       TextBox2.Text
                Application.DoEvents()
            End Try
            GC.Collect()
            GC.WaitForFullGCComplete()
        Next
        csvzeileSpeichern(zeile.ToString, puAusgabeStream)
        zeile.Clear()

        swfehlt.WriteLine(idok & "Teil2 fertig  -------" & Now.ToString & "-------------- " & igesamt)

        '####
        'swfehlt.Close()
        l("fertig  " & puFehler)
    End Sub

    Private Function schreibeEreignisdatei(outfile As String, outstring As String) As Boolean
        Try
            My.Computer.FileSystem.WriteAllText(outfile, outstring, False)
            Return True
        Catch ex As Exception
            l("fertig  " & ex.ToString)
            Return False
        End Try
    End Function

    Private Function erzeugeEreignisString(beschreibung As String, richtung As String, art As String,
                                           dbdatum As Date,
                                           vid As String, eid As String, typnr As String, notiz As String) As String
        Dim pu As New Text.StringBuilder
        Try
            pu.Append("vorgang: " & vid & " ereignis: " & eid & Environment.NewLine)
            pu.Append("richtung: " & cleanString(richtung) & " art: " & cleanString(art) & Environment.NewLine)
            pu.Append("datum: " & dbdatum.ToString("yyyyMMdd_hhmmss") & Environment.NewLine)
            pu.Append("typnr: " & cleanString(typnr) & Environment.NewLine)
            pu.Append("beschreibung: " & cleanString(beschreibung) & Environment.NewLine)
            pu.Append("notiz: " & cleanString(notiz) & Environment.NewLine)
            Return pu.ToString
        Catch ex As Exception
            l("fehler  " & ex.ToString)
            Return "fehler"
        End Try
    End Function



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
        '  Bearbeitungsart = modPrep.getVid()
        Dim maxobj As Integer = 0
        maxobj = setMaxObj(maxobj)
        If Bearbeitungsart = "fehler" Then End
        Dim Sql As String
        Sql = "SELECT * FROM dokumente where   ort<2000000 and ort>0  " &
                  " and (revisionssicher=1) order by ort desc "
        Sql = "SELECT * FROM dokumente " &
            " LEFT JOIN t08 " &
            " ON dokumente.DOKUMENTID = t08.DOKID " &
            " where  dokid =" & ausCheckDokumentid
        DT = RevSicherdokumentDatenHolen(Sql)
        dateinameext = CStr(DT.Rows(0).Item("Verfahrensart"))
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
        host = "kh-w-sql02" : schema = "Paradigma" : dbuser = "sgis" : dbpw = "Grunt8-Cornhusk-Reporter"
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
