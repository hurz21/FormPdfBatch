Imports System.IO
Imports System.Text

Module bplanpdfhtml


    Public Function makeHtml4Gemarkung(startverzeichnis As String, htmlausgabeverzeichnis As String) As Boolean
        Try
            ' Sicherstellen, dass das Basis-Ausgabeverzeichnis existiert
            If Not Directory.Exists(htmlausgabeverzeichnis) Then
                Directory.CreateDirectory(htmlausgabeverzeichnis)
            End If

            ' Alle Verzeichnisse im Startverzeichnis ermitteln, die mit "bplan" beginnen
            Dim bplanDirs As String() = Directory.GetDirectories(startverzeichnis, "bplan*", SearchOption.TopDirectoryOnly)

            For Each bplanDir As String In bplanDirs
                Dim dirName As String = Path.GetFileName(bplanDir).ToLower()

                ' Gemarkungsnamen extrahieren ("bplanbuchschlag" -> "Buchschlag")
                Dim gemarkungName As String = dirName
                If gemarkungName.StartsWith("bplan") Then
                    gemarkungName = gemarkungName.Substring(5)
                End If
                If gemarkungName.Length > 0 Then
                    gemarkungName = Char.ToUpper(gemarkungName(0)) & gemarkungName.Substring(1)
                End If

                ' Unterverzeichnisse im Bplan-Verzeichnis einlesen (z.B. bu_1-10, bu_b1)
                Dim subDirs As String() = Directory.GetDirectories(bplanDir, "*", SearchOption.TopDirectoryOnly)

                For Each subDir As String In subDirs
                    Dim subDirName As String = Path.GetFileName(subDir)

                    ' Ziel-Unterverzeichnis im Ausgabeverzeichnis definieren (z.B. L:\fkat\test\bplaene\bu_1-10)
                    Dim targetSubDir As String = Path.Combine(htmlausgabeverzeichnis, subDirName)

                    ' Unterverzeichnis kopieren (inkl. Dateien)
                    CopyDirectory(subDir, targetSubDir)

                    ' HTML-Datei direkt im Ausgabeverzeichnis anlegen (neben dem kopierten Ordner)
                    Dim htmlFileName As String = Path.Combine(htmlausgabeverzeichnis, subDirName & ".html")

                    ' Alle Dateien aus dem kopierten Verzeichnis einlesen, um die Links zu generieren
                    Dim copiedFiles As String() = Directory.GetFiles(targetSubDir)

                    Dim planName = subDirName
                    Dim a() As String
                    a = subDirName.Split("_"c)
                    planName = a(1)


                    ' HTML-Inhalt aufbauen
                    Dim sb As New StringBuilder()
                    sb.AppendLine("<!DOCTYPE html>")
                    sb.AppendLine("<html lang=""de"">")
                    sb.AppendLine("<head>")
                    sb.AppendLine("    <meta charset=""utf-8"">")
                    sb.AppendLine($"    <title>Bebauungsplan Gemarkung: {gemarkungName} - {subDirName}</title>")
                    sb.AppendLine("    <style>")
                    sb.AppendLine("        body { font-family: Arial, sans-serif; margin: 40px; color: #333; }")
                    sb.AppendLine("        h1 { color: #0056b3; }")
                    sb.AppendLine("        ul { line-height: 1.8; }")
                    sb.AppendLine("        a { text-decoration: none; color: #0066cc; }")
                    sb.AppendLine("        a:hover { text-decoration: underline; }")
                    sb.AppendLine("    </style>")
                    sb.AppendLine("</head>")
                    sb.AppendLine("<body>")
                    sb.AppendLine($"    <h1>Bebauungsplan Gemarkung: {gemarkungName}</h1>")
                    sb.AppendLine($"    <h2>B-Plan-Nr: {planName}</h2>")
                    sb.AppendLine("    <ul>")

                    For Each filePath As String In copiedFiles
                        Dim fileName As String = Path.GetFileName(filePath)

                        ' Da die HTML-Datei direkt eine Ebene über dem Ordner oder daneben liegt:
                        ' Wenn die HTML-Datei im gleichen Ordner wie die PDFs liegen soll, reicht der Dateiname.
                        ' Wenn die PDFs im kopierten Unterordner liegen (z.B. ausgabeverzeichnis\bu_1-10\datei.pdf) 
                        ' und die HTML-Datei im ausgabeverzeichnis\bu_1-10.html, dann verweist der Link in das Unterverzeichnis:
                        Dim relPath As String = $"{subDirName}/{fileName}"

                        ' Listeneintrag mit Link
                        sb.AppendLine($"        <li><a href=""{relPath}"" target=""_blank"">{fileName}</a></li>")
                    Next

                    sb.AppendLine("    </ul>")
                    sb.AppendLine("</body>")
                    sb.AppendLine("</html>")

                    ' HTML-Datei schreiben
                    File.WriteAllText(htmlFileName, sb.ToString(), Encoding.UTF8)
                Next
            Next

            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function


    ''' <summary>
    ''' Hilfsfunktion zum rekursiven Kopieren eines Verzeichnisses inklusive aller Dateien.
    ''' </summary>
    Private Sub CopyDirectory(sourceDir As String, destinationDir As String)
        ' Zielverzeichnis erstellen, falls nicht vorhanden
        If Not Directory.Exists(destinationDir) Then
            Directory.CreateDirectory(destinationDir)
        End If

        ' Dateien kopieren
        For Each filePath As String In Directory.GetFiles(sourceDir)
            Dim fileName As String = Path.GetFileName(filePath)
            Dim destFile As String = Path.Combine(destinationDir, fileName)
            File.Copy(filePath, destFile, True) ' True überschreibt vorhandene Dateien
        Next

        ' Unterverzeichnisse rekursiv kopieren
        For Each subDir As String In Directory.GetDirectories(sourceDir)
            Dim subDirName As String = Path.GetFileName(subDir)
            Dim destSubDir As String = Path.Combine(destinationDir, subDirName)
            CopyDirectory(subDir, destSubDir)
        Next
    End Sub


End Module
