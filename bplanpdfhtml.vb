Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Module bplanpdfhtml


    Public Function makeHtml4Gemarkung(startverzeichnis As String, htmlausgabeverzeichnis As String, bplandatatable As DataTable) As Boolean
        Try
            ' Sicherstellen, dass das Basis-Ausgabeverzeichnis existiert
            If Not Directory.Exists(htmlausgabeverzeichnis) Then
                Directory.CreateDirectory(htmlausgabeverzeichnis)
            End If
            Dim name3 As String
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
                    ' CopyDirectory(subDir, targetSubDir)

                    ' HTML-Datei direkt im Ausgabeverzeichnis anlegen (neben dem kopierten Ordner)
                    Dim htmlFileName As String = Path.Combine(htmlausgabeverzeichnis, subDirName & ".html")

                    ' Alle Dateien aus dem kopierten Verzeichnis einlesen, um die Links zu generieren
                    Dim copiedFiles As String() = Directory.GetFiles(targetSubDir)

                    Dim planName = subDirName
                    Dim a() As String
                    a = subDirName.Split("_"c)
                    planName = a(1)

                    name3 = GetName3ByPdf(bplandatatable, subDirName)

                    ' HTML-Inhalt aufbauen
                    Dim sb As New StringBuilder()
                    sb.AppendLine("<!DOCTYPE html>")
                    sb.AppendLine("<html lang=""de"">")
                    sb.AppendLine("<head>")
                    sb.AppendLine("    <meta charset=""utf-8"">")
                    sb.AppendLine($"    <title>Bebauungsplan Gemarkung: {gemarkungName} - {subDirName} </title>")
                    sb.AppendLine("    <style>")
                    sb.AppendLine("        body { font-family: Arial, sans-serif; margin: 40px; color: #333; }")
                    sb.AppendLine("        h1 { color: #0056b3; }")
                    sb.AppendLine("        ul { line-height: 1.8; }")
                    sb.AppendLine("        a { text-decoration: none; color: #0066cc; }")
                    sb.AppendLine("        a:hover { text-decoration: underline; }")
                    sb.AppendLine("    </style>")
                    sb.AppendLine("</head>")
                    sb.AppendLine("<body>")
                    sb.AppendLine($"    <h1>Bebauungsplan </h1>")
                    sb.AppendLine($"    <h3>Gemarkung: {gemarkungName}</h3>")
                    sb.AppendLine($"    <p>B-Plan-Nr: {planName}: {name3}</p>")
                    Dim tabelle = BuildBebUeberlagerHtmlByPdf(bplandatatable, subDirName)
                    sb.AppendLine(tabelle)
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

    Public Function GetName3ByPdf(dt As DataTable, pdfName As String) As String
        Dim name3 As String = String.Empty

        Try
            If dt Is Nothing OrElse String.IsNullOrWhiteSpace(pdfName) Then
                Return name3
            End If

            For Each row As DataRow In dt.Rows
                If Not row.IsNull("pdf") AndAlso
               String.Equals(row("pdf").ToString().Trim(), pdfName.Trim(), StringComparison.OrdinalIgnoreCase) Then

                    If Not row.IsNull("name3") Then
                        name3 = row("name3").ToString()
                    End If
                    Exit For
                End If
            Next

            Return name3
        Catch ex As Exception
            Return String.Empty
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



    Public Function makeMasterIndexHtml(ausgabeverzeichnis As String) As Boolean
        Try
            If Not Directory.Exists(ausgabeverzeichnis) Then
                Return False
            End If

            Dim indexPath As String = Path.Combine(ausgabeverzeichnis, "index.html")

            ' Struktur: Gemarkung -> Liste von BPlan-Namen (ohne .html)
            Dim gemarkungenDict As New SortedDictionary(Of String, SortedSet(Of String))(StringComparer.OrdinalIgnoreCase)

            ' Alle HTML-Dateien im Ausgabeverzeichnis einlesen (außer index.html)
            Dim htmlFiles As String() = Directory.GetFiles(ausgabeverzeichnis, "*.html", SearchOption.TopDirectoryOnly)

            For Each htmlFile As String In htmlFiles
                Dim bplanName As String = Path.GetFileNameWithoutExtension(htmlFile)
                If String.Equals(bplanName, "index", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                ' Gemarkungsnamen aus dem H1-Tag der kleinen HTML-Datei auslesen
                Dim gemarkungName As String = "Allgemein"
                Dim content As String = File.ReadAllText(htmlFile, Encoding.UTF8)
                Dim match As Match = Regex.Match(content, "<h1[^>]*>Bebauungsplan Gemarkung:\s*([^<]+)</h1>", RegexOptions.IgnoreCase)
                If match.Success Then
                    gemarkungName = match.Groups(1).Value.Trim()
                End If

                If Not gemarkungenDict.ContainsKey(gemarkungName) Then
                    gemarkungenDict(gemarkungName) = New SortedSet(Of String)(StringComparer.OrdinalIgnoreCase)
                End If

                gemarkungenDict(gemarkungName).Add(bplanName)
            Next

            ' Daten als JSON für das Frontend aufbereiten
            Dim jsonBuilder As New StringBuilder()
            jsonBuilder.AppendLine("{")
            Dim gIndex As Integer = 0
            For Each gemarkung In gemarkungenDict
                If gIndex > 0 Then jsonBuilder.AppendLine(",")
                jsonBuilder.AppendLine($"  ""{EscapeJson(gemarkung.Key)}"": [")

                Dim bIndex As Integer = 0
                For Each bplan In gemarkung.Value
                    If bIndex > 0 Then jsonBuilder.AppendLine(",")
                    jsonBuilder.Append($"    ""{EscapeJson(bplan)}""")
                    bIndex += 1
                Next

                jsonBuilder.AppendLine()
                jsonBuilder.Append("  ]")
                gIndex += 1
            Next
            jsonBuilder.AppendLine()
            jsonBuilder.Append("}")

            ' HTML, CSS und JavaScript zusammenbauen
            Dim sb As New StringBuilder()
            sb.AppendLine("<!DOCTYPE html>")
            sb.AppendLine("<html lang=""de"">")
            sb.AppendLine("<head>")
            sb.AppendLine("    <meta charset=""utf-8"">")
            sb.AppendLine("    <title>Bebauungsplan-Kataster Übersicht</title>")
            sb.AppendLine("    <style>")
            sb.AppendLine("        * { box-sizing: border-box; margin: 0; padding: 0; }")
            sb.AppendLine("        body { font-family: Arial, sans-serif; display: flex; height: 100vh; background-color: #f4f6f9; color: #333; overflow: hidden; }")
            sb.AppendLine("")
            sb.AppendLine("        /* Linke Spalte: Gemarkungen */")
            sb.AppendLine("        #sidebar { width: 300px; background: #2c3e50; color: #fff; display: flex; flex-direction: column; border-right: 1px solid #1a252f; }")
            sb.AppendLine("        #sidebar h2 { padding: 18px 20px; font-size: 1.1rem; background: #1a252f; border-bottom: 1px solid #34495e; }")
            sb.AppendLine("        #gemarkung-list { list-style: none; overflow-y: auto; flex: 1; }")
            sb.AppendLine("        #gemarkung-list li { padding: 12px 20px; cursor: pointer; border-bottom: 1px solid #34495e; transition: background 0.15s; font-size: 0.95rem; }")
            sb.AppendLine("        #gemarkung-list li:hover { background: #34495e; }")
            sb.AppendLine("        #gemarkung-list li.active { background: #3498db; font-weight: bold; }")
            sb.AppendLine("")
            sb.AppendLine("        /* Rechte Spalte: oben B-Pläne, unten IFrame */")
            sb.AppendLine("        #main-content { flex: 1; display: flex; flex-direction: column; height: 100vh; }")
            sb.AppendLine("        .pane-header { padding: 10px 20px; background: #ecf0f1; font-weight: bold; border-bottom: 1px solid #ddd; font-size: 0.95rem; color: #2c3e50; }")
            sb.AppendLine("")
            sb.AppendLine("        #top-pane { height: 40%; min-height: 160px; border-bottom: 2px solid #ccc; background: #fff; display: flex; flex-direction: column; }")
            sb.AppendLine("        #bplan-list { list-style: none; overflow-y: auto; flex: 1; }")
            sb.AppendLine("        #bplan-list li { padding: 9px 20px; cursor: pointer; border-bottom: 1px solid #f1f1f1; transition: background 0.15s; }")
            sb.AppendLine("        #bplan-list li:hover { background: #f8f9fa; }")
            sb.AppendLine("        #bplan-list li.active { background: #e3f2fd; color: #0d6efd; font-weight: bold; border-left: 4px solid #0d6efd; }")
            sb.AppendLine("")
            sb.AppendLine("        #bottom-pane { flex: 1; background: #fff; display: flex; flex-direction: column; min-height: 0; }")
            sb.AppendLine("        #frame-container { position: relative; flex: 1; width: 100%; height: 100%; }")
            sb.AppendLine("        #bplan-frame { width: 100%; height: 100%; border: none; display: none; background: #fff; }")
            sb.AppendLine("        .placeholder-text { padding: 20px; color: #888; font-style: italic; }")
            sb.AppendLine("    </style>")
            sb.AppendLine("</head>")
            sb.AppendLine("<body>")
            sb.AppendLine("")
            sb.AppendLine("    <div id=""sidebar"">")
            sb.AppendLine("        <h2>Gemarkungen</h2>")
            sb.AppendLine("        <ul id=""gemarkung-list""></ul>")
            sb.AppendLine("    </div>")
            sb.AppendLine("")
            sb.AppendLine("    <div id=""main-content"">")
            sb.AppendLine("        <div id=""top-pane"">")
            sb.AppendLine("            <div class=""pane-header"" id=""bplan-header"">Bebauungspläne (Bitte Gemarkung wählen)</div>")
            sb.AppendLine("            <ul id=""bplan-list""><li class=""placeholder-text"">Bitte links eine Gemarkung auswählen.</li></ul>")
            sb.AppendLine("        </div>")
            sb.AppendLine("")
            sb.AppendLine("        <div id=""bottom-pane"">")
            sb.AppendLine("            <div class=""pane-header"" id=""file-header"">Detailansicht</div>")
            sb.AppendLine("            <div id=""frame-container"">")
            sb.AppendLine("                <div id=""frame-placeholder"" class=""placeholder-text"">Bitte oben einen Bebauungsplan auswählen.</div>")
            sb.AppendLine("                <iframe id=""bplan-frame"" title=""Bebauungsplan Detailansicht""></iframe>")
            sb.AppendLine("            </div>")
            sb.AppendLine("        </div>")
            sb.AppendLine("    </div>")
            sb.AppendLine("")
            sb.AppendLine("    <script>")
            sb.AppendLine($"        const data = {jsonBuilder.ToString()};")
            sb.AppendLine("")
            sb.AppendLine("        const gemarkungListEl = document.getElementById('gemarkung-list');")
            sb.AppendLine("        const bplanListEl = document.getElementById('bplan-list');")
            sb.AppendLine("        const bplanHeaderEl = document.getElementById('bplan-header');")
            sb.AppendLine("        const fileHeaderEl = document.getElementById('file-header');")
            sb.AppendLine("        const bplanFrameEl = document.getElementById('bplan-frame');")
            sb.AppendLine("        const framePlaceholderEl = document.getElementById('frame-placeholder');")
            sb.AppendLine("")
            sb.AppendLine("        function resetDetailFrame(message) {")
            sb.AppendLine("            fileHeaderEl.textContent = 'Detailansicht';")
            sb.AppendLine("            bplanFrameEl.removeAttribute('src');")
            sb.AppendLine("            bplanFrameEl.style.display = 'none';")
            sb.AppendLine("            framePlaceholderEl.textContent = message;")
            sb.AppendLine("            framePlaceholderEl.style.display = 'block';")
            sb.AppendLine("        }")
            sb.AppendLine("")
            sb.AppendLine("        function initGemarkungen() {")
            sb.AppendLine("            const gemarkungen = Object.keys(data).sort((a, b) => a.localeCompare(b, 'de'));")
            sb.AppendLine("            gemarkungListEl.innerHTML = '';")
            sb.AppendLine("")
            sb.AppendLine("            gemarkungen.forEach(gem => {")
            sb.AppendLine("                const li = document.createElement('li');")
            sb.AppendLine("                li.textContent = gem;")
            sb.AppendLine("                li.onclick = () => selectGemarkung(gem, li);")
            sb.AppendLine("                gemarkungListEl.appendChild(li);")
            sb.AppendLine("            });")
            sb.AppendLine("        }")
            sb.AppendLine("")
            sb.AppendLine("        function selectGemarkung(gem, element) {")
            sb.AppendLine("            Array.from(gemarkungListEl.children).forEach(el => el.classList.remove('active'));")
            sb.AppendLine("            element.classList.add('active');")
            sb.AppendLine("")
            sb.AppendLine("            bplanHeaderEl.textContent = `Bebauungspläne in Gemarkung: ${gem}`;")
            sb.AppendLine("            resetDetailFrame('Bitte oben einen Bebauungsplan auswählen.');")
            sb.AppendLine("")
            sb.AppendLine("            const bplaene = [...data[gem]].sort((a, b) => a.localeCompare(b, 'de'));")
            sb.AppendLine("            bplanListEl.innerHTML = '';")
            sb.AppendLine("")
            sb.AppendLine("            if (bplaene.length === 0) {")
            sb.AppendLine("                bplanListEl.innerHTML = '<li class=""placeholder-text"">Keine Bebauungspläne gefunden.</li>';")
            sb.AppendLine("                return;")
            sb.AppendLine("            }")
            sb.AppendLine("")
            sb.AppendLine("            bplaene.forEach(bp => {")
            sb.AppendLine("                const li = document.createElement('li');")
            sb.AppendLine("                li.textContent = bp;")
            sb.AppendLine("                li.onclick = () => selectBPlan(bp, li);")
            sb.AppendLine("                bplanListEl.appendChild(li);")
            sb.AppendLine("            });")
            sb.AppendLine("        }")
            sb.AppendLine("")
            sb.AppendLine("        function selectBPlan(bp, element) {")
            sb.AppendLine("            Array.from(bplanListEl.children).forEach(el => el.classList.remove('active'));")
            sb.AppendLine("            element.classList.add('active');")
            sb.AppendLine("")
            sb.AppendLine("            fileHeaderEl.textContent = `Detailansicht: ${bp}.html`;")
            sb.AppendLine("            framePlaceholderEl.style.display = 'none';")
            sb.AppendLine("            bplanFrameEl.style.display = 'block';")
            sb.AppendLine("            bplanFrameEl.src = `${encodeURIComponent(bp)}.html`;")
            sb.AppendLine("        }")
            sb.AppendLine("")
            sb.AppendLine("        initGemarkungen();")
            sb.AppendLine("    </script>")
            sb.AppendLine("</body>")
            sb.AppendLine("</html>")

            File.WriteAllText(indexPath, sb.ToString(), Encoding.UTF8)
            Return True

        Catch ex As Exception
            Return False
        End Try


    End Function




    Private Function EscapeJson(str As String) As String
        If String.IsNullOrEmpty(str) Then Return String.Empty
        Return str.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "").Replace(vbLf, "")
    End Function


    Public Function BuildBebUeberlagerHtmlByPdf(
    dtSource As DataTable,
    pdfName As String,
    Optional pdfColumnName As String = "pdf"
) As String

        Dim sb As New StringBuilder()
        sb.Append("<table style=""width:300px; border-collapse:collapse; font-family:sans-serif; font-size:12px; table-layout:fixed;"">")

        Try
            If dtSource Is Nothing OrElse String.IsNullOrWhiteSpace(pdfName) OrElse Not dtSource.Columns.Contains(pdfColumnName) Then
                sb.Append("<tr><td style=""border:1px solid #ccc; padding:6px;"">Kein passender Eintrag gefunden.</td></tr>")
                sb.Append("</table>")
                Return sb.ToString()
            End If

            Dim mappings As (SourceCol As String, Label As String)() = {
            ("baulnutz", "Bauliche Nutzung"),
            ("aufstellun", "Datum Aufstellung"),
            ("rechtswirk", "Datum Rechtswirksamkeit"),
            ("wird_ueber", "Wird überlagert"),
            ("ueberlager", "Überlagert"),
            ("bemerkung", "Bemerkung"),
            ("ausgleich", "IstAusgleichsfläche"),
            ("flaeche_qm", "Fläche qm")
        }

            Dim foundAny As Boolean = False

            For Each row As DataRow In dtSource.Rows
                If row.IsNull(pdfColumnName) Then
                    Continue For
                End If

                If Not String.Equals(row(pdfColumnName).ToString().Trim(), pdfName.Trim(), StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                foundAny = True

                For Each mapItem In mappings
                    Dim labelText As String = EscapeHtml(mapItem.Label)
                    Dim valueText As String = EscapeHtml(GetSafeCellValue(row, mapItem.SourceCol))

                    sb.Append("<tr>")
                    sb.Append($"<td style=""width:135px; border:1px solid #ccc; padding:4px 6px; font-weight:bold; vertical-align:top; word-wrap:break-word; overflow-wrap:break-word;"">{labelText}</td>")
                    sb.Append($"<td style=""width:165px; border:1px solid #ccc; padding:4px 6px; vertical-align:top; word-wrap:break-word; overflow-wrap:break-word;"">{valueText}</td>")
                    sb.Append("</tr>")
                Next

                sb.Append("<tr><td colspan=""2"" style=""height:8px; border:none;""></td></tr>")
            Next

            If Not foundAny Then
                sb.Append("<tr><td style=""border:1px solid #ccc; padding:6px;"">Kein Eintrag für dieses PDF gefunden.</td></tr>")
            End If

        Catch ex As Exception
            sb.Clear()
            sb.Append("<table style=""width:300px; border-collapse:collapse; font-family:sans-serif; font-size:12px;"">")
            sb.Append($"<tr><td style=""border:1px solid #ccc; padding:6px; color:#b00020;"">Fehler: {EscapeHtml(ex.Message)}</td></tr>")
        End Try

        sb.Append("</table>")
        Return sb.ToString()
    End Function

    Private Function EscapeHtml(text As String) As String
        If String.IsNullOrEmpty(text) Then
            Return String.Empty
        End If

        Return text.Replace("&", "&amp;").
                Replace("<", "&lt;").
                Replace(">", "&gt;").
                Replace("""", "&quot;").
                Replace("'", "&#39;")
    End Function

    Private Function GetSafeCellValue(row As DataRow, colName As String) As String
        If Not row.Table.Columns.Contains(colName) OrElse row.IsNull(colName) Then
            Return String.Empty
        End If

        Dim val As Object = row(colName)
        If TypeOf val Is DateTime Then
            Return CDate(val).ToString("dd.MM.yyyy")
        End If

        Return val.ToString()
    End Function
End Module
