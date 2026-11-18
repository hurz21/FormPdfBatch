Imports System.IO
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Text

Module SqlToWktExportModule
    Public Sub ExportDatabaseToWktFile(outputFilePath As String)
        Dim connectionString As String = $"Server=kh-w-sql02;Database=Paradigma;User Id=sgis;Password=Grunt8-Cornhusk-Reporter;"

        Dim query As String = "SELECT [ID], [RAUMBEZUGSID], [VORGANGSID], [SERIALSHAPE], [TYP], [AREAQM], [SERIALUTM] " &
                                 "FROM [Paradigma].[dbo].[RAUMBEZUG2GEOPOLYGON] " &
                                 "WHERE SERIALSHAPE LIKE '5;%' " &
                                 "ORDER BY VORGANGSID DESC"

        Dim germanCulture As CultureInfo = CultureInfo.GetCultureInfo("de-DE")

        Using writer As New StreamWriter(outputFilePath, False, Encoding.UTF8)

            ' Kopfzeile schreiben
            writer.WriteLine("ID;RAUMBEZUGSID;VORGANGSID;TYP;AREAQM;SERIALUTM;WKT_GEOMETRY")

            Using connection As New SqlConnection(connectionString)
                Using command As New SqlCommand(query, connection)
                    Try
                        connection.Open()
                        Using reader As SqlDataReader = command.ExecuteReader()
                            While reader.Read()
                                Dim id = reader("ID")?.ToString()
                                Dim raumbezugsId = reader("RAUMBEZUGSID")?.ToString()
                                Dim vorgangsId = reader("VORGANGSID")?.ToString()
                                Dim serialShape = reader("SERIALSHAPE")?.ToString()
                                Dim typ = reader("TYP")?.ToString()
                                Dim areaQm = reader("AREAQM")?.ToString()
                                Dim serialUtm = reader("SERIALUTM")?.ToString()

                                Dim wktPolygon As String = ""

                                ' WKT-Konvertierung für SERIALSHAPE
                                If Not String.IsNullOrEmpty(serialShape) Then
                                    Dim parts = serialShape.Split(";"c)
                                    If parts.Length >= 3 Then
                                        Dim coordString = parts(2)
                                        Dim rawCoords = coordString.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)

                                        If rawCoords.Length >= 4 Then
                                            Dim wktPairs As New List(Of String)()

                                            For i = 0 To rawCoords.Length - 1 Step 2
                                                If i + 1 < rawCoords.Length Then
                                                    Dim xStr = rawCoords(i).Trim()
                                                    Dim yStr = rawCoords(i + 1).Trim()
                                                    Dim x, y As Double

                                                    If Double.TryParse(xStr, NumberStyles.Any, germanCulture, x) AndAlso
                                                       Double.TryParse(yStr, NumberStyles.Any, germanCulture, y) Then
                                                        wktPairs.Add($"{x.ToString(CultureInfo.InvariantCulture)} {y.ToString(CultureInfo.InvariantCulture)}")
                                                    End If
                                                End If
                                            Next

                                            If wktPairs.Count > 0 Then
                                                wktPolygon = $"POLYGON (({String.Join(", ", wktPairs)}))"
                                            End If
                                        End If
                                    End If
                                End If

                                ' Datensatz schreiben
                                Dim lineOut = $"{id};{raumbezugsId};{vorgangsId};{typ};{areaQm};{serialUtm};{wktPolygon}"
                                writer.WriteLine(lineOut)
                            End While
                        End Using

                        Console.WriteLine("Export erfolgreich abgeschlossen!")

                    Catch ex As Exception
                        Console.WriteLine($"Fehler beim Datenbankzugriff: {ex.Message}")
                    End Try
                End Using
            End Using
        End Using
    End Sub
    Sub ExportDatabaseToWktFile1(outputFilePath As String)
        ' Verbindungsschnur für den SQL Server zusammenbauen        host = "kh-w-sql02" : schema = "Paradigma" : dbuser = "sgis" : dbpw = "Grunt8-Cornhusk-Reporter"
        Dim connectionString As String = $"Server=kh-w-sql02;Database=Paradigma;User Id=sgis;Password=Grunt8-Cornhusk-Reporter;"

        ' Deine SQL-Abfrage
        Dim query As String = "SELECT [ID], [RAUMBEZUGSID], [VORGANGSID], [SERIALSHAPE], [TYP], [AREAQM], [SERIALUTM] " &
                              "FROM [Paradigma].[dbo].[RAUMBEZUG2GEOPOLYGON] " &
                              "WHERE SERIALSHAPE LIKE '5;%' " &
                              "ORDER BY VORGANGSID DESC"

        Dim outputLines As New List(Of String)()
        Dim germanCulture As CultureInfo = CultureInfo.GetCultureInfo("de-DE")

        ' CSV-Kopfzeile für die Ausgabedatei
        outputLines.Add("ID;RAUMBEZUGSID;VORGANGSID;TYP;AREAQM;SERIALUTM;WKT_GEOMETRY")

        Using connection As New SqlConnection(connectionString)
            Using command As New SqlCommand(query, connection)
                Try
                    connection.Open()
                    Using reader As SqlDataReader = command.ExecuteReader()
                        While reader.Read()
                            ' Attribute auslesen
                            Dim id = reader("ID")?.ToString()
                            Dim raumbezugsId = reader("RAUMBEZUGSID")?.ToString()
                            Dim vorgangsId = reader("VORGANGSID")?.ToString()
                            Dim serialShape = reader("SERIALSHAPE")?.ToString()
                            Dim typ = reader("TYP")?.ToString()
                            Dim areaQm = reader("AREAQM")?.ToString()
                            Dim serialUtm = reader("SERIALUTM")?.ToString()

                            Dim wktPolygon As String = ""

                            ' WKT-Konvertierung fürSERIALSHAPE (Format: 5;0;X1|Y1|X2|Y2...)
                            If Not String.IsNullOrEmpty(serialShape) Then
                                Dim parts = serialShape.Split(";"c)
                                If parts.Length >= 3 Then
                                    Dim coordString = parts(2)
                                    Dim rawCoords = coordString.Split(New Char() {"|"c}, StringSplitOptions.RemoveEmptyEntries)

                                    If rawCoords.Length >= 4 Then
                                        Dim wktPairs As New List(Of String)()

                                        ' Paarweise durchgehen (Step 2 für X und Y)
                                        For i = 0 To rawCoords.Length - 1 Step 2
                                            If i + 1 < rawCoords.Length Then
                                                Dim xStr = rawCoords(i).Trim()
                                                Dim yStr = rawCoords(i + 1).Trim()
                                                Dim x, y As Double

                                                ' Konvertierung unter Berücksichtigung des deutschen Dezimal-Kommas
                                                If Double.TryParse(xStr, NumberStyles.Any, germanCulture, x) AndAlso
                                                   Double.TryParse(yStr, NumberStyles.Any, germanCulture, y) Then

                                                    ' WKT benötigt zwingend den Punkt als Dezimaltrennzeichen
                                                    wktPairs.Add($"{x.ToString(CultureInfo.InvariantCulture)} {y.ToString(CultureInfo.InvariantCulture)}")
                                                End If
                                            End If
                                        Next

                                        If wktPairs.Count > 0 Then
                                            wktPolygon = $"POLYGON (({String.Join(", ", wktPairs)}))"
                                        End If
                                    End If
                                End If
                            End If

                            ' Datensatz als Zeile zusammenfügen
                            Dim lineOut = $"{id};{raumbezugsId};{vorgangsId};{typ};{areaQm};{serialUtm};{wktPolygon}"
                            outputLines.Add(lineOut)
                        End While
                    End Using

                    ' In Datei schreiben
                    File.WriteAllLines(outputFilePath, outputLines, Encoding.UTF8)
                    Console.WriteLine("Export erfolgreich abgeschlossen!")

                Catch ex As Exception
                    Console.WriteLine($"Fehler beim Datenbankzugriff oder Export: {ex.Message}")
                End Try
            End Using
        End Using
    End Sub

End Module
