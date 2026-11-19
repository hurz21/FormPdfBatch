Imports System.Data
Imports Npgsql

Public Module PostGisExport

    Public Function GetBebUeberlagerDataTable(
        sql As String
    ) As DataTable

        Dim connString As String =
            $"Host=gis;Port=5432;Database=postgis20;Username=postgres;Password=lkof4;"



        Dim dt As New DataTable("beb_ueberlager_f")

        Using conn As New NpgsqlConnection(connString)
            conn.Open()

            Using cmd As New NpgsqlCommand(sql, conn)
                Using reader As NpgsqlDataReader = cmd.ExecuteReader()
                    ' Spaltenstruktur anlegen (ohne geom)
                    For i As Integer = 0 To reader.FieldCount - 1
                        Dim colName As String = reader.GetName(i)
                        If Not String.Equals(colName, "geom", StringComparison.OrdinalIgnoreCase) Then
                            dt.Columns.Add(colName, reader.GetFieldType(i))
                        End If
                    Next

                    ' Datenzeilen übertragen
                    While reader.Read()
                        Dim row As DataRow = dt.NewRow()

                        For Each col As DataColumn In dt.Columns
                            Dim val As Object = reader(col.ColumnName)
                            row(col.ColumnName) = If(val Is Nothing OrElse TypeOf val Is DBNull, DBNull.Value, val)
                        Next

                        dt.Rows.Add(row)
                    End While
                End Using
            End Using
        End Using

        Return dt
    End Function

End Module