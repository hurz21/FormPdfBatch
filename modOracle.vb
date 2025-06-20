
Imports System.Data

Module modOracle


    Function getDT(sql As String) As DataTable
        'l("in doDatenbank -------------------------------------")
        Dim host, datenbank, schema, tabelle, dbuser, dbpw, dbport As String
        Dim dt As New DataTable
        Dim myoracle As SqlClient.SqlConnection
        host = "kh-w-sql02" : schema = "Paradigma" : dbuser = "sgis" : dbpw = "Grunt8-Cornhusk-Reporter" '"Grunt8-Cornhusk-Reporter"
        Dim conbuil As New SqlClient.SqlConnectionStringBuilder
        Dim v = "Data Source=" & host & ";User ID=" & dbuser & ";Password=" & dbpw & ";" +
                "Initial Catalog=" & schema & ";"

        myoracle = New SqlClient.SqlConnection(v)
        'MsgBox(myoracle.ToString)

        myoracle.Open()
        'MsgBox("nach open")

        'l(sql)

        Dim com As New SqlClient.SqlCommand(sql, myoracle)
        Dim da As New SqlClient.SqlDataAdapter(com)
        da.MissingSchemaAction = MissingSchemaAction.AddWithKey

        Dim mycount As Integer
        'MsgBox("vorfill")
        mycount = da.Fill(dt)
        'l("mycount:" & CStr(mycount))
        Try
            com.Dispose()
            da.Dispose()
            myoracle.Close()
            myoracle.Dispose()
            Return dt

        Catch oex As SqlClient.SqlException
            '  nachricht("Fehler in GetNewid&:" & oex.ToString & " / " & sql)
            ' MsgBox(oex.ToString)

            Return Nothing
        Catch ex As Exception
            '  nachricht("Fehler in GetNewid&:" & ex.ToString & " / " & sql)
            '  MsgBox(ex.ToString)
            Return Nothing
        Finally
            myoracle.Close()
        End Try


    End Function
    Function setExcelAttribute2(vid As Integer, relativpfad As String, dateinameext As String, typ As String, newsavemode As Boolean, dokumentid As Integer, drr As DataRow) As DataTable
        'l("in doDatenbank -------------------------------------")
        Dim host, datenbank, schema, tabelle, dbuser, dbpw, dbport As String
        Dim dt As New DataTable
        Dim myoracle As SqlClient.SqlConnection
        host = "kh-w-sql02" : schema = "Paradigma" : dbuser = "sgis" : dbpw = "Grunt8-Cornhusk-Reporter"
        Dim conbuil As New SqlClient.SqlConnectionStringBuilder
        Dim v = "Data Source=" & host & ";User ID=" & dbuser & ";Password=" & dbpw & ";" +
                "Initial Catalog=" & schema & ";"

        'MsgBox(v)
        myoracle = New SqlClient.SqlConnection(v)
        'MsgBox(myoracle.ToString)

        myoracle.Open()
        'MsgBox("nach open")
        Dim Sql = "update dokumente set typ='xlsx' where dokumentid=" & dokumentid
        'l(sql)
        Dim neuersddateiname = dateinameext.ToLower.Replace(".xls", ".xlsx")
        'Dim com As New OracleCommand(Sql, myoracle)
        Dim cmd As SqlClient.SqlCommand = myoracle.CreateCommand()
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "update dokumente set typ = :param1, dateinameext = :param2 where dokumentid = :dokumentid"
        cmd.Parameters.AddWithValue("param1", "xlsx")
        cmd.Parameters.AddWithValue("param2", neuersddateiname)
        cmd.Parameters.AddWithValue("dokumentid", dokumentid)
        cmd.ExecuteNonQuery()

        Dim mycount As Integer
        'MsgBox("vorfill")

        'l("mycount:" & CStr(mycount))
        Try
            cmd.Dispose()

            myoracle.Close()
            myoracle.Dispose()
            Return dt
        Catch oex As SqlClient.SqlException
            '  nachricht("Fehler in GetNewid&:" & oex.ToString & " / " & sql)
            ' MsgBox(oex.ToString)

            Return Nothing
        Catch ex As Exception
            '  nachricht("Fehler in GetNewid&:" & ex.ToString & " / " & sql)
            '  MsgBox(ex.ToString)
            Return Nothing
        Finally
            myoracle.Close()
        End Try


    End Function




End Module
