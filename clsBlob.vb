Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Runtime.Remoting.Messaging

Public Class clsBlob
    Public eing, ausg As String
    'Private Sub start()
    '    eing = "C:\Users\hurz\Pictures\Camera Roll\WIN_20150120_071442.jpg"
    '    ausg = "C:\Users\hurz\Pictures\Camera Roll\WIN_20150120_071442.jpg"
    '    eing = "C:\Users\hurz\Desktop\games\60min_Strategie.pdf"
    '    ausg = "C:\Users\hurz\Desktop\games\60min_Strategie2.pdf"
    '    'Dim newid = db_speichern(eing)
    '    'ausdbholen(ausg, newid)
    'End Sub

    'Private Sub btnspeichern_Click(sender As Object, e As RoutedEventArgs)
    '    Dim SqlConnectionTEMPDB = "Server=localhost\SQLEXPRESS;Database=paradigma;Trusted_Connection=True;"
    '    Dim con As New SqlConnection(SqlConnectionTEMPDB)
    '    'db_speichern(eing, con)
    'End Sub
    Shared Function db_speichern(dateianme As String, dokid As Integer, con As SqlConnection,
                                 EID As Integer, vid As Integer) As Long

        l("db_speichern  0")
        Dim aFS As New System.IO.FileStream(dateianme, System.IO.FileMode.Open, System.IO.FileAccess.Read)
        'Dim aData(Convert.ToInt32(aFS.Length )) As Byte  musste -1 ergänzen weil vb immer eins zu hoch init. dann ist word und excel kaputt
        Dim aData(Convert.ToInt32(aFS.Length - 1)) As Byte
        Dim newid As Long
        l("db_speichern  a")
        Try
            aFS.Read(aData, 0, Convert.ToInt32(aFS.Length))
        Finally
            aFS.Close()
        End Try
        Dim aBLObInsertCmd As New SqlClient.SqlCommand()
        l("db_speichern  1")
        With aBLObInsertCmd
            .CommandText = "INSERT INTO t08(imgfield,DOKID,EID,VID) VALUES (@imgfield,@DOKID,@EID,@VID)"
            .CommandText = .CommandText & ";SELECT CAST(scope_identity() AS int);"
            .Connection = con
            .Parameters.Add(New SqlClient.SqlParameter("@imgfield", SqlDbType.VarBinary, 2147483647, "imgfield"))
            .Parameters(0).Value = aData
            .Parameters.AddWithValue("DOKID", dokid)
            .Parameters.AddWithValue("EID", EID)
            .Parameters.AddWithValue("VID", vid)
        End With
        l("db_speichern  2")
        Try
            ' Schritt 3: Datensatz einfügen
            con.Open()
            l("db_speichern  3")
            Dim kobjssss = aBLObInsertCmd.ExecuteScalar
            l("db_speichern  4")
            If kobjssss Is Nothing Then
                newid = 0
            Else
                newid = CLng(kobjssss.ToString)
            End If

            l("db_speichern  5")
            Return newid
        Finally
            con.Close()
        End Try
    End Function


    Shared Function saveDokumenteTooltip(dokumentid As Integer, myoracle As SqlConnection,
                                 fullpath As String) As Long

        'myoracle = New SqlClient.SqlConnection(v)
        'MsgBox(myoracle.ToString)

        'MsgBox("nach open")
        'Dim Sql = "update dokumente set typ='xlsx' where dokumentid=" & dokumentid
        'l(sql)
        'Dim neuersddateiname = dateinameext.ToLower.Replace(".xls", ".xlsx")
        'Dim com As New OracleCommand(Sql, myoracle)
        Dim cmd As SqlClient.SqlCommand = myoracle.CreateCommand()
        cmd.CommandType = CommandType.Text
        'cmd.CommandText = "update dokumente set tooltip = :param1  where dokumentid = :dokumentid"
        cmd.CommandText = "update dokumente set tooltip ='" & fullpath & "'  where dokumentid = " & dokumentid & ""
        'cmd.Parameters.AddWithValue("param1", fullpath)
        'cmd.Parameters.AddWithValue("dokumentid", dokumentid)
        Try
            'cmd.ExecuteNonQuery()
            Dim mycount As Integer = cmd.ExecuteNonQuery()
            nachricht("Fehler in GetNewid&:" & cmd.CommandText)

            cmd.Dispose()

            'myoracle.Close()
            'myoracle.Dispose()
            Return mycount
        Catch oex As SqlClient.SqlException
            '  nachricht("Fehler in GetNewid&:" & oex.ToString & " / " & sql)
            ' MsgBox(oex.ToString)

            Return 0
        Catch ex As Exception
            '  nachricht("Fehler in GetNewid&:" & ex.ToString & " / " & sql)
            '  MsgBox(ex.ToString)
            Return 0
        Finally
            'myoracle.Close()
        End Try
    End Function
    Shared Function dokufull_speichern(dokid As Integer, con As SqlConnection,
                                 fullpath As String, vid As Integer, tabname As String) As Long

        Dim newid As Long

        Dim aBLObInsertCmd As New SqlClient.SqlCommand()
        'l("db_speichern  1")
        With aBLObInsertCmd
            .CommandText = "INSERT INTO " & tabname & "(DOKUMENTID,FULLNAME,VID) VALUES (@DOKUMENTID,@FULLNAME,@VID)"
            '.CommandText = .CommandText & ";SELECT CAST(scope_identity() AS int);"
            .Connection = con
            .Parameters.AddWithValue("DOKUMENTID", dokid)
            .Parameters.AddWithValue("FULLNAME", fullpath)
            .Parameters.AddWithValue("VID", vid)
        End With
        'l("db_speichern  2")
        Try
            ' Schritt 3: Datensatz einfügen
            con.Open()
            'l("db_speichern  3")
            Dim kobjssss = aBLObInsertCmd.ExecuteScalar
            'l("db_speichern  4")
            If kobjssss Is Nothing Then
                newid = 0
                'l("Nciht gespeichert ")
            Else
                newid = CLng(kobjssss.ToString)
                'l("gespeichert ")
            End If
            Return newid
        Finally
            con.Close()
        End Try
    End Function




    Public Shared Function ausBLOBdbholen(dateiname As String, dokid As Long, con As SqlConnection) As Boolean
        l("ausholen")
        Dim aData As Byte()
        Dim aBLObSelectCmd As New SqlClient.SqlCommand()
        With aBLObSelectCmd
            .CommandText = "SELECT imgfield FROM t08 WHERE (dokid = @dokid)"
            .Connection = con
            .Parameters.Add(New SqlClient.SqlParameter("@dokid", SqlDbType.Int, 4, "dokid"))
            .Parameters(0).Value = dokid
        End With
        Try
            ' Schritt 1: DOC-Inhalt als BLOb aus der Datenbank auslesen
            l("ausholen open: " & aBLObSelectCmd.CommandText)
            con.Open()
            aData = DirectCast(aBLObSelectCmd.ExecuteScalar, Byte())
        Finally
            con.Close()
        End Try
        ' Schritt 2: BLOb-Inhalt als Byte-Array in eine Datei schreiben 
        l("ausholen 2")
        Dim iBLObSize As Integer
        If aData IsNot Nothing Then
            iBLObSize = aData.Length
            Dim aFS As New System.IO.FileStream(dateiname, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Write)
            Try
                aFS.Write(aData, 0, iBLObSize)
            Finally
                aFS.Close()
                aFS = Nothing
            End Try
            Return True
        Else
            Return False
        End If

    End Function


End Class
