Module modDokExist
        Public Function ALLEdokumentDatenHolen() As DataTable
        Dim Sql As String
        Dim dt As New DataTable
        Try
            Sql = "SELECT * FROM dokumente where   dokumentid<2000000 and dokumentid>0  " &
                  "  order by dokumentid desc "
            dt = getDT(Sql)
            l("nach getDT")
            Return dt
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

End Module
