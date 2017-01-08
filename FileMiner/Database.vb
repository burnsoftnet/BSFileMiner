Imports System.Data
Imports System.Data.Odbc
Public Class Database
    Public Conn As OdbcConnection
    Public Function sConnect() As String
        Dim sAns As String = ""
        sAns = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Application.StartupPath & "\fileminer.mdb"
        Return sAns
    End Function
    Public Sub ConnectDB(Optional ByRef sError As String = "")
        Try
            Conn = New OdbcConnection(sConnect)
            Conn.Open()
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
    End Sub
    Public Sub CloseDB(Optional ByRef sError As String = "")
        Try
            Conn.Close()
            Conn = Nothing
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
    End Sub
    Public Sub ConnExec(ByVal sSQL As String, Optional ByRef sError As String = "")
        Try
            Call ConnectDB()
            Dim CMD As New OdbcCommand
            CMD.Connection = Conn
            CMD.CommandText = sSQL
            CMD.ExecuteNonQuery()
            CMD.Connection.Close()
            CMD = Nothing
            Conn = Nothing
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
    End Sub
    Public Function GetData(SQL As String, Optional ByRef sError As String = "") As DataTable
        Dim Table As New DataTable
        Try
            Call ConnectDB()
            Dim CMD As New OdbcCommand(SQL, Conn)
            Dim RS As New OdbcDataAdapter
            RS.SelectCommand = CMD
            RS.Fill(Table)
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
        Return Table
    End Function
    Public Function StringExists(sValue As String, Optional ByRef sError As String = "") As Boolean
        Dim bAns As Boolean = False
        Try
            Dim SQL As String = "SELECT * from SearchStrings where SearchString='" & sValue & "'"
            Call ConnectDB()
            Dim CMD As New OdbcCommand(SQL, Conn)
            Dim RS As OdbcDataReader
            RS = CMD.ExecuteReader
            bAns = RS.HasRows
            RS.Close()
            RS = Nothing
            CMD = Nothing
            Call CloseDB()
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
        Return bAns
    End Function
    Public Function SearchCategoryExists(sValue As String, Optional ByRef sError As String = "") As Boolean
        Dim bAns As Boolean = False
        Try
            Dim SQL As String = "SELECT * from SearchCategorys where SearchName='" & sValue & "'"
            Call ConnectDB()
            Dim CMD As New OdbcCommand(SQL, Conn)
            Dim RS As OdbcDataReader
            RS = CMD.ExecuteReader
            bAns = RS.HasRows
            RS.Close()
            RS = Nothing
            CMD = Nothing
            Call CloseDB()
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
        Return bAns
    End Function
End Class
