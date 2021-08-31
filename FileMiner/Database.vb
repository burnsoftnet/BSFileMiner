Imports System.Data.Odbc
Imports System.IO

''' <summary>
''' Class Database.
''' </summary>
Public Class Database
    Public Conn As OdbcConnection
    ''' <summary>
    ''' Database Connection String
    ''' </summary>
    ''' <returns></returns>
    Public Function SConnect() As String
        Dim appPath As String =Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\BurnSoft\FileMiner\" & "\fileminer.mdb"
        If Not File.Exists(appPath) Then appPath = Application.StartupPath & "\fileminer.mdb"
        Return "Driver={Microsoft Access Driver (*.mdb)};dbq=" & appPath
    End Function
    ''' <summary>
    ''' Connects to the database
    ''' </summary>
    ''' <param name="sError"></param>
    Public Sub ConnectDb(Optional ByRef sError As String = "")
        Try
            Conn = New OdbcConnection(SConnect)
            Conn.Open()
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
    End Sub
    ''' <summary>
    ''' Close the Database
    ''' </summary>
    ''' <param name="sError"></param>
    Public Sub CloseDb(Optional ByRef sError As String = "")
        Try
            Conn.Close()
            Conn = Nothing
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
    End Sub
    ''' <summary>
    ''' Simple T-SQL statement execution Sub
    ''' </summary>
    ''' <param name="sSql"></param>
    ''' <param name="sError"></param>
    Public Sub ConnExec(ByVal sSql As String, Optional ByRef sError As String = "")
        Try
            Call ConnectDb()
            Dim cmd As New OdbcCommand
            cmd.Connection = Conn
            cmd.CommandText = sSql
            cmd.ExecuteNonQuery()
            cmd.Connection.Close()
            Conn = Nothing
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
    End Sub
    ''' <summary>
    ''' Function to return the querey back in a datatable format
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <param name="sError"></param>
    ''' <returns>DataTable</returns>
' ReSharper disable once UnusedMember.Global
    Public Function GetData(sql As String, Optional ByRef sError As String = "") As DataTable
        Dim table As New DataTable
        Try
            Call ConnectDb()
            Dim cmd As New OdbcCommand(sql, Conn)
            Dim rs As New OdbcDataAdapter
            rs.SelectCommand = cmd
            rs.Fill(table)
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
        Return table
    End Function
    ''' <summary>
    ''' Checks to see if a string already exists in the searchstring table.
    ''' </summary>
    ''' <param name="sValue"></param>
    ''' <param name="scid"></param>
    ''' <param name="sError"></param>
    ''' <returns>Boolean</returns>
    Public Function StringExists(sValue As String, scid As Long, Optional ByRef sError As String = "") As Boolean
        Dim bAns As Boolean = False
        Try
            Dim sql As String = "SELECT * from SearchStrings where SearchString='" & sValue & "' and SCID=" & scid
            Call ConnectDb()
            Dim cmd As New OdbcCommand(sql, Conn)
            Dim rs As OdbcDataReader
            rs = cmd.ExecuteReader
            bAns = rs.HasRows
            rs.Close()
            Call CloseDb()
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
        Return bAns
    End Function
    ''' <summary>
    ''' Checks to see if there is already a category listed in the database with the same value
    ''' </summary>
    ''' <param name="sValue"></param>
    ''' <param name="sError"></param>
    ''' <returns>Boolean</returns>
    Public Function SearchCategoryExists(sValue As String, Optional ByRef sError As String = "") As Boolean
        Dim bAns As Boolean = False
        Try
            Dim sql As String = "SELECT * from SearchCategorys where SearchName='" & sValue & "'"
            Call ConnectDb()
            Dim cmd As New OdbcCommand(sql, Conn)
            Dim rs As OdbcDataReader
            rs = cmd.ExecuteReader
            bAns = rs.HasRows
            rs.Close()
            Call CloseDb()
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
        Return bAns
    End Function
End Class
