Imports System.Data
Imports System.Data.Odbc
Public Class Database
    Public Conn As OdbcConnection
    ''' <summary>
    ''' Database Connection String
    ''' </summary>
    ''' <returns></returns>
    Public Function sConnect() As String
        Dim sAns As String = ""
        sAns = "Driver={Microsoft Access Driver (*.mdb)};dbq=" & Application.StartupPath & "\fileminer.mdb"
        Return sAns
    End Function
    ''' <summary>
    ''' Connects to the database
    ''' </summary>
    ''' <param name="sError"></param>
    Public Sub ConnectDB(Optional ByRef sError As String = "")
        Try
            Conn = New OdbcConnection(sConnect)
            Conn.Open()
        Catch ex As Exception
            sError = Err.Number & "-" & ex.Message.ToString
        End Try
    End Sub
    ''' <summary>
    ''' Close the Database
    ''' </summary>
    ''' <param name="sError"></param>
    Public Sub CloseDB(Optional ByRef sError As String = "")
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
    ''' <param name="sSQL"></param>
    ''' <param name="sError"></param>
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
    ''' <summary>
    ''' Function to return the querey back in a datatable format
    ''' </summary>
    ''' <param name="SQL"></param>
    ''' <param name="sError"></param>
    ''' <returns>DataTable</returns>
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
    ''' <summary>
    ''' Checks to see if a string already exists in the searchstring table.
    ''' </summary>
    ''' <param name="sValue"></param>
    ''' <param name="SCID"></param>
    ''' <param name="sError"></param>
    ''' <returns>Boolean</returns>
    Public Function StringExists(sValue As String, SCID As Long, Optional ByRef sError As String = "") As Boolean
        Dim bAns As Boolean = False
        Try
            Dim SQL As String = "SELECT * from SearchStrings where SearchString='" & sValue & "' and SCID=" & SCID
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
    ''' <summary>
    ''' Checks to see if there is already a category listed in the database with the same value
    ''' </summary>
    ''' <param name="sValue"></param>
    ''' <param name="sError"></param>
    ''' <returns>Boolean</returns>
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
