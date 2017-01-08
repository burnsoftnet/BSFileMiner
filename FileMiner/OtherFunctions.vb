Imports FileMiner.Database
Imports System.Data
Imports System.Data.Odbc
Public Class OtherFunctions
    Public Function FC(sValue As String) As String
        Dim sAns As String = ""
        sAns = Replace(sValue, "'", "''")
        Return Trim(sAns)
    End Function
    Public Function BuildSearchString(lCATID As Long) As String
        Dim sAns As String = ""
        Dim MyValue As String = ""
        Dim Obj As New Database
        Call Obj.ConnectDB()
        Dim SQL As String = "SELECT * from SearchStrings where SCID=" & lCATID
        Dim CMD As New OdbcCommand(SQL, Obj.Conn)
        Dim RS As OdbcDataReader
        RS = CMD.ExecuteReader
        While RS.Read
            MyValue = RS("SearchString")
            If sAns.Length = 0 Then
                sAns = MyValue & "[^ ,]*"
            Else
                sAns &= "|" & MyValue & "[^ ,]*"
            End If
        End While
        RS.Close()
        RS = Nothing
        CMD = Nothing
        Obj.CloseDB()
        'sAns = "^.*\b(" & sAns & ")\b.*$"
        'sAns = "^.*\b(" & sAns & ")\b.*$"
        Return sAns
    End Function
    Public Function GetSearchListID(sName As String) As Long
        Dim lAns As Long = 0
        Dim Obj As New Database
        Dim SQL As String = "SELECT * from SearchCategory where SearchName='" & sName & "'"
        Call Obj.ConnectDB()
        Dim CMD As New OdbcCommand(SQL, Obj.Conn)
        Dim RS As OdbcDataReader
        RS = CMD.ExecuteReader
        While RS.Read
            lAns = RS("ID")
        End While
        RS.Close()
        RS = Nothing
        CMD = Nothing
        Obj.CloseDB()
        Obj = Nothing
        Return lAns
    End Function
End Class
