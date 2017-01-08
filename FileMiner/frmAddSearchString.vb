Imports FileMiner.Database
Public Class frmAddSearchString
    Public lCAT As Long
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox1.TextLength = 0 Then Exit Sub
            Dim Obj As New Database
            Dim ObjOF As New OtherFunctions
            Dim sString As String = ObjOF.FC(TextBox1.Text)
            Dim CID As Long = cmdSearchList.SelectedValue
            Dim SQL As String = "INSERT INTO SearchStrings (SearchString, SCID) VALUES('" & sString & "'," & CID & ")"
            If Not Obj.StringExists(sString) Then
                Obj.ConnExec(SQL)
            Else
                MsgBox(sString & " already exists!")
            End If
            TextBox1.Text = ""
            If Not CheckBox1.Checked Then Me.Close()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub frmAddSearchString_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If lCAT = 0 Then
            Me.SearchCategoryTableAdapter.Fill(Me.FileminerDataSet.SearchCategory)
        Else
            Me.SearchCategoryTableAdapter.FillBy_ID(Me.FileminerDataSet.SearchCategory, lCAT)
        End If

    End Sub
End Class