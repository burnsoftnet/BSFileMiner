Imports FileMiner.Database
''' <summary>
''' Class frmAddSearchString.
''' Implements the <see cref="System.Windows.Forms.Form" />
''' </summary>
''' <seealso cref="System.Windows.Forms.Form" />
Public Class frmAddSearchString
    ''' <summary>
    ''' The l cat
    ''' </summary>
    Public lCAT As Long
    ''' <summary>
    ''' Handles the Click event of the Button2 control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
    ''' <summary>
    ''' Handles the Click event of the Button1 control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox1.TextLength = 0 Then Exit Sub
            Dim Obj As New Database
            Dim ObjOF As New OtherFunctions
            Dim sString As String = ObjOF.FC(TextBox1.Text)
            Dim CID As Long = cmdSearchList.SelectedValue
            Dim SQL As String = "INSERT INTO SearchStrings (SearchString, SCID) VALUES('" & sString & "'," & CID & ")"
            If Not Obj.StringExists(sString, CID) Then
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