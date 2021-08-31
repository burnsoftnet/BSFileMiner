''' <summary>
''' Class frmAddSearchString.
''' Implements the <see cref="System.Windows.Forms.Form" />
''' </summary>
''' <seealso cref="System.Windows.Forms.Form" />
Public Class FrmAddSearchString
    ''' <summary>
    ''' The l cat
    ''' </summary>
    Public LCat As Long
    ''' <summary>
    ''' Handles the Click event of the Button2 control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Close()
    End Sub
    ''' <summary>
    ''' Handles the Click event of the Button1 control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            If TextBox1.TextLength = 0 Then Exit Sub
            Dim obj As New Database
            Dim objOf As New OtherFunctions
            Dim sString As String = objOf.FC(TextBox1.Text)
            Dim cid As Long = cmdSearchList.SelectedValue
            Dim sql As String = "INSERT INTO SearchStrings (SearchString, SCID) VALUES('" & sString & "'," & cid & ")"
            If Not obj.StringExists(sString, cid) Then
                obj.ConnExec(sql)
            Else
                MsgBox(sString & " already exists!")
            End If
            TextBox1.Text = ""
            If Not CheckBox1.Checked Then Close()

        Catch ex As Exception

        End Try
    End Sub
    ''' <summary>
    ''' Handles the Load event of the frmAddSearchString control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub frmAddSearchString_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If LCat = 0 Then
            SearchCategoryTableAdapter.Fill(FileminerDataSet.SearchCategory)
        Else
            SearchCategoryTableAdapter.FillBy_ID(FileminerDataSet.SearchCategory, LCat)
        End If

    End Sub
End Class