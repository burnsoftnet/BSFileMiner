Public Class frmViewList
    Public lCATID As Long
    Public UpdatePending As Boolean
    Private Sub frmViewList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.SearchStringsTableAdapter.FillBy_SCID(Me.FileminerDataSet.SearchStrings, lCATID)
    End Sub

    Private Sub SearchStringsBindingSource_ListChanged(sender As Object, e As System.ComponentModel.ListChangedEventArgs) Handles SearchStringsBindingSource.ListChanged
        If Me.FileminerDataSet.HasChanges Then
            Me.UpdatePending = True
        End If
    End Sub

    Private Sub DataGridView1_RowValidated(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.RowValidated
        If Me.UpdatePending Then
            Me.SearchStringsTableAdapter.Update(Me.FileminerDataSet)
            Me.UpdatePending = False
        End If
    End Sub
End Class