﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmViewList
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.IDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SearchStringDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SearchStringsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.FileminerDataSet = New FileMiner.fileminerDataSet()
        Me.SearchStringsTableAdapter = New FileMiner.fileminerDataSetTableAdapters.SearchStringsTableAdapter()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SearchStringsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FileminerDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IDDataGridViewTextBoxColumn, Me.SearchStringDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.SearchStringsBindingSource
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(284, 395)
        Me.DataGridView1.TabIndex = 0
        '
        'IDDataGridViewTextBoxColumn
        '
        Me.IDDataGridViewTextBoxColumn.DataPropertyName = "ID"
        Me.IDDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IDDataGridViewTextBoxColumn.Name = "IDDataGridViewTextBoxColumn"
        Me.IDDataGridViewTextBoxColumn.Visible = False
        '
        'SearchStringDataGridViewTextBoxColumn
        '
        Me.SearchStringDataGridViewTextBoxColumn.DataPropertyName = "SearchString"
        Me.SearchStringDataGridViewTextBoxColumn.HeaderText = "SearchString"
        Me.SearchStringDataGridViewTextBoxColumn.Name = "SearchStringDataGridViewTextBoxColumn"
        '
        'SearchStringsBindingSource
        '
        Me.SearchStringsBindingSource.DataMember = "SearchStrings"
        Me.SearchStringsBindingSource.DataSource = Me.FileminerDataSet
        '
        'FileminerDataSet
        '
        Me.FileminerDataSet.DataSetName = "fileminerDataSet"
        Me.FileminerDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'SearchStringsTableAdapter
        '
        Me.SearchStringsTableAdapter.ClearBeforeFill = True
        '
        'frmViewList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 395)
        Me.Controls.Add(Me.DataGridView1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmViewList"
        Me.Text = "View Search List"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SearchStringsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FileminerDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents FileminerDataSet As FileMiner.fileminerDataSet
    Friend WithEvents SearchStringsBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents SearchStringsTableAdapter As FileMiner.fileminerDataSetTableAdapters.SearchStringsTableAdapter
    Friend WithEvents IDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SearchStringDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
