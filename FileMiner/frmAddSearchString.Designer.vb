<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAddSearchString
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdSearchList = New System.Windows.Forms.ComboBox()
        Me.FileminerDataSet = New FileMiner.fileminerDataSet()
        Me.SearchCategoryBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.SearchCategoryTableAdapter = New FileMiner.fileminerDataSetTableAdapters.SearchCategoryTableAdapter()
        CType(Me.FileminerDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SearchCategoryBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 45)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "String:"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(107, 42)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(210, 20)
        Me.TextBox1.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(67, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Keep Open?"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(107, 68)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(44, 17)
        Me.CheckBox1.TabIndex = 3
        Me.CheckBox1.Text = "Yes"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(15, 98)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "&Add"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(198, 98)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "E&xit"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(12, 18)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(89, 13)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "Search Category:"
        '
        'cmdSearchList
        '
        Me.cmdSearchList.DataSource = Me.SearchCategoryBindingSource
        Me.cmdSearchList.DisplayMember = "SearchName"
        Me.cmdSearchList.FormattingEnabled = True
        Me.cmdSearchList.Location = New System.Drawing.Point(107, 15)
        Me.cmdSearchList.Name = "cmdSearchList"
        Me.cmdSearchList.Size = New System.Drawing.Size(210, 21)
        Me.cmdSearchList.TabIndex = 7
        Me.cmdSearchList.ValueMember = "ID"
        '
        'FileminerDataSet
        '
        Me.FileminerDataSet.DataSetName = "fileminerDataSet"
        Me.FileminerDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'SearchCategoryBindingSource
        '
        Me.SearchCategoryBindingSource.DataMember = "SearchCategory"
        Me.SearchCategoryBindingSource.DataSource = Me.FileminerDataSet
        '
        'SearchCategoryTableAdapter
        '
        Me.SearchCategoryTableAdapter.ClearBeforeFill = True
        '
        'frmAddSearchString
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(325, 136)
        Me.Controls.Add(Me.cmdSearchList)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAddSearchString"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add Search String."
        CType(Me.FileminerDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SearchCategoryBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdSearchList As System.Windows.Forms.ComboBox
    Friend WithEvents FileminerDataSet As FileMiner.fileminerDataSet
    Friend WithEvents SearchCategoryBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents SearchCategoryTableAdapter As FileMiner.fileminerDataSetTableAdapters.SearchCategoryTableAdapter
End Class
