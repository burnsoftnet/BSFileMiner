<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.ListBox2 = New System.Windows.Forms.ListBox()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.ProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.tsLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SaveLastFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.YesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AutoOutputFoundResultsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.CloseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AddToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StringToSearchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SetLastFileProcessToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ViewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SearchListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LastFileFoundToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtFileFormat = New System.Windows.Forms.TextBox()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbSearchList = New System.Windows.Forms.ComboBox()
        Me.SearchCategoryBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.FileminerDataSetBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.FileminerDataSet = New FileMiner.fileminerDataSet()
        Me.SearchCategoryTableAdapter = New FileMiner.fileminerDataSetTableAdapters.SearchCategoryTableAdapter()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.SearchCategoryBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FileminerDataSetBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FileminerDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(112, 35)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(297, 20)
        Me.TextBox1.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(112, 123)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Search Text"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.AccessibleDescription = "Browse to Directory"
        Me.Button2.Image = Global.FileMiner.My.Resources.Resources.mac_folder_32x32
        Me.Button2.Location = New System.Drawing.Point(415, 33)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(35, 23)
        Me.Button2.TabIndex = 8
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(0, 152)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(935, 488)
        Me.TabControl1.TabIndex = 9
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.ListView1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(927, 462)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "File Found"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'ListView1
        '
        Me.ListView1.Location = New System.Drawing.Point(9, 7)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(909, 485)
        Me.ListView1.TabIndex = 10
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.List
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.ListBox2)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(927, 462)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Errors"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'ListBox2
        '
        Me.ListBox2.FormattingEnabled = True
        Me.ListBox2.Location = New System.Drawing.Point(6, 6)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(915, 485)
        Me.ListBox2.TabIndex = 0
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProgressBar1, Me.tsLabel})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 643)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(934, 22)
        Me.StatusStrip1.SizingGrip = False
        Me.StatusStrip1.TabIndex = 10
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(100, 16)
        '
        'tsLabel
        '
        Me.tsLabel.Name = "tsLabel"
        Me.tsLabel.Size = New System.Drawing.Size(0, 17)
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(85, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "Path To Search:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(7, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(95, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Use Progress Bar?"
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Checked = True
        Me.CheckBox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox1.Location = New System.Drawing.Point(112, 64)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(44, 17)
        Me.CheckBox1.TabIndex = 13
        Me.CheckBox1.Text = "Yes"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(162, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(434, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "(Depending on the size of the directory, turning this off can speed it up, but yo" &
    "u lose status)"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(222, 123)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(106, 23)
        Me.Button3.TabIndex = 15
        Me.Button3.Text = "Save Results"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.AddToolStripMenuItem, Me.ViewToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(934, 24)
        Me.MenuStrip1.TabIndex = 16
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SaveLastFileToolStripMenuItem, Me.AutoOutputFoundResultsToolStripMenuItem, Me.ToolStripSeparator1, Me.CloseToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'SaveLastFileToolStripMenuItem
        '
        Me.SaveLastFileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.YesToolStripMenuItem, Me.NoToolStripMenuItem})
        Me.SaveLastFileToolStripMenuItem.Name = "SaveLastFileToolStripMenuItem"
        Me.SaveLastFileToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.SaveLastFileToolStripMenuItem.Text = "Save Last File"
        '
        'YesToolStripMenuItem
        '
        Me.YesToolStripMenuItem.Checked = True
        Me.YesToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.YesToolStripMenuItem.Name = "YesToolStripMenuItem"
        Me.YesToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.YesToolStripMenuItem.Text = "Yes"
        '
        'NoToolStripMenuItem
        '
        Me.NoToolStripMenuItem.Name = "NoToolStripMenuItem"
        Me.NoToolStripMenuItem.Size = New System.Drawing.Size(92, 22)
        Me.NoToolStripMenuItem.Text = "No"
        '
        'AutoOutputFoundResultsToolStripMenuItem
        '
        Me.AutoOutputFoundResultsToolStripMenuItem.Checked = True
        Me.AutoOutputFoundResultsToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked
        Me.AutoOutputFoundResultsToolStripMenuItem.Name = "AutoOutputFoundResultsToolStripMenuItem"
        Me.AutoOutputFoundResultsToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.AutoOutputFoundResultsToolStripMenuItem.Text = "Auto Output Found Results "
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(218, 6)
        '
        'CloseToolStripMenuItem
        '
        Me.CloseToolStripMenuItem.Name = "CloseToolStripMenuItem"
        Me.CloseToolStripMenuItem.Size = New System.Drawing.Size(221, 22)
        Me.CloseToolStripMenuItem.Text = "&Close"
        '
        'AddToolStripMenuItem
        '
        Me.AddToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StringToSearchToolStripMenuItem, Me.SetLastFileProcessToolStripMenuItem})
        Me.AddToolStripMenuItem.Name = "AddToolStripMenuItem"
        Me.AddToolStripMenuItem.Size = New System.Drawing.Size(41, 20)
        Me.AddToolStripMenuItem.Text = "&Add"
        '
        'StringToSearchToolStripMenuItem
        '
        Me.StringToSearchToolStripMenuItem.Name = "StringToSearchToolStripMenuItem"
        Me.StringToSearchToolStripMenuItem.Size = New System.Drawing.Size(175, 22)
        Me.StringToSearchToolStripMenuItem.Text = "&String To Search"
        '
        'SetLastFileProcessToolStripMenuItem
        '
        Me.SetLastFileProcessToolStripMenuItem.Name = "SetLastFileProcessToolStripMenuItem"
        Me.SetLastFileProcessToolStripMenuItem.Size = New System.Drawing.Size(175, 22)
        Me.SetLastFileProcessToolStripMenuItem.Text = "Set LastFile Process"
        '
        'ViewToolStripMenuItem
        '
        Me.ViewToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SearchListToolStripMenuItem, Me.LastFileFoundToolStripMenuItem})
        Me.ViewToolStripMenuItem.Name = "ViewToolStripMenuItem"
        Me.ViewToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.ViewToolStripMenuItem.Text = "&View"
        '
        'SearchListToolStripMenuItem
        '
        Me.SearchListToolStripMenuItem.Name = "SearchListToolStripMenuItem"
        Me.SearchListToolStripMenuItem.Size = New System.Drawing.Size(153, 22)
        Me.SearchListToolStripMenuItem.Text = "&Search List"
        '
        'LastFileFoundToolStripMenuItem
        '
        Me.LastFileFoundToolStripMenuItem.Name = "LastFileFoundToolStripMenuItem"
        Me.LastFileFoundToolStripMenuItem.Size = New System.Drawing.Size(153, 22)
        Me.LastFileFoundToolStripMenuItem.Text = "Last File Found"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(515, 41)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(61, 13)
        Me.Label4.TabIndex = 17
        Me.Label4.Text = "File Format:"
        '
        'txtFileFormat
        '
        Me.txtFileFormat.Location = New System.Drawing.Point(582, 38)
        Me.txtFileFormat.Name = "txtFileFormat"
        Me.txtFileFormat.Size = New System.Drawing.Size(51, 20)
        Me.txtFileFormat.TabIndex = 18
        Me.txtFileFormat.Text = "*.*"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(10, 90)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(63, 13)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Search List:"
        '
        'cmbSearchList
        '
        Me.cmbSearchList.DataSource = Me.SearchCategoryBindingSource
        Me.cmbSearchList.DisplayMember = "SearchName"
        Me.cmbSearchList.FormattingEnabled = True
        Me.cmbSearchList.Location = New System.Drawing.Point(112, 87)
        Me.cmbSearchList.Name = "cmbSearchList"
        Me.cmbSearchList.Size = New System.Drawing.Size(297, 21)
        Me.cmbSearchList.TabIndex = 20
        Me.cmbSearchList.ValueMember = "ID"
        '
        'SearchCategoryBindingSource
        '
        Me.SearchCategoryBindingSource.DataMember = "SearchCategory"
        Me.SearchCategoryBindingSource.DataSource = Me.FileminerDataSetBindingSource
        '
        'FileminerDataSetBindingSource
        '
        Me.FileminerDataSetBindingSource.DataSource = Me.FileminerDataSet
        Me.FileminerDataSetBindingSource.Position = 0
        '
        'FileminerDataSet
        '
        Me.FileminerDataSet.DataSetName = "fileminerDataSet"
        Me.FileminerDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'SearchCategoryTableAdapter
        '
        Me.SearchCategoryTableAdapter.ClearBeforeFill = True
        '
        'btnAdd
        '
        Me.btnAdd.AccessibleDescription = "Add New List"
        Me.btnAdd.BackgroundImage = CType(resources.GetObject("btnAdd.BackgroundImage"), System.Drawing.Image)
        Me.btnAdd.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnAdd.Location = New System.Drawing.Point(415, 85)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(23, 23)
        Me.btnAdd.TabIndex = 21
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        Me.btnEdit.AccessibleDescription = "Edit List"
        Me.btnEdit.BackgroundImage = CType(resources.GetObject("btnEdit.BackgroundImage"), System.Drawing.Image)
        Me.btnEdit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnEdit.Location = New System.Drawing.Point(444, 85)
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.Size = New System.Drawing.Size(23, 23)
        Me.btnEdit.TabIndex = 22
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.AccessibleDescription = "Delete List"
        Me.btnDelete.BackgroundImage = CType(resources.GetObject("btnDelete.BackgroundImage"), System.Drawing.Image)
        Me.btnDelete.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnDelete.Location = New System.Drawing.Point(473, 85)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(23, 23)
        Me.btnDelete.TabIndex = 23
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.ClientSize = New System.Drawing.Size(934, 665)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnEdit)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.cmbSearchList)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtFileFormat)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.Text = "File Miner String Search"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.SearchCategoryBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FileminerDataSetBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FileminerDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents ListBox2 As System.Windows.Forms.ListBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ToolStripProgressBar
    Friend WithEvents tsLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CloseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AddToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StringToSearchToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ViewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SearchListToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents SetLastFileProcessToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents LastFileFoundToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SaveLastFileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents YesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AutoOutputFoundResultsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtFileFormat As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents cmbSearchList As System.Windows.Forms.ComboBox
    Friend WithEvents FileminerDataSetBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents FileminerDataSet As FileMiner.fileminerDataSet
    Friend WithEvents SearchCategoryBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents SearchCategoryTableAdapter As FileMiner.fileminerDataSetTableAdapters.SearchCategoryTableAdapter
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents btnEdit As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button

End Class
