Imports System.IO
Imports System.Text.RegularExpressions
Imports iTextSharp
Imports System.Text
Imports System.Xml.XmlReader
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml
Imports System.Linq
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Public Class frmMain
    Public UseCounter As Boolean
    Public BuggerFile As String
    Public ErrorFile As String
    Public USEDEBUG As Boolean
    Public USEERROR As Boolean
    Public ContinueFromLast As Boolean
    Public MyLastFile As String
    Public SaveResults_filename As String
#Region "My.Settings Functions and Subs"
    Public Function LastFileExists() As Boolean
        Dim bAns As Boolean = False
        Dim sValue As String = My.Settings.LastFile
        If sValue.Length > 0 Then bAns = True
        Return bAns
    End Function
    Public Function GetLastFile() As String
        Dim sAns As String = ""
        sAns = My.Settings.LastFile
        Return sAns
    End Function
    Sub SaveLastFile(sValue As String)
        If YesToolStripMenuItem.Checked Then
            My.Settings.LastFile = sValue
            My.Settings.Save()
        End If
    End Sub
#End Region
    Sub INIT()
        BuggerFile = System.Windows.Forms.Application.StartupPath & "\fileminer.debug.log"
        ErrorFile = System.Windows.Forms.Application.StartupPath & "\fileminer.err.log"
        USEDEBUG = CBool(System.Configuration.ConfigurationManager.AppSettings("DEBUG"))
        USEERROR = CBool(System.Configuration.ConfigurationManager.AppSettings("ERROR"))

    End Sub
#Region "File Search and Decode Functions and Subs"
    Function ContentsExist(sContent As String, sSearchFor As String, Optional ByRef FoundAtIndex As Integer = 0) As Boolean
        Dim bAns As Boolean = False
        Dim Index As Integer = sContent.IndexOf(sSearchFor)
        If Index >= 0 Then
            bAns = True
        Else
            bAns = False
        End If
        Return bAns
    End Function
    Function ContentsExistRegex(sContent As String, sSearchFor As String) As Boolean
        Dim bAns As Boolean = False
        If (Regex.IsMatch(sContent, sSearchFor, RegexOptions.IgnoreCase)) Then
            bAns = True
        Else
            bAns = False
        End If

        Return bAns
    End Function
    ''' <summary>
    ''' Combine the Search strings relating to the selected category and format the string in a Regular expression format
    ''' Then Compair it to the Contents and return the word/words that matchup in the file.
    ''' </summary>
    ''' <param name="sContents"></param>
    ''' <returns>String</returns>
    Function GetWordsFound(sContents As String) As String
        Dim sAns As String = ""
        Try
            Dim Obj As New Database
            Dim sTemp As String = ""
            Call Obj.ConnectDb()
            Dim SQL As String = "SELECT * from SearchStrings where SCID=" & cmbSearchList.SelectedValue
            Dim sWord As String = ""
            Dim CMD As New Odbc.OdbcCommand(SQL, Obj.Conn)
            Dim RS As Odbc.OdbcDataReader
            RS = CMD.ExecuteReader
            While RS.Read
                sWord = RS("SearchString")
                If ContentsExistRegex(sContents, sWord & "[^ ,]*") Then
                    If sTemp.Length = 0 Then
                        sTemp = sWord
                    Else
                        sTemp &= "," & sWord
                    End If
                End If
            End While
            RS.Close()
            RS = Nothing
            CMD = Nothing
            Obj.CloseDb()
            sAns = sTemp
        Catch ex As Exception
            Call ErrorFound("GetWordsFound" & vbTab & Err.Number & vbTab & ex.Message.ToString)
        End Try
        Return sAns
    End Function
    ''' <summary>
    ''' Excel 2007 Processing
    ''' </summary>
    ''' <param name="sFile"></param>
    ''' <returns>String</returns>
    Function OpenFileAsExcel(sFile As String) As String
        Dim sAns As String = ""
        Try
            Dim spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(sFile, False)
            Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
            Dim shareStringPart As SharedStringTablePart = workbookPart.SharedStringTablePart
            Dim paragraphText As New StringBuilder()
            For Each Item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
                paragraphText.Append(Item.InnerText) 'should read all strings
            Next
            sAns = paragraphText.ToString()
            spreadsheetDocument.Close()
            spreadsheetDocument = Nothing
        Catch ex As Exception
            ListBox2.Items.Add(sFile & " " & Err.Description)
            Call ErrorFound("OpenFileAsExcel" & sFile & " " & vbTab & Err.Number & vbTab & ex.Message.ToString)
        End Try
        Return sAns
    End Function
    ''' <summary>
    ''' Word 2007 Processing
    ''' </summary>
    ''' <param name="sFile"></param>
    ''' <returns>String</returns>
    Function OpenFileAsWordProcessing(sFile As String) As String
        Dim sAns As String = ""
        Try
            Dim stream As Stream = File.Open(sFile, FileMode.Open)
            Dim wordprocessingDocument As WordprocessingDocument = WordprocessingDocument.Open(stream, False)
            Dim body As Body = wordprocessingDocument.MainDocumentPart.Document.Body
            sAns = body.InnerText.ToString
            wordprocessingDocument.Close()
            wordprocessingDocument = Nothing
        Catch ex As Exception
            ListBox2.Items.Add(sFile & " " & Err.Description)
            Call ErrorFound("OpenFileAsWordProcessing" & sFile & " " & vbTab & Err.Number & vbTab & ex.Message.ToString)
        End Try
        Return sAns
    End Function
    ''' <summary>
    ''' Word 2003 Processing and maybe older
    ''' </summary>
    ''' <param name="sFile"></param>
    ''' <returns>String</returns>
    Function OpenFileAsWordOlder(sFile As String) As String
        Dim sAns As String = ""
        Try
            Dim objApp As Microsoft.Office.Interop.Word.Application
            Dim objDoc As Microsoft.Office.Interop.Word.Document
            objApp = New Microsoft.Office.Interop.Word.Application()
            objDoc = objApp.Documents.Open(sFile, , True)
            sAns = objDoc.Content.Text
            objDoc.Close()
            objApp.Quit()
            objDoc = Nothing
            objDoc = Nothing
        Catch ex As Exception
            ListBox2.Items.Add(sFile & " " & Err.Description)
            Call ErrorFound("OpenFileAsWordOlder" & sFile & " " & vbTab & Err.Number & vbTab & ex.Message.ToString)
        End Try
        Return sAns
    End Function
    ''' <summary>
    ''' PDF Processing
    ''' </summary>
    ''' <param name="sFile"></param>
    ''' <returns></returns>
    Function OpenFileAsPDF(sFile As String) As String
        Dim sAns As String = ""
        Try
            Dim oReader As New iTextSharp.text.pdf.PdfReader(sFile)
            Dim stringOut As StringBuilder = New StringBuilder()
            If File.Exists(sFile) Then
                For i = 1 To oReader.NumberOfPages
                    Dim itsText As New iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
                    stringOut.Append(iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(oReader, i, itsText))
                Next
            End If
            sAns = stringOut.ToString()
            oReader.Close()
            oReader = Nothing
        Catch ex As Exception
            ListBox2.Items.Add(sFile & " " & Err.Description)
            Call ErrorFound("OpenFileAsPDF" & sFile & " " & vbTab & Err.Number & vbTab & ex.Message.ToString)
        End Try
        Return sAns
    End Function
    Function OpenFileAsText(sFile As String) As String
        Dim sAns As String = ""
        Try
            Dim Reader As StreamReader = New StreamReader(sFile)
            sAns = Reader.ReadToEnd
            Reader.Close()
            Reader = Nothing
        Catch ex As Exception
            ListBox2.Items.Add(sFile & " " & Err.Description)
        End Try
        Return sAns
    End Function
    ''' <summary>
    '''  Process the list of files found in the directory
    ''' </summary>
    ''' <param name="targetDirectory"></param>
    Public Sub ProcessDirectory(ByVal targetDirectory As String)
        Call BuggerMe("Started Processing " & targetDirectory)
        Try
            Dim fileEntries As String() = Directory.GetFiles(targetDirectory, txtFileFormat.Text)
            Dim TempStr(1) As String
            Dim TempNode As ListViewItem
            Dim fileName As String
            Dim FileExt As String = ""
            Dim ObjOF As New OtherFunctions
            Dim lCATID As Long = cmbSearchList.SelectedValue
            Dim sSearchFor As String = ObjOF.BuildSearchString(lCATID)
            Dim sContents As String = ""
            Dim FoundCrap As Boolean = False
            Dim FoundTheseWords As String
            For Each fileName In fileEntries
                If UseCounter Then ProgressBar1.Value += 1
                If Not ContinueFromLast Then
                    FileExt = GetExtOfFile(fileName)
                    Debug.Print(fileName)
                    Call BuggerMe(fileName)
                    tsLabel.Text = fileName
                    StatusStrip1.Refresh()
                    Call SaveLastFile(fileName)
                    Select Case LCase(FileExt)
                        Case ".pdf"
                            sContents = OpenFileAsPDF(fileName)
                            FoundCrap = ContentsExistRegex(sContents, sSearchFor)
                        Case ".docx"
                            sContents = OpenFileAsWordProcessing(fileName)
                            FoundCrap = ContentsExistRegex(sContents, sSearchFor)
                        Case ".doc"
                            sContents = OpenFileAsWordOlder(fileName)
                            FoundCrap = ContentsExistRegex(sContents, sSearchFor)
                        Case ".xlsx"
                            sContents = OpenFileAsExcel(fileName)
                            FoundCrap = ContentsExistRegex(sContents, sSearchFor)
                        Case Else
                            sContents = OpenFileAsText(fileName)
                            FoundCrap = ContentsExistRegex(sContents, sSearchFor)
                    End Select
                    If FoundCrap Then
                        FoundTheseWords = ""
                        FoundTheseWords = GetWordsFound(sContents)
                        Call BuggerMe("FOUND! " & fileName & vbTab & FoundTheseWords)
                        If AutoOutputFoundResultsToolStripMenuItem.Checked Then Call SaveResults(fileName & vbTab & FoundTheseWords)
                        TempStr(0) = fileName
                        TempStr(1) = FoundTheseWords
                        TempNode = New ListViewItem(TempStr)
                        ListView1.Items.Add(TempNode)

                    End If
                    ListView1.Refresh()
                    Me.Refresh()
                Else
                    Call BuggerMe("SKIPPING: " & fileName)
                    If fileName = MyLastFile Then ContinueFromLast = False
                End If
            Next fileName
            Dim subdirectoryEntries As String() = Directory.GetDirectories(targetDirectory)
            ' Recurse into subdirectories of this directory. 
            Dim subdirectory As String
            For Each subdirectory In subdirectoryEntries
                ProcessDirectory(subdirectory)
            Next subdirectory
        Catch ex As Exception
            Call ErrorFound("ACCESS DENIED! " & targetDirectory & " " & vbTab & Err.Number & vbTab & ex.Message.ToString)
        End Try
        Call BuggerMe("End Processing " & targetDirectory)
    End Sub 'ProcessDirectory
    ' Insert logic for processing found files here. 
    Public Shared Sub ProcessFile(ByVal path As String)
        Debug.Print(path)
    End Sub 'ProcessFile
    Public Function GetExtOfFile(ByVal sFile As String) As String
        Dim sAns As String = ""
        sAns = Path.GetExtension(sFile)
        Return sAns
    End Function
#End Region

    Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If TextBox1.TextLength = 0 Then MsgBox("Please put in a path!") : Exit Sub
        ProgressBar1.Visible = True
        Me.Cursor = Cursors.WaitCursor
        If AutoOutputFoundResultsToolStripMenuItem.Checked Then
            SaveResults_filename = GenerateSaveFileName()
        End If
        ListView1.ShowItemToolTips = True
        ListView1.View = System.Windows.Forms.View.Details
        ListView1.Columns.Clear()
        ListView1.Columns.Add("Path", 700)
        ListView1.Columns.Add("Key Word(s)", 200)
        ListView1.Items.Clear()
        Dim path As String = TextBox1.Text
        ContinueFromLast = LastFileExists()
        If ContinueFromLast Then MyLastFile = GetLastFile()
        If NoToolStripMenuItem.Checked Then ContinueFromLast = False
        If UseCounter Then
            Dim counter = IO.Directory.GetFiles(path, txtFileFormat.Text, SearchOption.AllDirectories)
            ProgressBar1.Minimum = 0
            ProgressBar1.Maximum = counter.Count
        End If
        If File.Exists(path) Then
            ' This path is a file.
            ProcessFile(path)
        Else
            If Directory.Exists(path) Then
                ' This path is a directory.
                Call ProcessDirectory(path)
            Else
                MsgBox("{0} is not a valid file or directory.", path)
            End If
        End If
        Me.Cursor = Cursors.Default
        MsgBox("end of processing")
        Call SaveLastFile("")
        ProgressBar1.Visible = False
        tsLabel.Text = ""
        If UseCounter Then ProgressBar1.Value = 0
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If

    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            UseCounter = True
        Else
            UseCounter = False
        End If
    End Sub
#Region "Error and Debugging Functions and Subs"
    Public Sub CreateFile(ByVal sPath As String)
        If Not File.Exists(sPath) Then
            Dim fs As New FileStream(sPath, FileMode.Append, FileAccess.Write, FileShare.Write)
            fs.Close()
        End If
    End Sub
    Public Sub AppendToFile(ByVal sPath As String, ByVal sNewLine As String)
        If Not File.Exists(sPath) Then Call CreateFile(sPath)
        Dim sw As New StreamWriter(sPath, True, Encoding.ASCII)
        sw.WriteLine(sNewLine)
        sw.Close()
    End Sub
    Sub BuggerMe(sValue As String)
        If USEDEBUG Then
            Dim sline As String = Now & vbTab & sValue
            Call AppendToFile(BuggerFile, sline)
        End If
    End Sub
    Sub ErrorFound(sValue As String)
        If USEERROR Then
            Dim sline As String = Now & vbTab & sValue
            Call AppendToFile(ErrorFile, sline)
        End If
    End Sub
    Public Function GenerateSaveFileName() As String
        Dim sAns As String = ""
        Dim CurTime As String = Replace(Now.ToLongDateString, ",", "")
        CurTime = Replace(CurTime, " ", "_")
        Dim sPath As String = Replace(Replace(TextBox1.Text, "\", "_", ), ":", "")
        Dim FileName As String = System.Windows.Forms.Application.StartupPath & "\FileMiner_Results_" & sPath & "_" & CurTime & ".log"
        sAns = FileName
        Return sAns
    End Function
    Public Sub SaveResults(sValue As String)    
        Call AppendToFile(SaveResults_filename, sValue)
    End Sub
#End Region
   
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim CurTime As String = Replace(Now.ToLongDateString, ",", "")
        CurTime = Replace(CurTime, " ", "_")
        Dim sPath As String = Replace(Replace(TextBox1.Text, "\", "_", ), ":", "")
        Dim FileName As String = "FileMiner_Results_" & sPath & "_" & CurTime & ".csv"
        SaveFileDialog1.FilterIndex = 1
        SaveFileDialog1.Filter = "Comma Seperated File File (.csv)|*.csv"
        SaveFileDialog1.Title = "Export Files Found to CSV File"
        SaveFileDialog1.FileName = "filename"
        If SaveFileDialog1.ShowDialog = System.Windows.Forms.DialogResult.Cancel Then Exit Sub
        Dim sFile As String = SaveFileDialog1.FileName

        Using writer As New StreamWriter(sFile)
            writer.Write("Search In " & TextBox1.Text)
            writer.WriteLine()
            writer.Write("File & Path,Word(s) Found")
            writer.WriteLine()
            For Each item As ListViewItem In ListView1.Items
                For index = 0 To item.SubItems.Count - 1
                    If index > 0 Then
                        writer.Write(",")
                    End If
                    writer.Write(item.SubItems(index).Text)
                Next
                writer.WriteLine()
            Next
        End Using
        MsgBox("Export Completed!")
    End Sub

    Private Sub StringToSearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StringToSearchToolStripMenuItem.Click
        Dim frmNew As New frmAddSearchString
        frmNew.MdiParent = Me.Parent
        frmNew.Show()
    End Sub

    Private Sub SearchListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchListToolStripMenuItem.Click
        Dim frmNew As New frmViewList
        frmNew.MdiParent = Me.Parent
        frmNew.Show()
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Me.Close()
    End Sub
    Sub LoadData()
        Me.SearchCategoryTableAdapter.Fill(Me.FileminerDataSet.SearchCategory)
        Me.Refresh()
    End Sub
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call LoadData()
        Call INIT()
        ProgressBar1.Visible = False
        If My.Settings.SaveLastFile Then
            YesToolStripMenuItem.Checked = True
            NoToolStripMenuItem.Checked = False
        Else
            NoToolStripMenuItem.Checked = True
            YesToolStripMenuItem.Checked = False
        End If
    End Sub

    Private Sub SetLastFileProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetLastFileProcessToolStripMenuItem.Click
        Dim sAns As String = InputBox("Set Last File")
        Call SaveLastFile(sAns)
    End Sub

    Private Sub LastFileFoundToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LastFileFoundToolStripMenuItem.Click
        Dim sAns As String = GetLastFile()
        If sAns.Length = 0 Then
            MsgBox("No File Found!")
        Else
            MsgBox(GetLastFile)
        End If
    End Sub

    Private Sub YesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles YesToolStripMenuItem.Click
        My.Settings.SaveLastFile = True
        My.Settings.Save()
        YesToolStripMenuItem.Checked = True
        NoToolStripMenuItem.Checked = False
    End Sub

    Private Sub NoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NoToolStripMenuItem.Click
        My.Settings.SaveLastFile = False
        My.Settings.Save()
        YesToolStripMenuItem.Checked = False
        NoToolStripMenuItem.Checked = True
        Call SaveLastFile("")
    End Sub

    Private Sub AutoOutputFoundResultsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoOutputFoundResultsToolStripMenuItem.Click
        If AutoOutputFoundResultsToolStripMenuItem.Checked Then
            AutoOutputFoundResultsToolStripMenuItem.Checked = False
        Else
            AutoOutputFoundResultsToolStripMenuItem.Checked = True
        End If
        My.Settings.AutoSaveResults = AutoOutputFoundResultsToolStripMenuItem.Checked
        My.Settings.Save()
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Dim frmNew As New frmViewList
        Dim lCAT As Long = cmbSearchList.SelectedValue
        frmNew.lCATID = lCAT
        frmNew.Parent = Me.Parent
        frmNew.Show()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim sAns As String = MsgBox("Are you sure that you wish to delete " & cmbSearchList.SelectedText & "?", vbYesNo)
        If sAns = vbYes Then
            Dim lCAT As Long = cmbSearchList.SelectedValue
            Dim SQL As String = "DELETE from SearchStrings where SCID=" & lCAT
            Dim Obj As New Database
            Obj.ConnExec(SQL)
            SQL = "DELETE from SearchCategory where ID=" & lCAT
            Obj.ConnExec(SQL)
            Obj = Nothing
            Call LoadData()
        End If
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim sAns As String = InputBox("New Search List Name:", "Add New Search List")
        Dim Obj As New Database
        Dim ObjOF As New OtherFunctions
        If Not Obj.SearchCategoryExists(ObjOF.FC(sAns)) Then
            Dim SQL As String = "INSERT INTO SearchCategory (SearchName) VALUES ('" & ObjOF.FC(sAns) & "')"
            Obj.ConnExec(SQL)
        End If
        Dim lCAT As Long = ObjOF.GetSearchListID(ObjOF.FC(sAns))
        Call LoadData()
        Dim frmNew As New frmAddSearchString
        frmNew.lCAT = lCAT
        frmNew.Parent = Me.Parent
        frmNew.Show()
        ObjOF = Nothing
        Obj = Nothing
    End Sub
End Class
