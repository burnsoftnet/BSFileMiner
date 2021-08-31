Imports System.Configuration
Imports System.Data.Odbc
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Text
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing
Imports DocumentFormat.OpenXml.Spreadsheet
Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.parser
Imports Microsoft.Office.Interop.Word

''' <summary>
''' Class frmMain.
''' </summary>
Public Class FrmMain
    ''' <summary>
    ''' The use counter
    ''' </summary>
    Public UseCounter As Boolean
    ''' <summary>
    ''' The bugger file
    ''' </summary>
    Public BuggerFile As String
    ''' <summary>
    ''' The error file
    ''' </summary>
    Public ErrorFile As String
    ''' <summary>
    ''' The usedebug
    ''' </summary>
    Public Usedebug As Boolean
    ''' <summary>
    ''' The useerror
    ''' </summary>
    Public Useerror As Boolean
    ''' <summary>
    ''' The continue from last
    ''' </summary>
    Public ContinueFromLast As Boolean
    ''' <summary>
    ''' My last file
    ''' </summary>
    Public MyLastFile As String
    ''' <summary>
    ''' The save results filename
    ''' </summary>
    Public SaveResultsFilename As String
#Region "My.Settings Functions and Subs"
    ''' <summary>
    ''' Lasts the file exists.
    ''' </summary>
    ''' <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
    Public Function LastFileExists() As Boolean
        Dim bAns = False
        Dim sValue As String = My.Settings.LastFile
        If sValue.Length > 0 Then bAns = True
        Return bAns
    End Function
    ''' <summary>
    ''' Gets the last file.
    ''' </summary>
    ''' <returns>System.String.</returns>
    Public Function GetLastFile() As String
' ReSharper disable once RedundantAssignment
        Dim sAns = ""
        sAns = My.Settings.LastFile
        Return sAns
    End Function
    ''' <summary>
    ''' Saves the last file.
    ''' </summary>
    ''' <param name="sValue">The s value.</param>
    Sub SaveLastFile(sValue As String)
        If YesToolStripMenuItem.Checked Then
            My.Settings.LastFile = sValue
            My.Settings.Save()
        End If
    End Sub
#End Region
    ''' <summary>
    ''' Initializes this instance.
    ''' </summary>
    Sub Init()
        BuggerFile = System.Windows.Forms.Application.StartupPath & "\fileminer.debug.log"
        ErrorFile = System.Windows.Forms.Application.StartupPath & "\fileminer.err.log"
        Usedebug = CBool(ConfigurationManager.AppSettings("DEBUG"))
        Useerror = CBool(ConfigurationManager.AppSettings("ERROR"))

    End Sub
#Region "File Search and Decode Functions and Subs"
    ''' <summary>
    ''' Contentses the exist.
    ''' </summary>
    ''' <param name="sContent">Content of the s.</param>
    ''' <param name="sSearchFor">The s search for.</param>
    ''' <param name="foundAtIndex">Index of the found at.</param>
    ''' <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
    Function ContentsExist(sContent As String, sSearchFor As String, Optional ByRef foundAtIndex As Integer = 0) As Boolean
        Dim bAns = False
        Dim index As Integer = sContent.IndexOf(sSearchFor)
        If index >= 0 Then
            bAns = True
        Else
            bAns = False
        End If
        Return bAns
    End Function
    ''' <summary>
    ''' Contentses the exist regex.
    ''' </summary>
    ''' <param name="sContent">Content of the s.</param>
    ''' <param name="sSearchFor">The s search for.</param>
    ''' <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
    Function ContentsExistRegex(sContent As String, sSearchFor As String) As Boolean
        Dim bAns = False
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
        Dim sAns = ""
        Try
            Dim obj As New Database
            Dim sTemp = ""
            Call obj.ConnectDb()
            Dim sql As String = "SELECT * from SearchStrings where SCID=" & cmbSearchList.SelectedValue
            Dim sWord = ""
            Dim cmd As New OdbcCommand(sql, obj.Conn)
            Dim rs As OdbcDataReader
            rs = cmd.ExecuteReader
            While rs.Read
                sWord = rs("SearchString")
                If ContentsExistRegex(sContents, sWord & "[^ ,]*") Then
                    If sTemp.Length = 0 Then
                        sTemp = sWord
                    Else
                        sTemp &= "," & sWord
                    End If
                End If
            End While
            rs.Close()
            rs = Nothing
            cmd = Nothing
            obj.CloseDb()
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
        Dim sAns = ""
        Try
            Dim spreadsheetDocument As SpreadsheetDocument = SpreadsheetDocument.Open(sFile, False)
            Dim workbookPart As WorkbookPart = spreadsheetDocument.WorkbookPart
            Dim shareStringPart As SharedStringTablePart = workbookPart.SharedStringTablePart
            Dim paragraphText As New StringBuilder()
            For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
                paragraphText.Append(item.InnerText) 'should read all strings
            Next
            sAns = paragraphText.ToString()
            spreadsheetDocument.Close()
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
        Dim sAns = ""
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
        Dim sAns = ""
        Try
            Dim objApp As Application
            Dim objDoc As Microsoft.Office.Interop.Word.Document
            objApp = New Application()
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
    Function OpenFileAsPdf(sFile As String) As String
        Dim sAns = ""
        Try
            Dim oReader As New PdfReader(sFile)
            Dim stringOut = New StringBuilder()
            If File.Exists(sFile) Then
                For i = 1 To oReader.NumberOfPages
                    Dim itsText As New SimpleTextExtractionStrategy
                    stringOut.Append(PdfTextExtractor.GetTextFromPage(oReader, i, itsText))
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
    ''' <summary>
    ''' Opens the file as text.
    ''' </summary>
    ''' <param name="sFile">The s file.</param>
    ''' <returns>System.String.</returns>
    Function OpenFileAsText(sFile As String) As String
        Dim sAns = ""
        Try
            Dim reader = New StreamReader(sFile)
            sAns = reader.ReadToEnd
            reader.Close()
            reader = Nothing
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
            Dim tempStr(1) As String
            Dim tempNode As ListViewItem
            Dim fileName As String
            Dim fileExt = ""
            Dim objOf As New OtherFunctions
            Dim lCatid As Long = cmbSearchList.SelectedValue
            Dim sSearchFor As String = objOf.BuildSearchString(lCatid)
            Dim sContents = ""
            Dim foundCrap = False
            Dim foundTheseWords As String
            For Each fileName In fileEntries
                If UseCounter Then ProgressBar1.Value += 1
                If Not ContinueFromLast Then
                    fileExt = GetExtOfFile(fileName)
                    Debug.Print(fileName)
                    Call BuggerMe(fileName)
                    tsLabel.Text = fileName
                    StatusStrip1.Refresh()
                    Call SaveLastFile(fileName)
                    Select Case LCase(fileExt)
                        Case ".pdf"
                            sContents = OpenFileAsPdf(fileName)
                            foundCrap = ContentsExistRegex(sContents, sSearchFor)
                        Case ".docx"
                            sContents = OpenFileAsWordProcessing(fileName)
                            foundCrap = ContentsExistRegex(sContents, sSearchFor)
                        Case ".doc"
                            sContents = OpenFileAsWordOlder(fileName)
                            foundCrap = ContentsExistRegex(sContents, sSearchFor)
                        Case ".xlsx"
                            sContents = OpenFileAsExcel(fileName)
                            foundCrap = ContentsExistRegex(sContents, sSearchFor)
                        Case Else
                            sContents = OpenFileAsText(fileName)
                            foundCrap = ContentsExistRegex(sContents, sSearchFor)
                    End Select
                    If foundCrap Then
                        foundTheseWords = ""
                        foundTheseWords = GetWordsFound(sContents)
                        Call BuggerMe("FOUND! " & fileName & vbTab & foundTheseWords)
                        If AutoOutputFoundResultsToolStripMenuItem.Checked Then Call SaveResults(fileName & vbTab & foundTheseWords)
                        tempStr(0) = fileName
                        tempStr(1) = foundTheseWords
                        tempNode = New ListViewItem(tempStr)
                        ListView1.Items.Add(tempNode)

                    End If
                    ListView1.Refresh()
                    Refresh()
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
    ''' <summary>
    ''' Processes the file.
    ''' </summary>
    ''' <param name="path">The path.</param>
    Public Shared Sub ProcessFile(ByVal path As String)
        Debug.Print(path)
    End Sub 'ProcessFile    
    ''' <summary>
    ''' Gets the ext of file.
    ''' </summary>
    ''' <param name="sFile">The s file.</param>
    ''' <returns>System.String.</returns>
    Public Function GetExtOfFile(ByVal sFile As String) As String
        Dim sAns = ""
        sAns = Path.GetExtension(sFile)
        Return sAns
    End Function
#End Region
    ''' <summary>
    ''' Handles the 1 event of the Button1_Click control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.TextLength = 0 Then MsgBox("Please put in a path!") : Exit Sub
        ProgressBar1.Visible = True
        Me.Cursor = Cursors.WaitCursor
        If AutoOutputFoundResultsToolStripMenuItem.Checked Then
            SaveResultsFilename = GenerateSaveFileName()
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
            Dim counter = Directory.GetFiles(path, txtFileFormat.Text, SearchOption.AllDirectories)
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
    ''' <summary>
    ''' Handles the Click event of the Button2 control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If

    End Sub
    ''' <summary>
    ''' Handles the CheckedChanged event of the CheckBox1 control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            UseCounter = True
        Else
            UseCounter = False
        End If
    End Sub
#Region "Error and Debugging Functions and Subs"
    ''' <summary>
    ''' Creates the file.
    ''' </summary>
    ''' <param name="sPath">The s path.</param>
    Public Sub CreateFile(ByVal sPath As String)
        If Not File.Exists(sPath) Then
            Dim fs As New FileStream(sPath, FileMode.Append, FileAccess.Write, FileShare.Write)
            fs.Close()
        End If
    End Sub
    ''' <summary>
    ''' Appends to file.
    ''' </summary>
    ''' <param name="sPath">The s path.</param>
    ''' <param name="sNewLine">The s new line.</param>
    Public Sub AppendToFile(ByVal sPath As String, ByVal sNewLine As String)
        If Not File.Exists(sPath) Then Call CreateFile(sPath)
        Dim sw As New StreamWriter(sPath, True, Encoding.ASCII)
        sw.WriteLine(sNewLine)
        sw.Close()
    End Sub
    ''' <summary>
    ''' Buggers me.
    ''' </summary>
    ''' <param name="sValue">The s value.</param>
    Sub BuggerMe(sValue As String)
        If Usedebug Then
            Dim sline As String = Now & vbTab & sValue
            Call AppendToFile(BuggerFile, sline)
        End If
    End Sub
    ''' <summary>
    ''' Errors the found.
    ''' </summary>
    ''' <param name="sValue">The s value.</param>
    Sub ErrorFound(sValue As String)
        If Useerror Then
            Dim sline As String = Now & vbTab & sValue
            Call AppendToFile(ErrorFile, sline)
        End If
    End Sub
    ''' <summary>
    ''' Generates the name of the save file.
    ''' </summary>
    ''' <returns>System.String.</returns>
    Public Function GenerateSaveFileName() As String
        Dim sAns = ""
        Dim curTime As String = Replace(Now.ToLongDateString, ",", "")
        curTime = Replace(curTime, " ", "_")
        Dim sPath As String = Replace(Replace(TextBox1.Text, "\", "_", ), ":", "")
        Dim fileName As String = System.Windows.Forms.Application.StartupPath & "\FileMiner_Results_" & sPath & "_" & curTime & ".log"
        sAns = fileName
        Return sAns
    End Function
    ''' <summary>
    ''' Saves the results.
    ''' </summary>
    ''' <param name="sValue">The s value.</param>
    Public Sub SaveResults(sValue As String)    
        Call AppendToFile(SaveResultsFilename, sValue)
    End Sub
#End Region
    ''' <summary>
    ''' Handles the Click event of the Button3 control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim curTime As String = Replace(Now.ToLongDateString, ",", "")
        curTime = Replace(curTime, " ", "_")
        Dim sPath As String = Replace(Replace(TextBox1.Text, "\", "_", ), ":", "")
        Dim fileName As String = "FileMiner_Results_" & sPath & "_" & curTime & ".csv"
        SaveFileDialog1.FilterIndex = 1
        SaveFileDialog1.Filter = "Comma Seperated File File (.csv)|*.csv"
        SaveFileDialog1.Title = "Export Files Found to CSV File"
        SaveFileDialog1.FileName = "filename"
        If SaveFileDialog1.ShowDialog = DialogResult.Cancel Then Exit Sub
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
    ''' <summary>
    ''' Handles the Click event of the StringToSearchToolStripMenuItem control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub StringToSearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StringToSearchToolStripMenuItem.Click
        Dim frmNew As New frmAddSearchString
        frmNew.MdiParent = Me.Parent
        frmNew.Show()
    End Sub
    ''' <summary>
    ''' Handles the Click event of the SearchListToolStripMenuItem control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub SearchListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchListToolStripMenuItem.Click
        Dim frmNew As New frmViewList
        frmNew.MdiParent = Parent
        frmNew.Show()
    End Sub
    ''' <summary>
    ''' Handles the Click event of the CloseToolStripMenuItem control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Close()
    End Sub
    ''' <summary>
    ''' Loads the data.
    ''' </summary>
    Sub LoadData()
        SearchCategoryTableAdapter.Fill(FileminerDataSet.SearchCategory)
        Refresh()
    End Sub
    ''' <summary>
    ''' Handles the Load event of the frmMain control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Call LoadData()
        Call Init()
        ProgressBar1.Visible = False
        If My.Settings.SaveLastFile Then
            YesToolStripMenuItem.Checked = True
            NoToolStripMenuItem.Checked = False
        Else
            NoToolStripMenuItem.Checked = True
            YesToolStripMenuItem.Checked = False
        End If
    End Sub
    ''' <summary>
    ''' Handles the Click event of the SetLastFileProcessToolStripMenuItem control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub SetLastFileProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetLastFileProcessToolStripMenuItem.Click
        Dim sAns As String = InputBox("Set Last File")
        Call SaveLastFile(sAns)
    End Sub
    ''' <summary>
    ''' Handles the Click event of the LastFileFoundToolStripMenuItem control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub LastFileFoundToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LastFileFoundToolStripMenuItem.Click
        Dim sAns As String = GetLastFile()
        If sAns.Length = 0 Then
            MsgBox("No File Found!")
        Else
            MsgBox(GetLastFile)
        End If
    End Sub
    ''' <summary>
    ''' Handles the Click event of the YesToolStripMenuItem control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub YesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles YesToolStripMenuItem.Click
        My.Settings.SaveLastFile = True
        My.Settings.Save()
        YesToolStripMenuItem.Checked = True
        NoToolStripMenuItem.Checked = False
    End Sub
    ''' <summary>
    ''' Handles the Click event of the NoToolStripMenuItem control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub NoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NoToolStripMenuItem.Click
        My.Settings.SaveLastFile = False
        My.Settings.Save()
        YesToolStripMenuItem.Checked = False
        NoToolStripMenuItem.Checked = True
        Call SaveLastFile("")
    End Sub
    ''' <summary>
    ''' Handles the Click event of the AutoOutputFoundResultsToolStripMenuItem control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub AutoOutputFoundResultsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoOutputFoundResultsToolStripMenuItem.Click
        If AutoOutputFoundResultsToolStripMenuItem.Checked Then
            AutoOutputFoundResultsToolStripMenuItem.Checked = False
        Else
            AutoOutputFoundResultsToolStripMenuItem.Checked = True
        End If
        My.Settings.AutoSaveResults = AutoOutputFoundResultsToolStripMenuItem.Checked
        My.Settings.Save()
    End Sub
    ''' <summary>
    ''' Handles the Click event of the btnEdit control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        Dim frmNew As New frmViewList
        Dim lCat As Long = cmbSearchList.SelectedValue
        frmNew.lCATID = lCat
        frmNew.Parent = Me.Parent
        frmNew.Show()
    End Sub
    ''' <summary>
    ''' Handles the Click event of the btnDelete control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Dim sAns As String = MsgBox("Are you sure that you wish to delete " & cmbSearchList.SelectedText & "?", vbYesNo)
        If sAns = vbYes Then
            Dim lCat As Long = cmbSearchList.SelectedValue
            Dim sql As String = "DELETE from SearchStrings where SCID=" & lCat
            Dim obj As New Database
            obj.ConnExec(sql)
            sql = "DELETE from SearchCategory where ID=" & lCat
            obj.ConnExec(sql)
            obj = Nothing
            Call LoadData()
        End If
    End Sub
    ''' <summary>
    ''' Handles the Click event of the btnAdd control.
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        Dim sAns As String = InputBox("New Search List Name:", "Add New Search List")
        Dim obj As New Database
        Dim objOf As New OtherFunctions
        If Not obj.SearchCategoryExists(objOf.FC(sAns)) Then
            Dim sql As String = "INSERT INTO SearchCategory (SearchName) VALUES ('" & objOf.FC(sAns) & "')"
            obj.ConnExec(sql)
        End If
        Dim lCat As Long = objOf.GetSearchListID(objOf.FC(sAns))
        Call LoadData()
        Dim frmNew As New frmAddSearchString
        frmNew.LCat = lCat
        frmNew.Parent = Me.Parent
        frmNew.Show()
        objOf = Nothing
        obj = Nothing
    End Sub
End Class
