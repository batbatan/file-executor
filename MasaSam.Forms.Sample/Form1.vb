Imports Scripting
Imports System
Imports System.Text
Imports System.Drawing
Imports System.Drawing.Image
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Text.RegularExpressions
Imports System.Data
Imports System.Windows.Forms
Imports System.IO
Imports System.Drawing.Drawing2D
Imports MasaSam
Imports System.Runtime.InteropServices
Imports SharpCompress.Archive
Imports SharpCompress.Common
Imports Microsoft.Win32

Public Class Form1

    Public Sub Form1()
        InitializeComponent()
    End Sub

    Private m_PanStartPoint As New Point
    Private WithEvents RTBWrapper As New cRTBWrapper

    '' GLOBAL CURRENT DIRECTORY
    Dim Path_directory As String
    '' GLOBAL FILE SELECTED
    Dim File_Selected As String
    Dim extension As String
    Dim dt As DataTable
    Dim pathTxt As String = Nothing
    Dim fileReader As String
    Dim selectedFileTxtEditor As String
    Dim docobj As Object
    '' GLOBAL FILE FOR COPY
    Dim global_file_to_copy
    Dim copy_file = False
    '' GLOBAL FILE FOR CUT
    Dim global_file_to_cut
    Dim cut_file = False
    ''IMAGE GLOBAL
    Dim imageGlobal = Nothing
    '' GLobal List View Index
    Dim nIndex As Integer = 0
    Dim boot As Integer = 0

#Region "FILE SYSTEM TREE"

    Private Sub emptyClipboard()
        'Empty file executor
        Me.Text = "File Executor"
        '' Empty Rich Text Box
        RichTextBox1.Text = ""
        RichTextBox1.Rtf = ""
        '' Empty Status Bar
        StatusBar1.Text = ""
    End Sub

    '' COLORIZE CODE
    Private Sub colorizedCode()
        Try
            If RichTextBox1.Text.Trim <> "" Then
                ''GET FILE IN UTF
                ToUTF(RichTextBox1.Text)
                '' COLORIZE CODE
                ''ColorizeHtmlCodes()
                '' CODE WRAPPER
                ''ColoredWrapper()
                ''ColorText.ColorAsYouType(RichTextBox1)
                Dim RTB = New RTB()
                RTB.AddColoredTextOld(RichTextBox1.Text, RichTextBox1)
                ''AddColoredText(RichTextBox1.Text, RichTextBox1)
                ''Me.AddColouredText(RichTextBox1.Text)
                '' ReadFile(File_Selected)     
                '' RichTextBox1.Text = RichTextBox1.Text.Insert(RichTextBox1.SelectionStart, ChrW(&H22))
                ''RTBWrapper.colorDocument()
                '' Encoding.UTF8.GetBytes(RichTextBox1.Text)
                '' UnicodeStringToRtfText(RichTextBox1.Text)
            End If
        Catch ex As Exception
            Exit Sub
        End Try
    End Sub

    Private Sub FileSystemTree1_FileSelected(sender As Object, e As Forms.Controls.FileInfoEventArgs) Handles FileSystemTree1.FileSelected
        ' txtSelected.Text = e.File.FullName
        'Dim extension As String
        emptyClipboard()
        '' Get Text
        Me.Text = e.File.FullName
        '' GET FILE SELECTED
        File_Selected = e.File.FullName

        Try
            ''STATUS BAR 2
            StatusBar2.Text = "File Name: " + System.IO.Path.GetFileName(File_Selected).ToString.Trim + " (" + GetFileSize(File_Selected) + ")"
            '' FILE EXTENTION
            ''extension = LCase(File_Selected.Substring(InStrRev(File_Selected, ".")).Trim())
        Catch ex As Exception
            Exit Sub
        End Try

        '' GET FILE INFORMATION
        boot = boot + 1
        If boot = 2 Then
            Try
                GetFile()
                Refresh_FileTree()
            Catch ex As Exception
                Exit Sub
            End Try
        End If
    End Sub

    Private Sub FileSystemTree1_DriveSelected(sender As Object, e As Forms.Controls.DriveInfoEventArgs) Handles FileSystemTree1.DriveSelected
        ' txtSelected.Text = e.Drive.Name
        ' Me.TextBox1.Text = e.Drive.Name
        emptyClipboard()
        boot = boot + 1
        If boot = 2 Then
            TabControl1.SelectTab(0)
            Path_directory = e.Drive.Name
            StatusBar2.Text = "Drive Name: " + e.Drive.Name
            Me.Text = "File Executor"
            CreateMyListView(e.Drive.Name)
        End If

    End Sub

    Private Sub FileSystemTree1_DirectorySelected(sender As Object, e As Forms.Controls.DirectoryInfoEventArgs) Handles FileSystemTree1.DirectorySelected
        ' txtSelected.Text = e.Directory.FullName
        ' Me.TextBox1.Text = e.Directory.FullName
        emptyClipboard()
        boot = boot + 1
        If boot = 2 Then
            TabControl1.SelectTab(0)
            Path_directory = e.Directory.FullName
            StatusBar2.Text = "Directory Name: " + e.Directory.FullName
            Me.Text = "File Executor"
            CreateMyListView(e.Directory.FullName)
        End If
    End Sub

#End Region

#Region "GET FILE"

    Private Sub streamReader()
        'Dim sw As StreamWriter = New StreamWriter(File_Selected)
        'sw.WriteLine("This")
        'sw.WriteLine("is some text")
        'sw.WriteLine("to test")
        'sw.WriteLine("Reading")
        'sw.Close()

        'Dim sr As StreamReader = New StreamReader(File_Selected)

        ''This allows you to do one Read operation.
        'RichTextBox1.Text = sr.ReadToEnd()
        'sr.Close()

        Me.RichTextBox1.LoadFile(File_Selected, RichTextBoxStreamType.PlainText Or RichTextBoxStreamType.RichText)
    End Sub

    Sub GetFile()

        pathTxt = Nothing

        'Dim knownExt As String = Replace(e.Node.Text, e.Node.Text.Substring(InStrRev(e.Node.Text, " (")), "")

        extension = LCase(File_Selected.Substring(InStrRev(File_Selected, ".")).Trim())
        Dim file_name_ext As String = File_Selected.Substring(InStrRev(File_Selected, "/"))
        Dim FileName As String = File_Selected
        selectedFileTxtEditor = FileName  ' // za file editora //
        'TextBox2.Text = System.IO.Path.GetFileName(File_Selected).ToString.Trim
        Me.Text = "File Executor - " + System.IO.Path.GetFileName(File_Selected).ToString.Trim + " ( " + GetFileSize(File_Selected) + " )"

        'Panel Settings
        Panel1.AutoScroll = True

        'Picture Box Settings
        'PictureBox1.SizeMode = PictureBoxSizeMode.CenterImage
        PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize
        'PictureBox1.BorderStyle = BorderStyle.Fixed3D
        Try
            Select Case extension
                Case "pdf"
                    If Not file_name_ext = extension Then
                        AxAcroPDF1.src = FileName
                        TabControl1.SelectTab(2)
                    End If
                Case "bat"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "log"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        colorizedCode()
                        TabControl1.SelectTab(1)
                    End If
                Case "exe"
                    If Not file_name_ext = extension Then
                        Dim result As DialogResult = MessageBox.Show("Do you want to execute this file?", _
                               "Executing File!", _
                               MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                        If result = DialogResult.Yes Then
                            Call Shell(FileName, 1)
                        End If
                    End If
                Case "txt"
                    If Not file_name_ext = extension Then
                        'fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        'System.Text.Encoding.Default)
                        'RichTextBox1.Text = fileReader

                        streamReader()
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "docx"
                    Dim result As DialogResult = MessageBox.Show("Do you want to execute this file?", _
                              "Executing File!", _
                              MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                    If result = DialogResult.Yes Then
                        If Not file_name_ext = extension Then
                            WebBrowser1.Navigate(FileName)
                            TabControl1.SelectTab(3)
                        End If
                    End If

                Case "doc"
                    Dim result As DialogResult = MessageBox.Show("Do you want to execute this file?", _
                              "Executing File!", _
                              MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                    If result = DialogResult.Yes Then
                        If Not file_name_ext = extension Then
                            WebBrowser1.Navigate(FileName)
                            TabControl1.SelectTab(3)
                        End If
                    End If
                Case "xlsx"
                    Dim result As DialogResult = MessageBox.Show("Do you want to execute this file?", _
                              "Executing File!", _
                              MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                    If result = DialogResult.Yes Then
                        If Not file_name_ext = extension Then
                            WebBrowser1.Navigate(FileName)
                            TabControl1.SelectTab(3)
                        End If
                    End If
                Case "xls"
                    Dim result As DialogResult = MessageBox.Show("Do you want to execute this file?", _
                             "Executing File!", _
                             MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                    If result = DialogResult.Yes Then
                        If Not file_name_ext = extension Then
                            WebBrowser1.Navigate(FileName)
                            TabControl1.SelectTab(3)
                        End If
                    End If
                Case "rtf"
                    If Not file_name_ext = extension Then
                        RichTextBox1.LoadFile(FileName, RichTextBoxStreamType.RichText)
                        TabControl1.SelectTab(1)
                    End If
                Case "htm"
                    If Not file_name_ext = extension Then
                        'WebBrowser1.Navigate(FileName)
                        ' TabControl1.SelectTab(2)
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "mfx"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "sql"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                    End If
                Case "aspx"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "vb"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "xml"
                    If Not file_name_ext = extension Then
                        'WebBrowser1.Navigate(FileName)
                        'fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        'System.Text.Encoding.Default)
                        'RichTextBox1.Text = fileReader

                        Dim sr As New StreamReader(File_Selected)
                        Dim str As String = sr.ReadToEnd()
                        Me.RichTextBox1.Text = str
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "jpg"
                    If Not file_name_ext = extension Then
                        ' Process.Start(FileName)
                        PictureBox1.Image = Image.FromFile(FileName)
                        imageGlobal = FileName
                        TabControl1.SelectTab(4)
                        ZoomImage(40)
                    End If
                Case "png"
                    If Not file_name_ext = extension Then
                        ' Process.Start(FileName)
                        PictureBox1.Image = Image.FromFile(FileName)
                        imageGlobal = FileName
                        TabControl1.SelectTab(4)
                    End If
                Case "ico"
                    If Not file_name_ext = extension Then
                        ' Process.Start(FileName)
                        PictureBox1.Image = Image.FromFile(FileName)
                        imageGlobal = FileName
                        TabControl1.SelectTab(4)
                    End If
                Case "bmp"
                    If Not file_name_ext = extension Then
                        'Process.Start(FileName)
                        PictureBox1.Image = Bitmap.FromFile(FileName)
                        imageGlobal = FileName
                        TabControl1.SelectTab(4)
                    End If
                Case "gif"
                    If Not file_name_ext = extension Then
                        'Process.Start(FileName)
                        PictureBox1.Image = Image.FromFile(FileName)
                        imageGlobal = FileName
                        TabControl1.SelectTab(4)
                    End If
                Case "tif"
                    If Not file_name_ext = extension Then
                        'Process.Start(FileName)
                        PictureBox1.Image = Image.FromFile(FileName)
                        imageGlobal = FileName
                        TabControl1.SelectTab(4)
                    End If
                Case "jpeg"
                    If Not file_name_ext = extension Then
                        'Process.Start(FileName)
                        PictureBox1.Image = Image.FromFile(FileName)
                        imageGlobal = FileName
                        TabControl1.SelectTab(4)
                    End If
                Case "html"
                    If Not file_name_ext = extension Then
                        'WebBrowser1.Navigate(FileName)
                        'TabControl1.SelectTab(2)
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "css"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "sub"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "srt"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                    End If
                Case "php"
                    If Not file_name_ext = extension Then
                        'System.Diagnostics.Process.Start(FileName)
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader.ToString
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "js"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "ini"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "config"
                    If Not file_name_ext = extension Then
                        fileReader = My.Computer.FileSystem.ReadAllText(FileName,
                        System.Text.Encoding.Default)
                        RichTextBox1.Text = fileReader
                        TabControl1.SelectTab(1)
                        colorizedCode()
                    End If
                Case "zip"
                    If Not file_name_ext = extension Then
                        Dim result As DialogResult = MessageBox.Show("Do you want to extract this file?", _
                                 "Extract File!", _
                                 MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                        If result = DialogResult.Yes Then
                            extractZipFile(FileName)
                        End If
                    End If

                    CreateMyListView(Path_directory)


                Case "rar"
                    If Not file_name_ext = extension Then
                        Dim result As DialogResult = MessageBox.Show("Do you want to extract this file?", _
                                 "Extract File!", _
                                 MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                        If result = DialogResult.Yes Then
                            extractRarFile(FileName)
                        End If
                    End If

                    CreateMyListView(Path_directory)

                Case Else
                    If Not file_name_ext = extension Then
                        'Dim intResult As String = Shell("Notepad.exe " & FileName, vbNormalFocus)
                        'System.Diagnostics.Process.Start(File_Selected)
                        'MsgBox("This format is not supported. Please try to push 'Open with...' button ! ")
                    End If
            End Select
            boot = 1
        Catch ex As Exception
            MsgBox("Error! File Cannot Opened! " + ex.Message.ToString)
        End Try


    End Sub


#End Region

#Region "TXT EDITOR"

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        If RichTextBox1.CanUndo Then RichTextBox1.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click
        If RichTextBox1.CanRedo Then RichTextBox1.Redo()
    End Sub

    Private Sub NewToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem1.Click
        ''CONTAINS IN COMMON FILES
        createNewTxtFile()
    End Sub

    Private Sub OpenFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenFileToolStripMenuItem.Click
        Dim fo As New OpenFileDialog
        fo.Title = "RTE - Open File"
        fo.DefaultExt = "rtf"
        'fo.Filter = "Text Files|*.txt"
        'fo.Filter = "Plain Text Files (*.txt)|*.txt|All files (*.*)|*.*"
        fo.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*"
        fo.FilterIndex = 1
        fo.ShowDialog()
        If (fo.FileName = Nothing) Then
            MsgBox("No file selected.")
        Else
            pathTxt = fo.FileName
            'TextBox2.Text = pathTxt.Trim
            Using sr As New System.IO.StreamReader(fo.FileName)
                RichTextBox1.Text = sr.ReadToEnd()
            End Using
            Me.Text = "Text Editor - " + System.IO.Path.GetFileName(pathTxt).ToString.Trim + " ( " + GetFileSize(pathTxt) + " )"
            File_Selected = pathTxt
        End If
    End Sub

    Private Sub SaveToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem3.Click
        Dim result As DialogResult = MessageBox.Show("Do you want to save this file?", _
                           "Save File!", _
                           MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                If (Not pathTxt = Nothing) Then
                    'Using sw As New System.IO.StreamWriter(pathTxt)                     
                    '    sw.Write(RichTextBox1.Text)
                    '    Encoding.UTF8.GetBytes(RichTextBox1.Text)                     
                    'End Using
                    SaveFile(pathTxt)
                    ColorizeHtmlCodes()
                    ''AddColoredText(RichTextBox1.Text, RichTextBox1)
                ElseIf (Not selectedFileTxtEditor = Nothing) Then
                    'Using sw As New System.IO.StreamWriter(selectedFileTxtEditor)
                    '    sw.Write(RichTextBox1.Text)
                    'End Using
                    SaveFile(selectedFileTxtEditor)
                    ColorizeHtmlCodes()
                    ''AddColoredText(RichTextBox1.Text, RichTextBox1)
                End If
                RichTextBox1.Modified = True
            Catch ex As Exception
                MsgBox("Error saving ! " + ex.Message.ToString)
            Finally
                MsgBox("The file was saved successful !")
            End Try
        End If
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        SaveTxtAs()
    End Sub

    Private Sub ExecuteFileToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ExecuteFileToolStripMenuItem.Click
        Dim result As DialogResult = MessageBox.Show("Do you want to execute this file ?", _
                       "Executing File", _
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                If (Not pathTxt = Nothing) Then
                    Call Shell(pathTxt, 1)
                ElseIf (Not selectedFileTxtEditor = Nothing) Then
                    Call Shell(selectedFileTxtEditor, 1)
                End If
            Catch ex As Exception
                MsgBox("Error Executing ! " + ex.Message.ToString)
            End Try
        End If
    End Sub

    Private Sub CreateBatToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateBatToolStripMenuItem.Click
        Dim Newtxt As String
        Dim result As DialogResult = MessageBox.Show("Do you want to create Bat file ?", _
                           "Create Bat File", _
                           MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                If (Not pathTxt = Nothing) Then
                    Using sr As New System.IO.StreamReader(pathTxt)
                        Newtxt = sr.ReadToEnd()
                    End Using
                    IO.File.WriteAllText(IO.Path.GetDirectoryName(pathTxt) + "\" + Replace(IO.Path.GetFileName(pathTxt), ".txt", ".bat"), Newtxt.ToString())
                ElseIf (Not selectedFileTxtEditor = Nothing) Then
                    Using sr As New System.IO.StreamReader(selectedFileTxtEditor)
                        Newtxt = sr.ReadToEnd()
                    End Using
                    IO.File.WriteAllText(IO.Path.GetDirectoryName(selectedFileTxtEditor) + "\" + Replace(IO.Path.GetFileName(selectedFileTxtEditor), ".txt", ".bat"), Newtxt.ToString())
                End If
                Refresh_FileTree()
                MsgBox("Bat file is created !")
            Catch ex As Exception
                MsgBox("Error creating Bat file ! " + ex.Message.ToString)
            End Try
        End If
    End Sub

    Private Sub FontsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FontsToolStripMenuItem.Click
        Dim Font As New FontDialog()
        Font.Font = RichTextBox1.Font
        Font.ShowDialog(Me)
        Try
            RichTextBox1.Font = Font.Font
        Catch ex As Exception
            ' Do nothing on Exception
        End Try
    End Sub

    Private Sub BackgroundToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles BackgroundToolStripMenuItem1.Click
        Dim Colour As New ColorDialog()
        Colour.Color = RichTextBox1.ForeColor
        Colour.ShowDialog(Me)
        Try
            RichTextBox1.ForeColor = Colour.Color
        Catch ex As Exception
            ' Do nothing on Exception
        End Try
    End Sub

    Private Sub BAckgroundToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles BAckgroundToolStripMenuItem2.Click
        Dim Colour As New ColorDialog()
        Colour.Color = RichTextBox1.BackColor
        Colour.ShowDialog(Me)
        Try
            RichTextBox1.BackColor = Colour.Color
        Catch ex As Exception
            ' Do nothing on Exception
        End Try
    End Sub

    Private Sub PrintToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem2.Click
        If Not selectedFileTxtEditor = Nothing Or selectedFileTxtEditor = "" Then
            Try
                Dim psi As New ProcessStartInfo
                psi.UseShellExecute = True
                psi.Verb = "print"
                psi.WindowStyle = ProcessWindowStyle.Hidden
                'psi.Arguments = PrintDialog1.PrinterSettings.PrinterName.ToString()
                psi.WorkingDirectory = IO.Path.GetDirectoryName(selectedFileTxtEditor)
                psi.FileName = IO.Path.GetFileName(selectedFileTxtEditor)
                Process.Start(psi)
            Catch ex As Exception
                MsgBox("Error Printing! " + ex.Message.ToString)
            End Try
        End If
    End Sub

    Private Sub LeftToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles LeftToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Left
    End Sub

    Private Sub RightToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles RightToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Right
    End Sub

    Private Sub CEnterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CEnterToolStripMenuItem.Click
        RichTextBox1.SelectionAlignment = HorizontalAlignment.Center
    End Sub

    Private Sub BoldToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BoldToolStripMenuItem.Click
        If Not RichTextBox1.SelectionFont Is Nothing Then
            Dim currentFont As System.Drawing.Font = RichTextBox1.SelectionFont
            Dim newFontStyle As System.Drawing.FontStyle
            If RichTextBox1.SelectionFont.Bold = True Then
                newFontStyle = FontStyle.Regular
            Else
                newFontStyle = FontStyle.Bold
            End If

            RichTextBox1.SelectionFont = New Font(currentFont.FontFamily,
            currentFont.Size, newFontStyle)
        End If
    End Sub

    Private Sub UnderlineToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UnderlineToolStripMenuItem.Click
        If Not RichTextBox1.SelectionFont Is Nothing Then

            Dim currentFont As System.Drawing.Font = RichTextBox1.SelectionFont
            Dim newFontStyle As System.Drawing.FontStyle

            If RichTextBox1.SelectionFont.Underline = True Then
                newFontStyle = FontStyle.Regular
            Else
                newFontStyle = FontStyle.Underline
            End If

            RichTextBox1.SelectionFont = New Font(currentFont.FontFamily, currentFont.Size, newFontStyle)

        End If
    End Sub

    Private Sub FindToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindToolStripMenuItem.Click
        Dim find_str As New frmFind()
        find_str.Show()
    End Sub

    Private Sub ToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem5.Click
        Dim f As New frmReplace()
        f.Show()
    End Sub

    Private Sub AddBulletsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddBulletsToolStripMenuItem.Click

        RichTextBox1.BulletIndent = 10
        RichTextBox1.SelectionBullet = True

    End Sub



    Private Sub RemoveBulletsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RemoveBulletsToolStripMenuItem.Click

        RichTextBox1.SelectionBullet = False

    End Sub



    Private Sub mnuIndent0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIndent0.Click

        RichTextBox1.SelectionIndent = 0

    End Sub



    Private Sub mnuIndent5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIndent5.Click

        RichTextBox1.SelectionIndent = 5

    End Sub



    Private Sub mnuIndent10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIndent10.Click

        RichTextBox1.SelectionIndent = 10

    End Sub



    Private Sub mnuIndent15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIndent15.Click

        RichTextBox1.SelectionIndent = 15

    End Sub



    Private Sub mnuIndent20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles mnuIndent20.Click

        RichTextBox1.SelectionIndent = 20

    End Sub




#End Region

#Region "WEB BROWSER"

    Private Sub BackToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BackToolStripMenuItem.Click
        WebBrowser1.GoBack()
    End Sub

    Private Sub ForwardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ForwardToolStripMenuItem.Click
        WebBrowser1.GoForward()
    End Sub

    Private Sub SAveToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles SAveToolStripMenuItem2.Click
        Try
            ' docobj = Me.WebBrowser1.ActiveXInstance.document
            ' docobj.SaveAs(selectedFileTxtEditor)
            GetFile()
        Catch ex As Exception
            MsgBox("Error Saving Document! " + ex.Message.ToString)
        End Try

    End Sub

    Private Sub SaveToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem1.Click
        WebBrowser1.ShowSaveAsDialog()
    End Sub

    Private Sub PageSetupToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PageSetupToolStripMenuItem.Click
        WebBrowser1.ShowPageSetupDialog()
    End Sub

    Private Sub PrintToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem1.Click
        WebBrowser1.ShowPrintDialog()
    End Sub

    Private Sub GoogleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GoogleToolStripMenuItem.Click
        Try
            Dim lngTag As String = LCase(My.Application.Culture.IetfLanguageTag.ToString.Substring(InStrRev(My.Application.Culture.IetfLanguageTag.ToString, "-")).Trim())
            WebBrowser1.Navigate("google." + lngTag)
        Catch ex As Exception
            MsgBox("Unable to open web browser!" + ex.Message.ToString)
        End Try

    End Sub

#End Region

#Region "TXT MENU Strip "

    Private Sub OpenWithNotepadToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles OpenWithNotepadToolStripMenuItem.Click
        If Not File_Selected = Nothing Then
            Try
                Dim intResult As String = Shell("Notepad.exe " & File_Selected, vbNormalFocus)
            Catch ex As Exception
                MsgBox("Unable open this file with selected option!")
            End Try
        End If
    End Sub


    Private Sub OpenWithDefaultProgramToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles OpenWithDefaultProgramToolStripMenuItem.Click
        If Not File_Selected = Nothing Then
            Try
                System.Diagnostics.Process.Start(File_Selected)
            Catch ex As Exception
                MsgBox("Unable open this file with selected option!")
            End Try
        End If
    End Sub

    Private Sub CutToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CutToolStripMenuItem1.Click
        RichTextBox1.Cut()
    End Sub

    Private Sub CopyToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem1.Click
        RichTextBox1.Copy()
    End Sub

    Private Sub PastToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PastToolStripMenuItem.Click
        RichTextBox1.Paste()
    End Sub

    Private Sub DeleteToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem1.Click
        RichTextBox1.SelectedText = ""
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        RichTextBox1.SelectAll()
    End Sub

#End Region

#Region "TEXT REDACTOR STRIP MENU - Text File"

    Private Sub NewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        ''CONTAINS IN COMMON FILES
        createNewTxtFile()
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Dim fo As New OpenFileDialog
        fo.Title = "RTE - Open File"
        fo.DefaultExt = "rtf"
        'fo.Filter = "Text Files|*.txt"
        'fo.Filter = "Plain Text Files (*.txt)|*.txt|All files (*.*)|*.*"
        fo.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*"
        fo.FilterIndex = 1
        fo.ShowDialog()
        If (fo.FileName = Nothing) Then
            MsgBox("No file selected.")
        Else
            pathTxt = fo.FileName
            'TextBox2.Text = pathTxt.Trim
            Using sr As New System.IO.StreamReader(fo.FileName)
                RichTextBox1.Text = sr.ReadToEnd()
            End Using
            Me.Text = "Text Editor - " + System.IO.Path.GetFileName(pathTxt).ToString.Trim + " ( " + GetFileSize(pathTxt) + " )"
            File_Selected = pathTxt
        End If
    End Sub

    Private Sub UndoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem1.Click
        If RichTextBox1.CanUndo Then RichTextBox1.Undo()
    End Sub

    Private Sub RedoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem1.Click
        If RichTextBox1.CanRedo Then RichTextBox1.Redo()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        Dim result As DialogResult = MessageBox.Show("Do you want to save this file?", _
                         "Save File!", _
                         MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                If (Not pathTxt = Nothing) Then
                    '' Using sw As New System.IO.StreamWriter(pathTxt)
                    ''sw.Write(RichTextBox1.Text)
                    '' End Using
                    SaveFile(pathTxt)
                    ColorizeHtmlCodes()
                    '' AddColoredText(RichTextBox1.Text, RichTextBox1)
                ElseIf (Not selectedFileTxtEditor = Nothing) Then
                    '' Using sw As New System.IO.StreamWriter(selectedFileTxtEditor)
                    ''sw.Write(RichTextBox1.Text)
                    ''End Using
                    SaveFile(selectedFileTxtEditor)
                    ColorizeHtmlCodes()
                    ''AddColoredText(RichTextBox1.Text, RichTextBox1)
                End If
                RichTextBox1.Modified = True
            Catch ex As Exception
                MsgBox("Error saving ! " + ex.Message.ToString)
            End Try
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        SaveTxtAs()
    End Sub

    Private Sub ExecuteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExecuteToolStripMenuItem.Click
        Dim result As DialogResult = MessageBox.Show("Do you want to execute this file ?", _
                       "Executing File", _
                       MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                If (Not pathTxt = Nothing) Then
                    Call Shell(pathTxt, 1)
                ElseIf (Not selectedFileTxtEditor = Nothing) Then
                    Call Shell(selectedFileTxtEditor, 1)
                End If
            Catch ex As Exception
                MsgBox("Error Executing ! " + ex.Message.ToString)
            End Try
        End If
    End Sub

    Private Sub CreateBatFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateBatFileToolStripMenuItem.Click
        Dim Newtxt As String
        Dim result As DialogResult = MessageBox.Show("Do you want to create Bat file ?", _
                           "Create Bat File", _
                           MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                If (Not pathTxt = Nothing) Then
                    Using sr As New System.IO.StreamReader(pathTxt)
                        Newtxt = sr.ReadToEnd()
                    End Using
                    IO.File.WriteAllText(IO.Path.GetDirectoryName(pathTxt) + "\" + Replace(IO.Path.GetFileName(pathTxt), ".txt", ".bat"), Newtxt.ToString())
                ElseIf (Not selectedFileTxtEditor = Nothing) Then
                    Using sr As New System.IO.StreamReader(selectedFileTxtEditor)
                        Newtxt = sr.ReadToEnd()
                    End Using
                    IO.File.WriteAllText(IO.Path.GetDirectoryName(selectedFileTxtEditor) + "\" + Replace(IO.Path.GetFileName(selectedFileTxtEditor), ".txt", ".bat"), Newtxt.ToString())
                End If
                Refresh_FileTree()
                MsgBox("Bat file is created !")
            Catch ex As Exception
                MsgBox("Error creating Bat file ! " + ex.Message.ToString)
            End Try
        End If
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        If Not selectedFileTxtEditor = Nothing Or selectedFileTxtEditor = "" Then
            Try
                Dim psi As New ProcessStartInfo
                psi.UseShellExecute = True
                psi.Verb = "print"
                psi.WindowStyle = ProcessWindowStyle.Hidden
                'psi.Arguments = PrintDialog1.PrinterSettings.PrinterName.ToString()
                psi.WorkingDirectory = IO.Path.GetDirectoryName(selectedFileTxtEditor)
                psi.FileName = IO.Path.GetFileName(selectedFileTxtEditor)
                Process.Start(psi)
            Catch ex As Exception
                MsgBox("Error Printing! " + ex.Message.ToString)
            End Try
        End If
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        If RichTextBox1.Modified Then
            Dim AnswerYesNo As Integer
            AnswerYesNo = MessageBox.Show("The current document has not been saved, would you like to continue without saving?", " Unsaved Document ", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If AnswerYesNo = Windows.Forms.DialogResult.No Then
                Exit Sub
            Else
                Application.Exit()
            End If
        Else
            Application.Exit()
        End If
    End Sub

    Private Sub FindToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles FindToolStripMenuItem1.Click
        Dim find_str As New frmFind()
        find_str.Show()
    End Sub

    Private Sub FindAndReplaceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FindAndReplaceToolStripMenuItem.Click
        Dim f As New frmReplace()
        f.Show()
    End Sub



#End Region

#Region "TREE STRIP MENU "

    '' CREATE FILE
    Private Sub CreateFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateFileToolStripMenuItem.Click
        If Path_directory = Nothing Then
            MsgBox("Please,choose a directory!")
            Return
        End If
        createAFile()
    End Sub

    '' MOVE FILE
    Private Sub MoveFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoveFileToolStripMenuItem.Click
        MoveFile("")
        Refresh_FileTree()
    End Sub

    ''REFRESH
    Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        Refresh_FileTree()
    End Sub

    ''COPY FILE
    Private Sub ToolStripMenuCopy_Click(sender As Object, e As EventArgs) Handles ToolStripMenuCopy.Click
        If File_Selected = Nothing Then
            MsgBox("Please,choose a file!")
            Return
        End If
        global_file_to_copy = New FileInfo(File_Selected)
        copy_file = True
        Refresh_FileTree()
    End Sub

    ''CUT FILE
    Private Sub CutFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CutFileToolStripMenuItem.Click
        If File_Selected = Nothing Then
            MsgBox("Please,choose a file!")
            Return
        End If
        global_file_to_cut = New FileInfo(File_Selected)
        cut_file = True
        Refresh_FileTree()
    End Sub

    ''PASTE FILE
    Private Sub PasteFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteFileToolStripMenuItem.Click
        If Path_directory = Nothing Then
            MsgBox("Please,choose a directory!")
            Return
        End If
        pasteFile()
        Refresh_FileTree()
    End Sub

    '' DELETE FILE
    Private Sub DeleteFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteFileToolStripMenuItem.Click
        If File_Selected <> Nothing Then
            Dim result As DialogResult = MessageBox.Show("Do you realy want to delete this file?", _
                         "Delete file!", _
                         MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Try
                    My.Computer.FileSystem.DeleteFile(File_Selected)
                Catch ex As Exception
                    MsgBox("Unable to delete file: " & File_Selected & " !")
                Finally
                    MsgBox("File: " & File_Selected & " was deleted !")
                    If Path_directory <> Nothing Then

                        CreateMyListView(Path_directory)

                    End If
                    Refresh_FileTree()
                End Try
            End If
        ElseIf Path_directory <> Nothing Then
            Dim result As DialogResult = MessageBox.Show("Do you realy want to delete this directory?", _
                                     "Delete directory!", _
                                     MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                deleteDirectory(Path_directory)
            End If
        Else
            MsgBox("Please, choose a directory or file to delete!")
            Return
        End If

        '' REFRESH TREE
        FileSystemTree1.Update()
        Refresh_FileTree()
    End Sub

    '' OPEN WITH PAINT
    Private Sub ToolStripMenuOpenWithPaint_Click(sender As Object, e As EventArgs) Handles ToolStripMenuOpenWithPaint.Click
        ''Shell("mspaint.exe " + File_Selected.Trim, vbNormalFocus)
        Try
            Dim strFilename = File_Selected
            Shell("mspaint.exe" & Space(1) & Chr(34) & strFilename & Chr(34))
        Catch ex As Exception
            MsgBox("Unable to open this file!")
        End Try
    End Sub

    '' CREATE DIRECTORY
    Private Sub CreateDirectoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateDirectoryToolStripMenuItem.Click
        CreateDirectory()
        Refresh_FileTree()
    End Sub

    '' OPEN WITH DEFAULT PROGRAMM
    Private Sub OpenWithToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenWithToolStripMenuItem.Click
        If File_Selected = Nothing Then
            MsgBox("Please,choose a file!")
            Return
        End If
        ShellExecute(File_Selected)
        Refresh_FileTree()
    End Sub

#End Region

#Region "FORM FUNCTIONS"

    Private Sub Form1_Click(sender As Object, e As EventArgs) Handles MyBase.Click
        RichTextBox1.Text = ""
        RichTextBox1.Rtf = ""
        StatusBar1.Text = ""
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Visible = False
        Panel1.AutoScroll = True
        pBar1.Visible = False
        'Picture Box Settings
        If System.Environment.Version.Major < 4 Then
            MessageBox.Show("Please update your Windows OS!")
            Return
        End If

        ''PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize
        '' PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage

        'Dim dpath As String = "C:\Program Files"
        'If (Environment.Is64BitOperatingSystem) Then
        '    dpath = dpath + " (x86)"
        'End If

        'If (Not System.IO.Directory.Exists(dpath + "\Kingsoft")) Then
        '    Me.TabControl1.Controls.Remove(Me.TabPage2)
        'End If

    End Sub

#End Region

#Region "SELECT FOLDER"
    'VBA select folder relies upon the Microsoft Shell Controls And Automation object library. 
    'You will need to set a Reference to it. In the VBA Editor's Tools menu, click References... scroll down to "Microsoft Shell Controls And Automation" and choose it. 
    'We have to add it to be able to use the Shell32 library.

    Function SelectFolder(ByVal Title As String, ByVal TopFolder As String) As String
        Dim objShell As New Shell32.Shell
        Dim objFolder As Shell32.Folder
        objFolder = objShell.BrowseForFolder(0, Title, 0, TopFolder)
        If Not objFolder Is Nothing Then
            Try
                Return objFolder.self.Path
            Catch ex As Exception
                MsgBox("Please,choose a folder! ")
            End Try
        End If
        objFolder = Nothing
        objShell = Nothing
        Return Nothing
    End Function

#End Region

#Region "System Information AND About"

    Private Sub SystemInformationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SystemInformationToolStripMenuItem.Click
        Dim objWMI As New clsWMI()
        With objWMI

            MessageBox.Show( _
              "Computer Name - " & .ComputerName & vbCrLf & _
              "Manufacturer - " & .Manufacturer & vbCrLf & _
              "Model - " & .Model & vbCrLf & _
              "OS Name - " & .OsName & vbCrLf & _
              "Version - " & .OSVersion & vbCrLf & _
              "System Type - " & .SystemType & vbCrLf & _
              "Physical Memory - " & .TotalPhysicalMemory & vbCrLf & _
              "Windows Folder - " & .WindowsDirectory, _
              "Operating System Information", MessageBoxButtons.OK, _
              MessageBoxIcon.Information)

        End With
    End Sub

    Private Sub AboutToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem1.Click
        'MessageBox.Show( _
        '    "® NSoft BG" & vbCrLf & vbCrLf & _
        '    "Site: http://www.webstarmax.com ", _
        '    "About", MessageBoxButtons.OK, _
        '    MessageBoxIcon.Asterisk)
        loadAbout()
    End Sub


#End Region

#Region "RICH TEXT BOX"

    Sub colorWord(ByVal word As String, ByVal color As Color)
        For i As Integer = 0 To RichTextBox1.TextLength
            Try
                If RichTextBox1.Text.ElementAt(i).ToString = word.ElementAt(0).ToString Then
                    Dim found As Boolean = False
                    For j As Integer = 1 To word.Count - 1
                        If RichTextBox1.Text.ElementAt(i + j) = word.ElementAt(j) Then
                            found = True
                        Else
                            found = False
                            Exit For
                        End If
                    Next
                    If found = True Then
                        RichTextBox1.Select(i, word.Length)
                        RichTextBox1.SelectionColor = color
                    End If
                End If
            Catch ex As Exception
                Continue For
            End Try
        Next
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged

    End Sub


#End Region

#Region "PICTURE BOX"

    Private Sub PictureBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseDown
        'Capture the initial point 
        m_PanStartPoint = New Point(e.X, e.Y)
    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox1.MouseMove
        'Verify Left Button is pressed while the mouse is moving
        If e.Button = Windows.Forms.MouseButtons.Left Then

            'Here we get the change in coordinates.
            Dim DeltaX As Integer = (m_PanStartPoint.X - e.X)
            Dim DeltaY As Integer = (m_PanStartPoint.Y - e.Y)

            'Then we set the new autoscroll position.
            'ALWAYS pass positive integers to the panels autoScrollPosition method
            Panel1.AutoScrollPosition = _
            New Drawing.Point((DeltaX - Panel1.AutoScrollPosition.X), _
                            (DeltaY - Panel1.AutoScrollPosition.Y))
        End If
    End Sub

#End Region

#Region "ZOOM CONTRPOLS"

    ''ZOOM  OUT CONTROL
    Private Sub ZoomOutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoomOutToolStripMenuItem.Click
        ZoomImage(90)
    End Sub

    ''ZOOM  IN CONTROL
    Private Sub ZoomInToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ZoomInToolStripMenuItem.Click
        ZoomImage(110)
    End Sub
    '' ZOOM IMAGE
    Public Sub ZoomImage(ByRef ZoomValue As Int32)
        Dim original As Image
        'Get our original image
        original = PictureBox1.Image

        'Create a new image based on the zoom parameters we require
        Dim zoomImage As New Bitmap(original, (Convert.ToInt32(original.Width * ZoomValue) / 100), (Convert.ToInt32(original.Height * ZoomValue / 100)))

        'Create a new graphics object based on the new image
        Dim converted As Graphics = Graphics.FromImage(zoomImage)

        'Clean up the image
        converted.InterpolationMode = InterpolationMode.HighQualityBicubic

        'Clear out the original image
        PictureBox1.Image = Nothing

        'Display the new "zoomed" image
        PictureBox1.Image = zoomImage
    End Sub

    '' Open With Paint
    Private Sub OpenWithPaintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenWithPaintToolStripMenuItem.Click
        openPaint()
    End Sub

    Private Sub OpenWithToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles OpenWithToolStripMenuItem1.Click
        If imageGlobal = Nothing Then
            MsgBox("Please,choose a file!")
            Return
        Else
            ShellExecute(imageGlobal)
        End If

    End Sub

#End Region

#Region "OTHER PROGRAMS"

    'For i = ProgressBar1.Minimum To ProgressBar1.Maximum
    '        ProgressBar1.Value = i
    '        ProgressBar1.Update()
    '        System.Threading.Thread.Sleep(55)
    '    Next

    '' Execute File withe default program
    Private Function ShellExecute(ByVal File As String) As Boolean
        Dim myProcess As New Process
        myProcess.StartInfo.FileName = File
        myProcess.StartInfo.UseShellExecute = True
        myProcess.StartInfo.RedirectStandardOutput = False
        myProcess.Start()
        myProcess.Dispose()
    End Function

    '' OPEN WITH PAINT
    Private Sub openPaint()
        Try
            If Not Me.PictureBox1.Image Is Nothing Then
                Dim s As String = System.IO.Path.GetTempFileName()
                PictureBox1.Image.Save(s, System.Drawing.Imaging.ImageFormat.Png)
                Dim p As Process = Process.Start("mspaint.exe", s)
                'maybe here's a better way to wait
                While p.MainWindowHandle = IntPtr.Zero
                    System.Threading.Thread.Sleep(100)
                End While
                System.IO.File.Delete(s)
            End If
        Catch ex As Exception
            MsgBox("Unable to open this file!")
        End Try
        'Try
        '    Dim sendimage As Bitmap = CType(PictureBox1.Image, Bitmap)
        '    Clipboard.SetDataObject(sendimage)
        '    Dim programid As Integer = Shell("mspaint", AppWinStyle.MaximizedFocus)
        '    System.Threading.Thread.Sleep(100)
        '    AppActivate(programid)
        '    SendKeys.Send("^v")
        'Catch ex As Exception
        '    MsgBox("Unable to execute this file !")
        'End Try
    End Sub


    Private Function ReadFile(ByVal myfile As String) As String
        Try
            Dim myFileStream As Stream = System.IO.File.Open(myfile, FileMode.Open)
            Dim encoding As System.Text.Encoding = System.Text.Encoding.UTF8
            ' Read string from binary file with UTF8 encoding
            Dim bytes(1000) As Byte
            Dim numBytesToRead As Integer = CInt(myFileStream.Length)
            'MessageBox.Show("Length& numBytesToRead)
            Dim numBytesRead As Integer = 0
            While numBytesToRead > 0
                ' Read may return anything from 0 to numBytesToRead
                Dim n As Integer = myFileStream.Read(bytes, numBytesRead, numBytesToRead)
                'MessageBox.Show("Numberytes : " & n & " -" & encoding.GetString(bytes))
                If n = 0 Then ' We're at EOF
                    Exit While
                End If
                numBytesRead += n
                numBytesToRead -= n
            End While
            myFileStream.Close()
            'Dim buffer(30) As Byte
            'binary_file.Read(buffer, 0, 30)
            'MessageBox.Show(encoding.GetString(bytes))
            Return encoding.GetString(bytes)
        Catch

        End Try
    End Function

    '' Save and read file after redaction
    Function SaveFile(ByVal myfile As String) As String
        Try
            Dim myFileStream As Stream = System.IO.File.Open(myfile, FileMode.Truncate)
            Dim getNewString As String = RichTextBox1.Text
            Dim encoding As System.Text.Encoding = System.Text.Encoding.Default
            myFileStream.Write(encoding.GetBytes(getNewString), 0, encoding.GetByteCount(getNewString))
            myFileStream.Close()
        Catch
        End Try
        RichTextBox1.Text = RichTextBox1.Text
        RichTextBox1.Refresh()
    End Function


    Private Function UnicodeStringToRtfText(str As String) As String
        Dim arrStr As Char() = str.ToCharArray()
        Dim retStr As String = ""
        For Each ch As String In arrStr
            If (AscW(ch) > 122) Then
                retStr &= "\u" & AscW(ch) & "?"
            Else
                retStr &= ch
            End If
        Next
        Return retStr
    End Function

    Function ToUTF(str As String)
        Dim bValue() As Byte
        bValue = System.Text.UTF8Encoding.Default.GetBytes(RichTextBox1.Text)
        Return str
    End Function

    '' Colorize HTML text 
    Private Sub ColorizeHtmlCodes()
        Dim patterns(,) As String = {{"\s.*?=", "#CC33CC"}, {"(?i-mxs)<[/\?]?.*?[\s>]", "blue"}, {"=.*?""?[\s>]", "teal"}, {""".+?""", "maroon"}, _
                                     {"(?ims-){.*?}", "gray"}, {"(?i-mxs)>.*?<", "Black"}}
        For i = 0 To patterns.Length / 2 - 1
            Dim ptrn As String = patterns(i, 0)
            Dim clr As String = patterns(i, 1)
            For Each Match As Match In Regex.Matches(RichTextBox1.Text, ptrn)
                RichTextBox1.Select(Match.Index, Match.Length)
                RichTextBox1.SelectionColor = ColorTranslator.FromHtml(clr)
                My.Application.DoEvents()
            Next
        Next

    End Sub

#End Region

#Region "STATUS BARS"

    Private Sub RichTextBox1Wrapper_position(ByVal PositionInfo As cRTBWrapper.cPosition) Handles RTBWrapper.position
        StatusBar1.Text = "Cursor: " & PositionInfo.Cursor & ",  Line: " & PositionInfo.CurrentLine & ", Position: " & PositionInfo.LinePosition
    End Sub

    Private Sub ColoredWrapper()
        With RTBWrapper
            .bind(RichTextBox1)
            '.rtfSyntax.add("<span.*?>", True, True, Color.Red.ToArgb)
            '.rtfSyntax.add("<p.*>", True, True, Color.DarkCyan.ToArgb)
            '.rtfSyntax.add("<a.*?>", True, True, Color.Blue.ToArgb)
            '.rtfSyntax.add("<table.*?>", True, True, Color.Tan.ToArgb)
            '.rtfSyntax.add("<tr.*?>", True, True, Color.Brown.ToArgb)
            '.rtfSyntax.add("<td.*?>", True, True, Color.Brown.ToArgb)
            '.rtfSyntax.add("<img.*?>", True, True, Color.Red.ToArgb)
        End With
    End Sub

#End Region

#Region "LIST VIEW"
    '' CHECK ALL BOXES
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim itm As ListViewItem
        For Each itm In ListView1.Items
            If CheckBox1.Checked Then
                itm.Checked = True
            Else
                itm.Checked = False
            End If
        Next
    End Sub

    '' CREATE LIST VIEW
    Private Sub CreateMyListView(ByVal dirpath As String)

        '' DECLARATIONS
        Dim hImgSmall As IntPtr
        Dim hImgLarge As IntPtr
        Dim objWMI As New clsWMI()
        Dim FilesInDir As String()
        Dim shinfo As FileInfoClass.SHFILEINFO = New FileInfoClass.SHFILEINFO()

        ListView1.Clear()
        ImageListSmall.Images.Clear()
        ImageListLarge.Images.Clear()
        nIndex = 0
        ListView1.Columns.Add("      File Name", 200, HorizontalAlignment.Right)
        ListView1.Columns.Add("Size", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Date", 100, HorizontalAlignment.Left)
        ListView1.Columns.Add("Type", 100, HorizontalAlignment.Center)
        'ListView1.Columns.Add("Attribute", 100, HorizontalAlignment.Left)

        ListView1.LargeImageList = ImageListLarge
        ListView1.SmallImageList = ImageListSmall
        ListView1.View = View.Details
        ListView1.LabelEdit = True
        ListView1.AllowColumnReorder = True
        ListView1.CheckBoxes = True
        ListView1.FullRowSelect = True
        ListView1.GridLines = True
        ListView1.Sorting = SortOrder.Ascending

        Try
            FilesInDir = Directory.GetFiles(dirpath, "*.*")
        Catch ex As Exception
            Exit Sub
        End Try

        Dim SFile As String
        Dim item1 As New ListViewItem("", 0)
        item1.SubItems.Add("1")
        item1.SubItems.Add("2")
        item1.SubItems.Add("3")
        ''item1.SubItems.Add("4")
        Dim ctdate As Date
        Dim fAttr As FileAttribute
        Dim common = New Common()
        Dim fType As String
        Dim counter As Integer = 0

        Try
            For Each SFile In FilesInDir
                'Çounter
                counter = counter + 1
                'Progress Bar
                common.ProgressBar(pBar1, counter, FilesInDir.Count)
                ctdate = IO.File.GetCreationTime(SFile)
                fAttr = IO.File.GetAttributes(SFile)
                fType = GetFileExtType(LCase(SFile.Substring(InStrRev(SFile, ".")).Trim()))
                shinfo.szDisplayName = New String(Chr(0), 260)
                shinfo.szTypeName = New String(Chr(0), 80)
                hImgSmall = FileInfoClass.SHGetFileInfo(SFile, 0, shinfo, _
                                        Marshal.SizeOf(shinfo), _
                                        FileInfoClass.SHGFI_ICON Or FileInfoClass.SHGFI_SMALLICON)
                ImageListSmall.Images.Add(System.Drawing.Icon.FromHandle(shinfo.hIcon))       'Add icon to smallimageList.
                hImgLarge = FileInfoClass.SHGetFileInfo(SFile, 0, shinfo, _
                                                Marshal.SizeOf(shinfo), _
                                                FileInfoClass.SHGFI_ICON Or FileInfoClass.SHGFI_LARGEICON)
                ImageListLarge.Images.Add(System.Drawing.Icon.FromHandle(shinfo.hIcon))       'Add icon to LargeimageList.
                ListView1.Items.Add(New ListViewItem(New String() {SFile, objWMI.ConvertSize(CStr(FileLen(SFile))), ctdate.ToString, fType.ToString}, nIndex)) '' , fAttr.ToString
                nIndex = nIndex + 1
            Next
            If ListView1.Items.Count > 0 Then
                CheckBox1.Visible = True
            Else
                CheckBox1.Visible = False
            End If
        Catch ex As Exception
            Exit Sub
        End Try
        boot = 1
    End Sub

    ''  LIST VIEW BY DOUBLE CLICK
    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        For Each File As ListViewItem In ListView1.Items
            '' Dim FilePath As String = File.SubItems(0).Text & "\" & File.
            Dim FilePath As String = File.SubItems(0).Text
            If File.Selected = True Then
                RichTextBox1.Text = ""
                RichTextBox1.Rtf = ""
                File_Selected = FilePath
                GetFile()
                If RichTextBox1.Text.Trim <> "" Then
                    ''GET FILE IN UTF
                    ToUTF(RichTextBox1.Text)
                    '' COLORIZE CODE
                    ColorizeHtmlCodes()
                    ''Dim RTB = New RTB()
                    ''RTB.AddColoredTextOld(RichTextBox1.Text, RichTextBox1)
                    ''ColorText.ColorAsYouType(RichTextBox1)
                    '' CODE WRAPPER
                    ColoredWrapper()
                End If
            End If
        Next
    End Sub

    ''  LIST VIEW RENAME FILE
    Private Sub ListView1_AfterLabelEdit(sender As Object, e As LabelEditEventArgs) Handles ListView1.AfterLabelEdit
        Dim Response As MsgBoxResult
        Response = MsgBox("Are you realy want to rename this file?", _
                          MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                          "Rename File")
        If Response = MsgBoxResult.Yes Then
            ' Determine if label is changed by checking to see if it is equal to Nothing. 
            If e.Label <> Nothing Then
                Try
                    Dim filePath As String = File_Selected
                    Dim myString As String = e.Label
                    Dim mySplit() As Char = "\".ToCharArray
                    Dim myResult() As String = myString.Split(mySplit, StringSplitOptions.RemoveEmptyEntries)
                    Dim count As Int16 = myResult.Length()
                    Dim newPath As String = myResult(count - 1)

                    If System.IO.File.Exists(filePath) Then
                        My.Computer.FileSystem.RenameFile(filePath, newPath)
                    End If

                    Refresh_FileTree()
                Catch ex As Exception
                    MsgBox(ex.Message)
                    Return
                End Try
            End If
        Else
            For Each File As ListViewItem In ListView1.Items
                '' Dim FilePath As String = File.SubItems(0).Text & "\" & File.
                Dim FilePath As String = File.SubItems(0).Text
                If File.Selected = True Then
                    ListView1.FocusedItem.Text = File.Text
                End If
            Next
        End If

    End Sub

    ''  LIST VIEW CLICK FILE
    Private Sub ListView1_Click(sender As Object, e As EventArgs) Handles ListView1.Click
        File_Selected = ListView1.FocusedItem.Text
        Me.Text = "File Executor - " + System.IO.Path.GetFileName(File_Selected).ToString.Trim + " ( " + GetFileSize(File_Selected) + " )"
    End Sub

   
#End Region

#Region "ABOUT FORM"
    Public Sub loadAbout()

        Dim frmAbout As New AboutBox()
        Dim txtMoreInfo As String = ""
        '-- add some demonstrations of the RichTextBox auto-URLs:
        txtMoreInfo &= Environment.NewLine
        txtMoreInfo &= "http://www.webstarmax.com/" & Environment.NewLine
        txtMoreInfo &= Environment.NewLine
        txtMoreInfo &= "mailto:office@webstarmax.com" & Environment.NewLine

        With New AboutBox
            .AppTitle = frmAbout.AppTitle
            .AppDescription = frmAbout.AppDescription
            .AppVersion = frmAbout.AppVersion
            .AppCopyright = frmAbout.AppCopyright
            .AppMoreInfo = txtMoreInfo
            .ShowDialog(Me)
        End With

    End Sub
#End Region

#Region "COMMON FUNCTIONS"

    '' REFRESH TREE 
    Sub Refresh_FileTree()
        '' REFRESH TREE
        FileSystemTree1.Refresh()
        Me.Refresh()
        Refresh()
    End Sub

    ' SAVE FILE AS...
    Private Sub SaveTxtAs()
        If RichTextBox1.Text.Trim = "" Then
            MsgBox("There's not a file to save!")
            Return
        End If
        Dim result As DialogResult = MessageBox.Show("Do you want to save this file ?", _
                     "Save file as..", _
                     MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Try
                Dim fs As New SaveFileDialog
                fs.Title = "RTE - Open File"
                fs.DefaultExt = "rtf"
                'fs.Filter = "Text Files|*.txt"
                'fs.Filter = "Plain Text Files (*.txt)|*.txt|All files (*.*)|*.*"
                'fs.Filter = "Word file(*.doc) |*.doc;*.rtf|Text file(*.txt) |*.txt|All files(*.*) |*.*"
                fs.Filter = "Rich Text Files|*.rtf|Text Files|*.txt|HTML Files|*.htm|All Files|*.*"
                fs.FilterIndex = 1
                fs.ShowDialog()
                If (Not fs.FileName = Nothing) Then
                    Using sw As New System.IO.StreamWriter(fs.FileName)
                        sw.Write(RichTextBox1.Text)
                    End Using
                End If
                RichTextBox1.Modified = True
                Me.Text = "File Executor - " + IO.Path.GetFileName(fs.FileName)
            Catch ex As Exception
                MsgBox("Error saving ! " + ex.Message.ToString)
            Finally
                MsgBox("The file was saved successful !")
            End Try
        End If
    End Sub

    ' EXTRACT ZIP PROCEDURE
    Sub extractZipFile(FileName)
        Try
            Const NOCONFIRMATION = &H10&
            Const NOERRORUI = &H400&
            Const SIMPLEPROGRESS = &H100&
            Dim cFlags As String = NOCONFIRMATION + NOERRORUI + SIMPLEPROGRESS
            Dim fso As New FileSystemObject
            'If the output folder does not exist create it.
            If Not fso.FolderExists(IO.Path.GetDirectoryName(FileName)) Then
                fso.CreateFolder(IO.Path.GetDirectoryName(FileName))
            End If
            Dim objShell As New Shell32.Shell
            objShell.NameSpace(FileName).Items()
            objShell.NameSpace(IO.Path.GetDirectoryName(FileName))
            objShell.NameSpace(IO.Path.GetDirectoryName(FileName)).CopyHere(objShell.NameSpace(FileName).Items(), cFlags)
        Catch ex As Exception
            MsgBox("Error! Extracting aborted !")
        Finally
            MsgBox("Extracting done successfully !")
        End Try
    End Sub

    ' EXTRACT RAR PROCEDURE
    Sub extractRarFile(FileName)
        Try
            Dim archive As IArchive = ArchiveFactory.Open(FileName)
            For Each entry In archive.Entries
                If Not entry.IsDirectory Then
                    Console.WriteLine(entry.FilePath)
                    entry.WriteToDirectory(IO.Path.GetDirectoryName(FileName), ExtractOptions.ExtractFullPath Or ExtractOptions.Overwrite)
                End If
            Next
        Catch ex As Exception
            MsgBox("Error! Extracting aborted !")
        Finally
            MsgBox("Extracting done successfully !")
        End Try
    End Sub

    ' GET FILE SIZE
    Function GetFileSize(FullName As String) As String
        Dim file_size
        Dim objWMI As New clsWMI()
        Try
            'fs = CreateObject("Scripting.FileSystemObject")
            Dim fs As New FileSystemObject
            file_size = fs.GetFile(FullName)
            Return objWMI.ConvertSize(file_size.Size)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ' CREATE NEW TXT FILE
    Sub createNewTxtFile()
        Dim Response As MsgBoxResult
        Response = MsgBox("Are you sure you want to start a New Document?", _
                          MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                          "Text Editor")
        If Response = MsgBoxResult.Yes Then
            RichTextBox1.Text = "Write hier..."
            Dim new_string As String = "Text Editor - Untitled"
            Me.Text = new_string
            StatusBar2.Text = new_string
            TabControl1.SelectTab(1)
        Else
            MessageBox.Show("The operation was aborted!")
        End If

    End Sub

    ''DELETE FILES
    Sub deleteFiles(selected_item)
        Try
            For Each itm As ListViewItem In ListView1.Items
                If itm.Checked Then
                    My.Computer.FileSystem.DeleteFile(selected_item)
                End If
            Next
        Catch ex As Exception
            MsgBox("Unable to delete file: " & selected_item & " !")
        Finally
            MsgBox("File: " & selected_item & " was deleted !")
            If Path_directory <> Nothing Then

                CreateMyListView(Path_directory)
               
            End If
            Refresh_FileTree()
        End Try
    End Sub

    ''DELETE DIRECTORIES
    Sub deleteDirectory(Path)
        Try
            System.IO.Directory.Delete(Path, True)
        Catch ex As Exception
            MsgBox("Unable to delete directory: " & Path & " !")
        Finally
            MsgBox("Directory was deleted!")
        End Try
    End Sub

    '' CREATE DIRECTORY
    Private Sub CreateDirectory()
        Dim inputPath As String
        Dim NewPath As String
        inputPath = InputBox("Type a new folder fame", "New directory name")
        If inputPath.Trim = "" Then
            '' MessageBox.Show("The operation was aborted!")
            Return
        End If

        NewPath = Path_directory + "\" + inputPath
        If My.Computer.FileSystem.DirectoryExists(NewPath) Then
            MessageBox.Show("The selected directory already exists!")
        Else
            Try
                My.Computer.FileSystem.CreateDirectory(NewPath)
                ''MessageBox.Show("The selected directory has been created!")
                FileSystemTree1.Refresh()
                Me.Refresh()
            Catch ex As Exception
                MessageBox.Show("The directory could not be created!  Error: " & ex.Message, "Error creating directory.", _
                                MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                MessageBox.Show("The selected directory has been created!")
            End Try
        End If

    End Sub

    '' CREATE FILE
    Sub createAFile()
        Try
            Dim newValue As String = InputBox("Please,choose a file name", "Create file", "New file.txt", 100, 100)
            If newValue.Trim <> "" Then
                Dim filepath As String = Path_directory + "\" + newValue
                If Not System.IO.File.Exists(filepath) Then
                    System.IO.File.Create(filepath).Dispose()
                Else
                    MsgBox("This file already exist!")
                End If
            End If
        Catch ex As Exception
            MsgBox("Error file creating! " + ex.Message.ToString)
        Finally
            MsgBox("File was created!")
        End Try

    End Sub

    '' PASTE FILE
    Sub pasteFile()
        '' COPY FILE
        If copy_file Then
            Try
                global_file_to_copy.CopyTo(Path.Combine(Path_directory, global_file_to_copy.Name), True)
                copy_file = False
            Catch ex As Exception
                MsgBox("Unable to paste this file!")
            End Try

        End If

        '' CUT FILE
        If cut_file Then
            global_file_to_cut.CopyTo(Path.Combine(Path_directory, global_file_to_cut.Name), True)
            Dim result As DialogResult = MessageBox.Show("Do you realy want to delete this file?", _
                            "Delete file!", _
                            MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Try
                    My.Computer.FileSystem.DeleteFile(File_Selected)
                Catch ex As Exception
                    MsgBox("Unable to cut this file!")
                End Try

            End If
            cut_file = False
        End If
    End Sub

    '' MOVE FILE
    Private Sub MoveFile(file_to_move As String)

        Dim MyPath As String
        Dim path As String

        MyPath = SelectFolder("Please,choose a directory!", "")
        If Len(MyPath) Then
            MyPath = MyPath + "\"
        Else
            Return
        End If

        ' If file to move is empty get global file selected
        If file_to_move = Nothing Or file_to_move = "" Then
            path = File_Selected
        Else
            path = file_to_move
        End If

        Dim path2 As String = MyPath + System.IO.Path.GetFileName(path).ToString
        Try
            If System.IO.File.Exists(path) = False Then
                ' This statement ensures that the file is created, 
                ' but the handle is not kept. 
                Dim fs As System.IO.FileStream = System.IO.File.Create(path)
                fs.Close()
            End If

            ' Ensure that the target does not exist. 
            If System.IO.File.Exists(path2) Then
                System.IO.File.Delete(path2)
            End If

            ' Move the file.
            System.IO.File.Move(path, path2)
            'Console.WriteLine("{0} moved to {1}", path, path2)

            ' See if the original file exists now. 
            If System.IO.File.Exists(path) Then
                MsgBox("This file already exist! ")
            Else
                MsgBox("This file is successfully moved! ")
            End If
        Catch ex As Exception
            MsgBox("Error! " + ex.Message.ToString)
        End Try
    End Sub

    ''  GET FILE TYPE BY FILE EXTENTION
    Public Shared Function GetFileExtType(ByVal fExt As String) As String
        Dim fileType As String = ""
        'Search all keys under HKEY_CLASSES_ROOT
        Try
            For Each subKey As String In Registry.ClassesRoot.GetSubKeyNames()
                If String.IsNullOrEmpty(subKey) Then
                    Continue For
                End If
                If subKey.CompareTo("." & fExt) = 0 Then
                    'File extension found. Get Default Value
                    Dim defaultValue As String = Registry.ClassesRoot.OpenSubKey(subKey).GetValue("").ToString()
                    If defaultValue.Length = 0 Then
                        'No File Type specified
                        Exit For
                    End If
                    If fileType.Length = 0 Then
                        'Get Initial File Type and search for the full File Type Description
                        fileType = defaultValue
                        fExt = fileType
                    Else
                        'File Type Description found
                        If defaultValue.Length > 0 Then
                            fileType = defaultValue
                        End If
                        Exit For
                    End If
                End If
            Next
        Catch ex As Exception
            '' return default extention
            Return getFriendliFileType(fExt)
        End Try
        '' return get name
        Return getFriendliFileType(fileType)
    End Function

    ''  GET FILE FRIENDLY TYPE
    Public Shared Function getFriendliFileType(ByVal fileType As String) As String

        Select Case fileType
            Case "txtfile"
                Return "Text File"
            Case "xslfile"
                Return "XSL File"
            Case "xmlfile"
                Return "XML File"
            Case "batfile"
                Return "Bat File"
            Case "php_auto_file"
                Return "PHP File"
            Case "jpegfile"
                Return "JPEG File"
            Case "pngfile"
                Return "PNG File"
            Case "icofile"
                Return "Ico File"
            Case "jpgfile"
                Return "JPG File"
            Case "ET.Xlsx.6"
                Return "XLS File"
            Case "WPS.Docx.6"
                Return "Word File"
            Case "ET.Xls.6"
                Return "XLS File"
            Case "WPS.Doc.6"
                Return "Word File"
            Case "inifile"
                Return "INI File"
            Case "exefile"
                Return "Exe File"
            Case "dllfile"
                Return "DLL File"
            Case "dbfile"
                Return "DB File"
            Case "htmlfile"
                Return "HTML File"
            Case "pdf"
                Return "PDF File"
            Case "sqlfile"
                Return "SQL File"
            Case "sql"
                Return "SQL File"
            Case "aspx"
                Return "Aspx File"
            Case "dat"
                Return "DAT File"
            Case "CSSfile"
                Return "CSS File"
            Case "JSFile"
                Return "JS File"
            Case "giffile"
                Return "GIF File"
            Case "regfile"
                Return "REG File"
            Case "ttffile"
                Return "TTF File"
            Case "sysfile"
                Return "SYS File"
        End Select

        Return fileType
    End Function

#End Region

#Region "LIST VIEW STRIP MENU"

    '' DELETE SELECTED ITEMS
    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Dim no_files_to_delete = True
        For Each itm As ListViewItem In ListView1.Items
            If itm.Checked Then
                no_files_to_delete = False
            End If
        Next
        If no_files_to_delete Then
            MsgBox("Please, choose a file!")
            Return
        End If
        Dim result As DialogResult = MessageBox.Show("Do you realy want to delete selected files?", _
                         "Delete file!", _
                         MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            For Each itm As ListViewItem In ListView1.Items
                If itm.Checked Then
                    'MsgBox("item checked: " & itm.Text)
                    deleteFiles(itm.Text)
                End If
            Next
        End If
        Refresh_FileTree()
    End Sub

    '' CREATE DIRECTORY
    Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        CreateDirectory()
        Refresh_FileTree()
    End Sub

    '' CREATE FILE
    Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click
        If Path_directory = Nothing Then
            MsgBox("Please, choose a directory!")
            Return
        End If
        Dim result As DialogResult = MessageBox.Show("Do you realy want to create a file?", _
                        "Create file!", _
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            createAFile()
            Refresh_FileTree()
        End If
    End Sub

    '' COPY FILE
    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem8.Click
        If File_Selected = Nothing Then
            MsgBox("Please,choose a file!")
            Return
        End If
        global_file_to_copy = New FileInfo(File_Selected)
        copy_file = True
        Refresh_FileTree()
    End Sub

    '' PASTE FILE
    Private Sub ToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem9.Click
        If Path_directory = Nothing Then
            MsgBox("Please,choose a directory!")
            Return
        End If
        pasteFile()
        Refresh_FileTree()
    End Sub

    '' MOVE FILE
    Private Sub ToolStripMenuItem10_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem10.Click
        If File_Selected = Nothing Then
            MsgBox("Please, choose a file!")
            Return
        End If
        For Each itm As ListViewItem In ListView1.Items
            If itm.Checked Then
                MoveFile(itm.Text)
            End If
        Next
        Refresh_FileTree()
    End Sub

    '' CUT FILE
    Private Sub ToolStripMenuItem11_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem11.Click
        If File_Selected = Nothing Then
            MsgBox("Please, choose a file!")
            Return
        End If
        global_file_to_cut = New FileInfo(File_Selected)
        cut_file = True
        Refresh_FileTree()
    End Sub

    '' OPEN WITH PAINT
    Private Sub ToolStripMenuItem12_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem12.Click
        If File_Selected = Nothing Then
            MsgBox("Please, choose a file!")
            Return
        End If
        Try
            Dim strFilename = File_Selected
            Shell("mspaint.exe" & Space(1) & Chr(34) & strFilename & Chr(34))
        Catch ex As Exception
            MsgBox("Unable to open this file!")
        End Try
    End Sub

    '' OPEN WITH ...
    Private Sub ToolStripMenuItem13_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem13.Click
        If File_Selected = Nothing Then
            MsgBox("Please, choose a file!")
            Return
        End If
        ShellExecute(File_Selected)
        Refresh_FileTree()
    End Sub

    '' REFRESH
    Private Sub ToolStripMenuItem14_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem14.Click
        Refresh_FileTree()
    End Sub

#End Region


End Class


'' FILE INFO CLASS
Public Class FileInfoClass
    Public Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)> _
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)> _
        Public szTypeName As String
    End Structure
    Public Declare Auto Function SHGetFileInfo Lib "shell32.dll" _
            (ByVal pszPath As String, _
             ByVal dwFileAttributes As Integer, _
             ByRef psfi As SHFILEINFO, _
             ByVal cbFileInfo As Integer, _
             ByVal uFlags As Integer) As IntPtr
    Public Const SHGFI_ICON = &H100
    Public Const SHGFI_SMALLICON = &H1
    Public Const SHGFI_LARGEICON = &H0    ' Large icon
End Class

'' FILE LIST CLASS
Public Class FileListView
    Public Shared Sub Main()
        Application.Run(New Form1)
    End Sub
End Class

