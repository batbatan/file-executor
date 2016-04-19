Imports System.IO

Friend Class FileNode
    Inherits DirectoryTreeNode

    Private m_file As FileInfo
    Private Shared ImageKeyValueLookupTable As Dictionary(Of String, String)

    Public Sub New(ByVal file As FileInfo)
        MyBase.New(FileSystemNodeType.File, file.Name)
        Me.m_file = file
        SetImageKeys()
        ''Get File Size
        Dim file_size
        fs = CreateObject("Scripting.FileSystemObject")
        Try
            file_size = fs.GetFile(Me.m_file.FullName)
            ''Toolltip
            Me.ToolTipText = Me.m_file.FullName + " (" + ConvertSize(file_size.Size) + ")"
        Catch ex As Exception

        End Try
      
    End Sub

    '' Convert Size
    Public Function ConvertSize(Size)
        Size = CSng(Replace(Size, ",", ""))

        If Not VarType(Size) = vbSingle Then
            ConvertSize = "SIZE INPUT ERROR"
            Exit Function
        End If

        Dim suffix As String = " Bytes"
        If Size >= 1024 Then suffix = " KB"
        If Size >= 1048576 Then suffix = " MB"
        If Size >= 1073741824 Then suffix = " GB"
        If Size >= 1099511627776 Then suffix = " TB"

        Select Case suffix
            Case " KB"
                Size = Math.Round(Size / 1024, 1)
            Case " MB"
                Size = Math.Round(Size / 1048576, 1)
            Case " GB"
                Size = Math.Round(Size / 1073741824, 1)
            Case " TB"
                Size = Math.Round(Size / 1099511627776, 1)

        End Select

        ConvertSize = Size & suffix
    End Function
    Shared Sub New()
        ImageKeyValueLookupTable = New Dictionary(Of String, String)

        ImageKeyValueLookupTable.Add("jpg;jpeg;gif;tif;tiff;png;bmp;JPG;JPEG;PNG;GIF;BMP", "file_image.png")
        ImageKeyValueLookupTable.Add("mp3;wma", "file_music.png")
        ImageKeyValueLookupTable.Add("doc;docx", "file_word.png")
        ImageKeyValueLookupTable.Add("xls;xlsx", "file_excel.png")
        ImageKeyValueLookupTable.Add("zip;rar", "file_archive.png")
        ImageKeyValueLookupTable.Add("ppt;pptx", "file_powerpoint.png")
        ImageKeyValueLookupTable.Add("ini;config;bat;cmd", "file_setting.png")
        ImageKeyValueLookupTable.Add("txt;xml;log;php;htm;html", "file_text.png")
        'ImageKeyValueLookupTable.Add("pdf", "file_pdf.png")

    End Sub

    Public ReadOnly Property File As FileInfo
        Get
            Return Me.m_file
        End Get
    End Property

    Private Property fs As Object

    Private Sub SetImageKeys()
        Dim key As String = GetImageKey(Path.GetExtension(File.FullName))
        Me.ImageKey = key
        Me.SelectedImageKey = key
    End Sub

    Private Shared Function GetImageKey(ByVal extension As String) As String
        extension = extension.Replace(".", "")
        For Each key As String In ImageKeyValueLookupTable.Keys
            Dim extensions() As String = key.Split(";")
            If (extensions.Contains(extension)) Then
                Return ImageKeyValueLookupTable.Item(key)
            End If
        Next
        Return "file.png"
    End Function

    Public Overrides ReadOnly Property FullName As String
        Get
            Return m_file.FullName
        End Get
    End Property

End Class
