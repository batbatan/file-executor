''' <summary>
''' This class provides two subroutines used to:
''' Find (find the first instance of a search term)
''' Find Next (find other instances of the search term after the first one is found)
''' </summary>
''' <remarks></remarks>

Public Class frmFind

    Private Sub btnFind_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFind.Click

        Dim StartPosition As Integer
        Dim SearchType As CompareMethod

        If chkMatchCase.Checked = True Then
            SearchType = CompareMethod.Binary
        Else
            SearchType = CompareMethod.Text
        End If

        StartPosition = InStr(1, Form1.RichTextBox1.Text, txtSearchTerm.Text, SearchType)

        If StartPosition = 0 Then
            MessageBox.Show("String: " & txtSearchTerm.Text.ToString() & " not found", "No Matches", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Exit Sub
        End If

        Form1.RichTextBox1.Select(StartPosition - 1, txtSearchTerm.Text.Length)
        Form1.RichTextBox1.ScrollToCaret()
        Form1.Focus()

    End Sub


    Private Sub btnFindNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindNext.Click

        Dim StartPosition As Integer = Form1.RichTextBox1.SelectionStart + 2
        Dim SearchType As CompareMethod

        If chkMatchCase.Checked = True Then
            SearchType = CompareMethod.Binary
        Else
            SearchType = CompareMethod.Text
        End If

        StartPosition = InStr(StartPosition, Form1.RichTextBox1.Text, txtSearchTerm.Text, SearchType)

        If StartPosition = 0 Then
            MessageBox.Show("String: " & txtSearchTerm.Text.ToString() & " not found", "No Matches", MessageBoxButtons.OK, MessageBoxIcon.Asterisk)
            Exit Sub
        End If

        Form1.RichTextBox1.Select(StartPosition - 1, txtSearchTerm.Text.Length)
        Form1.RichTextBox1.ScrollToCaret()
        Form1.Focus()

    End Sub


End Class