Public Class Common
    '' PROGRESS BAR
    Public Sub ProgressBar(pBar1 As ProgressBar, st As Integer, all_count As Integer)
        pBar1.Show()
        pBar1.Minimum = 0
        pBar1.Maximum = all_count
        ''pBar1.Increment(100 / (all_count + 1
        '' pBar1.Increment(st)
        ''  For j As Integer = 0 To all_count - 1
        'do network functions
        'pBar1.Increment(100 / (all_count + 1))
        ''pBar1.Increment(j)
        ''  If (all_count = j) Then
        ''pBar1.Visible = False
        ''  End If
        'force .NET to update the forground thread, so the progress bar looks like it should
        ''Application.DoEvents()
        '' Next
        Try
            pBar1.Value = st
            If pBar1.Value = pBar1.Maximum Then
                Application.DoEvents()
                pBar1.Visible = False
            End If
        Catch ex As Exception

        End Try
        pBar1.Hide()
    End Sub

End Class
