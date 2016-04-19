Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices

Public Class RTB
    Inherits RichTextBox
    Public Shared Sub AddColoredText(ByVal strTextToAdd As String, ByVal rtb As RichTextBox)
        'Use the RichTextBox to create the initial RTF code
        rtb.Clear()
        rtb.AppendText(strTextToAdd)
        Dim strRTF As String = rtb.Rtf
        rtb.Clear()

        '             
        '             * ADD COLOUR TABLE TO THE HEADER FIRST 
        '             * 

        ' Search for colour table info, if it exists (which it shouldn't)
        ' remove it and replace with our one
        Dim iCTableStart As Integer = strRTF.IndexOf("colortbl;")

        If iCTableStart <> -1 Then 'then colortbl exists
            'find end of colortbl tab by searching
            'forward from the colortbl tab itself
            Dim iCTableEnd As Integer = strRTF.IndexOf("}"c, iCTableStart)
            strRTF = strRTF.Remove(iCTableStart, iCTableEnd - iCTableStart)

            'now insert new colour table at index of old colortbl tag
            strRTF = strRTF.Insert(iCTableStart, "colortbl ;\red255\green0\blue0;\red0\green128\blue0;\red0\green0\blue255;}")
            ' CHANGE THIS STRING TO ALTER COLOUR TABLE

            'colour table doesn't exist yet, so let's make one
        Else
            ' find index of start of header
            Dim iRTFLoc As Integer = strRTF.IndexOf("\rtf")
            ' get index of where we'll insert the colour table
            ' try finding opening bracket of first property of header first                
            Dim iInsertLoc As Integer = strRTF.IndexOf("{"c, iRTFLoc)

            ' if there is no property, we'll insert colour table
            ' just before the end bracket of the header
            If iInsertLoc = -1 Then
                iInsertLoc = strRTF.IndexOf("}"c, iRTFLoc) - 1
            End If

            ' insert the colour table at our chosen location                
            strRTF = strRTF.Insert(iInsertLoc, "{\colortbl ;\red128\green0\blue0;\red0\green128\blue0;\red0\green0\blue255;}")
            ' CHANGE THIS STRING TO ALTER COLOUR TABLE
        End If

        '            
        '             * NOW PARSE THROUGH RTF DATA, ADDING RTF COLOUR TAGS WHERE WE WANT THEM
        '             * In our colour table we defined:
        '             * cf1 = red  
        '             * cf2 = green
        '             * cf3 = blue             
        '             * 

        For i As Integer = 0 To strRTF.Length - 1
            If strRTF.Chars(i) = "<"c Then
                'add RTF tags after symbol 
                'Check for comments tags 
                If strRTF.Chars(i + 1) = "!"c Then
                    strRTF = strRTF.Insert(i + 4, "\cf2 ")
                Else
                    strRTF = strRTF.Insert(i + 1, "\cf1 ")
                End If
                'add RTF before symbol
                strRTF = strRTF.Insert(i, "\cf3 ")

                'skip forward past the characters we've just added
                'to avoid getting trapped in the loop
                i += 6
            ElseIf strRTF.Chars(i) = ">"c Then
                'add RTF tags after character
                strRTF = strRTF.Insert(i + 1, "\cf0 ")
                'Check for comments tags
                If strRTF.Chars(i - 1) = "-"c Then
                    strRTF = strRTF.Insert(i - 2, "\cf3 ")
                    'skip forward past the 6 characters we've just added
                    i += 8
                Else
                    strRTF = strRTF.Insert(i, "\cf3 ")
                    'skip forward past the 6 characters we've just added
                    i += 6
                End If
            End If
        Next i
        rtb.Rtf = strRTF
    End Sub

    '' ADD COLORED TEXT
    Public Shared Sub AddColoredTextOld(ByVal strTextToAdd As String, ByVal rtb As RichTextBox)
        'Use the RichTextBox to create the initial RTF code
        rtb.Clear()
        rtb.AppendText(strTextToAdd)
        Dim strRTF As String = rtb.Rtf
        rtb.Clear()

        '             
        '             * ADD COLOUR TABLE TO THE HEADER FIRST 
        '             * 

        ' Search for colour table info, if it exists (which it shouldn't)
        ' remove it and replace with our one
        Dim iCTableStart As Integer = strRTF.IndexOf("colortbl;")

        If iCTableStart <> -1 Then 'then colortbl exists
            'find end of colortbl tab by searching
            'forward from the colortbl tab itself
            Dim iCTableEnd As Integer = strRTF.IndexOf("}"c, iCTableStart)
            strRTF = strRTF.Remove(iCTableStart, iCTableEnd - iCTableStart)

            'now insert new colour table at index of old colortbl tag
            strRTF = strRTF.Insert(iCTableStart, "colortbl ;\red255\green0\blue0;\red0\green128\blue0;\red0\green0\blue255;}")
            ' CHANGE THIS STRING TO ALTER COLOUR TABLE

            'colour table doesn't exist yet, so let's make one
        Else
            ' find index of start of header
            Dim iRTFLoc As Integer = strRTF.IndexOf("\rtf")
            ' get index of where we'll insert the colour table
            ' try finding opening bracket of first property of header first                
            Dim iInsertLoc As Integer = strRTF.IndexOf("{"c, iRTFLoc)

            ' if there is no property, we'll insert colour table
            ' just before the end bracket of the header
            If iInsertLoc = -1 Then
                iInsertLoc = strRTF.IndexOf("}"c, iRTFLoc) - 1
            End If

            ' insert the colour table at our chosen location                
            strRTF = strRTF.Insert(iInsertLoc, "{\colortbl ;\red128\green0\blue0;\red0\green128\blue0;\red0\green0\blue255;}")
            ' CHANGE THIS STRING TO ALTER COLOUR TABLE
        End If

        '            
        '             * NOW PARSE THROUGH RTF DATA, ADDING RTF COLOUR TAGS WHERE WE WANT THEM
        '             * In our colour table we defined:
        '             * cf1 = red  
        '             * cf2 = green
        '             * cf3 = blue             
        '             * 

        For i As Integer = 0 To strRTF.Length - 1
            If strRTF.Chars(i) = "<"c Then
                'add RTF tags after symbol 
                'Check for comments tags 
                If strRTF.Chars(i + 1) = "!"c Then
                    strRTF = strRTF.Insert(i + 4, "\cf2 ")
                Else
                    strRTF = strRTF.Insert(i + 1, "\cf1 ")
                End If
                'add RTF before symbol
                strRTF = strRTF.Insert(i, "\cf3 ")

                'skip forward past the characters we've just added
                'to avoid getting trapped in the loop
                i += 6
            ElseIf strRTF.Chars(i) = ">"c Then
                'add RTF tags after character
                strRTF = strRTF.Insert(i + 1, "\cf0 ")
                'Check for comments tags
                If strRTF.Chars(i - 1) = "-"c Then
                    strRTF = strRTF.Insert(i - 2, "\cf3 ")
                    'skip forward past the 6 characters we've just added
                    i += 8
                Else
                    strRTF = strRTF.Insert(i, "\cf3 ")
                    'skip forward past the 6 characters we've just added
                    i += 6
                End If
            End If
        Next i
        rtb.Rtf = strRTF
    End Sub
End Class

Module ColorText

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    Private html() As String = {"<!DOCTYPE html>", "<html>", "</html>", "<head>", "</head>", "<body>", "</body>", "pre>", "</pre>", "<!DOCTYPE>", "<title>", "</title>", "<a>",
                                    "<abbr>", "<address>", "<area>", "<article>", "<aside>", "<audio>", "<acronym>", "<applet>", "<b>", "<base>", "<bdi>", "<bdo>", "<blockquote>", "<body>", "<br>", "<button>", "<basefont>", "<bgsound>", "<big>", "<blink>", "alt", "</a>", "src", "href", "id", "class", "<div>", "</div>"}

    Public Sub ColorAsYouType(myEditor)
        Dim colorText = myEditor
        '' Dim colorText As RichTextBox = myEditor.TabControl1.SelectedTab.Controls(0)
        Dim current_cursor_position As Integer = colorText.SelectionStart
        'This is useful to get a hold of where is the current cursor at
        'this will be needed once all coloring is done, and we need to return 

        SendMessage(colorText.Handle, &HB, CType(0, IntPtr), IntPtr.Zero) 'To avoid flickering. 0xB = WM_SETREDRAW is set to false


        Dim pattern As String = "<(.)*?>"
        Dim matches As MatchCollection = Regex.Matches(colorText.Text, pattern)
        For Each match In matches

            colorText.Select(match.index + 1, match.length - 2)

            Dim lookFor As String = match.ToString

            If match.ToString.Contains(" ") Then   'Checking if tag contains properties

                Dim length As Integer = 0
                Dim endChar() As Char = {"/", ">"}
                Dim newStr As String = lookFor.TrimStart("<").TrimEnd(endChar)
                newStr = newStr.Replace("=", " ")
                Dim start As Integer = match.index + 1
                Dim newLookFor As New ArrayList
                newLookFor.AddRange(newStr.Split(" "))

                For Each usern As String In newLookFor
                    length = usern.Length
                    colorText.Select(start, length)
                    If (html.Contains(usern)) Then
                        colorText.SelectionColor = Color.Blue
                    Else
                        colorText.SelectionColor = Color.Green
                    End If
                    start = start + length + 1
                Next
                lookFor = match.ToString.Substring(0, match.ToString.IndexOf(" ")) & ">"
                colorText.Select(match.index + 1, lookFor.Length - 1)
                'This line will strip away any extra properties, values, and will
                ' close up the tag to be able to look for it in the allowed array
            End If

            If html.Contains(lookFor.ToString.ToLower) Then
                'The tag is part of the allowed tags, and can be colored green.
                colorText.SelectionColor = Color.Green
            Else
                'This tag is not recognized, and shall be colored black..
                colorText.SelectionColor = Color.Black
            End If

        Next

        colorText.SelectionStart = current_cursor_position
        'Returning cursor to its original position

        colorText.SelectionLength = 0
        'De-Selecting text (if it was selected)


        colorText.SelectionColor = Color.Black
        'new text will be colored black, until 
        'recognized as HTML tag.

        'We finish. Now paint
        SendMessage(colorText.Handle, &HB, CType(1, IntPtr), IntPtr.Zero) '0xB = WM_SETREDRAW is set to True
        colorText.Invalidate()
    End Sub

End Module