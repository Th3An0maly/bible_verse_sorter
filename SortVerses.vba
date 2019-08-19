Private Sub SortVersesButton_Click()
    On Error GoTo VerseError:
        Dim colStart, colEnd
        colStart = 2
        colEnd = 100
        
        Dim row
        For row = 1 To 100
            Dim book
            book = Cells(row, 1).Value
            If Len(Trim(book)) > 0 Then
                'Debug.Print (book)
                Dim chapterDict
                Set chapterDict = CreateObject("Scripting.Dictionary")
                Dim col
                For col = colStart To colEnd
                    Dim fullVerse
                    fullVerse = Cells(row, col).Value
                    If Len(Trim(fullVerse)) > 0 Then
                        Dim chapterVerse() As String
                        chapterVerse = Split(fullVerse, ":")
                        Dim chapter As Integer
                        chapter = Val(Trim(chapterVerse(0)))
                        Dim verse As String
                        verse = Trim(chapterVerse(1))
                        ' Pad the verse with 0s
                        verse = padVerse(verse)
                        
                        If chapterDict.Exists(chapter) Then
                            chapterDict.Item(chapter).Add verse
                        Else
                            Dim verseList
                            Set verseList = CreateObject("System.Collections.ArrayList")
                            verseList.Add verse
                            chapterDict.Add chapter, verseList
                        End If
                    End If
                Next
                
                If chapterDict.Count > 0 Then
                    Dim sortedChapters
                    Set sortedChapters = CreateObject("System.Collections.ArrayList")
                    Dim chapterDictKeys
                    chapterDictKeys = chapterDict.Keys
                    Dim chapterDictItems
                    chapterDictItems = chapterDict.Items
                    Dim sortInd
                    For sortInd = 0 To chapterDict.Count - 1
                        sortedChapters.Add chapterDictKeys(sortInd)
                        chapterDictItems(sortInd).Sort
                    Next
                    sortedChapters.Sort
                    
                    Dim cellColInd
                    cellColInd = colStart
                    Dim chapterInd
                    Dim chapterVerses
                    For chapterInd = 0 To sortedChapters.Count - 1
                        Dim chapterKey As Integer
                        chapterKey = sortedChapters(chapterInd)
                        Set chapterVerses = chapterDict.Item(chapterKey)
                        Dim verseInd
                        For verseInd = 0 To chapterVerses.Count - 1
                            Dim verseVal As String
                            verseVal = chapterVerses(verseInd)
                            verseVal = unpadVerse(verseVal)
                            Debug.Print ("Cell (" + CStr(row) + ":" + CStr(cellColInd) + ") set to " + CStr(chapterKey) + ":" + verseVal)
                            Cells(row, cellColInd).Value = CStr(chapterKey) + ":" + verseVal
                            Cells(row, cellColInd).NumberFormat = "@"
                            cellColInd = cellColInd + 1
                        Next
                    Next
                    Dim clearInd
                    For clearInd = cellColInd To colEnd
                        Cells(row, cellColInd).Clear
                        Cells(row, cellColInd).NumberFormat = "@"
                    Next
                End If
            End If
        Next
        
        Exit Sub
    
VerseError:
    MsgBox ("Only the following formats are allowed: Chapter:Verse, Chapter:VerseA-VerseZ, Chapter:VerseA,VerseB,VerseC")
End Sub

Public Function padVerse(strInput As String) As String
    Dim firstVerse As String
    firstVerse = CStr(Val(strInput))
    Dim padInd
    For padInd = Len(firstVerse) To 2
        strInput = "0" + strInput
    Next
    padVerse = strInput
End Function

Public Function unpadVerse(strInput As String) As String
    Dim unpadedVerse
    unpadedVerse = strInput
    Dim padInd
    For padInd = 1 To Len(strInput)
        Dim char
        char = Mid(strInput, padInd, 1)
        If char = "0" Then
            unpadedVerse = Replace(unpadedVerse, "0", "", , 1)
        Else
            Exit For
        End If
    Next
    unpadVerse = unpadedVerse
End Function
