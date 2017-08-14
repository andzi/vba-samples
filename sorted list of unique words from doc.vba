macro builds a sorted list of unique words in the document and then adds those words to the end of the document
https://wordribbon.tips.net/T007697_Generating_a_List_of_Unique_Words.html
Sub UniqueWordList()
    Dim wList As New Collection
    Dim wrd
    Dim chkwrd
    Dim sTemp As String
    Dim k As Long

    For Each wrd In ActiveDocument.Range.Words
        sTemp = Trim(LCase(wrd))
        If sTemp >= "a" And sTemp <= "z" Then
            k = 0
            For Each chkwrd In wList
                k = k + 1
                If chkwrd = sTemp Then GoTo nw
                If chkwrd > sTemp Then
                    wList.Add Item:=sTemp, Before:=k
                    GoTo nw
                End If
            Next chkwrd
            wList.Add Item:=sTemp
        End If
nw:
    Next wrd

    sTemp = "There are " & ActiveDocument.Range.Words.Count & " words "
    sTemp = sTemp & "in the document, before this summary, but there "
    sTemp = sTemp & "are only " & wList.Count & " unique words."

    ActiveDocument.Range.Select
    Selection.Collapse Direction:=wdCollapseEnd
    Selection.TypeText vbCrLf & sTemp & vbCrLf
    For Each chkwrd In wList
        Selection.TypeText chkwrd & vbCrLf
    Next chkwrd
End Sub
