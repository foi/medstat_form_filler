# medstat_form_filler
Макрос для заполнения ебучих форм ебаного медстата

Пизда

```
'Parse Test.txt into two dimensional array
Function ParseTestTxtIntoArray(PathToTestTxt As String)
    Dim sPath As String
    Dim dataArray() As String
    Dim ArrayOfStringsSize As Integer
    Dim BookmarkLength As Integer
    Dim ParsedArrayIndex As Integer
    Dim StringLength As Integer
    Dim ParsedArray()
    BookmarkLength = 12

    sPath = PathToTestTxt & "\Test.txt"
    'Parse txt to array of strings
    Open sPath For Input As #1
        dataArray = Split(Input$(LOF(1), #1), vbLf)
    Close #1
    ArrayOfStringsSize = UBound(dataArray())
    ReDim Preserve ParsedArray(0 To ArrayOfStringsSize - 1, 0 To 1)
    'Dim ParsedArray(,) As String, strings() As String
    ParsedArrayIndex = 0
    'Parse string into Two-Dimensional Array
    For i = 0 To ArrayOfStringsSize - 1
        StringLength = Len(dataArray(ParsedArrayIndex))
        ParsedArray(ParsedArrayIndex, 0) = Left(dataArray(ParsedArrayIndex), 12)
        ParsedArray(ParsedArrayIndex, 1) = Trim(Mid(dataArray(ParsedArrayIndex), 13, StringLength - 12))
        ParsedArrayIndex = ParsedArrayIndex + 1
    Next i
    ParseTestTxtIntoArray = ParsedArray()
End Function

'Function that open Test.txt and parse test into array and insert bookmarks
Sub Main()
    Dim ParsedArray()
    Application.ScreenUpdating = False
    ParsedArray = ParseTestTxtIntoArray(ThisDocument.Path)
    'Insert bookmark trough array
    Selection.HomeKey Unit:=wdStory
    For k = 0 To UBound(ParsedArray)
        Selection.Find.ClearFormatting
        Selection.Find.Font.Color = wdColorRed
         With Selection.Find
        .Text = ""
        .Replacement.Text = ParsedArray(k, 0)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Color = wdColorAutomatic
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:=ParsedArray(k, 0)
        '.DefaultSorting = wdSortByName
        '.ShowHidden = False
    End With
    Next k
End Sub

'Insert Values Into Bookmarks
Sub InsertValuesIntoBookmarks()
    Dim ParsedArray()
    Dim BMRange As Range
    Application.ScreenUpdating = False
    ParsedArray = ParseTestTxtIntoArray(ThisDocument.Path)
    For i = 0 To UBound(ParsedArray)
        If ActiveDocument.Bookmarks.Exists(ParsedArray(i, 0)) = True Then
            'Debug.Print ParsedArray(i, 0)
            'Set BMRange = ActiveDocument.Bookmarks(ParsedArray(i, 0)).Range
            'Debug.Print BMRange.Text
            'BMRange.Text = ParsedArray(i, 1)
             UpdateBookmark CStr(ParsedArray(i, 0)), CStr(ParsedArray(i, 1))
        End If
    Next i
End Sub

'http://word.mvps.org/faqs/macrosvba/InsertingTextAtBookmark.htm
Sub UpdateBookmark(BookmarkToUpdate As String, TextToUse As String)
    Dim BMRange As Range
    Set BMRange = ActiveDocument.Bookmarks(BookmarkToUpdate).Range
    BMRange.Text = TextToUse
    ActiveDocument.Bookmarks.Add BookmarkToUpdate, BMRange
End Sub
```
