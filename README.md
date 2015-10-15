# medstat_form_filler
Макрос для заполнения ебучих форм ебаного медстата

Пизда

```
Sub InsertBookmarks()
    Dim sPath As String
    Dim dataArray() As String
    Dim ArrayOfStringsSize As Integer
    Dim BookmarkLength As Integer
    Dim ParsedArrayIndex As Integer
    Dim StringLength As Integer
    Dim ParsedArray()
    BookmarkLength = 12
    
    sPath = "c:\medstat\" & "Test.txt"
    'Parse txt to array of strings
    Open sPath For Input As #1
        dataArray = Split(Input$(LOF(1), #1), vbLf)
    Close #1
    ArrayOfStringsSize = UBound(dataArray())
    ReDim Preserve ParsedArray(0 To ArrayOfStringsSize - 1, 0 To 1)
    'Dim ParsedArray(,) As String, strings() As String
    ParsedArrayIndex = 0
    'Parse string into Two-Dimensional Array
    For i = 0 To 499
        StringLength = Len(dataArray(ParsedArrayIndex))
        ParsedArray(ParsedArrayIndex, 0) = Left(dataArray(ParsedArrayIndex), 12)
        ParsedArray(ParsedArrayIndex, 1) = RTrim(Trim(Mid(dataArray(ParsedArrayIndex), 13, StringLength - 12)))
        ParsedArrayIndex = ParsedArrayIndex + 1
    Next i
    MsgBox ParsedArray(0, 0)
End Sub
```
