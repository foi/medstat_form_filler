# medstat_form_filler
Макрос для заполнения ебучих форм ебаного медстата

Пизда

```
Sub InsertBookmarks()
    Dim sPath As String
    Dim dataArray() As String
    Dim ArrayOfStringsSize As Integer
    Dim BookmarkLength As Integer
    Dim ParsedArray()
    BookmarkLength = 12
    
    sPath = "c:\medstat\" & "Test.txt"
    'Parse txt to array of strings
    Open sPath For Input As #1
        dataArray = Split(Input$(LOF(1), #1), vbLf)
    Close #1
    ArrayOfStringsSize = UBound(dataArray())
    ReDim Preserve ParsedArray(0 To ArrayOfStringsSize, 0 To 1)
    'Dim ParsedArray(,) As String, strings() As String
    For i = 0 To ArrayOfStringsSize
        Debug.Print dataArray(0)
    Next i
    'MsgBox UBound(ParsedArray)
End Sub
```
