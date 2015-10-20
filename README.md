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
    'Insert bookmark trough array, from 2 - because 0 & 1 is useless
    For k = 2 To UBound(ParsedArray)
        Selection.Find.ClearFormatting
        Selection.Find.Font.Color = wdColorRed
        With Selection.Find
        .Text = "": .Replacement.Text = ParsedArray(k, 0): .Forward = True: .Wrap = wdFindContinue: .Format = True: .MatchCase = False: .MatchWholeWord = False: .MatchWildcards = False: .MatchSoundsLike = False: .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceOne
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    Selection.Font.Color = wdColorAutomatic
    With ActiveDocument.Bookmarks
        .Add Range:=Selection.Range, Name:=ParsedArray(k, 0)
    End With
    Next k
End Sub

'Insert Values Into Bookmarks
Sub InsertValuesIntoBookmarks()
    Dim ParsedArray()
    Application.ScreenUpdating = False
    ParsedArray = ParseTestTxtIntoArray(ThisDocument.Path)
    'from 2 - because 0 & 1 is useless
    For i = 2 To UBound(ParsedArray)
        If ActiveDocument.Bookmarks.Exists(ParsedArray(i, 0)) = True Then
            UpdateBookmark CStr(ParsedArray(i, 0)), CStr(ParsedArray(i, 1))
        End If
    Next i
End Sub

'http://word.mvps.org/faqs/macrosvba/InsertingTextAtBookmark.htm
Sub UpdateBookmark(BookmarkToUpdate As String, TextToUse As String)
    Dim BMRange As Range
    Set BMRange = ActiveDocument.Bookmarks(BookmarkToUpdate).Range
    BMRange.Text = Replace(TextToUse, Chr(13), "")
    ActiveDocument.Bookmarks.Add BookmarkToUpdate, BMRange
End Sub

'https://support.microsoft.com/en-us/kb/184041
Sub StripAllBookmarks()
    Dim stBookmark As Bookmark
    ActiveDocument.Bookmarks.ShowHidden = True
    If ActiveDocument.Bookmarks.Count >= 1 Then
       For Each stBookmark In ActiveDocument.Bookmarks
          stBookmark.Delete
       Next stBookmark
    End If
End Sub
```

# Руководство для прогеров для работы с медстат

:neckbeard:

## Прибытие новых форм

Сначала статистикам приходят формы (каждый год 1 раз и более), они сравнивают старые формы и новые, выделяют что изменилось, по идее должны оставлять комментарий это добавилось, переименовалось или удалилось. Пример отмеченных изменений можно увидеть на картинке ниже. 

![pizdetz](http://i.imgur.com/k064ewi.jpg)

После прогеры берут и вносят изменения в поля и графы форм, которые есть в медстате.

Внесение изменений в поля и графы в медстате

Сделать.

## Создание закладок в шаблонах форм

Это надо для того, чтобы из убищного медстата выгружать данные в удобновоспринимаемые таблицы в вордовском формате. Смысл такой, нужно в вордовском документе расставить якоря (закладки), которые бы связывали данные с выгрузками медстата (файлик Test.txt - как правило в папке ARXГОД находится и изменяется при выполнениии из медстата "Распечатка из БД заполненных бланков отчетных форм" 1

Так, для начала в новую форму надо вставить макрос для заполнения закладок, что приведен выше. После этого работа с этим макросом происходит в следующем режиме: Необходимо выделить красным цветом те ячейки в таблицах, в которые будут вставляться закладки из файла Test.txt 2, после этого перемещаем в файлике Test.txt строки по порядке, потому что заполняться по таблицам они будут слева направо, просто потому нам это нужно, потому что сам медстат выгружает строки не по порядку, как они идут в самом медстате, а сортирует по цифрам. 

ИТ-отдел ЦНИИ ФИМОЗ

![Я же говорил](http://healthquality.ru/img_openhealthqu/3_1_copy.jpg)

Примечания

1 Встаёт вопрос, нахуя нам это делать каждый раз, когда минздрав и так присылает форму с закладками? Если говорить про исполнителей, кто делает эти формы, то это ЦНИИ ФИМОЗ - <http://mednet.ru/>, имеющие  своем штате печатные машинки в виде женщин под 40 без каких либо способностей к автоматизации труда.
2 В макросе захардкожено что Test.txt находится в той же папке что и сам документ, в который необходимо вставлять закладки.
