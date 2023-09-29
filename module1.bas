Attribute VB_Name = "Module1"
Option Explicit               ' требование явно объявлять переменные
Private headersArr As Object  'список имен маркеров, по которым мы отбираем текст
Private keysArr() As Variant  ' список найденных имен маркеров
Private posArr() As Variant   ' список позиций найденных маркеров
Public Doc As Document        ' ссылка на активный документ


Sub main()
  Dim i As Integer
  Set Doc = ActiveDocument
  Set headersArr = CreateObject("Scripting.Dictionary")
  ' записываем в словарь значение заголовков по которому применяем функции форматирования:
  ' 1-я функция для текста внутри заголовков, 2-я функция для заголовков
  headersArr.Add "НАЗВАНИЕ ЗАГОЛОВКА-1:", Array("TextFormatterFunc.Format_ID", "TextFormatterFunc.HeaderFormat_ID")
  headersArr.Add "НАЗВАНИЕ ЗАГОЛОВКА-2:", Array("TextFormatterFunc.Format_ID", "TextFormatterFunc.HeaderFormat_ID")
  headersArr.Add "НАЗВАНИЕ ЗАГОЛОВКА-3:", Array("TextFormatterFunc.Format_ID", "TextFormatterFunc.HeaderFormat_ID")
  
  ' поиск и сохранение позиций заголовков для применения функций форматирования к тексту внутри заголовков
  GetHeadersPosition True
  ' применяем функции форматирования к тексту внутри заголовков начиная снизу документа, чтобы не пересчитывать позиции заголовков
  For i = UBound(keysArr) To LBound(keysArr) Step -1
    Application.Run headersArr(keysArr(i))(0), posArr(i)
  Next i
  ' поиск и сохранение позиций заголовков для применения функций форматирования к заголовкам
  GetHeadersPosition False
  ' заменяем позиции end
  For i = LBound(posArr) To UBound(posArr)
    posArr(i)(1) = posArr(i)(0) + Len(keysArr(i))
  Next i
  ' применяем функции форматирования к заголовкам начиная снизу документа, чтобы не пересчитывать позиции заголовков
  For i = UBound(keysArr) To LBound(keysArr) Step -1
    Application.Run headersArr(keysArr(i))(1), posArr(i)
  Next i
End Sub

' ищем в тексте документа заголовки и их позиции. Сохраняем в keysArr и posArr соответственно
Sub GetHeadersPosition(offset As Boolean)
  Dim i As Integer
  Dim j As Integer
  Dim beginPos As Integer
  Dim endPos As Integer
  Dim tempVal As Variant
  Dim headersArr_keys() As Variant

  ' Поиск и сохранение позиций маркеров в тексте документа
  headersArr_keys = headersArr.keys
  For i = LBound(headersArr_keys) To UBound(headersArr_keys)
    beginPos = InStr(1, Doc.Range.Text, headersArr_keys(i), vbTextCompare)
    If beginPos <> 0 Then
      endPos = beginPos + Len(headersArr_keys(i))
      ReDim Preserve keysArr(i)
      ReDim Preserve posArr(i)
      keysArr(i) = headersArr_keys(i)
      posArr(i) = Array(beginPos - 1, endPos)
    End If
  Next i
  ' сортировка массивов
  For i = LBound(keysArr) To UBound(keysArr)
    For j = i + 1 To UBound(keysArr)
      If posArr(i)(0) > posArr(j)(0) Then
        ' sort keys
        tempVal = keysArr(i)
        keysArr(i) = keysArr(j)
        keysArr(j) = tempVal
        ' sort position
        tempVal = posArr(i)
        posArr(i) = posArr(j)
        posArr(j) = tempVal
      End If
    Next j
  Next i
  If offset Then
    ' смещаем элементы на 1 влево, и вставляем номер позиции последнего символа в документе в последний элемент массива
    For i = LBound(posArr) To UBound(posArr) - 1
      posArr(i) = Array(posArr(i)(1), posArr(i + 1)(0))
    Next i
    posArr(UBound(posArr)) = Array(posArr(UBound(posArr))(1), ActiveDocument.Range.End)
  End If
End Sub
