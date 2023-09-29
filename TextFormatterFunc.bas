Attribute VB_Name = "TextFormatterFunc"
Option Explicit

Sub Format_ID(ByVal curPos As Variant)
  Dim alignment As Integer
  Dim font_name As String
  Dim font_size As Integer
  alignment = 3
  font_name = "Times New Roman"
  font_size = 16
  SetParagraph_ curPos, alignment
  SetStyleFont_ curPos, font_name, font_size
  RemoveConsecutiveBlankLines_ curPos
  RemoveMultipleSpaces_ curPos
End Sub

Sub HeaderFormat_ID(ByVal curPos As Variant)
  Dim alignment As Integer
  Dim font_name As String
  Dim font_size As Integer
  alignment = 3
  font_name = "Times New Roman"
  font_size = 26
  SetParagraph_ curPos, alignment
  SetStyleFont_ curPos, font_name, font_size
  RemoveMultipleSpaces_ curPos
End Sub

' Ключевое слово                Число     Описание
' wdAlignParagraphLeft            0       Выравнивание по левому краю
' wdAlignParagraphCenter          1       Выравнивание по центру
' wdAlignParagraphRight           2       Выравнивание по правому краю
' wdAlignParagraphJustify         3       Выравнивание по ширине
' wdAlignParagraphJustifyLow      4       Выравнивание по ширине с последней строкой по левому краю
' wdAlignParagraphJustifyMedium   5       Выравнивание по ширине с последней строкой по центру
' wdAlignParagraphJustifyHigh     6       Выравнивание по ширине с последней строкой по правому краю
Sub SetParagraph_(ByVal curPos As Variant, alignment As Integer)
  Doc.Range(curPos(0), curPos(1)).ParagraphFormat.alignment = alignment
End Sub

' устанавливаем стиль шрифта и размер
Sub SetStyleFont_(ByVal curPos As Variant, font_name As String, font_size As Integer)
  If font_name <> "" Then Doc.Range(curPos(0), curPos(1)).Font.Name = font_name
  If font_size > 0 Then Doc.Range(curPos(0), curPos(1)).Font.Size = font_size
End Sub

' удаляем множественные пробелы
Sub RemoveMultipleSpaces_(ByVal curPos As Variant)
  Dim myRange As Range
  Set myRange = Doc.Range(curPos(0), curPos(1))
  Do
    myRange.Text = Replace(myRange.Text, "  ", " ")
    If InStr(1, myRange.Text, "  ") = 0 Then
      Exit Do
    End If
  Loop
End Sub

' удаляем множественные пустые строки
Sub RemoveConsecutiveBlankLines_(ByVal curPos As Variant)
  Dim myRange As Range
  Set myRange = Doc.Range(curPos(0), curPos(1))
  myRange.Text = Replace(myRange.Text, vbCr, " ")
  ' Добавляем пустую строку в конец диапазона
  myRange.Collapse Direction:=wdCollapseEnd
  myRange.InsertParagraphAfter
End Sub
