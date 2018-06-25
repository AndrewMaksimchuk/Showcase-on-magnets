Attribute VB_Name = "Module1"
Sub ВставкаНовогоРядкаМіжІснуючими()
Attribute ВставкаНовогоРядкаМіжІснуючими.VB_Description = "Вставка Нового Рядка Між Існуючими"
Attribute ВставкаНовогоРядкаМіжІснуючими.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ВставкаНовогоРядкаМіжІснуючими Макрос
' Вставка Нового Рядка Між Існуючими
'

'
    Rows("3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("C2").Select
    Selection.Cut
    Range("B3").Select
    ActiveSheet.Paste

End Sub
Sub ПеренесенняЗначенняЗКоміркиВІншуКомірку()
Attribute ПеренесенняЗначенняЗКоміркиВІншуКомірку.VB_Description = "Перенесення Значення З Комірки В Іншу Комірку"
Attribute ПеренесенняЗначенняЗКоміркиВІншуКомірку.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ПеренесенняЗначенняЗКоміркиВІншуКомірку Макрос
' Перенесення Значення З Комірки В Іншу Комірку
'

'
    Range("C2").Select
    Selection.Cut
    Range("B3").Select
    ActiveSheet.Paste

End Sub

Sub test1()

    Rows("3").Select
    Selection.Insert
    
    Range("C2").Cut
    Range("B3").Select
    ActiveSheet.Paste

    
        
End Sub

Sub test2()
' Стартові данні, адреса для старту, точка відліку
    xxx = 3
    ' a-рядок
    aaa = 2
    aaa1 = 3
    PriceRow = 2
    ' b-колонка
    bbb = 3
    bbb1 = 2
    bbb2 = 4
    bbb3 = 1
    bbb4 = 5
    
    For A = 0 To 100
' Вставка пустих рядків
        Rows(xxx).Select
        Selection.Insert
        xxx = xxx + 2
' Перенесення данних з комірок колонки під назвою "Габариты"
            Cells(aaa, bbb).Select
            Selection.Cut
            Cells(aaa1, bbb1).Select
            ActiveSheet.Paste
' Перенесення данних з комірок колонки під назвою "Усилие на отрыв"
                Cells(aaa, bbb2).Select
                Selection.Cut
                Cells(aaa1, bbb3).Select
                ActiveSheet.Paste
' Перенесення данних з комірок колонки під назвою "Розница"
                Cells(aaa, bbb4).Select
                Selection.Cut
                Cells(aaa1, bbb4).Select
                ActiveSheet.Paste
                With Selection
                .VerticalAlignment = xlTop
                End With
' Вставка тексту "Ціна, грн:"
                Cells(PriceRow, bbb4).Value = "Ціна, грн:"
                PriceRow = PriceRow + 2
' Виділити жовтим кольором код товара і встановити жирний шрифт
                Cells(aaa, 1).Select
                     With Selection.Interior
                     .Pattern = xlSolid
                     .PatternColorIndex = xlAutomatic
                     .Color = 65535
                     End With
                     With Selection.Font
                     .Name = "Calibri"
                     .Size = 14
                     End With
                     Selection.Font.Bold = True
' Створити границю для кожної позиції
    Range(Cells(aaa, 1), Cells(aaa1, 5)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
' Зміна на наступний рядок
                    aaa = aaa + 2
                    aaa1 = aaa1 + 2
    Next A
' Вирівнювання вмісту комірок по центру колонки "Ціна, грн:"
        Columns("E:E").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
    With Selection.Font
        .Name = "Calibri"
        .Size = 14
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
' Вирівнювання вмісту комірок по центру колонки "Код товара" + "зусилля на відрив"
        Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
' Видалення колонки D "Зусилля на відрив"
    Columns("D:D").Select
    Selection.Delete
' Установити ширину колонки 5см = 24,86
    Columns("C:C").Select
    Selection.ColumnWidth = 24.86
    
End Sub

Sub Beeps()
    Cells(3, 3).Select
    Selection.Cut
    Cells(3, 1).Select
    ActiveSheet.Paste
    
    
    
End Sub
