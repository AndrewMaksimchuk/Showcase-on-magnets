Attribute VB_Name = "Module1"
Sub ������������������̳���������()
Attribute ������������������̳���������.VB_Description = "������� ������ ����� ̳� ���������"
Attribute ������������������̳���������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ������������������̳��������� ������
' ������� ������ ����� ̳� ���������
'

'
    Rows("3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("C2").Select
    Selection.Cut
    Range("B3").Select
    ActiveSheet.Paste

End Sub
Sub ��������������������������²���������()
Attribute ��������������������������²���������.VB_Description = "����������� �������� � ������ � ���� ������"
Attribute ��������������������������²���������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��������������������������²��������� ������
' ����������� �������� � ������ � ���� ������
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
' ������� ����, ������ ��� ������, ����� �����
    xxx = 3
    ' a-�����
    aaa = 2
    aaa1 = 3
    PriceRow = 2
    ' b-�������
    bbb = 3
    bbb1 = 2
    bbb2 = 4
    bbb3 = 1
    bbb4 = 5
    
    For A = 0 To 100
' ������� ������ �����
        Rows(xxx).Select
        Selection.Insert
        xxx = xxx + 2
' ����������� ������ � ������ ������� �� ������ "��������"
            Cells(aaa, bbb).Select
            Selection.Cut
            Cells(aaa1, bbb1).Select
            ActiveSheet.Paste
' ����������� ������ � ������ ������� �� ������ "������ �� �����"
                Cells(aaa, bbb2).Select
                Selection.Cut
                Cells(aaa1, bbb3).Select
                ActiveSheet.Paste
' ����������� ������ � ������ ������� �� ������ "�������"
                Cells(aaa, bbb4).Select
                Selection.Cut
                Cells(aaa1, bbb4).Select
                ActiveSheet.Paste
                With Selection
                .VerticalAlignment = xlTop
                End With
' ������� ������ "ֳ��, ���:"
                Cells(PriceRow, bbb4).Value = "ֳ��, ���:"
                PriceRow = PriceRow + 2
' ������� ������ �������� ��� ������ � ���������� ������ �����
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
' �������� ������� ��� ����� �������
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
' ���� �� ��������� �����
                    aaa = aaa + 2
                    aaa1 = aaa1 + 2
    Next A
' ����������� ����� ������ �� ������ ������� "ֳ��, ���:"
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
' ����������� ����� ������ �� ������ ������� "��� ������" + "������� �� �����"
        Columns("A:A").Select
    With Selection
        .HorizontalAlignment = xlCenter
    End With
' ��������� ������� D "������� �� �����"
    Columns("D:D").Select
    Selection.Delete
' ���������� ������ ������� 5�� = 24,86
    Columns("C:C").Select
    Selection.ColumnWidth = 24.86
    
End Sub

Sub Beeps()
    Cells(3, 3).Select
    Selection.Cut
    Cells(3, 1).Select
    ActiveSheet.Paste
    
    
    
End Sub
