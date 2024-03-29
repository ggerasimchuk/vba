Attribute VB_Name = "Module1"
Sub priznak()

Dim ws As Worksheet
Dim i As Long, j As Long, LastRowA As Long, LastRowI As Long
Dim x As Long, m As Long, uid_right As Long, uid_left As Long
Dim priznak As String


Set ws = ThisWorkbook.Sheets("слайд 13")


LastRowA = ws.Range("A" & Rows.Count).End(xlUp).Row
LastRowI = ws.Range("I" & Rows.Count).End(xlUp).Row

priznak = ""
'Откуда начинаем итерировать столбец с ID из БД
id_I = 3
'размер подгруппы
subgroup_number = 0


For j = id_I To LastRowA
'Если ячейка не пустая, записываем в переменную priznak признак подгруппы
    For x = id_I To LastRowI
    If ws.Cells(x, "I").Value <> "" Then
            priznak = ws.Cells(x - 1, "J").Value
        Exit For
    End If
    Next x
'рассчитываем размер подгруппы
    For m = id_I To LastRowI
        If ws.Cells(m, "I").Value <> "" Then
            subgroup_number = subgroup_number + 1
        ElseIf ws.Cells(m, "I").Value = "" Then
            Exit For
        End If
    Next m
    
'проходим по подгруппе и сравниваем её уид с уид из столбца А
    For right_uid = id_I To (subgroup_number + id_I)
        For left_uid = 2 To LastRowA
            
            If ws.Cells(right_uid, "I").Value = ws.Cells(left_uid, "A").Value Then
                ws.Cells(left_uid, "G").Value = priznak
            Else
            End If
        Next left_uid
    Next right_uid
    
    
    id_I = id_I + subgroup_number + 1
    subgroup_number = 0
    priznak = ""

Next j


End Sub
