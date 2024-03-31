VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3264
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5136
   OleObjectBlob   =   "priznak_frm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim ws As Worksheet
Dim i As Long, j As Long, LastRowA As Long, LastRowI As Long
Dim X As Long, m As Long, uid_right As Long, uid_left As Long
Dim priznak As String

Dim feature_right As String
Dim feature_left As String

Set ws = ThisWorkbook.Sheets("Sheet1")

col_feature_left = left_feature.Text
col_uid_right = right_uid.Text
col_feature_right = Split(Cells(1, Columns(col_uid_right).Column + 1).Address, "$")(1)

'размер датасета в строчках
LastRowA = ws.Range("A" & Rows.Count).End(xlUp).Row
LastRowI = ws.Range(col_uid_right & Rows.Count).End(xlUp).Row

priznak = ""
'Откуда начинаем проходить столбец с ID из БД
id_I = 0
'размер подгруппы
subgroup_number = 0

'Строка первого уид
For Row = 1 To LastRowI
    If ws.Cells(Row, col_uid_right).Value <> "" Then
            id_I = Row
        Exit For
    End If
    If id_I = 0 Then
        id_I = Row
    End If
Next Row
    

For j = id_I To LastRowA
'Если ячейка не пустая, записываем в переменную priznak признак подгруппы
    For X = id_I To LastRowI
    If ws.Cells(X, col_uid_right).Value <> "" Then
            priznak = ws.Cells(X - 1, col_feature_right).Value
        Exit For
    End If
    Next X


'рассчитываем размер подгруппы
    For m = id_I To LastRowI
        If ws.Cells(m, col_uid_right).Value <> "" Then
            subgroup_number = subgroup_number + 1
        ElseIf ws.Cells(m, col_uid_right).Value = "" Then
            Exit For
        End If
    Next m
    
'проходим по подгруппе и сравниваем её уид с уид из столбца А
    For uid_right = id_I To (subgroup_number + id_I)
        For uid_left = 2 To LastRowA
            
            If ws.Cells(uid_right, col_uid_right).Value = ws.Cells(uid_left, "A").Value Then
                ws.Cells(uid_left, col_feature_left).Value = priznak
            Else
            End If
        Next uid_left
    Next uid_right
    
    
    id_I = id_I + subgroup_number + 1
    subgroup_number = 0
    priznak = ""

Next j
End Sub
