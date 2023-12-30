
Sub ShowMessage()
    MsgBox "你已經按下"
End Sub

Function GetSheetRowCount(ByRef ws As Worksheet) As Integer
    GetSheetRowCount = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
End Function

Function isInCollection(ByRef tests As Collection, ByVal test As String) As String
    For Each a1 In tests
        If a1 = test Then
            isInCollection = a1
            'Debug.Print isInCollection
            Exit Function
        End If
    Next a1
    isInCollection = ""
End Function
Function isInCollectionInStr(ByRef tests As Collection, ByVal test As String) As String
    ' test 會是摘要字串
    ' tests 會是人員姓名

    Dim pos As Integer

    For Each a1 In tests
        pos = InStr(1, test, a1)
        If pos > 0 Then
            isInCollectionInStr = a1
            Exit Function
        End If
    Next a1
    isInCollectionInStr = ""
End Function

Function GetPeople() As Collection
    Dim re As Collection
    Dim ws As Worksheet
    Dim rows As Integer
    
    Set re = New Collection
    Set ws = ThisWorkbook.Sheets("人員")
    
    rows = GetSheetRowCount(ws)
    For i = 2 To rows
        r1 = ws.Cells(i, 1).Value
        If Trim(r1) <> "" Then
            re.Add Trim(r1)
        End If
    Next i
    Set GetPeople = re
End Function
Function GetSubjects() As Collection
    Dim re As Collection
    Dim ws As Worksheet
    Dim rows As Integer
    
    Set re = New Collection
    Set ws = ThisWorkbook.Sheets("科目")
    
    rows = GetSheetRowCount(ws)
    For i = 2 To rows
        r1 = ws.Cells(i, 1).Value
        If Trim(r1) <> "" Then
            re.Add Trim(r1)
        End If
    Next i
    Set GetSubjects = re
    
    
'    For i = 1 To re.Count
'        Debug.Print re.Item(i)
'        Debug.Print i
'    Next i
    
End Function

Sub Button1Click()
    Dim ws As Worksheet
    Dim people As Collection
    Dim subjects As Collection
    
    Set subjects = GetSubjects()
    Set people = GetPeople()
    Debug.Print subjects.Count
    Debug.Print people.Count
    
    
    Dim aSub, aSub2 As String
    Dim aCom, aCom2 As String
    Dim aDate As Date
    
    
    
    
    Set ws = ThisWorkbook.Sheets("傳票")
    RowCount = GetSheetRowCount(ws)
    
    Set wsRe = ThisWorkbook.Sheets("輸出")
    Dim rowRe As Integer
    rowRe = 2
    
    For i = 1 To RowCount
        If IsEmpty(ws.Cells(i, 5)) Or IsEmpty(ws.Cells(i, 19)) Then
            ' 不作任何事，VBA 似乎沒有 Continue
        Else
            aSub2 = ws.Cells(i, 5).Value
            aCom2 = ws.Cells(i, 19).Value
            aSub = isInCollection(subjects, aSub2)
            aCom = isInCollectionInStr(people, aCom2)

            If aSub <> "" Then
                If aCom <> "" Then
                aSub2 = ws.Cells(i, 5).Value
                aMoney = ws.Cells(i, 14).Value
                
                wsRe.Cells(rowRe, 1) = aCom
                wsRe.Cells(rowRe, 2) = aSub2
                wsRe.Cells(rowRe, 3) = aCom2
                wsRe.Cells(rowRe, 4) = aMoney
                wsRe.Cells(rowRe, 5) = ws.Cells(i, 4).Value
                
                
                
                rowRe = rowRe + 1
                End If
            End If
        End If
    Next i
    
End Sub

Sub Button_Reset_Clicked()
    Dim wsRe As Worksheet
    Dim rows As Integer
    
    Set wsRe = ThisWorkbook.Sheets("輸出")
    rows = GetSheetRowCount(wsRe)
    wsRe.Range(wsRe.Cells(2, 1), wsRe.Cells(rows, 20)).Clear
    
    
End Sub

