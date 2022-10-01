
Sub PrepareQuery()

    Dim headers As New Collection
    Set headers = GetHeader()
    Dim rowsList As New Collection
    Set rowsList = GetExcelRows()
    
    
    Query = PrepareQueryString(headers, rowsList)
    Debug.Print Query
    
    
    'Dim groupedRows As Scripting.Dictionary
    'Set groupedRows = GroupByColumnsWithNonEmptyValue(rowsList)
    'Display groupedRows

End Sub
Public Function Display(ByRef groupedRows As Scripting.Dictionary)
    Dim key1 As Variant
    For Each key1 In groupedRows.Keys
        Debug.Print key1
        Dim i As Long
        For i = 1 To groupedRows(key1).Count
                        
            Dim key As Variant
            For Each key In groupedRows(key1)(i).Keys
                Debug.Print key, groupedRows(key1)(i)(key)
            Next key
        Next i
    Next key1
End Function


Public Function PrepareQueryString(ByRef headers As Collection, ByRef rowsList As Collection) As String
        
        Dim row As Long
        Dim TableName As String
        Dim SelectString As String
        Dim ColumnNames As String
        Dim FromTable As String
        Dim WhereClause As String
        
        
        TableName = "Employee"
        SelectString = "select  "
        ColumnNames = Join(Filter(CollectionToArray(headers), "rec_seq_number", False, vbTextCompare), ",")
        FromTable = " from " + TableName
        WhereClause = " where "
        
        Dim queries As Collection
        Set queries = New Collection
        
        For row = 1 To rowsList.Count
            Dim key As Variant
            Dim conditionsList As Collection
            Set conditionsList = New Collection
            
            Dim Conditions As String
            Conditions = ""
            Dim TempColumns As String
                        
            
            For Each key In rowsList(row).Keys
                If StrComp(key, "rec_seq_number") Then
                    If Not rowsList(row)(key) = "" Then
                        StrTemp = " " + key + "=" + Chr(39) + CStr(rowsList(row)(key)) + Chr(39) + " "
                        conditionsList.Add StrTemp
                    End If
                Else
                    TempColumns = CStr(rowsList(row)(key)) + "," + ColumnNames
                End If
            Next key
            
            Conditions = Join(CollectionToArray(conditionsList), "and")
            
            queries.Add SelectString + TempColumns + FromTable + WhereClause + Conditions
            
            
        Next row
        PrepareQueryString = Join(CollectionToArray(queries), " limit 5 " + vbNewLine + " union all " + vbNewLine) + "limit 5;"
End Function

Public Function GroupByColumnsWithNonEmptyValue(ByRef rowsList As Collection) As Scripting.Dictionary
        Dim groupedRows As Scripting.Dictionary
        Set groupedRows = New Scripting.Dictionary
        
        Dim row As Long
        For row = 1 To rowsList.Count
            Dim key As Variant
            Dim keysWithNonEmptyValues As Collection
            Set keysWithNonEmptyValues = New Collection
            
            For Each key In rowsList(row).Keys
                If Not rowsList(row)(key) = "" Then
                    keysWithNonEmptyValues.Add key
                Else
                    'Debug.Print "Remove", row, key
                    rowsList(row).Remove key
                End If
            
                'Debug.Print row, key, rowsList(row)(key), rowsList(row).Exists(key)
            Next key
            keysWithNonEmptyValuesAsString = Join(CollectionToArray(keysWithNonEmptyValues), "|")
            If Not groupedRows.Exists(keysWithNonEmptyValuesAsString) Then
                groupedRows.Add keysWithNonEmptyValuesAsString, New Collection
            End If
            groupedRows(keysWithNonEmptyValuesAsString).Add rowsList(row)
        Next row
        Set GroupByColumnsWithNonEmptyValue = groupedRows
End Function

Public Function GetExcelRows() As Collection
        Set StartCell = Range("A1")
        'Find Last Row and Column
        Dim LastRow As Long
        Dim LastColumn As Long
        
        LastRow = Cells(Rows.Count, StartCell.Column).End(xlUp).row
        LastColumn = Cells(StartCell.row, Columns.Count).End(xlToLeft).Column
        
        Dim headers As New Collection
        Set headers = GetHeader()

        Dim list As New Collection
        
        Dim row As Long
        Dim col As Long
        For row = 2 To LastRow Step 1
            Set dict = New Scripting.Dictionary
            
            For col = 1 To LastColumn Step 1
                If IsEmpty(Cells(row, col).Value) Then
                    dict.Add headers(col), ""
                Else
                    dict.Add headers(col), Cells(row, col).Value
                End If
                
            Next col
            
            list.Add dict
            
        Next row
        
        'Dim i As Long
        'For i = 1 To list.Count
            'Dim key As Variant
            'For Each key In list(i).Keys
                'Debug.Print i, key, list(i)(key)
            'Next key
        'Next i
        
        Set GetExcelRows = list
End Function


Public Function GetHeader() As Collection

        Dim LastRow As Long
        Dim LastColumn As Long
        Set StartCell = Range("A1")

        'Find Last Row and Column
        LastRow = Cells(Rows.Count, StartCell.Column).End(xlUp).row
        LastColumn = Cells(StartCell.row, Columns.Count).End(xlToLeft).Column


        Dim FirstRow As Boolean
        FirstRow = True

        Dim row As Long
        Dim col As Long
        
        For row = 1 To LastRow Step 1
            Set list = New Collection
            For col = 1 To LastColumn Step 1
                list.Add Trim(Cells(row, col))
            Next col
            
            If FirstRow Then
                Set GetHeader = list
                Exit Function
            End If
            
        Next row
End Function



Public Function CollectionToArray(myCol As Collection) As Variant

    Dim result  As Variant
    Dim cnt     As Long

    ReDim result(myCol.Count - 1)

    For cnt = 0 To myCol.Count - 1
        result(cnt) = myCol(cnt + 1)
    Next cnt

    CollectionToArray = result

End Function

