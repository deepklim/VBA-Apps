Private Function ExcelSQL(ByVal query As String, Optional ByVal result_location As Range, _
                          Optional ByVal header As Boolean = True, Optional ByVal as_array As Boolean = False) As Variant
    
    If query = "" Then Exit Function
    
    Dim cn As Object, rs As Object
    'Data tabs are within this workbook, so connects to itself
    'Use late binding so users don't need to manually import library under Tools -> References
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .ConnectionString = "Data Source=" & ThisWorkbook.Path & "\" & ThisWorkbook.Name & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        .Open
    End With
    
    'Run the SQL query
    Set rs = cn.Execute(query)
    If rs.EOF Then GoTo ExitFunction
    
    Dim i As Long, j As Long
    'Return to worksheet with header
    If Not as_array And header Then
        'Print header
        For i = 0 To rs.Fields.Count - 1
            result_location.Offset(0, i) = rs.Fields(i).Name
        Next i
        'Print records
        result_location.Offset(1, 0).CopyFromRecordset rs
    
    'Return to worksheet without header
    ElseIf Not as_array And Not header Then
        'Print records
        result_location.CopyFromRecordset rs
    
    'Return as array with header
    ElseIf as_array And header Then
        'Header
        Dim query_header As Variant: ReDim query_header(1 To rs.Fields.Count)
        For i = LBound(query_header) To UBound(query_header)
            query_header(i) = rs.Fields(i - 1).Name
        Next i
        'Rows
        Dim query_rows As Variant: query_rows = Application.Transpose(rs.GetRows)
        'Declare array for combined query result
        Dim return_array As Variant: ReDim return_array(1 To UBound(query_rows, 1) + 1, 1 To UBound(query_rows, 2))
        'Add header
        For i = LBound(query_header) To UBound(query_header)
            return_array(1, i) = query_header(i)
        Next i
        'Add rows
        For i = LBound(query_rows, 1) To UBound(query_rows, 1)
            For j = LBound(query_rows, 2) To UBound(query_rows, 2)
                return_array(i + 1, j) = query_rows(i, j)
            Next j
        Next i
        SQL = return_array
        
    'Return as array without header
    ElseIf as_array And Not header Then
        'Return records
        SQL = Application.Transpose(rs.GetRows)
    End If
    
ExitFunction:
    'Close connections
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    
End Function
