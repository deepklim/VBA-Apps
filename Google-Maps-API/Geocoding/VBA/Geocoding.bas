Attribute VB_Name = "Main"
Option Explicit
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

'API key found at https://developers.google.com/maps/documentation/javascript/get-api-key
Public Const API_KEY As String = "<API key>"

Sub Geocode()
    Application.ScreenUpdating = False
    Dim t As Long: t = Timer()
    
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Geocode")
    Dim r As Long: r = WS.Cells(Rows.Count, 1).End(xlUp).Row
    If r < 2 Then Exit Sub
    
    'Clear previous results
    WS.Range("C:C").ClearContents
    WS.Range("I15").ClearContents
    WS.Range("C1").Value = "Latitude / Longitude"
    
    'Pass to API
    Dim i As Long, city_state As String
    For i = 2 To r
        city_state = WS.Range("A" & i) & ", " & WS.Range("B" & i)
        WS.Range("C" & i) = LatLong(city_state)
        'Sleep for 300 miliseconds to avoid 'OVER_QUERY_LIMIT' error
        Sleep 300
    Next
    
    WS.Range("I15") = Round(Timer() - t, 2)
    Application.ScreenUpdating = True
End Sub


Function LatLong(ByVal my_loc As String) As String
    
    'Build query
    my_loc = URLEncode(StripAccents(my_loc))
    Dim my_query As String
    my_query = "https://maps.googleapis.com/maps/api/geocode/xml?" & "address=" & my_loc & "&key=" & API_KEY
    
    Dim xml_request As Object: Set xml_request = CreateObject("MSXML2.XMLHTTP.6.0")
    Dim xml_result As Object: Set xml_result = CreateObject("MSXML2.DOMDocument.6.0")
    
    'Check if received a response
    With xml_request
        'Syntax: .Open Method, URL, Async
        .Open "GET", my_query, False
        .send
        If .readyState = 4 And .Status = 200 Then
            xml_result.LoadXML .responseText
        Else
            LatLong = .Status
            GoTo Quit
        End If
    End With
    
    'Check if response contains status = 'OK'
    Dim api_status As String: api_status = xml_result.SelectSingleNode("//GeocodeResponse/status").Text
    If api_status <> "OK" Then
        LatLong = api_status
        GoTo Quit
    End If
    
    Dim lat As Object, lng As Object
    Set lat = xml_result.SelectSingleNode("//GeocodeResponse/result/geometry/location/lat")
    Set lng = xml_result.SelectSingleNode("//GeocodeResponse/result/geometry/location/lng")
    
    LatLong = CStr(lat.Text & ", " & lng.Text)
    
Quit:
    Set xml_request = Nothing
    Set xml_result = Nothing
    Set lat = Nothing
    Set lng = Nothing
    
End Function


Function StripAccents(ByVal my_string As String) As String
    
    Dim accents As String: accents = "áàâäãåçğéèêëíìîïñóòôöõšúùûüıÿÁÀÂÄÃÅÇĞÉÈÊËÍÌÎÏÑÓÒÔÖÕŠÚÙÛÜİŸ"
    Dim striped As String: striped = "aaaaaacdeeeeiiiinooooosuuuuyyzAAAAAACDEEEEIIIINOOOOOSUUUUYYZ"
    
    Dim i As Long
    For i = 1 To Len(accents)
        my_string = Replace(my_string, Mid(accents, i, 1), Mid(striped, i, 1))
    Next
    
    StripAccents = my_string
    
End Function


Function URLEncode(ByVal my_string As String) As String
    
    'Get length of string
    Dim len_string As Long: len_string = Len(my_string)
    If len_string < 1 Then
        URLEncode = ""
        Exit Function
    End If
    
    'Loop through string, keep ASCII characters else replace with hex
    Dim current_char As String, current_char_ascii As Long
    ReDim return_string(1 To len_string) As String
    Dim i As Long
    
    For i = 1 To len_string
        current_char = Mid(my_string, i, 1)
        current_char_ascii = Asc(current_char)
        
        Select Case current_char_ascii
            'Numbers, uppercase, lowercase, hyphen, period, underscore, tilde
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                return_string(i) = current_char
            'Remove spaces
            Case 32
                return_string(i) = ""
            'Hex for all others, padding 0-15 with 0
            Case 0 To 15
                return_string(i) = "%0" & Hex(current_char_ascii)
            Case Else
                return_string(i) = "%" & Hex(current_char_ascii)
        End Select
        
    Next
    
    'Join string and return
    URLEncode = Join(return_string, "")
    
End Function
