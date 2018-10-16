Attribute VB_Name = "Main"
Option Explicit

'Based on the Google Maps API documentation found at:
'https://developers.google.com/maps/documentation/javascript/adding-a-google-map

'API key found at https://developers.google.com/maps/documentation/javascript/get-api-key
Public Const API_KEY As String = "<API key>"

Sub CreateMap()
    Application.ScreenUpdating = False
    
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Pivot")
    Dim r As Long: r = WS.Cells(Rows.Count, 1).End(xlUp).Row
    Dim dot_size As Long: dot_size = Replace(WS.Range("I4").Value, " ", "")
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object: Set oFile = fso.CreateTextFile(ThisWorkbook.Path & "\Create Map.html")
    
    oFile.WriteLine "<!DOCTYPE html>"
    oFile.WriteLine "<html>"
    oFile.WriteLine "  <head>"
    oFile.WriteLine "    <meta name=""viewport"" content=""initial-scale=1.0, user-scalable=no"">"
    oFile.WriteLine "    <meta charset=""utf-8"">"
    oFile.WriteLine "    <title>Create Map</title>"
    oFile.WriteLine "    <style>"
    oFile.WriteLine "      #map {"
    oFile.WriteLine "        height: 100%;"
    oFile.WriteLine "      }"
    oFile.WriteLine "      html, body {"
    oFile.WriteLine "        height: 100%;"
    oFile.WriteLine "        margin: 0;"
    oFile.WriteLine "        padding: 0;"
    oFile.WriteLine "      }"
    oFile.WriteLine "    </style>"
    oFile.WriteLine "  </head>"
    oFile.WriteLine "  <body>"
    oFile.WriteLine "    <div id=""map""></div>"
    oFile.WriteLine "    <script>"
    oFile.WriteLine "      var i = 0;"
    oFile.WriteLine "      var cities_array = ["
    
    Dim i As Long
    Dim lat_val As String, long_val As String, count_val As String
    For i = 5 To r
        lat_val = WS.Range("B" & i)
        long_val = WS.Range("C" & i)
        count_val = WS.Range("D" & i)
        
        oFile.WriteLine "        [" & lat_val & ", " & long_val & ", " & count_val & "],"
    Next i
    
    'Add two additional markers for the legend
    oFile.WriteLine "        [37.8, -73, 1],"
    oFile.WriteLine "        [37.8, -70, 2]"
    
    oFile.WriteLine "      ];"
    oFile.WriteLine "      function initMap() {"
    oFile.WriteLine "        var map = new google.maps.Map(document.getElementById('map'), {"
    oFile.WriteLine "          zoom: 6,"
    oFile.WriteLine "          center: {lat: 43, lng: -80},"
    oFile.WriteLine "          mapTypeId: 'terrain'"
    oFile.WriteLine "        });"
    oFile.WriteLine "        for (i = 0; i < cities_array.length; i++) {"
    
    '1 President / Prime Minister
    oFile.WriteLine "          if (cities_array[i][2] == 1) {"
    oFile.WriteLine "            var cityCircle = new google.maps.Circle({"
    oFile.WriteLine "              strokeColor: '#68228b',"
    oFile.WriteLine "              strokeOpacity: 0.8,"
    oFile.WriteLine "              strokeWeight: 2,"
    oFile.WriteLine "              fillColor: '#68228b',"
    oFile.WriteLine "              fillOpacity: 0.6,"
    oFile.WriteLine "              map: map,"
    oFile.WriteLine "              center: {lat: cities_array[i][0], lng: cities_array[i][1]},"
    oFile.WriteLine "              radius: " & dot_size
    oFile.WriteLine "            });"
    oFile.WriteLine "          }"
    
    '2 Presidents / Prime Ministers
    oFile.WriteLine "          if (cities_array[i][2] == 2) {"
    oFile.WriteLine "            var cityCircle = new google.maps.Circle({"
    oFile.WriteLine "              strokeColor: '#ff00ff',"
    oFile.WriteLine "              strokeOpacity: 0.8,"
    oFile.WriteLine "              strokeWeight: 2,"
    oFile.WriteLine "              fillColor: '#ff00ff',"
    oFile.WriteLine "              fillOpacity: 0.6,"
    oFile.WriteLine "              map: map,"
    oFile.WriteLine "              center: {lat: cities_array[i][0], lng: cities_array[i][1]},"
    oFile.WriteLine "              radius: " & dot_size
    oFile.WriteLine "            });"
    oFile.WriteLine "          }"
    oFile.WriteLine "        }"
    oFile.WriteLine "        addMarker({lat: 38.1, lng: -71.5}, map, 'Legend');"
    oFile.WriteLine "        addMarker({lat: 37.2, lng: -73}, map, '1 President / PM');"
    oFile.WriteLine "        addMarker({lat: 37.2, lng: -70}, map, '2 Presidents / PMs');"
    oFile.WriteLine "      }"
    
    'Declare addMarker function that adds text labels
    oFile.WriteLine "      function addMarker(location, map, label) {"
    oFile.WriteLine "        var marker = new google.maps.Marker({"
    oFile.WriteLine "          icon: 'x',"
    oFile.WriteLine "          position: location,"
    oFile.WriteLine "          label: label,"
    oFile.WriteLine "          map: map"
    oFile.WriteLine "        });"
    oFile.WriteLine "      }"
    
    oFile.WriteLine "    </script>"
    oFile.WriteLine "    <script async defer"
    oFile.WriteLine "    src=""https://maps.googleapis.com/maps/api/js?key=" & API_KEY & "&callback=initMap"">"
    oFile.WriteLine "    </script>"
    oFile.WriteLine "  </body>"
    oFile.WriteLine "</html>"
    
    oFile.Close
    MsgBox "Map sucessfully saved in current folder."
    
    Set fso = Nothing
    Set oFile = Nothing
    Application.ScreenUpdating = True
End Sub


Sub OpenIE()
    'Use late binding so users don't need to manually import library under Tools -> References
    Dim IE As Object
    On Error GoTo ErrHandler
    Set IE = CreateObject("InternetExplorer.Application")
    With IE
        .Visible = True
        .Navigate ThisWorkbook.Path & "\Create Map.html"
    End With
    
ErrHandler:
    Set IE = Nothing
    On Error GoTo 0
End Sub
