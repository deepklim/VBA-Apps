Attribute VB_Name = "CreateMap"
Option Explicit

'Based on the Google Maps API documentation found at:
'https://developers.google.com/maps/documentation/javascript/adding-a-google-map

'API key found at https://developers.google.com/maps/documentation/javascript/get-api-key
Public Const API_KEY As String = "<API key>"

Sub CreateMap()
    Application.ScreenUpdating = False
    
    Dim dot_size As Long: dot_size = 15000
    
    Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Map Maker")
    Dim r As Long: r = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object: Set oFile = fso.CreateTextFile(Environ("USERPROFILE") & "\Desktop\my_map.html")
    
    oFile.WriteLine "<!DOCTYPE html>"
    oFile.WriteLine "<html>"
    oFile.WriteLine "  <head>"
    oFile.WriteLine "    <meta name=""viewport"" content=""initial-scale=1.0, user-scalable=no"">"
    oFile.WriteLine "    <meta charset=""utf-8"">"
    oFile.WriteLine "    <title>Data Impact Map Maker</title>"
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
    
    
    'Store cities in an array
    oFile.WriteLine "      var cities_array = ["
    Dim i As Long
    Dim lat_val As String, lng_val As String, icon As String, icon_colour As String
    For i = 2 To r
        lat_val = WS.Range("C" & i)
        lng_val = WS.Range("D" & i)
        icon = WS.Range("E" & i)
        icon_colour = WS.Range("F" & i)
        
        'Ignore rows with text in lat_val - caused by error messages OVER_QUERY_LIMIT, etc.
        If Not IsNumeric(lat_val) Then GoTo SkipCity
        
        oFile.WriteLine "        [" & lat_val & ", " & lng_val & ", '" & icon & "', '" & icon_colour & "'],"
SkipCity:
    Next i
    oFile.WriteLine "      ];"
    
    
    'Store colour names and associated hex in dictionary
    oFile.WriteLine "      var colours = {"
    oFile.WriteLine "        'Red': '#FD7567',"
    oFile.WriteLine "        'Orange': '#FF9900',"
    oFile.WriteLine "        'Yellow': '#FDF569',"
    oFile.WriteLine "        'Green': '#00E64D',"
    oFile.WriteLine "        'Blue': '#6991FD',"
    oFile.WriteLine "        'Purple': '#8E67FD'"
    oFile.WriteLine "      };"
    
    
    oFile.WriteLine "      function initMap() {"
    oFile.WriteLine "        var map = new google.maps.Map(document.getElementById('map'), {"
    oFile.WriteLine "          zoom: 4.0,"
    oFile.WriteLine "          center: {lat: 53.9, lng: -93.3},"
    oFile.WriteLine "          mapTypeId: 'terrain'"
    oFile.WriteLine "        });"
    oFile.WriteLine "        for (i = 0; i < cities_array.length; i++) {"
    
    'Icon = Circle
    oFile.WriteLine "          if (cities_array[i][2] == 'Circle') {"
    oFile.WriteLine "            var cityCircle = new google.maps.Circle({"
    oFile.WriteLine "              strokeColor: colours[cities_array[i][3]],"
    oFile.WriteLine "              strokeOpacity: 0.8,"
    oFile.WriteLine "              strokeWeight: 2,"
    oFile.WriteLine "              fillColor: colours[cities_array[i][3]],"
    oFile.WriteLine "              fillOpacity: 0.6,"
    oFile.WriteLine "              map: map,"
    oFile.WriteLine "              center: {lat: cities_array[i][0], lng: cities_array[i][1]},"
    oFile.WriteLine "              radius: " & dot_size
    oFile.WriteLine "            });"
    oFile.WriteLine "          }"
    
    'Icon = Pin
    oFile.WriteLine "          if (cities_array[i][2] == 'Pin') {"
    oFile.WriteLine "            var cityCircle = new google.maps.Marker({"
    oFile.WriteLine "              map: map,"
    oFile.WriteLine "              position: {lat: cities_array[i][0], lng: cities_array[i][1]},"
    oFile.WriteLine "              icon: 'http://maps.google.com/mapfiles/ms/icons/' + cities_array[i][3].toLowerCase() + '-dot.png'"
    oFile.WriteLine "            });"
    oFile.WriteLine "          }"
    
    oFile.WriteLine "        }"
    oFile.WriteLine "      }"
    
    oFile.WriteLine "    </script>"
    oFile.WriteLine "    <script async defer"
    oFile.WriteLine "    src=""https://maps.googleapis.com/maps/api/js?key=" & API_KEY & "&callback=initMap"">"
    oFile.WriteLine "    </script>"
    oFile.WriteLine "  </body>"
    oFile.WriteLine "</html>"
    
    oFile.Close
    MsgBox "Map sucessfully saved on desktop."
    
    Set fso = Nothing
    Set oFile = Nothing
    Application.ScreenUpdating = True
End Sub


Sub OpenMap()
    
    CreateObject("Shell.Application").Open Environ("USERPROFILE") & "\Desktop\my_map.html"
    
End Sub
