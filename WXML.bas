Attribute VB_Name = "WXML"
Option Explicit
'chr(176) is "degrees" symbol
Const img31Path = "http://image.weather.com/web/common/wxicons/31/"
Const img52Path = "http://image.weather.com/web/common/wxicons/52/"

Private objXML As MSXML.XMLHTTPRequest

Public Type lWX
  Name As String
  IMG As Integer
  Cond As String
  Temp As String
  flTemp As String
  UVIndex As String
  Wind(1 To 2) As String
  Humidity As String
  Pressure As String
  PressureIMG As Integer
  Dew As String
  Visibility As String
  Udated As String
  SR0 As String
  SR1 As String
  SS0 As String
  SS1 As String
  Daylight As String
  f1Title As String
  f1 As String
  f1IMG As Integer
  f2Title As String
  f2 As String
  f2IMG As Integer
  f3Title As String
  f3 As String
  f3IMG As Integer
End Type

Public localWX As lWX

Public Function GetWX(mZIP As String) As Boolean
  On Error Resume Next
  Dim strURL As String
  Dim strHTM As String
  Dim mStart As Integer
  Dim mEnd As Integer
  Set objXML = New MSXML.XMLHTTPRequest
  
  'set the URL to get local weather for the given zip code
  strURL = "http://www.weather.com/weather/local/" & mZIP & "?lswe=" & mZIP & "&lwsa=WeatherLocalUndeclared"
  
  'get the html page
  objXML.open "GET", strURL
  'pause for 1 second to allow for lag time
  waitFor 1
  'send the request to open the page
  objXML.send
  'pause for 1 second to allow for lag time
  waitFor 1
  'check for status
  If objXML.Status <> "200" Then
    MsgBox "Could not find weather for Zip Code " & mZIP & vbCrLf & _
           "Please check the Zip Code and try again.", vbCritical, mZIP & " failed"
    Exit Function
  End If
  'get the HTML code for the URL requested
  strHTM = objXML.responseText
  'make sure we found a hit
  If InStr(1, strHTM, "No items found.") Then
    MsgBox "Could not find weather for Zip Code " & mZIP & vbCrLf & _
           "Please check the Zip Code and try again.", vbCritical, "No results for " & mZIP
    Exit Function
  End If
  'now we have to parse the data and get the information we need out of it
  
  'widdle down the HTM to the start of the data
  strHTM = Mid(strHTM, InStr(1, strHTM, "CLASS=""moduleTitleBarGML"""))
  'find the start and end for the city name
  mStart = InStr(1, strHTM, "<BR>") + 4
  mEnd = InStr(1, strHTM, "<BR><FONT")
  localWX.Name = Mid(strHTM, mStart, mEnd - mStart)
  
  'widdle down the HTM some more to the first image
  strHTM = Mid(strHTM, InStr(1, strHTM, "<IMG SRC"))
  'now get the Image for current conditions
  localWX.IMG = GetPic(Mid(strHTM, 1, InStr(1, strHTM, ">") - 1), 52)
  'get conditions text
  mStart = InStr(1, strHTM, "A>") + 2
  mEnd = InStr(1, strHTM, "</B>")
  localWX.Cond = Mid(strHTM, mStart, mEnd - mStart)
  'widdle down the HTML some more
  strHTM = Mid(strHTM, mEnd)
  'get the temperature
  mStart = InStr(1, strHTM, "A>") + 2
  mEnd = InStr(1, strHTM, "&")
  localWX.Temp = Mid(strHTM, mStart, mEnd - mStart) & Chr(176) & "F"
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "Feels Like"))
  'get the "feels like" temp
  mStart = InStr(1, strHTM, "<BR>") + 4
  mEnd = InStr(1, strHTM, "&")
  localWX.flTemp = Trim(Mid(strHTM, mStart, mEnd - mStart) & Chr(176) & "F")
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "Updated"))
  mStart = 1
  mEnd = InStr(1, strHTM, "</TD>")
  localWX.Udated = Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "UV"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "A"">"))
  'get the UV index
  mStart = 4
  mEnd = InStr(mStart, strHTM, "</td>")
  localWX.UVIndex = Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "Wind"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "A"">"))
  'get the wind
  mStart = 4
  mEnd = InStr(mStart, strHTM, "</td>")
  localWX.Wind(2) = Mid(strHTM, mStart, mEnd - mStart)
  If InStr(1, localWX.Wind(2), "<BR>") Then
    localWX.Wind(1) = Mid(localWX.Wind(2), 1, InStr(1, localWX.Wind(2), "<BR>") - 2)
    localWX.Wind(2) = Mid(localWX.Wind(2), InStr(1, localWX.Wind(2), "<BR>") + 5)
    frmMain.tmrWind.Enabled = True
  Else
    localWX.Wind(1) = localWX.Wind(2)
    frmMain.tmrWind.Enabled = False
  End If
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "Humidity"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "A"">"))
  'get the humidity
  mStart = 4
  mEnd = InStr(mStart, strHTM, "</td>")
  localWX.Humidity = Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "Pressure"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "A"">"))
  'get the barometric pressure
  mStart = 4
  mEnd = InStr(mStart, strHTM, "<IMG")
  localWX.Pressure = Mid(strHTM, mStart, mEnd - mStart)
  localWX.Pressure = Replace(localWX.Pressure, " ", "")
  localWX.Pressure = Replace(localWX.Pressure, "&nbsp;", " ")
  localWX.Pressure = Replace(localWX.Pressure, Chr(9), "")
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "<IMG") + 4)
  strHTM = Mid(strHTM, InStr(1, strHTM, "<IMG"))
  'get the pressure direction
  localWX.PressureIMG = GetPressure(Mid(strHTM, 1, InStr(1, strHTM, ">")))
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "Dew"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "A"">"))
  'get the Dew Point
  mStart = 4
  mEnd = InStr(1, strHTM, "&")
  localWX.Dew = Mid(strHTM, mStart, mEnd - mStart) & Chr(176) & "F"
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "Visibility"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "A"">"))
  'get the visibility
  mStart = 4
  mEnd = InStr(mStart, strHTM, "</td>")
  localWX.Visibility = Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "<script"))
  'get sunrise for f1
  mStart = InStr(1, strHTM, "srise0=") + 7
  mEnd = InStr(mStart, strHTM, "&") - 1
  localWX.SR0 = Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, mEnd)
  'get sunrise for tomorrow
  mStart = InStr(1, strHTM, "srise1=") + 7
  mEnd = InStr(mStart, strHTM, "&") - 1
  localWX.SR1 = Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, mEnd)
  'get sunset for f1
  mStart = InStr(1, strHTM, "sset0=") + 6
  mEnd = InStr(mStart, strHTM, "&") - 1
  localWX.SS0 = Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, mEnd)
  'get sunset for tomorrow
  mStart = InStr(1, strHTM, "sset1=") + 6
  mEnd = InStr(mStart, strHTM, "&") - 1
  localWX.SS1 = Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, mEnd)
  'get daylight remaining
  mStart = InStr(1, strHTM, "dlight0=") + 8
  mEnd = InStr(mStart, strHTM, "&") - 1
  localWX.Daylight = Mid(strHTM, mStart, mEnd - mStart)
  'widdle down to 36-Hour Forcast
  strHTM = Mid(strHTM, InStr(1, strHTM, "36-Hour"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "<!-- day buttons -->"))
  'widdle down to get f1,f2,f3 titles
  strHTM = Mid(strHTM, InStr(1, strHTM, "<B>") + 3)
  mEnd = InStr(1, strHTM, "<") - 1
  localWX.f1Title = Mid(strHTM, 1, mEnd)
  'widdle down to get f1,f2,f3 titles
  strHTM = Mid(strHTM, InStr(1, strHTM, "<B>") + 3)
  mEnd = InStr(1, strHTM, "<") - 1
  localWX.f2Title = Mid(strHTM, 1, mEnd)
  'widdle down to get f1,f2,f3 titles
  strHTM = Mid(strHTM, InStr(1, strHTM, "<B>") + 3)
  mEnd = InStr(1, strHTM, "<") - 1
  localWX.f3Title = Mid(strHTM, 1, mEnd)
  
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "<!-- if alert .. use this code -->"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "<IMG") + 4)
  strHTM = Mid(strHTM, InStr(1, strHTM, "<img"))
  'get f1 image
  localWX.f1IMG = GetPic(Mid(strHTM, 1, InStr(1, strHTM, ">")), 31)
  'widdle to next <FONT> tag
  strHTM = Mid(strHTM, InStr(1, strHTM, "font"))
  'get contidion for f1
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f1 = Mid(strHTM, mStart, mEnd - mStart)
  'widdle to next <FONT> tag
  strHTM = Mid(strHTM, mEnd + 7)
  strHTM = Mid(strHTM, InStr(1, strHTM, "font"))
  'get temp for f1
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f1 = localWX.f1 & vbCrLf & Mid(strHTM, mStart, mEnd - mStart)
  
  
  'widdle to next <FONT> tag
  strHTM = Mid(strHTM, mEnd + 7)
  strHTM = Mid(strHTM, InStr(1, strHTM, "font"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "<nobr>") + 1)
  'get temp for f1
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f1 = localWX.f1 & " " & Replace(Mid(strHTM, mStart, mEnd - mStart), "&deg;", Chr(176))
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "<!-- Precipitation line will always be displayed -->"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "<table"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f1 = localWX.f1 & vbCrLf & "Precip: " & Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "<DIV") + 1)
  'get comments for f1
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f1 = localWX.f1 & vbCrLf & vbCrLf & Mid(strHTM, mStart, mEnd - mStart)
  'widdle down to next 31 pix size URL string
  strHTM = Mid(strHTM, InStr(1, strHTM, img31Path))
  'get f2 image
  localWX.f2IMG = GetPic(Mid(strHTM, 1, InStr(1, strHTM, ">")), 31)
  'widdle to next <FONT> tag
  strHTM = Mid(strHTM, InStr(1, strHTM, "font"))
  'get contidion for f2
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f2 = Mid(strHTM, mStart, mEnd - mStart)
  'widdle to next <FONT> tag
  strHTM = Mid(strHTM, mEnd + 7)
  strHTM = Mid(strHTM, InStr(1, strHTM, "font"))
  'get temp for f2
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f2 = localWX.f2 & vbCrLf & Mid(strHTM, mStart, mEnd - mStart)
  'widdle to next <FONT> tag
  strHTM = Mid(strHTM, mEnd + 7)
  strHTM = Mid(strHTM, InStr(1, strHTM, "font"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "<nobr>") + 1)
  'get temp for f2
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f2 = localWX.f2 & " " & Replace(Mid(strHTM, mStart, mEnd - mStart), "&deg;", Chr(176))
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "<!-- Precipitation line will always be displayed -->"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "<table"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f2 = localWX.f2 & vbCrLf & "Precip: " & Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "<DIV") + 1)
  'get comments for f2
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f2 = localWX.f2 & vbCrLf & vbCrLf & Mid(strHTM, mStart, mEnd - mStart)
  localWX.f2 = Replace(localWX.f2, Chr(9), "")
  
  'widdle down to next 31 pix size URL string
  strHTM = Mid(strHTM, InStr(1, strHTM, img31Path))
  'get f3 image
  localWX.f3IMG = GetPic(Mid(strHTM, 1, InStr(1, strHTM, ">")), 31)
  'widdle to next <FONT> tag
  strHTM = Mid(strHTM, InStr(1, strHTM, "font"))
  'get contidion for f3
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f3 = Mid(strHTM, mStart, mEnd - mStart)
  'widdle to next <FONT> tag
  strHTM = Mid(strHTM, mEnd + 7)
  strHTM = Mid(strHTM, InStr(1, strHTM, "font"))
  'get temp for f3
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f3 = localWX.f3 & vbCrLf & Mid(strHTM, mStart, mEnd - mStart)
  'widdle to next <FONT> tag
  strHTM = Mid(strHTM, mEnd + 7)
  strHTM = Mid(strHTM, InStr(1, strHTM, "font"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "<nobr>") + 1)
  'get temp for f3
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f3 = localWX.f3 & " " & Replace(Mid(strHTM, mStart, mEnd - mStart), "&deg;", Chr(176))
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "<!-- Precipitation line will always be displayed -->"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "<table"))
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  strHTM = Mid(strHTM, InStr(1, strHTM, "TD") + 1)
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f3 = localWX.f3 & vbCrLf & "Precip: " & Mid(strHTM, mStart, mEnd - mStart)
  'widdle
  strHTM = Mid(strHTM, InStr(1, strHTM, "<DIV") + 1)
  'get comments for f3
  mStart = InStr(1, strHTM, ">") + 1
  mEnd = InStr(1, strHTM, "<")
  localWX.f3 = localWX.f3 & vbCrLf & vbCrLf & Mid(strHTM, mStart, mEnd - mStart)
  localWX.f3 = Replace(localWX.f3, Chr(9), "")
  GetWX = True
  Set objXML = Nothing
End Function

Private Function GetPic(STR As String, mSize As Integer) As Integer
  'get the picture index number out of the IMG SRC HTML
  On Error GoTo ErrHandler
  Dim mStart As Integer
  Dim mEnd As Integer
  Select Case mSize
    Case 52
      mStart = InStr(1, STR, img52Path) + Len(img52Path)
      mEnd = InStrRev(STR, ".")
      GetPic = CInt(Mid(STR, mStart, mEnd - mStart))
    Case 31
      mStart = InStr(1, STR, img31Path) + Len(img31Path)
      mEnd = InStrRev(STR, ".")
      GetPic = CInt(Mid(STR, mStart, mEnd - mStart))
  End Select
  Exit Function
ErrHandler:
  'if this fails return the "N/A" index of 44
  GetPic = 44
End Function

Private Function GetPressure(STR As String) As Integer
  'pull the image name form the HTML and set a picture
  'index based on the image name
  Dim mStart As Integer
  Dim mEnd As Integer
  Dim tmp As String
  mStart = InStrRev(STR, "/") + 1
  mEnd = InStrRev(STR, "_") - mStart
  tmp = Mid(STR, mStart, mEnd)
  Select Case tmp
    Case "up"
      GetPressure = 1
    Case "down"
      GetPressure = 2
    Case "steady"
      GetPressure = 3
  End Select
End Function

Private Function TrimFont(STR As String) As String
  
End Function
