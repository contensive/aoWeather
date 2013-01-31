VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeatherClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Main As Object

Public Sub Init(MainObject As Object)

    Set Main = MainObject

    Exit Sub
    
ErrorTrap:
    Call HandleError("WeatherClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap

    Dim HTTP As New WinHttp.WinHttpRequest
    Dim objXML As New MSXML2.DOMDocument
    Dim NOAAResponse As String
    Dim CurrentTemperature As String
    Dim IconString As String
    Dim StationID As String
    Dim DataRequestPage As String
    Dim WeatherString As String
    Dim Stream As String
    Dim startPos As Long
    Dim endPos As Long
    Dim subString As String
    
    StationID = Main.GetAggrOption("StationID", OptionString)
'Call Main.TestPoint("100")
    If StationID <> "" Then
    
        DataRequestPage = "http://www.weather.gov/data/current_obs/" & StationID & ".xml"
    
        objXML.setProperty "ServerHTTPRequest", True
        objXML.async = False
        
        If objXML.Load(DataRequestPage) Then
            Call objXML.Load(DataRequestPage)
        End If
'Call Main.TestPoint("200")
        If objXML.hasChildNodes Then
            CurrentTemperature = objXML.getElementsByTagName("temperature_string").Item(0).Text
            IconString = objXML.getElementsByTagName("icon_url_base").Item(0).Text
            IconString = IconString & objXML.getElementsByTagName("icon_url_name").Item(0).Text
            WeatherString = objXML.getElementsByTagName("weather").Item(0).Text
        End If
'Call Main.TestPoint("300")
        Set objXML = Nothing
        
        CurrentTemperature = Replace(CurrentTemperature, " F", "&deg; F", 1, 99, vbTextCompare)
        CurrentTemperature = Replace(CurrentTemperature, " C", "&deg; C", 1, 99, vbTextCompare)
'Call Main.TestPoint("400")
        If InStr(1, CurrentTemperature, "(", vbTextCompare) <> 0 Then
            startPos = InStr(1, CurrentTemperature, "(")
            endPos = InStr(1, CurrentTemperature, ")")
            subString = Mid(CurrentTemperature, startPos, (endPos - startPos) + 1)
            CurrentTemperature = Replace(CurrentTemperature, subString, "", , , vbTextCompare)
        End If
'Call Main.TestPoint("500")
        Stream = "<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>"
        Stream = Stream & "<TR>"
        Stream = Stream & "<TD class=""weather""><nobr>" & CurrentTemperature & " - </nobr></TD>"
        Stream = Stream & "<TD class=""weather"">&nbsp;<nobr>" & WeatherString & "</nobr></TD>"
        Stream = Stream & "<TD width=100%>"
        If IconString <> "" Then
            Stream = Stream & "&nbsp;&nbsp;<img src=""" & IconString & """ width=""16"" height=""16"">"
        End If
        Stream = Stream & "</TD>"
        Stream = Stream & "</TR>"
        Stream = Stream & "</TABLE>"
'Call Main.TestPoint("600")
    Else
        Stream = "A Surface Weather Observation Station ID is required to display current weather conditions. <br /><br />" _
               & "A list of current observation stations can found here <a target=""_blank"" href=""http://www.faa.gov/AIRPORTS_AIRTRAFFIC/WEATHER/ASOS/"">http://www.faa.gov/airports_airtraffic/weather/asos/</a>.<br /><br />" _
               & "To set the Station ID, turn on Advanced Edit, and hover over the settings icon at the top of the Add-on on the page."
               
        Stream = Main.GetAdminHintWrapper(Stream)
        
    End If
    
    GetContent = "<div class=""containerWeather"">" & Stream & "</div>"

    Exit Function
ErrorTrap:
    Call HandleError("WeatherClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function

