Attribute VB_Name = "Afstand"
Attribute VB_Description = "tekst hier"

''''''''''''''''''''''''''''''''''
' Reisafstand berekenen
'
' start = startlocatie
' eind = eindlocatie
' vervoer = manier waarop te reizen
' eenheid = mogelijk in meters of kilometers
'
''''''''''''''''''''''''''''''''''

Public Function G_AFSTAND(start As String, eind As String, Optional vervoer As Variant, Optional eenheid As Variant) As Variant
Attribute G_AFSTAND.VB_Description = "tekst hier 1"

Dim Verv As String
Dim Eenh As String
Dim Link As String
Dim Bestemming As String
Dim Mode As String
Dim Taal As String
Dim APIKEY As String
Dim totaallink As String


   
''' Link opbouw '''
    Link = "https://maps.googleapis.com/maps/api/distancematrix/json?&origins="
    Bestemming = "&destinations="
    Mode = "&mode="
    Taal = "&language=nl"
    APIKEY = "&Key=AIzaSyCiqWhbdKzPjJwkVsq2GPSGrvW-xPONDhw"
    totaallink = Link & Replace(start, " ", "+") & Bestemming & Replace(eind, " ", "+") & Mode & Verv & Taal & APIKEY

''' Controleren op waarde in vervoer '''
' Openbaar vervoer is een registratienummer voor nodig bij google '
    If IsMissing(vervoer) = True Or IsEmpty(vervoer) = True Then
        Verv = "driving"
    Else
        If vervoer > 2 Then
          Verv = "driving"
        Else
          Select Case vervoer
             Case 0: Verv = "driving"
             Case 1: Verv = "walking"
             Case 2: Verv = "bicycling"
          End Select
        End If
    End If

''' Eenheid display '''
    If IsMissing(eenheid) = True Or IsEmpty(eenheid) = True Then
        Eenh = 0
    Else
        Eenh = eenheid
    End If

''' Oproepen informatie '''
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    URL = Link & Replace(start, " ", "+") & Bestemming & Replace(eind, " ", "+") & Mode & Verv & Taal & APIKEY
    objHTTP.Open "GET", URL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ("")
    
''' Als POST tekst niet klopt '''
    If InStr(objHTTP.responseText, """distance"" : {") = 0 Then GoTo Error
    
''' Als eenheid in meters '''
    meters = Right(objHTTP.responseText, Len(objHTTP.responseText) - InStr(objHTTP.responseText, """value"" : ") - 9)
    
''' Als eenheid in kilometers '''
    kilometers = Right(objHTTP.responseText, Len(objHTTP.responseText) - InStr(objHTTP.responseText, """text"" : """) - 9)
    
''' Eindresultaat maken '''
    If Eenh = 1 Then
    G_AFSTAND = CDbl(Replace(Split(meters)(0), ".", ","))
    Else
    G_AFSTAND = CDbl(Replace(Split(kilometers, " km""")(0), ".", ","))
    End If
    Exit Function
    
''' Error uitgang '''
Error:
    G_AFSTAND = totaallink

End Function



