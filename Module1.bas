Attribute VB_Name = "Module1"
Function GetTract(ByVal a As String, b As String, c As String, d As String) As String
Dim hReq As Object, JSON As Dictionary
Dim sht As Worksheet: Set sht = ActiveSheet
Dim tract As String
Dim response As String
Dim totaladdress As String
Dim address As String
Dim city As String
Dim state As String
Dim zip As String
Dim splitstring As String

Dim strUrl As String
    strUrl = "https://geocoding.geo.census.gov/geocoder/geographies/address?street=" & a & "&city=" & b & "&state=" & c & "&zip=" & d & "&benchmark=Public_AR_Current&vintage=Current_Current&format=json" & """"
    strUrl = Left(strUrl, Len(strUrl) - 1)
    
Set hReq = CreateObject("MSXML2.XMLHTTP")
    
hReq.Open "GET", strUrl, False
hReq.Send
'MsgBox (hReq.ResponseText)

response = hReq.ResponseText

If InStr(1, response, "GEOID") = 0 Then
    tract = ""
Else
    tract = Mid(response, InStr(1, response, "GEOID"), 21)
    tract = Mid(tract, 9, 11)
    
    totaladdress = Mid(response, InStr(1, response, "matchedAddress"), 200)
    totaladdress = Mid(totaladdress, 18)
    
    address = Split(totaladdress, ",")(0)
    
    city = Split(totaladdress, ",")(1)
    city = Mid(city, 2)
    
    state = Split(totaladdress, ",")(2)
    state = Mid(state, 2)
    
    zip = Split(totaladdress, ",")(3)
    zip = Mid(zip, 2, 5)
    
End If

GetTract = tract
   
End Function

Sub Button2_Click()

UserForm1.Show

End Sub

Sub code()
Dim w As Worksheet: Set w = ActiveSheet
Dim Last As Integer: Last = w.Range("A1000").End(xlUp).Row
Dim address As String
Dim city As String
Dim state As String
Dim zip As String
Dim CTract As String
Dim i As Integer
Dim pctComp0 As Single
Dim pctComp1 As Single

For i = 2 To Last

address = w.Range("B" & i).Value
city = w.Range("C" & i).Value
state = w.Range("D" & i).Value
zip = w.Range("E" & i).Value
CTract = GetTract(address, city, state, zip)
w.Cells(i, 6).Value = CTract
    
pctComp0 = ((i / Last) * 100)
pctComp1 = Round(pctComp0, 0)
progress pctComp1

Next i

UserForm1.Hide

End Sub

Sub progress(pctCompl As Single)

UserForm1.Text.Caption = pctCompl & "% Completed"
UserForm1.Bar.Width = pctCompl * 2

DoEvents

End Sub
