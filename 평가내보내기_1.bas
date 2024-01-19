Attribute VB_Name = "Module18"
Sub SendJSON()
    Dim JsonString As String
    Dim xmlhttp As Object
    
    ' Define the JSON string directly in the code
    JsonString = StrConv("test=test1", vbFromUnicode)
    
    ' Create a new XML HTTP request
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' The URL to send the request to
    Dim url As String
    url = "http://urosys-web.juroinstruments.com/app/createValWebJob"
    
    ' Open the HTTP request as a POST method
    xmlhttp.Open "POST", url, False
    
    ' Set the request content-type header to application/json
    xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    
    ' Send the request with the JSON string
    xmlhttp.Send "officeCd=BO&name=TEST11&valDate=20231228&valTypeCode=P&greekLevel=&contextIds=BO&dataSetIds=official&simId=&priority=4&itemCodes=ELS3588"
    
    ' Check the status of the request
    If xmlhttp.Status = 200 Then
        ' If the request was successful, output the response
        MsgBox xmlhttp.responseText
    Else
        ' If the request failed, output the status
        MsgBox "Error: " & xmlhttp.Status & " - " & xmlhttp.statusText
    End If
    
    ' Clean up
    Set xmlhttp = Nothing
End Sub

'평가작업'
