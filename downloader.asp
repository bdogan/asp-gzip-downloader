<%

Function Download(Uri, FileName, ByRef WithGZIP)
	Dim gZipExec : gZipExec = Server.Mappath("gzip.exe") ' Location of GZIP exe
	If (IsEmpty(WithGZIP)) Then WithGZIP = True
	Download = False
	
	Set xmlHTTP = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	xmlHTTP.setTimeouts 900000, 900000, 900000, 900000
	xmlHTTP.open "GET", Uri, false
	xmlHTTP.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36"
	xmlHTTP.setRequestHeader "Accept", "text/xml;charset=ISO-8859-9"
	xmlHTTP.setRequestHeader "Accept-Charset", "ISO-8859-9"
	If (WithGZIP = True) Then xmlHTTP.setRequestHeader "Accept-Encoding", "gzip"
	xmlHTTP.send()
	
	If (xmlHttp.Status <> "200" OR Len(xmlHTTP.ResponseBody) = 0) Then Set xmlHTTP = Nothing : Exit Function
	
	On Error Resume Next
	
	If (xmlHTTP.GetResponseHeader("Content-Encoding") = "gzip") Then
		FileName = FileName & ".gz"
	End If
	
	Set WriteStream = Server.CreateObject("ADODB.Stream")
	WriteStream.Open
	WriteStream.Type = 1
	WriteStream.Write xmlHTTP.ResponseBody
	WriteStream.SaveToFile Server.Mappath(FileName), 2
	Set WriteStream = Nothing
	
	If (Err.Number <> 0) Then Exit Function
	Download = True
	
	If (xmlHTTP.GetResponseHeader("Content-Encoding") = "gzip") Then
		WithGZIP = True
		Dim WshShell : Set WshShell = CreateObject("WScript.Shell")		
		Dim Result : Result = WshShell.Run(gZipExec & " -q -f -d " & Server.Mappath(FileName), 0, True)
		If Result = 0 Then Download = True
	Else
		WithGZIP = False
	End If
	
	Set xmlHTTP = Nothing
End Function

%>