<!--#include file="downloader.asp" -->
<%
Dim uri : uri = Request.Form("uri")
If (IsEmpty(uri)) Then uri = "http://yagami.net"

Dim TotalTime, WithGzip
If (Request.Form("send") = "ok") Then
	TotalTime = 0
	Dim StartTime : StartTime = timer()
	If (Request.Form("gzip") = "1") Then WithGzip = True Else WithGzip = False
	Download uri, "test.xml", WithGzip
	TotalTime = (timer() - StartTime)
	Dim GzipMode : GzipMode = "off"
	If (WithGzip = True) Then GzipMode = "on"
	Response.Redirect "?totaltime=" & Server.UrlEncode(TotalTime) & "&gzip=" & GzipMode
	Response.End
Else 
	TotalTime = Request.QueryString("totaltime")
	WithGzip = Request.QueryString("gzip")
End If

%>
<html>
	<head>
		<title>Classic ASP Gzip Downloader</title>
	</head>
	<body>
		<h3>Classic asp GZIP Downloader</h3>
		<hr />
		<form method="POST">
			<label for="uri">Url: </label>
			<input name="uri" id="uri" value="<%=uri%>" />
			<label>
				<input type="checkbox" name="gzip" value="1" checked="checked"/>gzip
			</label>
			<input name="send" value="ok" type="hidden" style="width: 30%;"/>
			<button type="submit" onclick="this.disabled = true;document.getElementById('result').remove();document.getElementById('wait').hidden = false;document.forms[0].submit();">Download</button>
		</form>
		<p id="result">
		<% If (NOT IsEmpty(TotalTime)) Then %>			
			<a href="test.xml">File Downloaded.</a> Download Time: <%=FormatNumber(TotalTime)%> secs. GZip Mode: <%=WithGzip%>
		<% End If %>
		</p>
		<p id="wait" hidden>
			Downloading...
		</p>
	</body>
</html>