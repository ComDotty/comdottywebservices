<%

Dim strSecureProtocol
Dim strSecureServerName
Dim strSecureURL
Dim strRedirectURL

If Request.ServerVariables("SERVER_PORT")=80 Then
	strSecureProtocol = "https://"
	strSecureServerName = Request.ServerVariables("SERVER_NAME")
	strSecureURL = Request.ServerVariables("URL")
	
	strRedirectURL = strSecureProtocol & strSecureServerName & strSecureURL
	
	Response.Redirect strRedirectURL
	
End If

%>