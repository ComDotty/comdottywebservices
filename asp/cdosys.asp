<%

On Error Resume Next

' Spam flag
Dim blnSpam

' Server verification
Dim intComp, strReferer, strServer

' Mail message
Dim sSubject, sTo, sFrom, sName, sMail, sPhone, iSelect, sEnquiry, sMailBody, sReadReceipt, sMsg

' SMTP Host configuration
Dim oMail, oConfig, oConfigFields

CONST SMTPSendUsing = 2 ' Send using Port (SMTP over the network)
CONST SMTPServer = "auth.smtp.1and1.co.uk"

' For local testing only
' CONST SMTPServer = "localhost"

CONST SMTPServerPort = 25
CONST SMTPConnectionTimeout = 10 'seconds

blnSpam = False

' Check if this page is being called from a another page on the server
strReferer = Request.ServerVariables("HTTP_REFERER")
strServer = Replace(Request.ServerVariables("SERVER_NAME"), "www.", "")

intComp = inStr(strReferer, strServer)

If intComp = 0 Then
' Block spam attempt
	blnSpam = True
End If


iSelect = Request.Form("frmsubject")

Select Case iSelect
	
	Case 1
	
		sSubject = "Application Development"
		
	Case 2
	
		sSubject = "Database Development"
		
	Case 3
	
		sSubject = "Domains & Hosting"
				
	Case 4
	
		sSubject = "eCommerce - Shopify"
		
	Case 5
	
		sSubject = "eCommerce - Shopify/Tagalog Language Conversion"
		
	Case 6
	
		sSubject = "Marketing & Analytics"
		
	Case 7
	
		sSubject = "Progressive Web Apps"
		
	Case 8
	
		sSubject = "Server Management"
		
	Case 9
	
		sSubject = "Web Design & Development"

	Case 10
	
		sSubject = "Something Else"
		
	Case 11
	
		sSubject = "Free Consultancy Offer"
		
	Case Else
	
		blnSpam = True
		
End Select

' Build mail message
If CStr(Request.Form("frmname")) = "" Then
' Block spam attempt
	blnSpam = True
Else
	sName = CStr(Request.Form("frmname"))
End If

If CStr(Request.Form("frmemail")) = "" Then
' Block spam attempt
	blnSpam = True
Else
	sMail = CStr(Request.Form("frmemail"))
End If

If iSelect <> 11 Then

	If CStr(Request.Form("frmphone")) = "" Then
	' Block spam attempt
		blnSpam = True
	Else
		sPhone = CStr(Request.Form("frmphone"))
	End If

End If

If CStr(Request.Form("frmenquiry")) = "" Then
' Block spam attempt
	blnSpam = True
ElseIf InStr(CStr(Request.Form("frmenquiry")), "http") Then
' Block spam attempt
	blnSpam = True
ElseIf InStr(CStr(Request.Form("frmenquiry")), "feedback form") Then
' Block spam attempt
	blnSpam = True
ElseIf InStr(CStr(Request.Form("frmenquiry")), "Feedback") Then
' Block spam attempt
	blnSpam = True
ElseIf InStr(CStr(Request.Form("frmenquiry")), "explainer video") Then
' Block spam attempt
	blnSpam = True
Else
	sEnquiry = CStr(Request.Form("frmenquiry"))
End If

' Redirect if spam flag set true
If blnSpam Then
	Response.Redirect "../pages/trap.html"
Else

	sTo = "les.piper@comdotty.com"
	sFrom = "enquiries@comdotty.com"
	
	sMailBody = "Name: " & sName & vbCrLf & vbCrLf
	sMailBody = sMailBody & "E-mail: " & sMail & vbCrLf
	sMailBody = sMailBody & "Phone: " & sPhone & vbCrLf
	sMailBody = sMailBody & "Subject: " & sSubject & vbCrLf
	sMailBody = sMailBody & "Enquiry: " & sEnquiry
	sMailBody = sMailBody & vbCrLf & vbCrLf
	sMailBody = sMailBody & "=================================" & vbCrLf
	sMailBody = sMailBody & "Form Submission Source:" & vbCrLf
	sMailBody = sMailBody & "Server Name: " & Request.ServerVariables("SERVER_NAME") & vbCrLf
	sMailBody = sMailBody & "Client IP: " & Request.ServerVariables("REMOTE_ADDR") & vbCrLf
	sMailBody = sMailBody & "Server IP: " & Request.ServerVariables("LOCAL_ADDR") & vbCrLf
	sMailBody = sMailBody & "Request URI: " & Request.ServerVariables("REQUEST_URI") & vbCrLf
	
	sReadReceipt = true
	
	sMsg = sMailBody
	
	' Create CDO message object and SMTP host config
	Set oMail = Server.CreateObject("CDO.Message")
	Set oConfig = Server.CreateObject("CDO.Configuration")
	Set oConfigFields = oConfig.Fields
	
	With oConfigFields
		.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = SMTPSendUsing
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTPServer
		.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTPServerPort
		.Update
	End with
	
	oMail.Configuration = oConfig
	
	oMail.Subject = "Website Enquiry: " & sSubject
	oMail.From = sFrom
	oMail.To = sTo
	oMail.TextBody = sMsg
	
	' Send the mail
	oMail.Send
	
	' Clean up
	Set oMail = Nothing

	If iSelect = 11 Then
	
		' If FREE consultancy mail successful redirect to offer thank you page
		Response.Redirect "../pages/offer.html"

	Else
	
		' Otherwise redirect to standard thank you page
		Response.Redirect "../pages/sent.html"
	
	End If
	
End If

' Redirect on error
If Err.Number <> 0  Then
	Response.Redirect "../pages/error.html"
End If

%> 
