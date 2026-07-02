<!DOCTYPE html>
<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"


Dim objCon
Dim objRS
Dim x
Dim lngEmailID
Dim strEmailTemplateName
Dim strEmailSubject
Dim strEmailHeader
Dim strEmailBody
Dim strEmailFooter
Dim strEmailAttachment
Dim strEmailImportance
Dim strEmailSensitivity
Dim strFromAddress
Dim strSelected
Dim lngErrorEmailID
Dim strDaysSince
Dim strDaysText

'Open Database Connection
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objCon = Server.CreateObject("ADODB.Connection")

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("EmailID")) Then
	lngEmailID = Request.QueryString("EmailID")
Else
	lngEmailID = 0
End If

	'Create the initial DIV tags
	Response.Write "<div class=""container""><div class=""form-row"">"
	
	'Description:	Loads Email Sent Details in modal on Email Admin screen
	objRS.Open "SELECT * FROM tblCAPSEmail WITH(NOLOCK) WHERE EmailID = " & lngEmailID & "",objCon
	
		If Not objRS.EOF Then

			'Response.Write "<div class=""col-md-12 mb-3""><label for=""EmailID"">Email ID</label>" & _
			'	"<input type=""text"" class=""form-control"" name=""EmailDetailID"" id=""EmailID"" value=""" & objRS("EmailID") & """ ></div></div>"
			
			If IsNull(objRS("DateSent")) or objRS("DateSent")="" Then
				strDaysSince = ""'"<span class=""badge badge-pill badge-danger"">N/A</span>"
			Else
				strDaysText = DateDiff("d",objRS("DateSent"),Now())
				IF clng(strDaysText) = 1 Then 
					strDaysText = "Sent Yesterday"
				ElseIf clng(strDaysText) > 365 Then
					strDaysText = DateDiff("yyyy",objRS("DateSent"),Now())
					IF clng(strDaysText) >1 Then
						strDaysText = "Sent " & DateDiff("yyyy",objRS("DateSent"),Now()) & " Years Ago"
					Else
						strDaysText = "Sent " & DateDiff("yyyy",objRS("DateSent"),Now()) & " Year Ago"
					End If
				Else
					strDaysText = "Sent " & strDaysText & " Days Ago"
				End IF
				
				strDaysSince = "<span class=""badge badge-pill badge-success"">" & strDaysText & "</span>"
			End If
			
			Response.Write "<table class=""table table-compact text-left""><thead><tr>" & _
			"<th> Email ID </th><td>" & objRS("EmailID") & "</td><tr>" & _
			"<th> Employee ID </th><td>" & objRS("EmployeeID") & "</td><tr>" & _
			"<th> Recipient </th><td>" & objRS("RecipientTitle") & " " & objRS("RecipientFirstName") & " " & objRS("RecipientLastName") & "</td><tr>" & _
			"<th> Receipient Email </th><td>" & objRS("EmailToAddress") & "</td><tr>" & _
			"<th> Sent By  </th><td>" & objRS("FromAddress") & "</td><tr>" & _
			"<th> Status </th><td>" & objRS("Status") & "</td><tr>" & _
			"<th> Date Sent </th><td>" & objRS("DateSent") & "<span style=""float:right"">" & strDaysSince & "</span></td><tr>" & _
			"<th> Subject </th><td>" & objRS("EmailSubject") & "</td><tr>" & _
			"<th style=""vertical-align:top;""> Body </th><td>" & objRS("EmailBody") & "</td><tr>" 
			
		Else
			Response.Write "No Email Found..."
		End If		

	objRS.Close
	

	Response.Write "</div>"
	
%>



