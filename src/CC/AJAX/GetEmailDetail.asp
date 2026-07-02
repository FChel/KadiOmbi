<!DOCTYPE html>
<%

Dim objCon
Dim objRS
Dim x
Dim lngEmailDetailID
Dim strEmailTemplateName
Dim strEmailSubject
Dim strEmailHeader
Dim strEmailBody
Dim strEmailFooter
Dim strEmailAttachment
Dim strEmailImportance
Dim strEmailSensitivity
Dim strFromAddress
Dim arrImportance(3)
Dim arrSensitivity(8)
Dim strSelected
Dim lngErrorEmailID

arrImportance(1) = "Low"
arrImportance(2) = "Normal"
arrImportance(3) = "High"

arrSensitivity(1) = "[SEC=UNOFFICIAL]"
arrSensitivity(2) = "[SEC=OFFICIAL]"
arrSensitivity(3) = "[SEC=OFFICIAL: Sensitive, ACCESS=Personal-Privacy]"
arrSensitivity(4) = "[SEC=OFFICIAL: Sensitive, ACCESS=Legal-Privilege]"
arrSensitivity(5) = "[SEC=OFFICIAL: Sensitive, ACCESS=Legislative-Secrecy]"
arrSensitivity(6) = "[SEC=PROTECTED: Sensitive, ACCESS=Personal-Privacy]"
arrSensitivity(7) = "[SEC=PROTECTED: Sensitive, ACCESS=Legal-Privilege]"
arrSensitivity(8) = "[SEC=PROTECTED: Sensitive, ACCESS=Legislative-Secrecy]"
 
'Open Database Connection
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objCon = Server.CreateObject("ADODB.Connection")

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("EmailDetailID")) Then
	lngEmailDetailID = Request.QueryString("EmailDetailID")
	Session("EmailDetailID") = lngEmailDetailID
End If



Public Sub Get_Record()

 'Description:	Loads tblCAPSEmailDetail

	objRS.Open "SELECT * FROM tblCAPSEmailDetail WHERE EmailDetailID = " & Session("EmailDetailID") & "",objCon
	
		If Not objRS.EOF Then

				lngEmailDetailID = objRS("EmailDetailID")
			  	strEmailTemplateName = objRS("EmailTemplateName")
			  	strEmailSubject = objRS("EmailSubject")
				strEmailHeader = objRS("EmailHeader")
				strEmailBody = objRS("EmailBody")
				strEmailFooter = objRS("EmailFooter")
				strEmailAttachment = objRS("EmailAttachment")
				strEmailImportance = objRS("EmailImportance")
				strEmailSensitivity = objRS("EmailSensitivity")
				strFromAddress = objRS("FromAddress")
				lngErrorEmailID = objRS("EmailErrorID")
		End If		

	objRS.Close
	
End Sub

Call Get_Record()

%>
<html>
  <head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta charset="utf-8" />
    <title>MyFi - Bootstrap Components</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <link href="fontawesome/css/all.css" rel="stylesheet" />
    <link rel="stylesheet" href="css/myfi-bootstrap.css" />
  </head>
 
  <body>
    <form class="needs-validation" novalidate>
	
	
	<div class="container">	
      	<div class="form-row">
        	<div class="col-md-12 mb-3">
				<label for="EmailTemplateName">Template Name</label>
				<input type="text" class="form-control" name="EmailTemplateName" id="EmailTemplateName" value="<%=strEmailTemplateName%>" maxlength="50" required>
				<input type="hidden" class="form-control" name="EmailDetailID" id="EmailDetailID" value="<%=lngEmailDetailID%>" >
          		<div class="valid-feedback">Valid </div>
		  		<div class="invalid-feedback">Template name must be entered.</div>
			</div>
		</div>

		<div class="form-row">
        	<div class="col-md-12 mb-3">
          		<label for="EmailSubject">Subject</label>
          		<input type="text" class="form-control" name="EmailSubject" placeholder="Subject" value="<%=strEmailSubject%>" maxlength="200" required>
          		<div class="valid-feedback">Valid</div>
		  		<div class="invalid-feedback">Subject must be entered.</div>
			</div>
		</div> 

		<div class="form-row">
			<div class="col-md-6 mb-3">
				<label for="EmailImportance">Importance</label>
				<SELECT class="form-control" name="EmailImportance" id="EmailImportance" >
				<%
					For x = 1 to 3
						If arrImportance(x) = cstr(strEmailImportance) Then
							strSelected = " SELECTED "
						Else
							strSelected = ""
						End If
							Response.Write "<option Value=""" & arrImportance(x) & """" & strSelected & ">" & arrImportance(x) & "</OPTION>"
					Next
				%>
				</Select>
			</div>
		</div>

		<div class="form-row">
			<div class="col-md-6 mb-3">
				<label for="EmailSensitivity">Sensitivity</label>
				<SELECT class="form-control" name="EmailSensitivity" id="EmailSensitivity" >
					<%
					For x = 1 to 8
						If arrSensitivity(x) = cstr(strEmailSensitivity) Then
							strSelected = " SELECTED "
						Else
							strSelected = ""
						End If
							Response.Write "<option Value=""" & arrSensitivity(x) & """" & strSelected & ">" & arrSensitivity(x) & "</OPTION>"
					Next
					%>
				</Select>
			</div>
		</div>

		<div class="form-row">
			<div class="col-md-12 mb-3">
				<label for="EmailHeader">Header</label>
				<textarea class="form-control" name="EmailHeader" id="EmailHeader" rows="15" cols="1" wrap="hard" required><%=strEmailHeader%></textarea>
		  		<div class="valid-feedback">Valid</div>
			</div>
		</div>

		<div class="form-row">
			<div class="col-md-12 mb-3">
				<label for="EmailBody">Body</label>
				<textarea class="form-control" name="EmailBody" id="EmailBody" rows="15" cols="1" ><%=strEmailBody%></textarea>
			  	<div class="valid-feedback">Valid</div>
			</div>
		</div>	

		<div class="form-row">
			<div class="col-md-12 mb-3">
				<label for="EmailFooter">Footer</label>
				<textarea class="form-control" name="EmailFooter" id="EmailFooter" rows="15" cols="1" ><%=strEmailFooter%></textarea>
		  		<div class="valid-feedback">Valid</div>
			</div>
		</div>		

		<div class="form-row">
			<div class="col-md-12 mb-3">
	  			<label for="FromAddress">From Address</label>
	  			<input type="text" class="form-control" name="FromAddress" id="FromAddress" maxlength="500" value="<%=strFromAddress%>" required>
	  			<div class="invalid-feedback">Please enter a valid email address.</div>
			</div>
		</div>

		<div class="form-row">
        	<div class="col-md-12 mb-3">
          		<label for="EmailAttachments">Attachments</label>
          		<input type="file" class="form-control" name="EmailAttachments" id="EmailAttachments" maxlength="500" value="<%=strEmailAttachment%>" >
        	</div>
		</div>		
		<div class="form-row">
        	<div class="col-md-12 mb-3">
          		<label for="EmailAttachments">Email Error ID</label>
          		<input type="number" class="form-control" name="EmailErrorID" id="EmailErrorID" value="<%=lngErrorEmailID%>" >
        	</div>
		</div>	

		<div class="modal-footer">
			<button type="button" class="btn btn-link" data-dismiss="modal">Cancel</button>
			<button type="submit" class="btn btn-primary">Save changes</button>
  		</div>

		<div class="row">
	  		<div class="col-md-12 mb-3">		
				<div class="alert alert-success" role="alert" id="AlertSuccess" style="display:none">Record has been successfully saved.</div>
				<div class="alert alert-danger" role="alert" id="AlertDanger" style="display:none">Invalid entries exist.  Please correct and resubmit.</div>
		</div>
		
	</div>

</form>
</body>
</html>

