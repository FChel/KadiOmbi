<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strCSFromDinersID
Dim strStatus
Dim strName
Dim strSubStatPos

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("CSFromDinersID")) Then
	strCSFromDinersID = " WHERE CSFromDinersID = " & Request.QueryString("CSFromDinersID") & ""
End If


Public Sub ShowDetails()

	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM qryCAPSCSFromDinersAuditLog " & strCSFromDinersID,objCon

			  
		If Not objRS.EOF Then
			
			'Write the Header
			'Response.write "<div class=""panel-content row""><div class=""form-row col-md-6"" style=""font-weight:bold;"">CS From Diners</div><div class=""form-row col-md-6"" style=""font-weight:bold;""><A HREF=""#"" onClick=""loadCard(" & objRS(0) & ")"">Audit Log Details</a> </div></div>"
			
			'Response.write "<div class=""row""><div class=""col-md-6""><h6 class=""mb-3"">CS From Diners</h6><table class=""table table-compact"">"

			Response.write "<div class=""row""><div class=""col-md-6""><div class=""row mb-3""><div class=""col-md-6""><h6 class=""mb-3"">CS From Diners</h6></div>" & _
				"<div class=""col-md-6""><select class=""form-control""><option>Record 1</option><option>Record 2</option><option>Record 3</option></select></div></div><table class=""table table-compact"">"
				
			'Create the Status list badge based on the status field
			If IsNull(objRS("Status")) Then
				strStatus = ""
			Else
				If objRS("Status") = "Processed" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Processed</span>"
				ElseIf objRS("Status") = "Imported" Then
					strStatus = "<span class=""badge badge-pill badge-warning"">Imported</span>"
				Else
					strStatus = objRS("Status")
				End If
			End If
		
			If IsNull(objRS("Surname")) Then
				strName = ""
			Else
				strName = Trim(Trim(objRS("Title")) & " " & Trim(objRS("GivenNames")) & " " & Trim(objRS("Surname")))
			End If
			
			If IsNull(objRS("Suburb")) Then
				strSubStatPos = ""
			Else
				strSubStatPos = Trim(Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode")))
			End If
			
			'Response.write "<tr onClick=""SelectFileSeqNum(String('" & objRS("CSFromDinersID") & "'));""><td>" & objRS("CSFromDinersID") & "</td><td>" & objRS("FileDateTime") & "</td><td>" & objRS("Status") & "</td>" & _
			'			"<td>" & objRS("EIDNo") & "</td><td>" & objRS("Title") & " " & objRS("GivenNames") & " " & objRS("Surname") & "</td></tr>"
			
		
			Response.Write "<tr><td>CS ID</td><td>" & objRS("CSFromDinersID") & "</td></tr>" & _
				"<tr><td>Status</td><td>" & strStatus & "</td></tr>" & _
				"<tr><td>File Seq Num</td><td>" & objRS("FileSeqNum") & "</td></tr>" & _
				"<tr><td>EID No</td><td>" & objRS("EIDNo") & "</td></tr>" & _
				"<tr><td>Cardholder</td><td>" & strName & "</td></tr>" & _
				"<tr><td>Card Number</td><td>" & MaskCard(objRS("CardNo")) & "</td></tr>" & _
				"<tr><td>CardUpdateInd</td><td>" & objRS("CardUpdateInd") & "</td></tr>" & _
				"<tr><td>Card Expiry Date</td><td>" & objRS("CardExpiryDate") & "</td></tr>" & _
				"<tr><td>Name On Card</td><td>" & Trim(objRS("NameOnCard")) & "</td></tr>" & _
				"<tr><td>Address 1</td><td>" & Trim(objRS("Address1")) & "</td></tr>" & _
				"<tr><td>Address 2</td><td>" & Trim(objRS("Address2")) & "</td></tr>" & _
				"<tr><td>Address 3</td><td>" & Trim(objRS("Address3")) & "</td></tr>" & _
				"<tr><td>Suburb/ State/ PostCode</td><td>" & strSubStatPos & "</td></tr>" & _
				"<tr><td>Home Phone</td><td>" & objRS("HomePhone") & "</td></tr>" & _
				"<tr><td>Work Phone</td><td>" & objRS("WorkPhone") & "</td></tr>" & _
				"<tr><td>Mobile Phone</td><td>" & objRS("MobilePhone") & "</td></tr>" & _
				"<tr><td>Email</td><td>" & Trim(objRS("Email")) & "</td></tr>" & _
				"<tr><td>Report Group</td><td>" & objRS("ReportGroup") & "</td></tr>" & _
				"<tr><td>Credit Limit</td><td>" & objRS("CreditLimit") & "</td></tr>" & _
				"<tr><td>Relationship</td><td>" & objRS("Relationship") & "</td></tr>" & _
				"<tr><td>Cat2</td><td>" & objRS("Cat2") & "</td></tr>" & _
				"<tr><td>Account Number</td><td>" & objRS("AccountNumber") & "</td></tr>" & _
				"<tr><td>ActivationFlag</td><td>" & objRS("ActivationFlag") & "</td></tr>" & _
				"<tr><td>PlasticID</td><td>" & objRS("PlasticID") & "</td></tr>" & _
				"<tr><td>Companion</td><td>" & MaskCard(objRS("Companion")) & "</td></tr>" & _
				"<tr><td>Status</td><td>" & objRS("Status") & "</td></tr>" 
			
			
			Response.Write "</table></div><div class=""col-md-6""><div class=""row mb-3""><div class=""col-md-6 my-auto""><button type=""button"" class=""btn btn-outline-secondary"" onClick=""loadCard(" & objRS(0) & ")"">Audit Log Details</button></div>" & _
				"<div class=""col-md-6""><select class=""form-control""><option>Record 1</option><option>Record 2</option><option>Record 3</option></select>" & _
				"</div></div><table class=""table table-compact"">"
				
			'Response.Write "</table></div><div class=""col-md-6""><div class=""row mb-3""><div class=""col-md-6 my-auto""><h6><A HREF=""#"" onClick=""loadCard(" & objRS(0) & ")"">Audit Log details</A></h6></div>" & _
			'	"<div class=""col-md-6""><select class=""form-control""><option>Record 1</option><option>Record 2</option><option>Record 3</option></select>" & _
			'	"</div></div><table class=""table table-compact"">"

			Response.Write "<tr><td>Audit Log ID</td><td>" & objRS("AuditLogID") & "</td></tr>" & _
				"<tr><td>File Date Time</td><td>" & objRS("FileDateTime") & "</td></tr>" & _
				"<tr><td>Change Type</td><td>" & objRS("Type") & "</td></tr>" & _
				"<tr><td>Sub Type</td><td>" & objRS("SubType") & "</td></tr>" & _
				"<tr><td>EID</td><td>" & objRS("EID") & "</td></tr>" & _
				"<tr><td>Card Type</td><td>" & objRS("CardType") & "</td></tr>" & _
				"<tr><td>Card Number</td><td>" & objRS("CardNumber") & "</td></tr>" & _
				"<tr><td>Actioned By</td><td>" & objRS("ActionedBy") & "</td></tr>" & _
				"<tr><td>Value Before</td><td>" & objRS("ValueBefore") & "</td></tr>" & _
				"<tr><td>Value After</td><td>" & objRS("ValueAfter") & "</td></tr>" & _
				"<tr><td>Source File</td><td>" & objRS("SourceFile") & "</td></tr>" & _
				"<tr><td>Change Details</td><td>" & objRS("ChangeDetails") & "</td></tr>" & _
				"<tr><td>Card ID</td><td>" & objRS("CardID") & "</td></tr>" & _
				"<tr><td>Application ID</td><td>" & objRS("ApplicationID") & "</td></tr>" & _
				"<tr><td>CSFromDiners ID</td><td>" & objRS("CSFromDinersID2") & "</td></tr>" & _
				"<tr><td>CSToDiners ID</td><td>" & objRS("CSToDinersID") & "</td></tr>" & _
				"<tr><td>Updated By</td><td>" & objRS("UpdatedBy") & "</td></tr>" & _
				"<tr><td>Process</td><td>" & objRS("Process") & "</td></tr>" 
			
			
			'Response.Write "<table class=""table table-compact"">" & _
			'	"<tr><td>CS ID</td><td><input class=""ModText"" type=""text"" id=""CSID"" name=""CSID"" class=""form-control input-md"" value=""" & objRS("CSFromDinersID") & """></td>" & _
			'	"<td>Audit Log ID</td><td><input class=""ModTextAudit"" type=""text"" id=""AuditLogID"" name=""AuditLogID"" class=""form-control input-md"" value=""" & objRS("AuditLogID") & """></td>" & _
			'	"<tr><td>Status</label></td></tr><div class=""form-row col-md-2""><input class=""ModText"" type=""text"" id=""Status"" name=""Status"" class=""form-control input-md"" value=""" & objRS("Status") & """>" & strStatus & "</td>" & _
			'	"<td>File Date Time</td><td><input class=""ModText"" type=""text"" id=""FileDateTime"" name=""FileDateTime"" class=""form-control input-md"" value=""" & objRS("FileDateTime") & """></td>" & _
			'	"<tr><td>File Seq Num</td><td><input class=""ModText"" type=""text"" id=""FileSeqNum"" name=""FileSeqNum"" class=""form-control input-md"" value=""" & objRS("FileSeqNum") & """></td>" & _
			'	"<td>Sub Type</td><td><input class=""ModTextAudit"" type=""text"" id=""SubType"" name=""SubType"" class=""form-control input-md"" value=""" & objRS("SubType") & """></td>" & _
			'	"<tr><td>EID No</td><td><input class=""ModText"" type=""text"" id=""EIDNo"" name=""EIDNo"" class=""form-control input-md"" value=""" & objRS("EIDNo") & """></td>" & _
			'	"<td>EID</td><td><input class=""ModTextAudit"" type=""text"" id=""EID"" name=""EID"" class=""form-control input-md"" value=""" & objRS("EID") & """></td>" & _
			'	"<tr><td>Cardholder</td><td><input class=""ModText"" type=""text"" id=""Cardholder"" name=""Cardholder"" class=""form-control input-md"" value=""" & strName & """></td>" & _
			'	"<td>Card Type</td><td><input class=""ModTextAudit"" type=""text"" id=""CardType"" name=""CardType"" class=""form-control input-md"" value=""" & objRS("CardType") & """></td>" & _
			'	"<tr><td>Card Number</td><td><input class=""ModText"" type=""text"" id=""CardNo"" name=""CardNo"" class=""form-control input-md"" value=""" & MaskCard(objRS("CardNo")) & """></td>" & _
			'	"<td>Card Number</td><td><input class=""ModTextAudit"" type=""text"" id=""CardNumber"" name=""CardNumber"" class=""form-control input-md"" value=""" & MaskCard(objRS("CardNumber")) & """></td>" & _
			'	"<tr><td>CardUpdateInd</td><td><input class=""ModText"" type=""text"" id=""CardUpdateInd"" name=""CardUpdateInd"" class=""form-control input-md"" value=""" & objRS("CardUpdateInd") & """></td>" & _
			'	"<td>Actioned By</td><td><input class=""ModTextAudit"" type=""text"" id=""ActionedBy"" name=""ActionedBy"" class=""form-control input-md"" value=""" & objRS("ActionedBy") & """></td>" & _
			'	"<tr><td>Card Expiry Date</td><td><input class=""ModText"" type=""text"" id=""CardExpiryDate"" name=""CardExpiryDate"" class=""form-control input-md"" value=""" & objRS("CardExpiryDate") & """></td>" & _
			'	"<td>Value Before</td><td><input class=""ModTextAudit"" type=""text"" id=""ValueBefore"" name=""ValueBefore"" class=""form-control input-md"" value=""" & objRS("ValueBefore") & """></td>" & _
			'	"<tr><td>Name On Card</td><td><input class=""ModText"" type=""text"" id=""NameOnCard"" name=""NameOnCard"" class=""form-control input-md"" value=""" & Trim(objRS("NameOnCard")) & """></td>" & _
			'	"<td>Value After</td><td><input class=""ModTextAudit"" type=""text"" id=""ValueAfter"" name=""ValueAfter"" class=""form-control input-md"" value=""" & objRS("ValueAfter") & """></td>" & _ 
			'	"<tr><td>Address 1</td><td><input class=""ModText"" type=""text"" id=""Address1"" name=""Address1"" class=""form-control input-md"" value=""" & Trim(objRS("Address1")) & """></td>" & _
			'	"<td>Source File</td><td><input class=""ModTextAudit"" type=""text"" id=""SourceFile"" name=""SourceFile"" class=""form-control input-md"" value=""" & objRS("SourceFile") & """></td>" & _
			'	"<tr><td>Address 2</td><td><input class=""ModText"" type=""text"" id=""Address2"" name=""Address2"" class=""form-control input-md"" value=""" & Trim(objRS("Address2")) & """></td>" & _
			'	"<td>Change Details</td><td><input class=""ModTextAudit"" type=""text"" id=""ChangeDetails"" name=""ChangeDetails"" class=""form-control input-md"" value=""" & objRS("ChangeDetails") & """></td>" & _
			'	"<tr><td>Address 3</td><td><input class=""ModText"" type=""text"" id=""Address3"" name=""Address3"" class=""form-control input-md"" value=""" & objRS("Address3") & """></td>" & _
			'	"<td>Card ID</td><td><input class=""ModTextAudit"" type=""text"" id=""CardID"" name=""CardID"" class=""form-control input-md"" value=""" & objRS("CardID") & """></td>" & _
			'	"<tr><td>Suburb/ State/ PostCode</td><td><input class=""ModText"" type=""text"" id=""SubStatPos"" name=""SubStatPos"" class=""form-control input-md"" value=""" & strSubStatPos & """></td>" & _
			'	"<td>Application ID</td><td><input class=""ModTextAudit"" type=""text"" id=""ApplicationID"" name=""ApplicationID"" class=""form-control input-md"" value=""" & objRS("ApplicationID") & """></td>" & _
			'	"<tr><td>Home Phone</td><td><input class=""ModText"" type=""text"" id=""HomePhone"" name=""HomePhone"" class=""form-control input-md"" value=""" & objRS("HomePhone") & """></td>" & _
			'	"<td>CSFromDinersID</td><td><input class=""ModTextAudit"" type=""text"" id=""CSFromDinersID2"" name=""CSFromDinersID2"" class=""form-control input-md"" value=""" & objRS("CSFromDinersID2") & """></td>" & _
			'	"<tr><td>Work Phone</td><td><input class=""ModText"" type=""text"" id=""WorkPhone"" name=""WorkPhone"" class=""form-control input-md"" value=""" & objRS("WorkPhone") & """></td>" & _
			'	"<td>CSToDinersID</td><td><input class=""ModTextAudit"" type=""text"" id=""CSToDinersID"" name=""CSToDinersID"" class=""form-control input-md"" value=""" & objRS("CSToDinersID") & """></td>" & _
			'	"<tr><td>Mobile Phone</td><td><input class=""ModText"" type=""text"" id=""MobilePhone"" name=""MobilePhone"" class=""form-control input-md"" value=""" & objRS("MobilePhone") & """></td>" & _
			'	"<td>UpdatedBy</td><td><input class=""ModTextAudit"" type=""text"" id=""UpdatedBy"" name=""UpdatedBy"" class=""form-control input-md"" value=""" & objRS("UpdatedBy") & """></td>" & _
			'	"<tr><td>Email</td><td><input class=""ModText"" type=""text"" id=""Email"" name=""Email"" class=""form-control input-md"" value=""" & Trim(objRS("Email")) & """></td>" & _
			'	"<td>Process</td><td><input class=""ModTextAudit"" type=""text"" id=""Process"" name=""Process"" class=""form-control input-md"" value=""" & objRS("Process") & """></td>" & _
			'	"<tr><td>Report Group</td><td><input class=""ModText"" type=""text"" id=""ReportGroup"" name=""ReportGroup"" class=""form-control input-md"" value=""" & objRS("ReportGroup") & """></td>" & _
			'	"<td></td><td></td>" & _
			'	"<tr><td>Credit Limit</td><td><input class=""ModText"" type=""text"" id=""CreditLimit"" name=""CreditLimit"" class=""form-control input-md"" value=""" & objRS("CreditLimit") & """></td>" & _
			'	"<td></td><td></td>" & _
			'	"<tr><td>Relationship</td><td><input class=""ModText"" type=""text"" id=""Relationship"" name=""Relationship"" class=""form-control input-md"" value=""" & objRS("Relationship") & """></td>" & _
			'	"<td></td><td></td>" & _
			'	"<tr><td>Cat2</td><td><input class=""ModText"" type=""text"" id=""Cat2"" name=""Cat2"" class=""form-control input-md"" value=""" & objRS("Cat2") & """></td>" & _
			'	"<td></td><td></td>" & _
			'	"<tr><td>Account Number</td><td><input class=""ModText"" type=""text"" id=""AccountNumber"" name=""AccountNumber"" class=""form-control input-md"" value=""" & objRS("AccountNumber") & """></td>" & _
			'	"<td></td><td></td>" & _
			'	"<tr><td>ActivationFlag</td><td><input class=""ModText"" type=""text"" id=""ActivationFlag"" name=""ActivationFlag"" class=""form-control input-md"" value=""" & objRS("ActivationFlag") & """></td>" & _
			'	"<td></td><td></td>" & _
			'	"<tr><td>PlasticID</td><td><input class=""ModText"" type=""text"" id=""PlasticID"" name=""PlasticID"" class=""form-control input-md"" value=""" & objRS("PlasticID") & """></td>" & _
			'	"<td></td><td></td>" & _
			'	"<tr><td>Companion</td><td><input class=""ModText"" type=""text"" id=""Companion"" name=""Companion"" class=""form-control input-md"" value=""" & MaskCard(objRS("Companion")) & """></td>" & _
			'	"<td></td><td></td>" & _
			'	"<tr><td>Status</td><td><input class=""ModText"" type=""text"" id=""Status"" name=""Status"" class=""form-control input-md"" value=""" & objRS("Status") & """></td>" & _
			'	"<td></td><td></td></tr>"
			
		Else
			Response.write "No Record for " & Request.QueryString("CSFromDinersID")
	   End If

	objRS.Close
	
	Response.write "</table>"
	
	
End Sub
%>
<head>
  
</head>

<body>
  <%
  Call ShowDetails()
  
  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>