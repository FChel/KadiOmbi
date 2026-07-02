<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strANZCardlistID
Dim strStatus
Dim strName
Dim strSubStatPos

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("ANZCardlistID")) Then
	strANZCardlistID = " WHERE ANZCardlistID = " & Request.QueryString("ANZCardlistID") & ""
End If


Public Sub ShowDetails()

	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM qryCAPSANZCardlistAuditLog WITH(NOLOCK) " & strANZCardlistID,objCon

			  
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
		
			If IsNull(objRS("NameOnCard")) Then
				strName = ""
			Else
				strName = Trim(objRS("NameOnCard"))
			End If
			
			If IsNull(objRS("Suburb")) Then
				strSubStatPos = ""
			Else
				strSubStatPos = Trim(Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode")))
			End If
			
			Response.Write "<tr><td>ANZ Cardlist ID</td><td>" & objRS("ANZCardlistID") & "</td></tr>" & _
				"<tr><td>Status</td><td>" & strStatus & "</td></tr>" & _
				"<tr><td>File ID</td><td>" & objRS("FileID") & "</td></tr>" & _
				"<tr><td>EID No</td><td>" & objRS("EmployeeID") & "</td></tr>" & _
				"<tr><td>Cardholder</td><td>" & strName & "</td></tr>" & _
				"<tr><td>Card Number</td><td>" & MaskCard(objRS("CardNumber")) & "</td></tr>" & _
				"<tr><td>Card Expiry Date</td><td>" & objRS("Expiry") & "</td></tr>" & _
				"<tr><td>Name On Card</td><td>" & Trim(objRS("NameOnCard")) & "</td></tr>" & _
				"<tr><td>Address 1</td><td>" & Trim(objRS("Address1")) & "</td></tr>" & _
				"<tr><td>Address 2</td><td>" & Trim(objRS("Address2")) & "</td></tr>" & _
				"<tr><td>Address 3</td><td>" & Trim(objRS("Address3")) & "</td></tr>" & _
				"<tr><td>Suburb/ State/ PostCode</td><td>" & strSubStatPos & "</td></tr>" & _
				"<tr><td>Phone</td><td>" & objRS("Phone") & "</td></tr>" & _
				"<tr><td>Credit Limit</td><td>" & objRS("CreditLimit") & "</td></tr>" & _
				"<tr><td>OTC Limit</td><td>" & objRS("OTCLimit") & "</td></tr>" & _
				"<tr><td>Transaction Limit</td><td>" & objRS("TransactionLimit") & "</td></tr>" & _
				"<tr><td>ATM Limit</td><td>" & objRS("ATMLimit") & "</td></tr>" & _
				"<tr><td>Relationship</td><td>" & objRS("Relationship") & "</td></tr>" & _
				"<tr><td>Billing Account</td><td>" & objRS("BillingAccount") & "</td></tr>" & _
				"<tr><td>Status</td><td>" & objRS("Status") & "</td></tr>" 
			
			
			Response.Write "</table></div><div class=""col-md-6""><div class=""row mb-3""><div class=""col-md-6 my-auto""><button type=""button"" class=""btn btn-outline-secondary"" onClick=""loadCard(" & objRS(0) & ")"">Audit Log Details</button></div>" & _
				"<div class=""col-md-6""><select class=""form-control""><option>Record 1</option><option>Record 2</option><option>Record 3</option></select>" & _
				"</div></div><table class=""table table-compact"">"
				

			Response.Write "<tr><td>Audit Log ID</td><td>" & objRS("AuditLogID") & "</td></tr>" & _
				"<tr><td>Change Date</td><td>" & objRS("ChangeDate") & "</td></tr>" & _
				"<tr><td>Change Type</td><td>" & objRS("Type") & "</td></tr>" & _
				"<tr><td>Sub Type</td><td>" & objRS("SubType") & "</td></tr>" & _
				"<tr><td>EID</td><td>" & objRS("EID") & "</td></tr>" & _
				"<tr><td>Card Type</td><td>" & objRS("CardType") & "</td></tr>" & _
				"<tr><td>Card Number</td><td>" & objRS("CardNumber") & "</td></tr>" & _
				"<tr><td>Actioned By</td><td>" & objRS("ActionedBy") & "</td></tr>" & _
				"<tr><td>Value Before</td><td>" & objRS("ValueBefore") & "</td></tr>" & _
				"<tr class=""updated""><td>Value After</td><td>" & objRS("ValueAfter") & "</td></tr>" & _
				"<tr><td>Source File</td><td>" & objRS("SourceFile") & "</td></tr>" & _
				"<tr><td>Change Details</td><td>" & objRS("ChangeDetails") & "</td></tr>" & _
				"<tr><td>Card ID</td><td>" & objRS("CardID") & "</td></tr>" & _
				"<tr><td>Application ID</td><td>" & objRS("ApplicationID") & "</td></tr>" & _
				"<tr><td>CSFromDiners ID</td><td>" & objRS("CSFromDinersID2") & "</td></tr>" & _
				"<tr><td>CSToDiners ID</td><td>" & objRS("CSToDinersID") & "</td></tr>" & _
				"<tr><td>Updated By</td><td>" & objRS("UpdatedBy") & "</td></tr>" & _
				"<tr><td>Process</td><td>" & objRS("Process") & "</td></tr>" 
			
		Else
			Response.write "No Record for " & Request.QueryString("ANZCardlistID")
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