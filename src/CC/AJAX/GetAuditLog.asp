<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strAuditLogID
Dim strStatus
Dim strName
Dim strSubStatPos

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("AuditLogID")) Then
	strAuditLogID = " WHERE AuditLogID = " & Request.QueryString("AuditLogID") & ""
End If


Public Sub ShowDetails()

	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM tblCAPSAuditLog WITH(NOLOCK) " & strAuditLogID,objCon

			  
		If Not objRS.EOF Then
			
			'Write the Header
			Response.write "<div class=""row""><div class=""col-md-12""><div class=""row mb-3""><div class=""col-md-6""><h6 class=""mb-3"">Audit Log Details</h6></div>" & _
				"<div class=""col-md-6""><select class=""form-control""><option>Record 1</option><option>Record 2</option><option>Record 3</option></select></div></div><table class=""table table-compact"">"
					

			Response.Write "<tr><td style=""Font-weight:bold;"">Audit Log ID</td><td>" & objRS("AuditLogID") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Change Date</td><td>" & objRS("ChangeDate") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Change Type</td><td>" & objRS("Type") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Sub Type</td><td>" & objRS("SubType") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">EID</td><td>" & objRS("EID") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Card Type</td><td>" & objRS("CardType") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Card Number</td><td>" & objRS("CardNumber") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Actioned By</td><td>" & objRS("ActionedBy") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Value Before</td><td>" & objRS("ValueBefore") & "</td></tr>" & _
				"<tr class=""updated""><td style=""Font-weight:bold;"">Value After</td><td>" & objRS("ValueAfter") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Source File</td><td>" & objRS("SourceFile") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Change Details</td><td>" & objRS("ChangeDetails") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Card ID</td><td>" & objRS("CardID") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Application ID</td><td>" & objRS("ApplicationID") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">CSFromDiners ID</td><td>" & objRS("CSFromDinersID") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">CSToDiners ID</td><td>" & objRS("CSToDinersID") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Updated By</td><td>" & objRS("UpdatedBy") & "</td></tr>" & _
				"<tr><td style=""Font-weight:bold;"">Process</td><td>" & objRS("Process") & "</td></tr>" 
			
		Else
			Response.write "No Record for " & Request.QueryString("AuditLogID")
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