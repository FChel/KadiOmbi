<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

'Response.Expires = -1500
'response.buffer = false

Dim objCon
Dim objRS
Dim x
Dim strNABCMID
Dim strStatus
Dim strName
Dim strSubStatPos

Dim strFileSeqNum
Dim strEIDNo
Dim strFileDateTime
Dim strCardNo
Dim strCardUpdateInd
Dim strCardExpiryDate
Dim strNameOnCard
Dim strSource
Dim strSQL
Dim strDateUpdate


'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")

    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

'Set the Where clause to a value so it doesn;t try to return the full table and time out if there is no cardno
'strNABCMID = "WHERE 1=2"
'The detail below has been moved to within the Procedure below to get the type first (as the procedure is called twice, once for archive and once for new audits)

If Not IsEmpty(Request.QueryString("CardNo")) Then
	strCardNo = Request.QueryString("CardNo")
End If

Public Sub ShowDetails()

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
		strEIDNo = Request.QueryString("EmployeeID")
		'strNABCMID = " WHERE CardNo = '" & strCardNo & "' ORDER By CAPSNABCMID DESC"
End If
	
	strSQL = "SELECT TOP 100 * FROM tblCAPSNABCM WITH(NOLOCK) WHERE CardNumber = '" & strCardNo & "' ORDER BY CAPSNABCMID DESC"
	
	'Description:	Loads CS To Diners details onto the Card Audit Modal by EmployeeId NOt Card.
	objRS.Open strSQL,objCon
		'Response.Write strSQL

		If Not objRS.EOF Then
			'Response.write "Yes"
			Response.Write "<table class=""table table-responsive""><b>CM to NAB</b><tr>" & _
			"<th style=""font-size:12px;"">Change</th>" & _
			"<th  style=""font-size:12px;"">Card</th>" & _
			"<th  style=""font-size:12px;"">Card Status</th>" & _
			"<th style=""font-size:12px;"">Expiry</th>" & _
			"<th  style=""font-size:12px;"">Name On Card</th>" & _
			"<th  style=""font-size:12px;"">Address 1</th>" & _
			"<th  style=""font-size:12px;"">Address 2</th>" & _
			"<th  style=""font-size:12px;"">Address 3</th>" & _
			"<th  style=""font-size:12px;"">Address 4</th>" & _
			"<th  style=""font-size:12px;"">Work Phone</th>" & _
			"<th  style=""font-size:12px;"">Mobile Phone</th>" & _
			"<th  style=""font-size:12px;"">Email</th>" & _
			"<th  style=""font-size:12px;"">Credit Limit</th>"  & _
			"<th style=""font-size:12px;"">Transaction Limit Code</th>" & _
			"<th style=""font-size:12px;"">Date Updated</th>" & _
			"</tr>"
		Else
			Response.write "No CM to NAB for " & Request.QueryString("EmployeeID") & " Card: " & MaskCard(strCardNo) & "</br>"
	   End If
	   
		Do Until objRS.EOF
		
		If IsNull(objRS("DateUpdated")) OR IsEmpty(objRS("DateUpdated")) Then
			strDateUpdate = ""
		Else
			strDateUpdate = Left(objRS("DateUpdated"),10)
		End If
		
		If objRS("CardStatus") = "" OR IsNull(objRS("CardStatus")) Then
			strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
		Else 
			strStatus = "<span class=""badge badge-pill badge-danger"">" & objRS("CardStatus") & "</span>"					
		End If
		
		If IsNull(objRS("FileSeqNum")) OR objRS("FileSeqNum") = "" Then
			strFileSeqNum = NULL
		Else	
			strFileSeqNum = objRS("fileSeqNum")
		End If
		
		If IsNull(objRS("Status")) OR objRS("Status") = "" Then
			Response.Write "<tr><td style=""font-size:12px;"" title=""Error, contact Admin""><a href=""/Admin/CSTransactionsToNAB.asp?FileLoadID=" & strFileSeqNum & "&AuditFor=" & strEIDNo & """><span class=""badge badge-pill badge-danger"">" & objRS("ChangeGroup") & "</span></a></td>"
		ElseIf objRS("Status") = "Awaiting Export" Then
			Response.Write "<tr><td style=""font-size:12px;"" title=""Awaiting Export""><a href=""/Admin/CSTransactionsToNAB.asp?FileLoadID=" & strFileSeqNum & "&AuditFor=" & strEIDNo & """><span class=""badge badge-pill badge-warning"">" & objRS("ChangeGroup") & "</span></a></td>"
		ElseIf objRS("Status") = "Exported" Then
			Response.Write "<tr><td style=""font-size:12px;"" title=""Exported""><a href=""/Admin/CSTransactionsToNAB.asp?FileLoadID=" & strFileSeqNum & "&AuditFor=" & strEIDNo & """><span class=""badge badge-pill badge-success"">" & objRS("ChangeGroup") & "</span></a></td>"
		Else
			Response.Write "<tr><td style=""font-size:12px;"" title=""Exported""><a href=""/Admin/CSTransactionsToNAB.asp?FileLoadID=" & strFileSeqNum & "&AuditFor=" & strEIDNo & """>" & objRS("ChangeGroup") & "</a></td>"
		End If

			'Response.Write objRS("Status")
			Response.Write "<td style=""font-size:12px;"">" & MaskCard(objRS("CardNumber")) & "</td>" & _
				"<td style=""font-size:12px;"">" & strStatus & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("CardExpiry") & "</td>" & _
				"<td style=""font-size:12px;"">" & Trim(objRS("EmbossingName")) & "</td>" & _
				"<td style=""font-size:12px;"">" & Trim(objRS("Address1")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address2")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address3")) & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("Address4") & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("WorkPhone") & "</td><td style=""font-size:12px;"">" & objRS("MobilePhone") & "</td>" & _
				"<td style=""font-size:12px;"">" & Trim(objRS("EmailAddress")) & "</td>" & _
				"<td style=""font-size:12px;"">" & Trim(objRS("CardCreditLimit")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("TransLimitCode")) & "</td>" & _
				"<td style=""font-size:12px;"">" & strDateUpdate & "</td>" & _
				"</tr>"

			objRS.Movenext
			
		Loop
		
		
	objRS.Close
	'Response.Write "end of loop"
	Response.Write "</table>"
	
End Sub
%>
<head>
  
</head>

<body>
  <%
  
  'Call the same procedure for New CS To Diners records
	Call ShowDetails()

  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>