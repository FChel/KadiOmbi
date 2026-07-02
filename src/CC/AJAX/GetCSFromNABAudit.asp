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
	
	strSQL = "SELECT TOP 50 * FROM tblCAPSNABCS WITH(NOLOCK) WHERE CardNumber = '" & strCardNo & "' ORDER BY CAPSNABCSID DESC"
	
	'Description:	Loads CS To Diners details onto the Card Audit Modal by EmployeeId NOt Card.
	objRS.Open strSQL,objCon
			  
		If Not objRS.EOF Then

			Response.Write "<table class=""table table-responsive""><b>CS from NAB</b><tr>" & _
			"<th style=""font-size:12px;"">Date</th>" & _
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
			"</tr>"
		Else
			Response.write "No CS from NAB for " & Request.QueryString("EmployeeID") & " Card: " & MaskCard(strCardNo) & "</br>"
	   End If
	   
		Do Until objRS.EOF
					
		If IsNull(objRS("DateUpdated")) OR IsEmpty(objRS("DateUpdated")) Then
			strDateUpdate = ""
		Else
			strDateUpdate = Left(objRS("DateUpdated"),10)
		End If
		
		If objRS("CardStatus") = "  " Then
			strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
		Else 
			strStatus = "<span class=""badge badge-pill badge-danger"">" & objRS("CardStatus") & "</span>"					
		End If
		'response.write objRS("filedatetime")
			Response.Write "<tr><td style=""font-size:12px;""><a href=""/Admin/CSTransactionsNAB.asp?FileLoadID=" & objRS("FileDateTime") & """>" & strDateUpdate & "</a></td>" & _
				"<td style=""font-size:12px;"">" & MaskCard(objRS("CardNumber")) & "</td>" & _
				"<td style=""font-size:12px;"">" & strStatus & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("CardExpiry") & "</td>" & _
				"<td style=""font-size:12px;"">" & Trim(objRS("EmbossingName")) & "</td>" & _
				"<td style=""font-size:12px;"">" & Trim(objRS("Address1")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address2")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address3")) & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("Address4") & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("WorkPhone") & "</td><td style=""font-size:12px;"">" & objRS("MobilePhone") & "</td>" & _
				"<td style=""font-size:12px;"">" & Trim(objRS("EmailAddress")) & "</td>"  & _
				"<td style=""font-size:12px;"">" & FormatCurrency(objRS("CardCreditLimit"),0) & "</td><td style=""font-size:12px;"">" & Trim(objRS("TransLimitCode")) & "</td>" & _
				"</tr>"
			objRS.Movenext
			
		Loop
		
		
	objRS.Close

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