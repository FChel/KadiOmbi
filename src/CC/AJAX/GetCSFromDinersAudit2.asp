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
Dim strSource
Dim strCardNo
Dim strCardStatus

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")

    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")


If Not IsEmpty(Request.QueryString("CardNo")) Then
	'strCSFromDinersID = strCSFromDinersID & " AND CardNo = '" & Request.QueryString("CardNo") & "' ORDER By CSFromDinersID DESC"
	
	'response.write "! qs=" & Request.QueryString("CardNo") & "! "
	
	'Remove the leading number (1) which has been added in the screen calling this function (CardDetail.asp) as this gets changed to a number in Javascript otherwise
	'strCardNo = Right(Request.QueryString("CardNo"),Len(Request.QueryString("CardNo"))-1)
	strCardNo = Request.QueryString("CardNo")
End If

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	strCSFromDinersID = " WHERE CardNo = '" & strCardNo & "' ORDER By CardNo,CSFromDinersID DESC"
	'strCSFromDinersID = " WHERE Cast(CardNo as bigint) = " & Request.QueryString("CardNo") & " ORDER By CardNo,CSFromDinersID DESC"
	'strCSFromDinersID = " WHERE EIDNo = '" & Request.QueryString("EmployeeID") & "' ORDER By CSFromDinersID DESC"
End If


Public Sub ShowDetails(strType)

Dim strActivationFlag

	'Description:	Loads CS To Diners details onto the Card Audit Modal by EmployeeId NOt Card.
	'objRS.Open "SELECT * FROM qryCAPSCSFromDinersAuditLog " & strCSFromDinersID,objCon
	If strType = "Archive" Then
		'objRS.Open "SELECT * FROM qryCAPSCSFromDinersAuditLog " & strCSFromDinersID,objCon
		objRS.Open "SELECT * FROM tblCAPSCSFromDiners_Archive WITH(NOLOCK) " & strCSFromDinersID,objCon
		'Response.Write " SELECT * FROM tblCAPSCSFromDiners_Archive WITH(NOLOCK) " & strCSFromDinersID
		
		strSource = " - Old CAPS"
	Else
		objRS.Open "SELECT * FROM tblCAPSCSFromDiners WITH(NOLOCK) " & strCSFromDinersID,objCon
		'Response.Write " SELECT * FROM tblCAPSCSFromDiners WITH(NOLOCK) " & strCSFromDinersID
		
		strSource = " - New CAPS"
	End If
			  
		If Not objRS.EOF Then
			
			Response.Write "<table class=""table table-compact""><b>CS From Diners " & strSource & "</b>" & _
				"<tr><th style=""font-size:12px;"">CS ID</th><th  style=""font-size:12px;"">Status</th><th  style=""font-size:12px;"">File Seq Num</th><th  style=""font-size:12px;"">File Date Time</th><th  style=""font-size:12px;"">EID No</th><th  style=""font-size:12px;"">Cardholder</th><th  style=""font-size:12px;"">Card Number</th><th  style=""font-size:12px;"">Card Status</th><th  style=""font-size:12px;"">CardUpdateInd</th>" & _
				"<th style=""font-size:12px;"">Activation Flag</th><th style=""font-size:12px;"">Card Expiry Date</th><th  style=""font-size:12px;"">Name On Card</th><th  style=""font-size:12px;"">Address 1</th><th  style=""font-size:12px;"">Address 2</th><th  style=""font-size:12px;"">Address 3</th><th  style=""font-size:12px;"">Suburb/ State/ PostCode</th>" & _
				"<th style=""font-size:12px;"">Home Phone</th><th  style=""font-size:12px;"">Work Phone</th><th  style=""font-size:12px;"">Mobile Phone</th><th  style=""font-size:12px;"">Email</th><th  style=""font-size:12px;"">Report Group</th><th  style=""font-size:12px;"">Credit Limit</th>"  & _
				"<th style=""font-size:12px;"">Status</th>"
		
		Else
			Response.write " No CS FROM Diners for " & Request.QueryString("EmployeeID") & " Card: " & strCardNo & "</br>"'Request.QueryString("CardNo")'Request.QueryString("CSFromDinersID")
	   End If

	   
		Do Until objRS.EOF
			'Create the Status list badge based on the status field - Process Status
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
		
			'Create the Status list badge based on the status field - Card Status
			If IsNull(objRS("CardStatus")) Then
				strCardStatus = ""
			Else
				If objRS("CardStatus") = "00" Then
					strCardStatus = "<span class=""badge badge-pill badge-success"">00</span>"
				'ElseIf objRS("CardStatus") = "Imported" Then
				'	strCardStatus = "<span class=""badge badge-pill badge-warning"">Imported</span>"
				Else
					strCardStatus = "<span class=""badge badge-pill badge-danger"">" & objRS("CardStatus") & "</span>"
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
			
			If IsNull(objRS("ActivationFlag")) Then
				strActivationFlag = ""
			Else
				strActivationFlag = Trim(objRS("ActivationFlag"))
					If strActivationFlag = "Y" Then
						strActivationFlag = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">Y</span>"
					Else
						strActivationFlag = "<span class=""badge badge-pill badge-secondary"" style=""font-size:12px;"">" & strActivationFlag & "</span>"
					End If
			End If
			
			Response.Write "<tr><td style=""font-size:12px;"">" & objRS("CSFromDinersID") & "</td><td style=""font-size:12px;"">" & strStatus & "</td><td style=""font-size:12px;"">" & objRS("FileSeqNum") & "</td><td style=""font-size:12px;"">" & objRS("FileDateTime") & "</td><td style=""font-size:12px;"">" & objRS("EIDNo") & "</td><td style=""font-size:12px;"">" & strName & "</td><td>" & MaskCard(objRS("CardNo")) & "</td>" & _
				"<td style=""font-size:12px;"">" & strCardStatus & "</td><td style=""font-size:12px;"">" & objRS("CardUpdateInd") & "</td><td style=""font-size:12px;"">" & strActivationFlag & "</td><td style=""font-size:12px;"">" & objRS("CardExpiryDate") & "</td><td style=""font-size:12px;"">" & Trim(objRS("NameOnCard")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address1")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address2")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address3")) & "</td>" & _
				"<td style=""font-size:12px;"">" & strSubStatPos & "</td><td style=""font-size:12px;"">" & objRS("HomePhone") & "</td><td style=""font-size:12px;"">" & objRS("WorkPhone") & "</td><td style=""font-size:12px;"">" & objRS("MobilePhone") & "</td><td style=""font-size:12px;"">" & Trim(objRS("Email")) & "</td><td style=""font-size:12px;"">" & objRS("ReportGroup") & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("CreditLimit") & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("Status") & "</td></tr>"
			
			
			objRS.Movenext
			
		Loop
		
		
	objRS.Close
	
	Response.write "</table>"
	
	
End Sub
%>
<head>
  
</head>

<body>
  <%
   'Call the same procedure to get Audit logs from the new tables
  Call ShowDetails("New")
  
  'Call the same procedure to get Audit logs from the new tables
  Call ShowDetails("Archive")

  
  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>