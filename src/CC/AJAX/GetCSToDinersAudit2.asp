<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

'Response.Expires = -1500
'response.buffer = false

Dim objCon
Dim objRS
Dim x
Dim strCSFromDinersID
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


'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")

    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

'Set the Where clause to a value so it doesn;t try to return the full table and time out if there is no cardno
strCSFromDinersID = "WHERE 1=2"
'The detail below has been moved to within the Procedure below to get the type first (as the procedure is called twice, once for archive and once for new audits)
'If Not IsEmpty(Request.QueryString("EmployeeID")) Then
'	If strType = "Archive" Then
'		strCSFromDinersID = " WHERE EID_No = '" & Request.QueryString("EmployeeID") & "' ORDER By Card_No,CSToDinersID DESC"
'		'strCSFromDinersID = " WHERE EID_No = '" & Request.QueryString("EmployeeID") & "' ORDER By CSToDinersID DESC"
'	Else
'		strCSFromDinersID = " WHERE EIDNo = '" & Request.QueryString("EmployeeID") & "' ORDER By CardNo,CSToDinersID DESC"
'	End If
	
'End If

If Not IsEmpty(Request.QueryString("CardNo")) Then
	'strCSFromDinersID = strCSFromDinersID & " AND Card_No = '" & Request.QueryString("CardNo") & "' ORDER By CSToDinersID DESC"
	
	'response.write " !" & Request.QueryString("CardNo") & "! "
	'Remove the leading number (1) which has been added in the screen calling this function (CardDetail.asp) as this gets changed to a number in Javascript otherwise
	'strCardNo = Right(Request.QueryString("CardNo"),Len(Request.QueryString("CardNo"))-1)
	strCardNo = Request.QueryString("CardNo")
	'response.write " !" & strCardNo & "! "
End If


Public Sub ShowDetails(strType)


If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	If strType = "Archive" Then
		strCSFromDinersID = " WHERE Card_No = '" & strCardNo & "' ORDER By Card_No,CSToDinersID DESC"
		'strCSFromDinersID = " WHERE Card_No = '" & Request.QueryString("CardNo") & "' ORDER By Card_No,CSToDinersID DESC"
		'strCSFromDinersID = " WHERE EID_No = '" & Request.QueryString("EmployeeID") & "' ORDER By CSToDinersID DESC"
		
	Else
		strCSFromDinersID = " WHERE CardNo = '" & strCardNo & "' ORDER By CardNo,CSToDinersID DESC"
		'strCSFromDinersID = " WHERE CardNo = '" & Request.QueryString("CardNo") & "' ORDER By CardNo,CSToDinersID DESC"
	End If
	
End If


	'Description:	Loads CS To Diners details onto the Card Audit Modal by EmployeeId NOt Card.
	'objRS.Open "SELECT * FROM qryCAPSCSFromDinersAuditLog " & strCSFromDinersID,objCon
	If strType = "Archive" Then
		objRS.Open "SELECT top 100 * FROM tblCAPSCSToDiners_Archive2 WITH(NOLOCK) " & strCSFromDinersID,objCon
		'Response.Write "SELECT * FROM tblCAPSCSToDiners_Archive2 WITH(NOLOCK) " & strCSFromDinersID
		
		strSource = " - Old CAPS"
	Else
		objRS.Open "SELECT top 100 * FROM tblCAPSCSToDiners WITH(NOLOCK) " & strCSFromDinersID,objCon
		'Response.Write "SELECT * FROM tblCAPSCSToDiners WITH(NOLOCK) " & strCSFromDinersID
		
		strSource = " - New CAPS"
	End If
			  
		If Not objRS.EOF Then
			

			Response.Write "<table class=""table table-compact""><b>CS To Diners " & strSource & "</b>" & _
				"<tr><th style=""font-size:12px;"">CS ID</th><th  style=""font-size:12px;"">Status</th><th  style=""font-size:12px;"">File Seq Num</th><th  style=""font-size:12px;"">File Date Time</th><th  style=""font-size:12px;"">EID No</th><th  style=""font-size:12px;"">Cardholder</th><th  style=""font-size:12px;"">Card Number</th><th  style=""font-size:12px;"">Card Status</th><th  style=""font-size:12px;"">Card Update Ind</th>" & _
				"<th style=""font-size:12px;"">Card Expiry Date</th><th  style=""font-size:12px;"">Name On Card</th><th  style=""font-size:12px;"">Address 1</th><th  style=""font-size:12px;"">Address 2</th><th  style=""font-size:12px;"">Address 3</th><th  style=""font-size:12px;"">Suburb/ State/ PostCode</th>" & _
				"<th style=""font-size:12px;"">Home Phone</th><th  style=""font-size:12px;"">Work Phone</th><th  style=""font-size:12px;"">Mobile Phone</th><th  style=""font-size:12px;"">Email</th><th  style=""font-size:12px;"">Report Group</th><th  style=""font-size:12px;"">Credit Limit</th>"  & _
				"<th style=""font-size:12px;"">Status</th>"
		
		Else
			Response.write "No CS TO Diners (" & strType & ") for " & Request.QueryString("EmployeeID") & " Card: " & strCSFromDinersID & "</br>"'Request.QueryString("CardNo")'Request.QueryString("CSFromDinersID")
	   End If

	   
		Do Until objRS.EOF
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
		
			'Create the Status list badge based on the status field - Card Status
			If strType = "Archive" Then
				If IsNull(objRS("Card_Status")) Then
					strCardStatus = ""
				Else
					strCardStatus = objRS("Card_Status")
				End If
			
			Else
				
				If IsNull(objRS("CardStatus")) Then
					strCardStatus = ""
				Else
					strCardStatus = objRS("CardStatus")
				End If
			End If
			
			If IsNull(strCardStatus) Then
				strCardStatus = ""
			Else
				If strCardStatus = "00" Then
					strCardStatus = "<span class=""badge badge-pill badge-success"">00</span>"
				'ElseIf objRS("CardStatus") = "Imported" Then
				'	strCardStatus = "<span class=""badge badge-pill badge-warning"">Imported</span>"
				Else
					strCardStatus = "<span class=""badge badge-pill badge-danger"">" & strCardStatus & "</span>"
				End If
			End If
			
			If IsNull(objRS("Surname")) Then
				strName = ""
			Else
				If strType = "Archive" Then
					strName = Trim(Trim(objRS("Title")) & " " & Trim(objRS("Given_Names")) & " " & Trim(objRS("Surname")))
				Else
					strName = Trim(Trim(objRS("Title")) & " " & Trim(objRS("GivenNames")) & " " & Trim(objRS("Surname")))
				End If
			End If
			
			If IsNull(objRS("Suburb")) Then
				strSubStatPos = ""
			Else
				strSubStatPos = Trim(Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode")))
			End If
			
			If strType = "Archive" Then
			Response.Write "<tr><td style=""font-size:12px;"">" & objRS("CSToDinersID") & "</td><td style=""font-size:12px;"">" & strStatus & "</td><td style=""font-size:12px;"">" & objRS("File_Seq_Num") & "</td><td style=""font-size:12px;"">" & objRS("File_Date_Time") & "</td><td style=""font-size:12px;"">" & objRS("EID_No") & "</td><td style=""font-size:12px;"">" & strName & "</td><td>" & MaskCard(objRS("Card_No")) & "</td>" & _
				"<td style=""font-size:12px;"">" & strCardStatus & "</td><td style=""font-size:12px;"">" & objRS("Card_Update_Ind") & "</td><td style=""font-size:12px;"">" & objRS("Card_Expiry_Date") & "</td><td style=""font-size:12px;"">" & Trim(objRS("Name_On_Card")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address_1")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address_2")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address_3")) & "</td>" & _
				"<td style=""font-size:12px;"">" & strSubStatPos & "</td><td style=""font-size:12px;"">" & objRS("Home_Phone") & "</td><td style=""font-size:12px;"">" & objRS("Work_Phone") & "</td><td style=""font-size:12px;"">" & objRS("Mobile_Phone") & "</td><td style=""font-size:12px;"">" & Trim(objRS("Email")) & "</td><td style=""font-size:12px;"">" & objRS("Report_Group") & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("Credit_Limit") & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("Status") & "</td></tr>"
			
			
			Else
				Response.Write "<tr><td style=""font-size:12px;"">" & objRS("CSToDinersID") & "</td><td style=""font-size:12px;"">" & strStatus & "</td><td style=""font-size:12px;"">" & objRS("FileSeqNum") & "</td><td style=""font-size:12px;"">" & objRS("FileDateTime") & "</td><td style=""font-size:12px;"">" & objRS("EIDNo") & "</td><td style=""font-size:12px;"">" & strName & "</td><td>" & MaskCard(objRS("CardNo")) & "</td>" & _
				"<td style=""font-size:12px;"">" & strCardStatus & "</td><td style=""font-size:12px;"">" & objRS("CardUpdateInd") & "</td><td style=""font-size:12px;"">" & objRS("CardExpiryDate") & "</td><td style=""font-size:12px;"">" & Trim(objRS("NameOnCard")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address1")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address2")) & "</td><td style=""font-size:12px;"">" & Trim(objRS("Address3")) & "</td>" & _
				"<td style=""font-size:12px;"">" & strSubStatPos & "</td><td style=""font-size:12px;"">" & objRS("HomePhone") & "</td><td style=""font-size:12px;"">" & objRS("WorkPhone") & "</td><td style=""font-size:12px;"">" & objRS("MobilePhone") & "</td><td style=""font-size:12px;"">" & Trim(objRS("Email")) & "</td><td style=""font-size:12px;"">" & objRS("ReportGroup") & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("CreditLimit") & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("Status") & "</td></tr>"
			End If
			
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
  
  'Call the same procedure for New CS To Diners records
  Call ShowDetails("New")
  
	'Call the procedure to load details onto the screen from the Archive table then the new table
  Call ShowDetails("Archive")
  
  
  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>