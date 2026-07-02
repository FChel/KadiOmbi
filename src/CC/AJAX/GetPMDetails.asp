<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strCardNumber
Dim strStatus
Dim strName
Dim strSubStatPos
Dim lngCardID

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("CardID")) Then
	'strCardNumber = " WHERE CardNumber = '" & Request.QueryString("CardNumber") & "'"
	
	lngCardID = Request.QueryString("CardID")
	
	'Call the function to get the CardNumber from the CardID passed in
	strCardNumber = GetCardNoShort(lngCardID)
	
	strCardNumber = " WHERE CardAccountNumber = '" & strCardNumber & "'"
End If


Public Sub ShowDetails()
'response.write strCardNumber
	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM qryCAPSProMasterAccountsUserDecode " & strCardNumber,objCon

		If Not objRS.EOF Then
			
			'Write the Header
			'Response.write "<div class=""panel-content row""><div class=""form-row col-md-6"" style=""font-weight:bold;"">CS From Diners</div><div class=""form-row col-md-6"" style=""font-weight:bold;""><A HREF=""#"" onClick=""loadCard(" & objRS(0) & ")"">Audit Log Details</a> </div></div>"
			
			'Response.write "<div class=""row""><div class=""col-md-6""><h6 class=""mb-3"">CS From Diners</h6><table class=""table table-compact"">"

			Response.write "<div class=""row""><div class=""col-md-6""><div class=""row mb-3""><div class=""col-md-6""><h6 class=""mb-3"">Card Details in ProMaster</h6></div>" & _
				"</div><table class=""table table-compact"">"
				
			'Create the Status list badge based on the status field
			If IsNull(objRS("card_status")) Then
				strStatus = ""
			Else
				If objRS("card_status") = "Processed" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Processed</span>"
				ElseIf objRS("card_status") = "Imported" Then
					strStatus = "<span class=""badge badge-pill badge-warning"">Imported</span>"
				Else
					strStatus = objRS("card_status")
				End If
			End If
		
			If IsNull(objRS("Surname")) Then
				strName = ""
			Else
				strName = Trim(Trim(objRS("first_name")) & " " & Trim(objRS("Surname")))
				'strName = Trim(Trim(objRS("Title")) & " " & Trim(objRS("GivenNames")) & " " & Trim(objRS("Surname")))
			End If
			
			'If IsNull(objRS("Suburb")) Then
			'	strSubStatPos = ""
			'Else
			'	strSubStatPos = Trim(Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode")))
			'End If
			
			'Response.write "<tr onClick=""SelectFileSeqNum(String('" & objRS("CSFromDinersID") & "'));""><td>" & objRS("CSFromDinersID") & "</td><td>" & objRS("FileDateTime") & "</td><td>" & objRS("Status") & "</td>" & _
			'			"<td>" & objRS("EIDNo") & "</td><td>" & objRS("Title") & " " & objRS("GivenNames") & " " & objRS("Surname") & "</td></tr>"
			

			Response.Write "<tr><th>Account Ref</th><td>" & objRS("AccountRefNo") & "</td></tr>" & _
				"<tr><th>Status</th><td>" & strStatus & "</td></tr>" & _
				"<tr><th>User</th><td>" & objRS("User_Name") & "</td></tr>" & _
				"<tr><th>EID No</th><td>" & objRS("Employee_id") & "</td></tr>" & _
				"<tr><th>Cardholder</th><td>" & strName & "</td></tr>" & _
				"<tr><th>Card Number</th><td>" & MaskCard(objRS("CardAccountNumber")) & "</td></tr>" & _
				"<tr><th>CardType</th><td>" & objRS("CardType") & "</td></tr>" &  _
				"<tr><th>CardAccountNumber</th><td>" & objRS("CardAccountNumber") & "</td></tr>" &  _
				"<tr><th>card_type</th><td>" & objRS("card_type") & "</td></tr>" &  _
				"<tr><th>user_name</th><td>" & objRS("user_name") & "</td></tr>" &  _
				"<tr><th>unit_id</th><td>" & objRS("unit_id") & "</td></tr>" &  _
				"<tr><th>company</th><td>" & objRS("company") & "</td></tr>" &  _
				"<tr><th>gl_code</th><td>" & objRS("gl_code") & "</td></tr>" &  _
				"<tr><th>cost_ctr</th><td>" & objRS("cost_ctr") & "</td></tr>" &  _
				"<tr><th>internal_order</th><td>" & objRS("internal_order") & "</td></tr>" &  _
				"<tr><th>wbs_element</th><td>" & objRS("wbs_element") & "</td></tr>" &  _
				"<tr><th>post_code</th><td>" & objRS("post_code") & "</td></tr>" &  _
				"<tr><th>account_ref_no</th><td>" & objRS("account_ref_no") & "</td></tr>" &  _
				"<tr><th>create_date</th><td>" & objRS("create_date") & "</td></tr>" &  _
				"<tr><th>created_by</th><td>" & objRS("created_by") & "</td></tr>" &  _
				"<tr><th>issue_date</th><td>" & objRS("issue_date") & "</td></tr>" &  _
				"<tr><th>expiry_date</th><td>" & objRS("expiry_date") & "</td></tr>" &  _
				"<tr><th>card_status</th><td>" & objRS("card_status") & "</td></tr>" &  _
				"<tr><th>Attention</th><td>" & objRS("Attention") & "</td></tr>" &  _
				"<tr><th>addr1</th><td>" & objRS("addr1") & "</td></tr>" &  _
				"<tr><th>addr2</th><td>" & objRS("addr2") & "</td></tr>" &  _
				"<tr><th>addr3</th><td>" & objRS("addr3") & "</td></tr>" &  _
				"<tr><th>addr_postcode</th><td>" & objRS("addr_postcode") & "</td></tr>" &  _
				"<tr><th>addr_state</th><td>" & objRS("addr_state") & "</td></tr>" &  _
				"<tr><th>Auto_Approve</th><td>" & objRS("Auto_Approve") & "</td></tr>" &  _
				"<tr><th>Plastic_Type</th><td>" & objRS("Plastic_Type") & "</td></tr>" &  _
				"<tr><th>monthly spend limit</th><td>" & objRS("monthly spend limit") & "</td></tr>" &  _
				"<tr><th>name on account</th><td>" & objRS("name on account") & "</td></tr>" &  _
				"<tr><th>cardholder eid</th><td>" & objRS("cardholder eid") & "</td></tr>" &  _
				"<tr><th>cardholder email</th><td>" & objRS("cardholder email") & "</td></tr>" &  _
				"<tr><th>report group</th><td>" & objRS("report group") & "</td></tr>" &  _
				"<tr><th>Extract_date</th><td>" & objRS("Extract_date") & "</td></tr>" &  _
				"<tr><th>True_Account_Ref</th><td>" & objRS("True_Account_Ref") & "</td></tr>" &  _
				"<tr><th>ExtractDatePMUser</th><td>" & objRS("ExtractDatePMUser") & "</td></tr>" &  _
				"<tr><th>employee_id</th><td>" & objRS("employee_id") & "</td></tr>" &  _
				"<tr><th>contractor_ind</th><td>" & objRS("contractor_ind") & "</td></tr>" &  _
				"<tr><th>UserNamePMUser</th><td>" & objRS("UserNamePMUser") & "</td></tr>" &  _
				"<tr><th>first_name</th><td>" & objRS("first_name") & "</td></tr>" &  _
				"<tr><th>surname</th><td>" & objRS("surname") & "</td></tr>" &  _
				"<tr><th>location_name</th><td>" & objRS("location_name") & "</td></tr>" &  _
				"<tr><th>admin_ctr</th><td>" & objRS("admin_ctr") & "</td></tr>" &  _
				"<tr><th>admin_ctr_name</th><td>" & objRS("admin_ctr_name") & "</td></tr>" &  _
				"<tr><th>active_indicator</th><td>" & objRS("active_indicator") & "</td></tr>" &  _
				"<tr><th>locked</th><td>" & objRS("locked") & "</td></tr>" &  _
				"<tr><th>inactive_reason</th><td>" & objRS("inactive_reason") & "</td></tr>" &  _
				"<tr><th>admin_centre_controller</th><td>" & objRS("admin_centre_controller") & "</td></tr>" &  _
				"<tr><th>enterprise_controller</th><td>" & objRS("enterprise_controller") & "</td></tr>" &  _
				"<tr><th>email_address</th><td>" & objRS("email_address") & "</td></tr>" &  _
				"<tr><th>Work_Phone</th><td>" & objRS("Work_Phone") & "</td></tr>" &  _
				"<tr><th>Mobile</th><td>" & objRS("Mobile") & "</td></tr>" &  _
				"<tr><th>review_date</th><td>" & objRS("review_date") & "</td></tr>" &  _
				"<tr><th>CreateDatePMUser</th><td>" & objRS("CreateDatePMUser") & "</td></tr>" &  _
				"<tr><th>CreatedByPMUser</th><td>" & objRS("CreatedByPMUser") & "</td></tr>" &  _
				"<tr><th>last_logon</th><td>" & objRS("last_logon") & "</td></tr>" &  _
				"<tr><th>unprocessed_transactions</th><td>" & objRS("unprocessed_transactions") & "</td></tr>" &  _
				"<tr><th>active_cards</th><td>" & objRS("active_cards") & "</td></tr>" &  _
				"<tr><th>Supervisor</th><td>" & objRS("Supervisor") & "</td></tr>"
	
		Else
			Response.write "No Record for " & strCardNumber
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