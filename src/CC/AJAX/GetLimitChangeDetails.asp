<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strApplicationID
Dim lngEmployeeID
Dim lngApplicationID

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	lngApplicationID = Request.QueryString("ApplicationID")
End If

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	lngEmployeeID = " WHERE EmployeeID = " & Request.QueryString("EmployeeID") & ""
	lngEmployeeID = Request.QueryString("EmployeeID") 
End If

Public Sub ShowDetails()
'Procedure to write all Credit Limit applications to the screen for the selected employee
Dim strApplication 
Dim x
Dim dteLimitDateFrom
Dim dteLimitDateTo
Dim strDurationofLimit
Dim strApplicationStatus

	'Description:	Loads CDMC Details onto the page called from
	objRS.Open "SELECT * FROM tblCAPSLimitChange WITH(NOLOCK) WHERE [EmployeeID]='" & lngEmployeeID & "' ORDER By LimitChangeID DESC",objCon


		If Not objRS.EOF Then
			
			'Write the Header			
			Response.write "<div class=""row""><div class=""col-md-12"">" & _
				"<table class=""table table-compact"" ><thead>"
			
			'Write the header
			Response.write "<tr><th style=""font-weight:bold;"">Limit Change ID</th><th style=""font-weight:bold;"">Employee ID</th><th style=""font-weight:bold;"">Application ID</th>" & _
						"<th style=""font-weight:bold;"">XML Application ID</th><th style=""font-weight:bold;"">Card ID</th><th style=""font-weight:bold;"">Card Number</th>" & _
						"<th style=""font-weight:bold;"">Card Type</th><th style=""font-weight:bold;"">Limit Date From</th><th style=""font-weight:bold;"">Limit Date To</th><th style=""font-weight:bold;"">Limit Duration</th>" & _
						"<th style=""font-weight:bold;"">Credit Limit Original</th><th style=""font-weight:bold;"">Credit Limit New</th><th style=""font-weight:bold;"">Transaction Limit Original</th>" & _
						"<th style=""font-weight:bold;"">Transaction Limit New</th><th style=""font-weight:bold;"">OTC Limit Original</th><th style=""font-weight:bold;"">OTC Limit New</th>" & _
						"<th style=""font-weight:bold;"">ATM Limit Original</th><th style=""font-weight:bold;"">ATM Limit New</th><th style=""font-weight:bold;"">Process Status</th>" & _
						"<th style=""font-weight:bold;"">Changes Permanent</th><th style=""font-weight:bold;"">Application Status</th><th style=""font-weight:bold;"">Credit Limit Original VarChar</th>" & _
						"<th style=""font-weight:bold;"">Transaction Limit Original VarChar</th><th style=""font-weight:bold;"">Date Reduced</th><th style=""font-weight:bold;"">Date Updated</th>" & _
						"<th style=""font-weight:bold;"">Updated By</th></tr></thead><tbody>"
						
			Do Until objRS.Eof
				
				x = x + 1
				
				If clng(objRS("ApplicationID")) = clng(lngApplicationID) Then
					strApplication = "class=""selectedX"" title=""Currently Selected Limit Change Application. Other Limit Change applications for " & objRS("EmployeeID") & " have a white background."""
				Else
					strApplication = "class=""NotselectedX"" "
				End If
				
				'If x > 1 Then
				'	Response.write "<tr style=""border-top: black solid 2px""><td></td><td></td><td></td><td></td><tr>"
				'End If
				
				'Format LimitDateFrom
				If Not IsNull(objRS("LimitDateFrom")) Then
					'Date of birth id formatted different by differenet applications
					dteLimitDateFrom = Right(Trim(objRS("LimitDateFrom")),2) & "/" & Mid(Trim(objRS("LimitDateFrom")),6,2) & "/" & Left(Trim(objRS("LimitDateFrom")),4)
					
					If IsDate(dteLimitDateFrom) Then
						dteLimitDateFrom = dteLimitDateFrom' & "&nbsp;&nbsp;&nbsp; <span class=""badge badge-pill badge-info"">" & DateDiff("d",objRS("LimitDateFrom"),objRS("LimitDateTo")) & " days</span>"
					Else
						dteLimitDateFrom = objRS("LimitDateFrom")
					End If
				Else
					dteLimitDateFrom = ""
				End If
				
				'Format LimitDateTo
				If Not IsNull(objRS("LimitDateTo")) Then
					'Date of birth id formatted different by differenet applications
					dteLimitDateTo = Right(Trim(objRS("LimitDateTo")),2) & "/" & Mid(Trim(objRS("LimitDateTo")),6,2) & "/" & Left(Trim(objRS("LimitDateTo")),4)
					
					If IsDate(dteLimitDateTo) Then
						dteLimitDateTo = dteLimitDateTo' & "&nbsp;&nbsp;&nbsp; <span class=""badge badge-pill badge-info"">" & DateDiff("d",objRS("LimitDateFrom"),objRS("LimitDateTo")) & " days</span>"
						strDurationofLimit = "<span class=""badge badge-pill badge-info"">" & DateDiff("d",objRS("LimitDateFrom"),objRS("LimitDateTo")) & " days</span>"
					Else
						dteLimitDateTo = objRS("LimitDateTo")
					End If
				Else
					dteLimitDateTo = ""
				End If
	
				'Get the Application Status
				strApplicationStatus = objRS("ApplicationStatus")
				
				'Select the Action and Status buttons/pills based on the application status
				Select Case strApplicationStatus
				
				
				Case "Submitted"
					strApplicationStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">Submitted to GCFO</span>"
				Case "Deleted"
					strApplicationStatus = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">Deleted</span>"
				Case "Rejected"
					strApplicationStatus = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">Rejected</span>"
				Case "ASFIN Approved"
					strApplicationStatus = "<span class=""badge badge-pill badge-success"">Approved by ASFIN</span>"
				Case  "Awaiting Review"
					strApplicationStatus = "<span class=""badge badge-pill badge-warning"" style=""font-size:12px;"">" & strApplicationStatus & "</span>"
				Case "On Hold"
					strApplicationStatus = "<span class=""badge badge-pill badge-secondary"" style=""font-size:12px;"">" & strApplicationStatus & "</span>"
				Case "Temp Hold"
					strApplicationStatus = "<span class=""badge badge-pill badge-danger "" >" & strApplicationStatus & "</span>"
				Case  "Done"
					strApplicationStatus = "<span class=""badge badge-pill badge-info"" style=""font-size:12px;"">" & strApplicationStatus & "</span>"
				Case  "Awaiting Export"
					strApplicationStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">" & strApplicationStatus & "</span>"
				Case  "Awaiting Issue"
					strApplicationStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">" & strApplicationStatus & "</span>"
				Case  "Card Received"
					strApplicationStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">" & strApplicationStatus & "</span>"
				Case "Old App Version"
					strApplicationStatus = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">" & strApplicationStatus & "</span>"
				Case Else
					strApplicationStatus = "<span class=""badge badge-pill badge-secondary"" style=""font-size:12px;"">" & strApplicationStatus & "</span>"
				End Select
				
				'Vertical display
			'Response.Write "<tr><td style=""font-weight:bold;"" " & strApplication & ">LimitChangeID</td><td " & strApplication & ">" & objRS("LimitChangeID") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">EmployeeID</td><td " & strApplication & ">" & objRS("EmployeeID") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">ApplicationID</td><td " & strApplication & ">" & objRS("ApplicationID") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">XMLApplicationID</td><td " & strApplication & ">" & objRS("XMLApplicationID") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">CardID</td><td " & strApplication & ">" & objRS("CardID") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">CardNumber</td><td " & strApplication & ">" & objRS("CardNumber") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">CardType</td><td " & strApplication & ">" & objRS("CardType") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">LimitDateFrom</td><td " & strApplication & ">" & objRS("LimitDateFrom") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">LimitDateTo</td><td " & strApplication & ">" & objRS("LimitDateTo") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">CreditLimitOriginal</td><td " & strApplication & ">" & objRS("CreditLimitOriginal") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">CreditLimitNew</td><td " & strApplication & ">" & objRS("CreditLimitNew") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">TransactionLimitOriginal</td><td " & strApplication & ">" & objRS("TransactionLimitOriginal") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">TransactionLimitNew</td><td " & strApplication & ">" & objRS("TransactionLimitNew") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">OTCLimitOriginal</td><td " & strApplication & ">" & objRS("OTCLimitOriginal") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">OTCLimitNew</td><td " & strApplication & ">" & objRS("OTCLimitNew") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">ATMLimitOriginal</td><td " & strApplication & ">" & objRS("ATMLimitOriginal") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">ATMLimitNew</td><td " & strApplication & ">" & objRS("ATMLimitNew") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">ProcessStatus</td><td " & strApplication & ">" & objRS("ProcessStatus") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">ChangesPermanent</td><td " & strApplication & ">" & objRS("ChangesPermanent") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">ApplicationStatus</td><td " & strApplication & ">" & objRS("ApplicationStatus") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">CreditLimitOriginalVarChar</td><td " & strApplication & ">" & objRS("CreditLimitOriginalVarChar") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">TransactionLimitOriginalVarChar</td><td " & strApplication & ">" & objRS("TransactionLimitOriginalVarChar") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">DateReduced</td><td " & strApplication & ">" & objRS("DateReduced") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">DateUpdated</td><td " & strApplication & ">" & objRS("DateUpdated") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">UpdatedBy</td><td " & strApplication & ">" & objRS("UpdatedBy") & "</td>" & _
			'	"<td style=""font-weight:bold;"" " & strApplication & ">TransactionLimitDateFrom</td><td " & strApplication & ">" & objRS("TransactionLimitDateFrom") & "</td></tr>" & _
			'	"<tr><td style=""font-weight:bold;"" " & strApplication & ">TransactionLimitDateTo</td><td " & strApplication & ">" & objRS("TransactionLimitDateTo") & "</td>" & _ 
			'	"<td style=""font-weight:bold;"" " & strApplication & ">--- End of Application</td><td " & strApplication & ">" & objRS("ApplicationID") & " ----</td></tr>"
			
			'---Horizontal display
			Response.write "<tr><td  " & strApplication & ">" & objRS("LimitChangeID") & "</td><td  " & strApplication & ">" & objRS("EmployeeID") & "</td><td  " & strApplication & ">" & objRS("ApplicationID") & "</td>" & _ 
				"<td " & strApplication & ">" & objRS("XMLApplicationID") & "</td><td  " & strApplication & ">" & objRS("CardID") & "</td><td  " & strApplication & ">" & objRS("CardNumber") & "</td>" & _ 
				"<td " & strApplication & ">" & objRS("CardType") & "</td><td  " & strApplication & ">" & dteLimitDateFrom & "</td><td  " & strApplication & ">" & dteLimitDateTo & "</td><td  " & strApplication & ">" & strDurationofLimit & "</td>" & _ 
				"<td  " & strApplication & ">" & objRS("CreditLimitOriginal") & "</td><td  " & strApplication & ">" & objRS("CreditLimitNew") & "</td><td  " & strApplication & ">" & objRS("TransactionLimitOriginal") & "</td>" & _ 
				"<td  " & strApplication & ">" & objRS("TransactionLimitNew") & "</td><td  " & strApplication & ">" & objRS("OTCLimitOriginal") & "</td><td  " & strApplication & ">" & objRS("OTCLimitNew") & "</td>" & _ 
				"<td  " & strApplication & ">" & objRS("ATMLimitOriginal") & "</td><td  " & strApplication & ">" & objRS("ATMLimitNew") & "</td><td  " & strApplication & ">" & objRS("ProcessStatus") & "</td>" & _ 
				"<td  " & strApplication & ">" & objRS("ChangesPermanent") & "</td><td  " & strApplication & ">" & strApplicationStatus & "</td><td  " & strApplication & ">" & objRS("CreditLimitOriginalVarChar") & "</td>" & _ 
				"<td  " & strApplication & ">" & objRS("TransactionLimitOriginalVarChar") & "</td><td  " & strApplication & ">" & objRS("DateReduced") & "</td><td  " & strApplication & ">" & objRS("DateUpdated") & "</td>" & _ 
				"<td  " & strApplication & ">" & objRS("UpdatedBy") & "</td></tr>"
				
				
			objRS.MoveNext
			Loop
			
		Else
			Response.write "No Limit Change Application Record for Employee ID " & Request.QueryString("EmployeeID")
	   End If

	objRS.Close
	
	Response.write "</tbody></div></table>"
	
	
End Sub
%>
<head>
 <style>
 .selectedX {
    background-color: #f2f8fb;
	color: #0079ad;
	border:1px solid rgba(0, 0, 0, 0.2);
}
.NotselectedX {
	border:1px solid rgba(0, 0, 0, 0.2);
}
 </style>
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