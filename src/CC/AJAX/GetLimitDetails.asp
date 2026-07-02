<!-- #Include file=../CAPSFunctions.asp -->

<html lang="en">
<%
Dim objCon
Dim objRS
Dim x
Dim appID
Dim strSQL
Dim objCmd
Dim strMonthlyStartDate
Dim strTXNStartDate
Dim strMonthlyEndDate
Dim strTXNEndDate
Dim strMonthlyPermLimit
Dim strTXNPermLimitCode
Dim strTXNPermLimitAmount
Dim strMonthlyNewLimit
Dim strTXNNewLimitCode
Dim strTXNNewLimitAmount
Dim strMonthlyStatus
Dim strTXnStatus
Dim strMonthlyTemp
Dim strTXNTemp
Dim strMonthlyDateApproved
Dim strTXNDateApproved
Dim strMonthlyLimitData
Dim strTXNLimitData
Dim strMonthlyApplicationID
Dim strTXNApplicationID
Dim strMonthlyAction
Dim strTXnAction
Dim strTXNPermCashAmount
Dim strTXNTempCashAmount
Dim strMonthlyCardID
Dim strTXNCardID
Dim clID

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("AppID")) Then
	appID = Request.QueryString("AppID")
	strSQL = "SELECT * FROM qryCAPSCLCreditLimits WHERE MonthlyLimitAppID = " & appID & " OR TXNLimitAppID = " & appID & ""
End If

If Not IsEmpty(Request.QueryString("CLID")) Then
	clID = Request.QueryString("CLID")
	strSQL = "SELECT * FROM qryCAPSCLCreditLimits WHERE CLMonthlyLimitID = " & clID & " OR CLTXNLimitID = " & clID & ""
End If
'strSQL = "SELECT TOP 3 MonthlyLimitAppID FROM qryCAPSCLCreditLimits"

Public Sub FindLimitDetails()

objRS.Open strSQL,objCon

IF objRS.EOF Then
	Response.Write "No Credit Limit Detail has been found. Contact Admin if you believe this is an error."
Else	
	If objRS("MonthlyLimitAppID") <> 0 OR objRS("CLMonthlyLimitID") <> 0 Then
		'write monthly limit table and data
		strMonthlyLimitData = "<div class=""row row-cols-1 pt-3 pb-3 m-auto overflow-auto w-100 align-items-end""><h4 class=""text-center"">Monthly Limit</h4><table class=""table flex-fill w-100 table-striped"">"
		strMonthlyTemp = objRS("MonthlyPermTempFlag")
		strMonthlyStartDate = objRS("MonthlyLimitStartDate")
		strMonthlyEndDate = objRS("MonthlyLimitEndDate")
		strMonthlyApplicationID = objRS("MonthlyLimitAppID")
		strMonthlyStatus = objRS("MonthlyLimitStatus")
		strMonthlyPermLimit = FormatCurrency(objRS("MonthlyPermLimit"),2)
		strMonthlyNewLimit = FormatCurrency(objRS("MonthlyActiveLimit"),2)
		strMonthlyDateApproved = objRS("MonthlyLimitDateActioned")
		strMonthlyCardID = objRS("MonthlyLimitCardID")
		strMonthlyAction = "<a type=""button"" title=""View Application"" class=""btn btn-warning text-light btn-md"" href=""ApplicationDetail.asp?ApplicationID=" & strMonthlyApplicationID  & """><i class=""fa fa-eye""></i></a>"
		strMonthlyAction = strMonthlyAction & "<a type=""button"" title=""View Card"" class=""btn btn-success text-light btn-md"" href=""CardDetail.asp?CardID=" & strMonthlyCardID  & """><i class=""fa fa-credit-card""></i></a>"
			
			If strMonthlyTemp = "Y  " Then
				strMonthlyTemp = "Permanent"
			Else
				strMonthlyTemp = "Temporary"
			End If
			
			'If (FormatDateTime(Now, vbShortDate)) > strMonthlyEndDate Then
			If IsDate(strMonthlyEndDate) Then
				If Date() > FormatDateTime(strMonthlyEndDate, vbShortDate) Then
					If strMonthlyStatus = "Ended" OR strMonthlyStatus = "Done" Then
					Else
					strMonthlyEndDate = "<span class=""badge badge-pill badge-danger"" title=""The end date for this temporary limit has passed, please check that the limit reduction has occurred"">" & strMonthlyEndDate & "</span>"
					End If
				End If
			End If
			
			Select Case strMonthlyStatus 
				Case "Created"
					strMonthlyStatus = "<span class=""badge badge-pill badge-dark"">" & strMonthlyStatus & "</span>"
				Case "Awaiting Export"
					strMonthlyStatus = "<span class=""badge badge-pill badge-warning"">" & strMonthlyStatus & "</span>"
				Case "Exported"
					strMonthlyStatus = "<span class=""badge badge-pill badge-success"">" & strMonthlyStatus & "</span>"
				Case "Started"
					strMonthlyStatus = "<span class=""badge badge-pill badge-success"">" & strMonthlyStatus & "</span>"
				Case "Ended"
					strMonthlyStatus = "<span class=""badge badge-pill badge-warning"">" & strMonthlyStatus & "</span>"
				Case "Done"
					strMonthlyStatus = "<span class=""badge badge-pill badge-primary"">" & strMonthlyStatus & "</span>"
				Case Else
					strMonthlyStatus = strMonthlyStatus
			End Select
			
			If strMonthlyTemp = "Permanent" Then
				strMonthlyLimitData = strMonthlyLimitData & "<thead><tr><th scope=""col"" class=""w-50""></th><th scope=""col""></th></tr></thead>"
				strMonthlyLimitData = strMonthlyLimitData & "<tbody><tr><th scope=""row"" class=""w-50"">ApplicationID</th><th class=""text-center"">" & strMonthlyApplicationID & "</th></tr><tr><th scope=""row"">Limit Status</th><td class=""text-center"">" & strMonthlyStatus & "</td></tr><tr><th scope=""row"">Permanent or Temporary?</th><td class=""text-center"">" & strMonthlyTemp & "</td></tr><tr><th scope=""row"">Previous Monthly Limit</th><td class=""text-center"">" & strMonthlyPermLimit & "</td></tr><tr><th scope=""row"">New Monthly Limit</th><td class=""text-center"">" & strMonthlyNewLimit & "</td></tr><tr><th scope=""row"">Date Application Released</th><td class=""text-center"">" & strMonthlyDateApproved & "</td></tr><tr><th scope=""row"">View Details</th><td class=""text-center"">" & strMonthlyAction & "</td></tr></tbody>"
			Else
				strMonthlyLimitData = strMonthlyLimitData & "<thead><tr><th scope=""col"" class=""w-50""></th><th scope=""col""></th></tr></thead>"
				strMonthlyLimitData = strMonthlyLimitData & "<tbody><tr><th scope=""row"" class=""w-50"">ApplicationID</th><th class=""text-center"">" & strMonthlyApplicationID & "</th></tr><tr><th scope=""row"">Limit Status</th><td class=""text-center"">" & strMonthlyStatus & "</td></tr><tr><th scope=""row"">Permanent or Temporary?</th><td class=""text-center"">" & strMonthlyTemp & "</td></tr><tr><th scope=""row"">Permanent Monthly Limit</th><td class=""text-center"">" & strMonthlyPermLimit & "</td></tr><tr><th scope=""row"">Temporary Monthly Limit</th><td class=""text-center"">" & strMonthlyNewLimit & "</td></tr><tr><th scope=""row"">Start Date</th><td class=""text-center"">" & strMonthlyStartDate & "</td></tr><tr><th scope=""row"">End Date</th><td class=""text-center"">" & strMonthlyEndDate & "</td></tr><tr><th scope=""row"">Date Application Released</th><td class=""text-center"">" & strMonthlyDateApproved & "</td></tr><tr><th scope=""row"">View Details</th><td class=""text-center"">" & strMonthlyAction & "</td></tr></tbody>"			
			End If
			
			strMonthlyLimitData = strMonthlyLimitData & "</table></div>"
	Else 
		strMonthlyLimitData = ""
	End If
	
	If objRS("TXNLimitAppID") <> 0 OR objRS("CLTXNLimitID") <> 0 Then
		'write txn limit table and data
		strTXNLimitData = "<div class=""row row-cols-1 pt-3 pb-3 m-auto overflow-auto w-100 align-items-end ""><h4 class=""text-center"">Transaction Limit</h4><table class=""table flex-fill w-100 table-striped"">"
		strTXNTemp = objRS("TXNPermTempFlag")
		strTXNStartDate = objRS("TXNLimitStartDate")
		strTXNEndDate = objRS("TXNLimitEndDate")
		strTXNApplicationID = objRS("TXNLimitAppID")
		strTXNStatus = objRS("TXNLimitStatus")
		strTXNPermLimitCode = objRS("TXNPermLimitCode")
		strTXNPermLimitAmount = FormatCurrency(objRS("TXNPermLimitAmount"),2)
		strTXNPermCashAmount = FormatCurrency(objRS("TXNPermCashAmount"),2)
		strTXNNewLimitCode = objRS("TXNTempLimitCode")
		strTXNNewLimitAmount = FormatCurrency(objRS("TXNTempLimitAmount"),2)
		strTXNNewCashAmount = FormatCurrency(objRS("TXNTempCashAmount"),2)
		strTXNDateApproved = objRS("TXNLimitDateActioned")
		strTXNCardID = objRS("TXNLimitCardID")
		strTXNAction = "<a type=""button"" title=""View Application"" class=""btn btn-warning text-light btn-md"" href=""ApplicationDetail.asp?ApplicationID=" & strTXNApplicationID  & """><i class=""fa fa-eye""></i></a>"
		strTXNAction = strTXnAction & "<a type=""button"" title=""View Card"" class=""btn btn-success text-light btn-md"" href=""CardDetail.asp?CardID=" & strTXNCardID  & """><i class=""fa fa-credit-card""></i></a>"
			
			If strTXNTemp = "Y  " Then
				strTXNTemp = "Permanent"
			Else
				strTXNTemp = "Temporary"
			End If
			
			'If (FormatDateTime(Now, vbShortDate)) > strTXNEndDate Then
			If Date() > (FormatDateTime(strTXNEndDate, vbShortDate))  Then
				If strTXNStatus = "Ended" OR strTXNStatus = "Done" Then
				Else
				strTXNEndDate = "<span class=""badge badge-pill badge-danger"" title=""The end date for this temporary limit has passed, please check that the limit reduction has occurred"">" & strTXNEndDate & "</span>"
				End If
			End If
			
			Select Case strTXNStatus 
				Case "Created"
					strTXNStatus = "<span class=""badge badge-pill badge-dark"">" & strTXNStatus & "</span>"
				Case "Awaiting Export"
					strTXNStatus = "<span class=""badge badge-pill badge-warning"">" & strTXNStatus & "</span>"
				Case "Exported"
					strTXNStatus = "<span class=""badge badge-pill badge-success"">" & strTXNStatus & "</span>"
				Case "Started"
					strTXNStatus = "<span class=""badge badge-pill badge-success"">" & strTXNStatus & "</span>"
				Case "Ended"
					strTXNStatus = "<span class=""badge badge-pill badge-warning"">" & strTXNStatus & "</span>"
				Case "Done"
					strTXNStatus = "<span class=""badge badge-pill badge-primary"">" & strTXNStatus & "</span>"
				Case Else
					strTXNStatus = strTXNStatus
			End Select
			
			If strTXNTemp = "Permanent" Then
				'strTXNLimitData = strTXNLimitData & "<thead><tr><th scope=""col"">Application ID</th><th scope=""col"">Credit Limit Status</th><th scope=""col"">Type</th><th scope=""col"">Previous TXN Limit Code</th><th scope=""col"">New TXN Limit</th><th scope=""col"">Date Application Released</th><th scope=""col"">Action</th></tr></thead>"
				'strTXNLimitData = strTXNLimitData & "<tbody><tr><th scope=""row"">" & strTXNApplicationID & "</th><td>" & strTXNStatus & "</td><td>" & strTXNTemp & "</td><td>" & strTXNPermLimitCode & "</td><td>" & strTXNNewLimit & "</td><td>" & strTXNDateApproved & "</td><td>" & strTXNAction & "</td></tr></tbody>"		
			Else
				strTXNLimitData = strTXNLimitData & "<thead><tr><th scope=""col""></th><th scope=""col""></th></thead>"
				strTXNLimitData = strTXNLimitData & "<tbody><tr><th scope=""row"" class=""w-50"">ApplicationID</th><th class=""text-center"">" & strTXNApplicationID & "</th></tr><tr><th scope=""row"">Limit Status</th><td class=""text-center"">" & strTXNStatus & "</td></tr><tr><th scope=""row"">Permanent or Temporary?</th><td class=""text-center"">" & strTXNTemp & "</td></tr><tr><th scope=""row"">Permanent TXN Code</th><td class=""text-center"">" & strTXNPermLimitCode & "</td></tr><tr><th scope=""row"">Permanent Limit / Cash</th><td class=""text-center"">" & strTXNPermLimitAmount & " / " & strTXNPermCashAmount & "</td></tr><tr><th scope=""row"">Temporary TXN Code</th><td class=""text-center"">" & strTXNNewLimitCode & "</td></tr><th scope=""row"">Temporary TXN Limit / Cash</th><td class=""text-center"">" & strTXNNewLimitAmount & " / " & strTXNNewCashAmount & "</td></tr><th scope=""row"">Start Date</th><td class=""text-center"">" & strTXNStartDate & "</td></tr><tr><th scope=""row"">End Date</th><td class=""text-center"">" & strTXNEndDate & "</td></tr><tr><th scope=""row"">Date Application Released</th><td class=""text-center"">" & strTXNDateApproved & "</td></tr><tr><th scope=""row"">View Details</th><td class=""text-center"">" & strTXNAction & "</td></tr></tbody>"			
			End If
			
			strTXNLimitData = strTXNLimitData & "</table></div>"
	Else
		strTXNLimitData = ""
	End If
	
	response.write strMonthlyLimitData & strTXNLimitData

End If
objRS.Close
End Sub


%>
<head>
  
</head>

<body>
  <%
  
  	'Call the procedure to write the Limit information
	Call FindLimitDetails()

  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>