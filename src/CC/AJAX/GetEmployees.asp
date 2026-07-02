<!DOCTYPE html>
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strFName
Dim strEmployeeID
Dim strLName
Dim strNameOnCard
Dim strSearchNameOnCard
Dim strSQL

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("FName")) Then
	strFName = " AND FirstName Like '%" & Request.QueryString("FName") & "%'"
'Else
'	strFName = " WHERE EmployeeID = '" & Session("EmployeeID") & "'"
End If

If Not IsEmpty(Request.QueryString("LName")) Then
	strLName = " AND Surname Like '%" & Request.QueryString("LName") & "%'"
End If

If Not IsEmpty(Request.QueryString("EmpID")) Then
	strEmployeeID = strLName & " AND EmployeeID Like '%" & Request.QueryString("EmpID") & "%'"
End If

If Not IsEmpty(Request.QueryString("NameOnCard")) Then
	strNameOnCard = " AND [name on account] Like '%" & Request.QueryString("NameOnCard") & "%'"
End If

'strEmployeeID = strFName & strLName & strEmployeeID & strNameOnCard

''''NEW October 2023 to allow for Name on card search which uses the Promaster Account table instead of the CDMC
'If the SearchType is for name on card then change the query being used
strSearchNameOnCard = ""

If Not IsEmpty(Request.QueryString("SearchType")) Then
	If Request.QueryString("SearchType") = "NameOnCard" Then
		strEmployeeID = strNameOnCard
		strSearchNameOnCard = "Yes"
	Else
		strEmployeeID = strFName & strLName & strEmployeeID
	End If
Else
	strEmployeeID = strFName & strLName & strEmployeeID
End If

'Remove the first AND from the SQL statement
If isNull(strEmployeeID) OR strEmployeeID= "" Then
	strEmployeeID = ""
Else
	strEmployeeID = " WHERE " & Right(strEmployeeID,Len(strEmployeeID) - 5)
End If

'''Set the SQL string with the table being searched
If strSearchNameOnCard = "Yes" Then
	strSQL = "SELECT Top 11 * FROM tblCAPSProMasterAccount " & strEmployeeID
Else
	strSQL = "SELECT Top 11 * FROM qryCAPSCDMCHistorySearch " & strEmployeeID
End If

Public Sub ShowDetails()


 'Description:	Loads Position details into page if applicable.
	objRS.Open strSQL,objCon
	'objRS.Open "SELECT * FROM tblCDMC WHERE EmployeeID = '" & Session("EmployeeID") & "'" & strFName,objCon

		Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">"
            '"<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th>EmployeeID</th><th>FirstName</th><th>LastName</th>"
			 '"<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th>EmployeeID</th><th>FirstName</th><th>LastName</th></tr>"
			 
		'Add the Name on Card if searched
		If strSearchNameOnCard = "Yes" Then
			Response.Write "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th>EmployeeID</th><th>CMS User Name</th><th>Name On Card</th>"
		Else
			Response.Write "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th>EmployeeID</th><th>FirstName</th><th>LastName</th>"
		End If
		
		Response.Write "</tr>"
		
		If Not objRS.EOF Then
		   
			Do Until objRS.EOF
				x = x + 1
				
				If x < 10 Then
					'Add the Name on Card if searched
					If strSearchNameOnCard = "Yes" Then
						Response.write "<tr onClick=""SelectEmp('" & objRS("cardholder eid") & "');""><td>" & objRS("cardholder eid") & "</td><td>" & objRS("user_name") & "</td><td>" & objRS("name on account") & "</td></tr>"
					Else
						Response.write "<tr onClick=""SelectEmp('" & objRS("EmployeeID") & "');""><td>" & objRS("EmployeeID") & "</td><td>" & objRS("FirstName") & "</td><td>" & objRS("Surname") & "</td></tr>"
					End If
					'Response.write "<tr onClick=""SelectEmp('" & objRS("EmployeeID") & "');""><td>" & objRS("EmployeeID") & "</td><td>" & objRS("FirstName") & "</td><td>" & objRS("Surname") & "</td><td>" & objRS("NameOnCard") & "</td></tr>"
				End If
			
			objRS.Movenext
			
			Loop
			'strEmployeeID = objRS("EmployeeID")
			'strFirstName = objRS("FirstName")
			'strLastName  = objRS("Surname")
			'strAddress1 = objRS("Address1")
			'strAddress2 = objRS("Address2")
			'strAddress3 = objRS("Address3")
			'strAddress4 = objRS("Address4")
			'strSuburb = objRS("Suburb")
			'strState = objRS("State")
			'strPostCode = objRS("PostCode")
			'dteDateReceived = objRS("DateReceived")
			'strStatus = objRS("Status")
			'strReviewedBy = objRS("ReviewedBy")
			'dteDateReviewed = objRS("DateReviewed")
			'lngCreditLimit = objRS("CreditLimit")
			
			'if there are more than 10 records then add an additional row to note this to the user
			If x > 10 Then Response.write "<tr><td colspan=""3""><i>Top 10 results displayed only. Use more detail (EID or FName and LName) to filter.</i></td></tr>"
			
		Else
			Response.write "empID" = Session("EmployeeID") 
			Response.write "<tr><td colspan=""3""><i>No results for " & strEmployeeID & "</i></td></tr>"
	   End If

	objRS.Close
	
	Response.write "</table></div></div></div>"
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
