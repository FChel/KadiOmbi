<!DOCTYPE html>
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strFName
Dim strEmployeeID
Dim strLName

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

strEmployeeID = strFName & strLName & strEmployeeID

If isnull(strEmployeeID) OR strEmployeeID= "" Then
	strEmployeeID = ""
Else
	strEmployeeID = " WHERE " & Right(strEmployeeID,Len(strEmployeeID) - 5)
End If

If Not IsEmpty(Request.QueryString("CardsSearch")) Then
	strSQL = "With tblCards AS (SELECT COUNT(*) as Cards,EmployeeID as CardEID FROM tblCAPSCard Group By EmployeeID) SELECT EmployeeID,FirstName,Surname,Cards FROM tblCAPSCDMC LEFT OUTER JOIN tblCards ON tblCards.CardEID=tblCAPSCDMC.EmployeeID "  & strEmployeeID
Else
	strSQL = "SELECT * FROM tblCAPSCDMC " & strEmployeeID
End If

Public Sub ShowDetails()


 'Description:	Loads Position details into page if applicable.
	objRS.Open strSQL,objCon
	'objRS.Open "SELECT * FROM tblCAPSCDMC " & strEmployeeID,objCon
	'objRS.Open "SELECT * FROM tblCDMC WHERE EmployeeID = '" & Session("EmployeeID") & "'" & strFName,objCon

	If Not IsEmpty(Request.QueryString("CardsSearch")) Then
		Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">" & _
            "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th>EmployeeID</th><th>FirstName</th><th>LastName</th><th>Cards</th></tr>"
	Else
		Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">" & _
            "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th>EmployeeID</th><th>FirstName</th><th>LastName77</th></tr>"
	End If
			  
		If Not objRS.EOF Then
		   
			Do Until objRS.EOF
				x = x + 1
				
				If x < 10 Then
					If Not IsEmpty(Request.QueryString("CardsSearch")) Then
						Response.write "<tr onClick=""SelectEmp(" & objRS("EmployeeID") & ");""><td>" & objRS("EmployeeID") & "</td><td>" & objRS("FirstName") & "</td><td>" & objRS("Surname") & "</td><td>" & objRS("Cards") & "</td></tr>"
					Else
						Response.write "<tr onClick=""SelectEmp(" & objRS("EmployeeID") & ");""><td>" & objRS("EmployeeID") & "</td><td>" & objRS("FirstName") & "</td><td>" & objRS("Surname") & "</td></tr>"
					End If
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
		Else
			Response.write " EmpID" = Session("EmployeeID") 
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
