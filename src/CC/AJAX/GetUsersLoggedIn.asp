<!DOCTYPE html>
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strUsers

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

'If Not IsEmpty(Request.QueryString("FileSeqNum")) Then
'	strBatchNo = " WHERE FileSeqNum Like '%" & Request.QueryString("FileSeqNum") & "%' AND [FileType] = 'CSFromDiners' AND [Deleted] = 'N'"
'End If

If IsNull(Application("NamedUsers")) Then
	strUsers = " WHERE UserID=0"
Else
	strUsers = " WHERE UserID IN (" & Application("NamedUsers") & ")"
End If

Public Sub ShowDetails()


	'Description:	Loads Position details into page if applicable.
	'objRS.Open "SELECT [UserID],[FName],[LName],[EmployeeID],[UserTypeID] FROM tblUsers " & strUsers,objCon
	objRS.Open "SELECT [UserID],[FName],[LName],[EmployeeID],[tblUsers].[UserTypeID] AS UserTypeID,[tblUserTypes].[UserTypeName] from tblUsers WITH(NOLOCK) RIGHT JOIN tblUserTypes ON tblUsers.UserTypeID = tblUserTypes.UserTypeID " & strUsers,objCon
 
		Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">" & _
            "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0"">" & _
			"<thead><tr><th>User ID</th><th>User Name</th><th>EmployeeID</th><th>User Type</th></tr>"
			  
			  
		If Not objRS.EOF Then
		   
			Do Until objRS.EOF
				x = x + 1
				
				Response.write "<tr><td>" & objRS("UserID") & "</td><td>" & objRS("FName") & " " & objRS("LName") & "</td><td>" & objRS("EmployeeID") & "</td><td>" & objRS("UserTypeID") & " - " & objRS("UserTypeName") & "</td></tr>"
				
			objRS.Movenext
			
			Loop
			
			Response.write "<tr><td colspan=""3"" style=""font-weight:bold;"">Total Users Logged In:</td><td style=""font-weight:bold;"">" & x & "</td></tr>"
			
		Else
			Response.write "No Users Logged in"
	   End If

	objRS.Close
	
	Response.write "</table></div></div></div>"
End Sub
%>
<head>
  
</head>

<body>
  <%
   'response.write "sdsfgs thghgffhf "
  Call ShowDetails()
 
  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>