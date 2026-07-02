<!DOCTYPE html>
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("GDriveSet")) Then
	
	Session("ReadGDrive") = Request.QueryString("GDriveSet")
	
End If

Public Sub DisplaySettings()

Dim strReadGDrive
Dim strReadGDriveActive(2)

	If IsNull(Session("ReadGDrive")) OR Session("ReadGDrive") = "" Then
		strReadGDriveActive(1) = "active"
	Else
		If Session("ReadGDrive") = "On" Then
			'strReadGDrive = ""
			strReadGDriveActive(1) = "active"
		Else
			'strReadGDrive = ""
			strReadGDriveActive(2) = "active"
		End If
	End If
	
	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & strReadGDriveActive(1) & """ onClick=""SaveReadGDrive('On');""><i class=""fa fa-check""></i> Read G Drive On</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & strReadGDriveActive(2) & """ onClick=""SaveReadGDrive('Off');""><i class=""fa fa-pause-circle""></i> Read G Drive Off</button>" & _
				"</div>"
	
End Sub

Public Sub ShowDetails()


	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM tblCAPSFileLoad " & strBatchNo,objCon

		Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">" & _
            "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0"">" & _
			"<thead><tr><th>Batch No</th><th>File Date Time</th><th>Status</th></tr>"
			  
			  
		If Not objRS.EOF Then
		   
			Do Until objRS.EOF
				x = x + 1
				
				
				Response.write "<tr onClick=""SelectFileSeqNum(String('" & objRS("FileSeqNum") & "'));""><td>" & objRS("FileSeqNum") & "</td><td>" & objRS("FileDateTime") & "</td><td>" & objRS("Status") & "</td></tr>"
				
			
			objRS.Movenext
			
			Loop
			
		Else
			Response.write "No File Seq Num for " & Request.QueryString("FileSeqNum")
	   End If

	objRS.Close
	
	Response.write "</table></div></div></div>"
End Sub
%>
<head>
  
</head>

<body>
  <%
  'Call ShowDetails()
  Call DisplaySettings()
  
  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>