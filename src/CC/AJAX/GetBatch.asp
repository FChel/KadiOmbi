<!DOCTYPE html>
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strBatchNo

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("FileSeqNum")) Then
	strBatchNo = " WHERE FileSeqNum Like '%" & Request.QueryString("FileSeqNum") & "%' AND [FileType] = 'CSFromDiners' AND [Deleted] = 'N'"
End If


Public Sub ShowDetails()


	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT [FileSeqNum],[FileDateTime],[Status] FROM tblCAPSFileLoad " & strBatchNo,objCon

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
  Call ShowDetails()
  
  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>