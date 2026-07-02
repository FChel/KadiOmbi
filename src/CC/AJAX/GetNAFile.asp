<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strWhere
Dim strStatus
Dim strName
Dim strSubStatPos
Dim strNameCard
Dim strSubStatPosCard
Dim strCardType

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("CardType")) Then

	strCardType = Request.QueryString("CardType")

End If

Response.Write strCardType

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	strWhere = " WHERE CardType = '" & strCardType & "' AND EmployeeID = '" & Request.QueryString("EmployeeID") & "'"
Else
	strWhere = " WHERE CardType = '" & strCardType & "' AND BatchNumber = 0"
End If


Public Sub ShowDetails()

Dim strStatus
Dim strStatusDisplay
Dim strUpdateDate

	strWhere = " WHERE CardType = '" & strCardType & "' AND BatchNumber = 0 AND [Status] <> 'Deleted'"
	
	
	'Description:	Loads CDMC Details onto the page called from
	objRS.Open "SELECT * FROM qryCAPSNAToDiners WITH(NOLOCK) " & strWhere,objCon
			  
		If Not objRS.EOF Then
			
			'Write the Header			
			Response.write "<div class=""row""><div class=""col-md-12"">" & _
				"<table class=""table table-compact"">"
			
			'Write the table Headers
			Response.Write "<tr><th>ID</th><th>Record Type</th><th>NA Detail</th><th>Status</th><th>Batch Number</th><th>Date Updated</th><th>Updated By</th><th>Employee ID</th><th width=""100px;""></th></tr>"
			
			Do Until objRS.EOF
			
			If IsNull(objRS("Status")) Then 
				strStatus = "Added to NA"
			Else
				strStatus = objRS("Status")
			End If
			
			If IsNull(objRS("DateUpdated")) Then
				strUpdateDate = ""
			Else
				If IsDate(objRS("DateUpdated")) Then
					strUpdateDate = FormatDateTime(objRS("DateUpdated"), vbShortDate)
				Else
					strUpdateDate = objRS("DateUpdated")
				End If
			End If
			
			If trim(strStatus) = "Added To NA" Then
				strStatusDisplay = "<span class=""badge badge-pill badge-success"">" & strStatus & "</span>"
				strStatus = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ExportNA.asp?Action=CancelNA&NAToDinersID=" & objrs("NAToDinersID") & "&NAEID=" & objRS("EmployeeID") & "&Status=Deleted'""; title=""Click to Remove from NA File""><i class=""fa fa-times""></i> Remove</button>"
				
			ElseIf strStatus = "Deleted" Then
				strStatusDisplay = "<span class=""badge badge-pill badge-danger"">" & strStatus & "</span>"
				strStatus = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='ExportNA.asp?Action=CancelNA&NAToDinersID=" & objrs("NAToDinersID") & "&NAEID=" & objRS("EmployeeID") & "&Status=Added To NA'""; title=""Click to Add to NA File""><i class=""fa fa-plus""></i> Add</button>"

			Else
				strStatusDisplay = "<span class=""badge badge-pill badge-warning"">" & strStatus & "</span>"
				strStatus = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ExportNA.asp?Action=CancelNA&NAToDinersID=" & objrs("NAToDinersID") & "&NAEID=" & objRS("EmployeeID") & "&Status=Deleted'""; title=""Click to Remove from NA File""><i class=""fa fa-times""></i> Remove</button>"
			End If
			
			Response.Write "<tr><td>" & objRS("NAToDinersID") & "</td><td>" & objRS("RecordType") & "</td><td>" & objRS("RecordText") & "</td><td>" & strStatusDisplay & "</td><td>" & objRS("BatchNumber") & "</td>" & _
				"<td title=""" & objRS("DateUpdated") & """>" & strUpdateDate & "</td><td>" & objRS("UpdatedByName") & "</td><td>" & objRS("EmployeeID") & "</td><td>" & strStatus & "</td></tr>"
			
			
				objRS.Movenext
			Loop
						
		Else
			Response.write "No NA File to Export (where there are Applications with no Batch Number)"
	   End If

	objRS.Close
	
	Response.write "</div></table>"
	
	
End Sub
%>
<head>
  
</head>

<body>
  <%
  
	'Call the error checking process
	Call CheckNAIssues()
	
	'Call the procedure to display the NA File details
	Call ShowDetails()
  
  %>
</body>

</html>
<%

Public Sub CheckNAIssues()
'Procedure to check for issues in file and display on screen

Dim strSQL
Dim x

	For x = 1 to 2
	
		Select Case x
		
			Case 1
				strSQL = "SELECT * from tblCAPSNAToDiners WHERE BatchNumber = 0 AND (RecordText = '' OR RecordText Is Null)"
			Case 2
				strSQL = "SELECT * from tblCAPSNAToDiners WHERE BatchNumber = 0 AND (RecordType = '' OR RecordType Is Null)"
			
		End Select
		
		objRS.Open strSQL, objCon
		
			If objRS.EOF then
			Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR in NA File: " & strSQL & "</div>"
			End If
	
		objRS.Close
		
	Next
	

End Sub

Set objRS = Nothing
Set objCon = Nothing
  
  %>