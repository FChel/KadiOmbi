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
Dim strDeleted
Dim strCardType

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("CardType")) Then

	strCardType = Request.QueryString("CardType")
	
End If

If Not IsEmpty(Request.QueryString("Deleted")) Then
	If Request.QueryString("Deleted") = "Y" Then
		strDeleted = ""
	Else
		strDeleted = "AND Status='Awaiting Export'"
	End If
Else
	strDeleted = "AND Status='Awaiting Export'"
End If

'If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	strWhere = " WHERE Left(CardTypeSub,3) <> 'NAB' AND (FileSeqNum = '0' OR FileSeqNum Is Null OR FileSeqNum ='') " & strDeleted
	'strWhere = " WHERE CardType = '" & strCardType & "' AND (FileSeqNum = '0' OR FileSeqNum Is Null OR FileSeqNum ='') " & strDeleted
'End If

Public Sub ShowDetails()

Dim strStatus
Dim strStatusDisplay
Dim strUpdateDate
Dim lngRecords

	'Description:	Loads CDMC Details onto the page called from
	objRS.Open "SELECT * FROM qryCAPSCSToDiners WITH(NOLOCK) " & strWhere,objCon,3,2
	
	Response.Write "<span style=""font-style: italic; color: gray;"">SELECT * FROM qryCAPSCSToDiners WITH(NOLOCK) " & strWhere & "</span>"
	
	
		If Not objRS.EOF Then
			
			objRS.Movelast
			objRS.Movefirst
			
			lngRecords = objRS.Recordcount
			
			'Write the Header			
			Response.write "<div class=""row""><div class=""col-md-12"">" & _
				"<table class=""table table-compact"">"
			
			'Total Records
			Response.Write "<tr><th colspan=""7"" style=""text-align:right; border:none;"">Total CS Record to be Exported:</th><th>" & lngRecords & "</th>" & _
				"<td style=""font-size:14px;""><button type=""button"" class=""btn btn-secondary btn-xs""onClick=""loadCSDeleted();"" Title=""Click to View all records not sent including DELETED records not in CS File to Add to CS File""><i class=""fa fa-times""></i> View Deleted</button></td></tr>"
			
			'Write the table Headers
			Response.Write "<tr><th>CS ID</th><th>Card Type</th><th>Card Type Sub</th><th>EID</th><th>Name</th><th>Card Status</th><th>Process Status</th><th>Card No</th><th>Action</th><th width=""100px;""></th></tr>"
			
			Do Until objRS.EOF
			
			If IsNull(objRS("Status")) Then 
				strStatus = "Added to CS"
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
			
			If trim(strStatus) = "Added To CS" Then
				strStatusDisplay = "<span class=""badge badge-pill badge-success""  style=""font-size:12px;"">" & strStatus & "</span>"
				strStatus = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ExportCS.asp?Action=CancelCS&CSToDinersID=" & objrs("CSToDinersID") & "&CSEID=" & objRS("EIDNo") & "&Status=Deleted'""; title=""Click to Remove from CS File""><i class=""fa fa-times""></i> Remove</button>"
				
			ElseIf strStatus = "Deleted" Then
				strStatusDisplay = "<span class=""badge badge-pill badge-danger""  style=""font-size:12px;"">" & strStatus & "</span>"
				strStatus = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='ExportCS.asp?Action=CancelCS&CSToDinersID=" & objrs("CSToDinersID") & "&CSEID=" & objRS("EIDNo") & "&Status=Awaiting Export'""; title=""Click to Add to CS File""><i class=""fa fa-plus""></i> Add</button>"

			Else
				strStatusDisplay = "<span  style=""font-size:12px;"" class=""badge badge-pill badge-warning"">" & strStatus & "</span>"
				strStatus = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ExportCS.asp?Action=CancelCS&CSToDinersID=" & objrs("CSToDinersID") & "&CSEID=" & objRS("EIDNo") & "&Status=Deleted'""; title=""Click to Remove from CS File""><i class=""fa fa-times""></i> Remove</button>"
			End If
			
			Response.Write "<tr><td style=""font-size:12px;"">" & objRS("CSToDinersID") & "</td><td style=""font-size:12px;"">" & objRS("CardType") & "</td><td style=""font-size:12px;"">" & objRS("CardTypeSub") & "</td><td style=""font-size:12px;"">" & objRS("EIDNo") & "</td><td style=""font-size:12px;"">" & objRS("NameOnCard") & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("CardStatus") & "</td><td style=""font-size:12px;"">" & strStatusDisplay & "</td><td style=""font-size:12px;"">" & MaskCard(objRS("CardNo")) & "</td>" & _
				"</td><td style=""font-size:12px;"">" & strStatus & "</td></tr>"
			
			
				objRS.Movenext
			Loop
						
		Else
			Response.write "No CS File to Export (with no Batch Number)"
	   End If

	objRS.Close
	
	Response.write "</div></table>"
	
	
	
End Sub
%>
<head>
  
</head>

<body>
  <%
  
	'Call the procedure to check for CS Issues
	Call CheckCSIssues()
	
	'Call the procedure to show CS file details
	Call ShowDetails()
  
  %>
</body>

</html>
<%

Public Sub CheckCSIssues()
'Procedure to check for issues in file and display on screen

Dim strSQL
Dim x

	For x = 1 to 8
	
		Select Case x
		
			Case 1
				strSQL = "SELECT * FROM tblCAPSCSToDiners WHERE Left(CardTypeSub,3) <> 'NAB' AND CardType = '" & strCardType & "' AND Status ='Awaiting Export' AND FileSeqNum = '' AND (LEN(Address1)>40 OR LEN(Address2)>40 OR LEN(Address3)>40)"
			Case 2
				strSQL = "SELECT * FROM tblCAPSCSToDiners WHERE Left(CardTypeSub,3) <> 'NAB' AND CardType = '" & strCardType & "' AND Status ='Awaiting Export' AND FileSeqNum = '' AND (LEN(Address1)>40 OR LEN(Address2)>40 OR LEN(Address3)>40)"
			Case 3
				strSQL = "SELECT * FROM tblCAPSCSToDiners WHERE Left(CardTypeSub,3) <> 'NAB' AND CardType = '" & strCardType & "' AND Status ='Awaiting Export' AND FileSeqNum = '' AND (EIDNo='' OR EIDNo Is Null)"
			Case 4
				strSQL = "SELECT * FROM tblCAPSCSToDiners WHERE Left(CardTypeSub,3) <> 'NAB' AND CardType = '" & strCardType & "' AND Status ='Awaiting Export' AND FileSeqNum = '' AND (CardNo='' OR CardNo Is Null)"
			Case 5
				'If there are cards to be deleted on the CS File then display a warning message
				strSQL = "SELECT * FROM tblCAPSCSToDiners WHERE Left(CardTypeSub,3) <> 'NAB' AND CardType = '" & strCardType & "' AND Status ='Awaiting Export' AND FileSeqNum = '' AND CardStatus <> '00'"
			Case 6
				strSQL = "SELECT * FROM tblCAPSCSToDiners WHERE Left(CardTypeSub,3) <> 'NAB' AND CardType = '" & strCardType & "' AND Status ='Awaiting Export' AND FileSeqNum = '' AND ((GivenNames='' OR GivenNames Is Null) OR (Surname='' OR Surname Is Null))"
			Case 7
				strSQL = "SELECT * FROM tblCAPSCSToDiners WHERE Left(CardTypeSub,3) <> 'NAB' AND CardType = '" & strCardType & "' AND Status ='Awaiting Export' AND FileSeqNum = '' AND (WorkPhone = '' OR WorkPhone Is Null)"
			Case 8
				strSQL = "SELECT * FROM tblCAPSCSToDiners WHERE Left(CardTypeSub,3) <> 'NAB' AND CardType = '" & strCardType & "' AND Status ='Awaiting Export' AND FileSeqNum = '' AND (CreditLimit = '' OR CreditLimit Is Null)"
		
		End Select
		
		objRS.Open strSQL, objCon
		
			If objRS.EOF then
			Else
				If x = 5 Then
					Response.Write "<div class=""alert alert-warning"" role=""alert"">Cards to be deleted in CS File: " & strSQL & "</div>"
				Else
					Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR in CS File: " & strSQL & "</div>"
				End If
			End If
	
		objRS.Close
		
	Next
	

End Sub


Set objRS = Nothing
Set objCon = Nothing
  
  %>