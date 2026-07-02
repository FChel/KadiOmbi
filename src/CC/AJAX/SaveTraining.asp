<!-- #Include file=../ADOVBS.inc -->
<%
If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp?State=Expired")

Dim objCon
Dim objRST
Dim objCmdT
Dim strEmployeeID
Dim strCourseID
Dim strBusinessArea
Dim strSelected
Dim intRecord
Dim strExemptDate

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRST = Server.CreateObject("ADODB.Recordset")
Set objCmdT = Server.CreateObject("ADODB.Command")

objCon.Open Session("DBConnection")	

Response.Expires=-1

If IsEmpty(Request.QueryString("EmployeeID")) Then
	'Response.Write "Error. No EmployeeID."
Else
	strEmployeeID = Request.QueryString("EmployeeID")
End If

If IsEmpty(Request.QueryString("EmployeeID")) Then
	'Response.Write "Error. No EmployeeID."
Else
	strEmployeeID = Request.QueryString("EmployeeID")
End If

If IsEmpty(Session("CourseID")) Then
	'Response.Write "Error. No CourseID."
Else
	strCourseID = Session("CourseID")'Request.QueryString("CourseID")
End If

If IsEmpty(Session("BusinessArea")) Then
	'Response.Write "Error. No BusinessArea."
	Session("BusinessArea") = "Compliance"
Else
	strBusinessArea = Session("BusinessArea")
End If

If IsEmpty(Session("ExemptDate")) Then
	'Response.Write "Error. No BusinessArea."
	strExemptDate = "NULL"
Else
	strExemptDate = Request.QueryString("ExemptDate")
End If

If IsEmpty(Request.QueryString("Selected")) Then
	'Response.Write "Error. No Selected."
Else
	If Request.QueryString("Selected")= "true" Then
		strSelected = "Y"
	Else
		strSelected = "N"
	End If
	'strSelected = Request.QueryString("Selected")
End If

'strEmployeeID = Replace(strEmployeeID,"'","''")

	'Call the procedure to save the details passed in if there are details in the Query string
	If IsNull(strEmployeeID) or strEmployeeID = "" Then
	Else
		Call SaveTrainingAction()
	End If
	
'Response.Write intRecord

	'Then call the procedure to load the selected employees for the current user logged in
	Call LoadTraining()

Sub SaveTrainingAction()

	
	With objCmdT

	.CommandType = 4
		.CommandText = "spCAPSTrainingActionSave"

		.Parameters.Append objCmdT.CreateParameter("CourseID", adVarChar, adParamInput, 10)
		.Parameters.Append objCmdT.CreateParameter("EmployeeID", adVarChar, adParamInput, 20)
		.Parameters.Append objCmdT.CreateParameter("BusinessArea", adVarChar, adParamInput, 20)
		.Parameters.Append objCmdT.CreateParameter("ExemptDate", adDate, adParamInput)
		.Parameters.Append objCmdT.CreateParameter("Selected", adChar, adParamInput, 1)              
		.Parameters.Append objCmdT.CreateParameter("UpdatedBy", adInteger, adParamInput)
		
		.Parameters.Append objCmdT.CreateParameter("TrainingActionIDOutput", adInteger, adParamOutput)
		
		.Parameters("CourseID") = strCourseID
		.Parameters("EmployeeID") = strEmployeeID
		.Parameters("BusinessArea") = strBusinessArea
		If strExemptDate = "NULL" Then
		Else
		.Parameters("ExemptDate") = strExemptDate		
		End If
		.Parameters("Selected") = strSelected
		.Parameters("UpdatedBy") = Session("UserID")
		
		.ActiveConnection = objCon
		 
	End With
   
	objCmdT.Execute        
  
	intRecord = objCmdT.Parameters.Item("TrainingActionIDOutput")	
	
End Sub

'Procedure to load the selected training records for the current user
Public Sub LoadTraining()

Dim strDateExempt

	Response.write "<div class=""card mb-3""><div class=""table-responsive table-compact"">" & _
            "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th style=""text-align:center; font-size:12px;"">EmpID</th><th style=""text-align:center; font-size:12px;"">Exempt</th><th style=""text-align:center; font-size:12px;"">Email</th></tr>"
					
	objRST.Open "SELECT * FROM tblCAPSTrainingAction WITH(NOLOCK) WHERE [UpdatedBy] = " & Session("UserID") & " AND [Selected] = 'Y' AND [BusinessArea] = '" & Session("BusinessArea") & "'",objCon

	If Not objRST.EOF Then
		   
		Do Until objRST.EOF
		
			If IsNull(objRST("ExemptDate")) or objRST("ExemptDate") = "" Then
				strDateExempt = ""
			Else
				strDateExempt = FormatDateTime(objRST("ExemptDate"),vbShortDate)
			End If
			
			'strEmails = strEmails & "<option id=""" & objRST("EmailDetailID") & """ value=""" & objRST("EmailDetailID") & """>" & objRST("EmailTemplateName") & "</option>"
			Response.Write "<tr class='clickable-row' data-href='Training.asp?EmployeeID=" & objRST("EmployeeID") & "' data-target='_blank'>" & _
					"<td style=""text-align:center; font-size:12px;""><a href='Training.asp?SearchInput=" & objRST("EmployeeID") & "'>" & objRST("EmployeeID") & "</a></td><td style=""text-align:center; font-size:12px;"">" & strDateExempt & "</td><td style=""text-align:center; font-size:12px;"">" & objRST("EmailID") & "</td></tr>"
			
		objRST.Movenext
		Loop
	Else
		'strEmails = "No email Templates available"
	End If
	
	objRST.Close
	
	Response.Write "</table></div>"
	
	'Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">" & _
    '        "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th>Exempt</th><th>Exempt Until</th><th>Email</th></tr>" & _
	'		"<tr><td style=""text-align:center;""><input type=""checkbox"" id=""chkTrain""></td><td style=""text-align:center;""><input type=""date"" id=""dteTrain""></td>" & _
	'		"<td style=""text-align:center;""><select id=""SelEmail"" name=""SelEmail""><option id=""0"" value=""0"">Select an Email Template</option>" & strEmails & "</select></td></tr>"&_
	'		"</table></div></div>"

End Sub

%>