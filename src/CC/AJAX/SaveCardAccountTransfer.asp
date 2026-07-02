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
Dim strConfirm
Dim strReturn

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRST = Server.CreateObject("ADODB.Recordset")
Set objCmdT = Server.CreateObject("ADODB.Command")

objCon.Open Session("DBConnection")	

Response.Expires=-1

If IsEmpty(Request.QueryString("EmployeeID")) Then
	'Response.Write "Error. No EmployeeID."
	'strEmployeeID = "0"
Else
	strEmployeeID = Request.QueryString("EmployeeID")
	
	If strEmployeeID = "" Then strEmployeeID = "0"
End If

If IsEmpty(Request.QueryString("CourseID")) Then
	'Response.Write "Error. No CourseID."
Else
	strCourseID = Request.QueryString("CourseID")'Session("CourseID")'
End If

'If IsEmpty(Session("BusinessArea")) Then
	'Response.Write "Error. No BusinessArea."
'	Session("BusinessArea") = "CardAccountTransfer"
'Else
	strBusinessArea = "CardAccountTransfer"'Session("BusinessArea")
'End If

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

'If this page is called by the confirm modal then display details only (no actions, buttons)
If IsEmpty(Request.QueryString("Confirm")) Then
	strConfirm = ""
Else
	strConfirm = "Y"
End If

'If this page is called by the confirm modal then display details only (no actions, buttons)
If IsEmpty(Request.QueryString("Return")) Then
	strReturn = ""
Else
	strReturn = Request.QueryString("Return")
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
		.CommandText = "spCAPSCardAccountTransferActionSave"

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

Dim strName
Dim strClear
Dim strCardMasked

	objRST.Open "SELECT * FROM tblCAPSTrainingAction WITH(NOLOCK) WHERE [UpdatedBy] = " & Session("UserID") & " AND [Selected] = 'Y' AND [BusinessArea] = 'CardAccountTransfer'",objCon
	
	Response.write "<div class=""card mb-3"">"
		
	'If the recordset has records then write the Clear Selected button
	If objRST.EOF Then
		strClear = ""
	Else
		If strConfirm = "Y" Then
		Else
			strClear = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onClick=""self.location.href='CardAccountTransfer.asp?Action=ClearSelected';""><i class=""fa fa-times""></i> Clear Selected</button>"
		End If
	End If
	
	'Allow the user to clear the selected batch if one is selected
	If IsNull(Session("CATEmployeeID")) Or Session("CATEmployeeID") = "0"  Then
		strClear = strClear & ""
	Else
		If strConfirm = "Y" Then
			strClear = strClear & "<b>Accounts Transferring From</b>"
		Else
			strClear = strClear & " <button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModalConfirm"" OnClick=""ConfirmCATAction('')""><i class=""fa fa-check""></i> Transfer Now </button>"
			'strClear = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onClick=""self.location.href='CardAccountTransfer.asp?Action=Transfer';""><i class=""fa fa-check""></i> Transfer Now</button>"
		End If
	End If
	
	Response.Write "<div class=""panel panel-light mb-0""><div class=""panel-header"" style=""padding-right:1px; padding-left:1px;""><h4></h4>" & strClear & "</div></div>"

	 'If the Confirmation modal has called this page then add the Comments text area
	If strConfirm = "Y" Then
		Response.Write "<div class=""table-responsive table-compact"">" & _
					"<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0"">" & _
					"<thead><tr style=""background-color:#EAECEE;""><td style=""text-align:center; font-size:12px; font-weight:bold; background-color:#EAECEE;"">Comments</td>" & _
					"<td style=""text-align:center; font-size:12px; background-color:grey;"" colspan=""2""><input type=""text"" autofocus id=""TransferComments"" Name=""TransferComments"" maxlength=""100"" placeholder=""Enter Comments...."" style=""width:100%; z-index:-1;"" /></td></tr>" & _
					"<tr><th colspan=""2""><input type=""hidden"" id=""ReturnType"" Name=""ReturnType"" value=""" & strReturn & """ /></th></tr>" & _
					"<tr><th style=""text-align:center; font-size:12px;"">EmpID</th><th style=""text-align:center; font-size:12px;"">CardNo.</th><th style=""text-align:center; font-size:12px;"">Name</th></tr>"
	Else
		Response.write "<div class=""table-responsive table-compact"">" & _
            "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th style=""text-align:center; font-size:12px;"">EmpID</th><th style=""text-align:center; font-size:12px;"">CardNo.</th><th style=""text-align:center; font-size:12px;"">Name</th></tr>"
	
	End If
	
	'Recordset open at the top of this procedure to assist with writing dyamic buttons
	If Not objRST.EOF Then
		Do Until objRST.EOF
		
			If IsNull(objRST("FirstName")) or objRST("FirstName") = "" Then
				strName = ""
			Else
				strName = objRST("FirstName") & " " & objRST("LastName")
				
				strName = Left(strName,15)
			End If
			
			If IsNull(objRST("CourseTitle")) or objRST("CourseTitle") = "" Then
				strCardMasked = ""
			Else
				If Len(objRST("CourseTitle"))>4 Then
					strCardMasked = Left(objRST("CourseTitle"),2) & "****" & Right(objRST("CourseTitle"),4)
				Else
					strCardMasked = objRST("CourseTitle")
				End If
				
			End If
			
			'If this page was called by the Confirm modal then do not display anchors
			If strConfirm = "Y" Then
				
				Response.Write "<tr><td style=""text-align:center; font-size:12px;"">" & objRST("EmployeeID") & "</td><td style=""text-align:center; font-size:12px;"">" & strCardMasked & "</td><td style=""text-align:center; font-size:12px;"">" & strName & "</td></tr>"
			Else
			
				Session("TrainingSelectedNow") = Session("TrainingSelectedNow") & "," & objRST("EmployeeID")
				
				'strEmails = strEmails & "<option id=""" & objRST("EmailDetailID") & """ value=""" & objRST("EmailDetailID") & """>" & objRST("EmailTemplateName") & "</option>"
				Response.Write "<tr class='clickable-row' data-href='Training.asp?EmployeeID=" & objRST("EmployeeID") & "' data-target='_blank'>" & _
						"<td style=""text-align:center; font-size:12px;""><a href=""#"" data-id=""" & objRST("EmployeeID") & """ data-CourseID=""" & objRST("CourseID") & """ onClick=""SaveCATAction(this);"">" & objRST("EmployeeID") & "</a></td><td style=""text-align:center; font-size:12px;"">" & strCardMasked & "</td><td style=""text-align:center; font-size:12px;"">" & strName & "</td></tr>"
			End If
			
		objRST.Movenext
		Loop
		
	Else
		'strEmails = "No email Templates available"
	End If
	
	objRST.Close
	
	Response.Write "</table></div>"
	

End Sub

%>