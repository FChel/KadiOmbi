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
		Call SaveEmployeeAction()
	End If
	
Sub SaveEmployeeAction()

	
	With objCmdT

	.CommandType = 4
		.CommandText = "spCAPSDontCancel"

		.Parameters.Append objCmdT.CreateParameter("EmployeeActionID", adInteger, adParamInput)
		.Parameters.Append objCmdT.CreateParameter("EmployeeID", adVarChar, adParamInput, 20)
		.Parameters.Append objCmdT.CreateParameter("Action", adVarChar, adParamInput, 20)
		.Parameters.Append objCmdT.CreateParameter("Active", adVarChar, adParamInput, 10)
		.Parameters.Append objCmdT.CreateParameter("Detail", adVarChar, adParamInput, 200)
		.Parameters.Append objCmdT.CreateParameter("Notes", adVarChar, adParamInput, 500)             
		.Parameters.Append objCmdT.CreateParameter("UpdatedBy", adInteger, adParamInput)
		
		.Parameters.Append objCmdT.CreateParameter("EmployeeActionIDDOutput", adInteger, adParamOutput)
		
		.Parameters("EmployeeActionID") = 0
		.Parameters("EmployeeID") = strEmployeeID
		.Parameters("Action") = "DontCancel"
		.Parameters("Active") = strSelected		
		.Parameters("Detail") = "Employee will not be cancelled if Active is Y"
		.Parameters("Notes") = ""
		.Parameters("UpdatedBy") = Session("UserID")
		
		.ActiveConnection = objCon
		 
	End With
   
	objCmdT.Execute        
  
	intRecord = objCmdT.Parameters.Item("EmployeeActionIDDOutput")	
	
	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Employee " & strEmployeeID & " NOT updated! An Error has occurred. See System Admin with Employee ID: " & strEmployeeID & " </div>"
	Else
		If strSelected = "N" Then
			Response.Write "<div class=""alert alert-warning"" role=""alert"">Employee " & strEmployeeID& " Updated to: Don't Cancel = " & strSelected & "!</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Employee " & strEmployeeID& " Updated to: Don't Cancel = " & strSelected & "!</div>"
		End If
	End If
		
		
End Sub

%>