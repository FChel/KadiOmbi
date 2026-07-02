<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

Dim objCon
Dim objCon2
Dim objRS
Dim objRS2
Dim x
Dim strCMSAccount
Dim strEmployeeID
Dim strLName
Dim strCMSFilePath
Dim lngCardID
Dim strLoadEmployee

on error resume next
	'ProMaster Connection details
	Set objCon2 = Server.CreateObject("ADODB.Connection")
	Session("DBConnection2") = "File Name=" & Server.MapPath("../../Database/ProMaster.udl") & ";"
	'Get the File Path for the CMS UDL from System Parameters
	If IsEmpty(Session("DBConnection2")) Then
		strCMSFilePath = GetSystemAdmin("CMSServerFilePath")
		
		'Session("DBConnection2") = "File Name=" & strCMSFilePath & ";"
		Session("DBConnection2") = "File Name=" & Server.MapPath("../../Database/ProMaster.udl") & ";"
	End If
	
	'Session("DBConnection2") = "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	objCon2.ConnectionTimeout=2
	objCon2.Open Session("DBConnection2")
	
on error goto 0

	
'Open Database Connection
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS2 = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("AccountCMS")) Then
	strCMSAccount = " AND [user_name] Like '%" & Request.QueryString("AccountCMS") & "%'"
End If

If Not IsEmpty(Request.QueryString("EmpID")) Then
	If Request.QueryString("EmpID") = "" Then strLoadEmployee = "N"
	
	strEmployeeID = strCMSAccount & " AND employee_id Like '%" & Replace(Request.QueryString("EmpID"), "'", "") & "%'"
End If

If Not IsEmpty(Request.QueryString("CardID")) Then
	lngCardID = Request.QueryString("CardID")
End If

'strEmployeeID = strFName & strLName & strEmployeeID

'If isnull(strEmployeeID) OR strEmployeeID= "" Then
'	strEmployeeID = ""
'Else
'	strEmployeeID = " WHERE " & Right(strEmployeeID,Len(strEmployeeID) - 5)
'End If
'response.write strEmployeeID
Public Sub ShowDetails()
'response.write "SELECT Top 10 [user_name],[employee_id] FROM procharge_user WITH(NoLock) WHERE [active_indicator ] = 'Y' " & strEmployeeID & "" 
'If there is no connection to ProMaster (Card Management System) then do not try to use the connection

'response.write objCon2.State
	If objCon2.State = 1 Then
		'Open a recordset in the ProMaster (CMS) database to check the Employee has a CMS Account
		objRS.Open "SELECT Top 10 [user_name],[employee_id],[first_name],[surname] FROM procharge_user WITH(NoLock) WHERE [active_indicator ] = 'Y' " & strEmployeeID & "",objCon2
		
 'Description:	Loads Position details into page if applicable.
	'objRS.Open "SELECT * FROM tblCAPSCDMCHistory " & strEmployeeID,objCon
	'objRS.Open "SELECT * FROM tblCDMC WHERE EmployeeID = '" & Session("EmployeeID") & "'" & strFName,objCon

		'Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">" & _
        '    "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th>CMS Account</th><th>EmployeeID</th></tr>"
			  
		If Not objRS.EOF Then
		   
			Response.write "CMS Account in <span class=""badge badge-pill badge-primary"">ProMaster</span> for EmployeeID: <span style=""font-weight:bold;"">" & Replace(Request.QueryString("EmpID"), "'", "") & " - " & objRS("first_name") & " " & objRS("surname") & "</span></br></br>"
			
			Response.write "<div class=""table-responsive"">" & _
				"<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th style=""padding-top:5px; padding-bottom:5px;"">CMS Account</th><th style=""padding-top:5px; padding-bottom:5px;"">EmployeeID</th></tr>"
			'Response.write "<div class=""card mb-3""><div class=""card-body"" style=""padding:5px; margin:5px;""><div class=""table-responsive"">" & _
				'"<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th>CMS Account</th><th>EmployeeID</th></tr>"
			
			Do Until objRS.EOF
				x = x + 1
				
				If x < 10 Then
				Response.write "<tr onClick=""SelectEmp('" & objRS("user_name") & "');""><td>" & objRS("user_name") & "</td><td>" & objRS("employee_id") & "</td></tr>"
				End If
			
			objRS.Movenext
			
			Loop

		Else
			Response.write "<span class=""badge badge-pill badge-danger"">No CMS Account in ProMaster for EmployeeID: " & Request.QueryString("EmpID") & "</span></br></br>"
	   End If

	objRS.Close
	
	Response.write "</table></div>"
	'Response.write "</table></div></div></div>"

	End If
		
End Sub


Public Sub ShowCardDetails()
'Write the CMS Details form the Card table for the Card selected so the user can view and edit
Dim strCompany
Dim strSQL
Dim objRS3
Dim strSelected
Dim strCostCentreNumber
Dim strDefaultCostCentre

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")

Set objRS3 = Server.CreateObject("ADODB.Recordset")

    objCon.Open Session("DBConnection")
		
		'Open a recordset to get the Card CMS details
		objRS2.Open "SELECT [CMSUser],[DefaultCompany],[DefaultCostCentre],[PMLoadStatus],[PMLoadDate],[ReportGroup],[CardNumberShort] FROM tblCAPSCard WITH(NoLock) WHERE [CardID] = " & lngCardID & "",objCon

		'Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">" & _
        '    "<table class=""table table-bordered table-hover header-fixed"" id=""dataTable"" width=""100%"" cellspacing=""0""><thead><tr><th>CMS Account</th><th>EmployeeID</th></tr>"
		
		
		Response.write "<span style=""font-weight:bold;"">Card CMS Details in <span class=""badge badge-pill badge-success"">CAPS</span> for Card: " & objRS2("CardNumberShort") & "</span></br></br>"
		
		If Not objRS2.EOF Then
		   
			Response.write "<div class=""card mb-3""><div class=""card-body""><div class=""table-responsive"">" & _
            "<table class=""table table-bordered table-hover header-fixed"" id=""CardCMSDataTable"" width=""100%"" cellspacing=""0"">"
			
			'Write the Header for each column
			Response.write "<tr><th style=""padding-top:5px; padding-bottom:5px;""></th><th style=""padding-top:5px; padding-bottom:5px;"">Currently in CAPS</th><th style=""background-color:#e6eeff; padding-top:5px; padding-bottom:5px;"">Change To <input type=""HIDDEN"" id=""CardIDCMS"" name=""CardIDCMS"" value=""" & lngCardID & """></th></tr>"
			
			
			'Response.write "<tr onClick=""SelectEmp('" & objRS("user_name") & "');""><td>" & objRS("user_name") & "</td><td>" & objRS("employee_id") & "</td></tr>"
			Response.write "<tr><th>CMS User (CAPS)</th><td>" & objRS2("CMSUser") & "</td><td style=""background-color:#e6eeff;""><input type=""text"" id=""CMSUser"" name=""CMSUser"" value=""" & objRS2("CMSUser") & """></td></tr>" & _
				"<tr><th>Default Company</th><td>" & objRS2("DefaultCompany") & "</td>"
				
				'Insert the Company Code drop down	
			strSQL = "SELECT CompanyCode FROM tblCAPSCostCentre WITH(NOLOCK) Group By CompanyCode ORDER BY CompanyCode"
			
			objRS3.Open strSQL, objCon
			
				If objRS3.EOF Then
					Response.Write "<td style=""background-color:#e6eeff;""><input type=""text"" id=""DefaultCompanyT"" name=""DefaultCompanyT"" value=""" & objRS2("DefaultCompany") & """></td></tr>"
				Else
				
					Do Until objRS3.EOF

						If objRS3("CompanyCode") = objRS2("DefaultCompany") Then
							strSelected = " SELECTED "
						Else
							strSelected = ""
						End If
						
						strCompany = strCompany & "<OPTION Value=""" & objRS3("CompanyCode") & """ " & strSelected & ">" & objRS3("CompanyCode") & "</OPTION>"
						
					objRS3.Movenext
					Loop
				   
					Response.Write "<td><SELECT id=""DefaultCompany"" name=""DefaultCompany"" class=""form-control"">" & strCompany & "</SELECT></tr>"
						
				End If
			
			objRS3.Close
				
				
			Response.write"<tr><th>Default Cost Centre</th><td>" & objRS2("DefaultCostCentre") & "</td>"
			
			'Clear the variable for use below
			strCompany = ""
			
			If IsNull(objRS2("DefaultCostCentre")) Then
				strCostCentreNumber = ""
			Else
				strCostCentreNumber = trim(objRS2("DefaultCostCentre"))
			End If			
			
			
			'Insert the Cost Centre drop down	
			strSQL = "SELECT CostCentreNumber FROM tblCAPSCostCentre WITH(NOLOCK) Group By CostCentreNumber ORDER BY CostCentreNumber"
			
			objRS3.Open strSQL, objCon
			
				If objRS3.EOF Then
					Response.Write "<td style=""background-color:#e6eeff;""><input type=""text"" id=""DefaultCostCentreT"" name=""DefaultCostCentreT"" value=""" & objRS2("DefaultCostCentre") & """></td></tr>"
				Else
				
					Do Until objRS3.EOF

						If IsNull(objRS3("CostCentreNumber")) Then
							strDefaultCostCentre = ""
						Else
							strDefaultCostCentre = trim(objRS3("CostCentreNumber"))
						End If
			
						If strDefaultCostCentre = cstr(strCostCentreNumber) Then
							strSelected = " SELECTED "
						Else
							strSelected = ""
						End If
						
						strCompany = strCompany & "<OPTION Value=""" & objRS3("CostCentreNumber") & """ " & strSelected & ">" & objRS3("CostCentreNumber") & "</OPTION>"
						
					objRS3.Movenext
					Loop
				   
					Response.Write "<td><SELECT id=""DefaultCostCentre"" name=""DefaultCostCentre"" class=""form-control"">" & strCompany & "</SELECT></tr>"
						
				End If
			
			objRS3.Close
		
			Response.write "<tr><th>PM Load Status</th><td>" & objRS2("PMLoadStatus") & "</td><td style=""background-color:#e6eeff;""><input type=""text"" id=""PMLoadStatus"" name=""PMLoadStatus"" value=""" & objRS2("PMLoadStatus") & """></td></tr>" & _
				"<tr><th>PM Load Date</th><td>" & objRS2("PMLoadDate") & "</td><td style=""background-color:#e6eeff;""><input type=""text"" DISABLED id=""PMLoadDate"" name=""PMLoadDate"" value=""" & objRS2("PMLoadDate") & """></td></tr>" & _
				"<tr><th>Report Group</th><td>" & objRS2("ReportGroup") & "</td><td style=""background-color:#e6eeff;""><input type=""text"" id=""ReportGroup"" name=""ReportGroup"" value=""" & objRS2("ReportGroup") & """></td></tr>"
			
			
		Else
			Response.write "No CMS Details for Card: " & Request.QueryString("CardID")
	   End If

	objRS2.Close
	
	Response.write "</table></div></div></div>"

		
End Sub

'SET objRS = Nothing
'SET objRS2 = Nothing
%>
<head>
  
</head>

<body>
  <%
	'Only load the Employee Details if the account has an EmployeeID
	If strLoadEmployee = "N" Then
		Response.write "<span class=""badge badge-pill badge-info"">No EmployeeID to search for in CMS (OK for CTS Account)</span></br></br>"
	Else
		Call ShowDetails()
	End If
 Call ShowCardDetails()

  %>
</body>

</html>
