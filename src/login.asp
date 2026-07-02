<%@ Language=VBScript %>
<%  
    'File:			Default.asp
	'Written By:	MG
	'Written On:	November 2011
	'Edit History:	
	'Purpose:	    Application Login Screen.
	
	Response.Expires = -1500
	
	Dim objCon
	Dim objRS
	Dim objRS2
	Dim strMessage
    Dim strPassword
    Dim strUA 
    Dim strRegistered
    Dim strClientType
    Dim strName
	Dim arrUType(5)
	Dim strSelected
	
	arrUType(1) = "Manager"
	arrUType(2) = "Employee"
	arrUType(3) = "CreditCards"
	arrUType(4) = "Other"
	arrUType(5) = "FBT"
	
    'Get the User's browser to change some of the freeze pane screens and warn them about the best browser at the moment
    strUA = lcase(Request.ServerVariables("HTTP_USER_AGENT"))

    'By default the browser is set to all others (FF, Opera, Safari, etc..)
    Session("UBrowser") = "FF"
    'If instr(strUA,"msie") Then Session("UBrowser") = "IE"
    'MS IE Version 10 forces compatability so displays like all other browsers
    If instr(strUA,"msie 10") Then Session("UBrowser") = "FF"
   
	Set objCon = Server.CreateObject("ADODB.Connection")
	Set objRS = Server.CreateObject("ADODB.Recordset")
	Set objRS2 = Server.CreateObject("ADODB.Recordset")

	'Catch the Session Expired details
	If Request.QueryString("State") = "Expired" Then
	    strMessage = "<span style=""color:red; font-weight:bold;"">Your Session Expired. Please login again...</span>"
	End If
	
	'Set Database Connection
	Session("DBConnection") = "File Name=" & Server.MapPath("Database/IsidoreGOL.udl") & ";"
	
	'Session("DBConnection") = "Provider=SQLOLEDB.1;DRIVER=SQL Server;SERVER=SONY-BULAND;UID=sa;PASSWORD=8887;DATABASE=BERTv70;"

	objCon.Open Session("DBConnection")	    
			
	If Request.QueryString("Action") = "Login" Then
	    
	    If Request.Form("UType") = "Employee" OR Request.Form("UType") = "Manager" Then
	    objRS.Open "SELECT Top 1 * FROM qryUserLogon WHERE UserLogon = '" & Request.Form("username") & "'",objCon,3,1
		''objRS.Open "SELECT Top 1 * FROM qryUserLogon WHERE EmailAddress = '" & Session("Logon") & "' AND Password = '" & strPassword & "' AND Active = 'Y'",objCon,3,1
	    Else
			objRS.Open "SELECT Top 1 * FROM qryUserLogon WHERE EmailAddress = '" & Request.Form("username") & "'",objCon,3,1
		End If
	        If Not objRS.EOF Then
	        
    	       strMessage = "<FONT color=""green"">GREAT. </FONT><a href=""index.asp"" target=""_parent"">Logon Successful</a>"
			   Session("Logon") = objRS("UserLogon")
			   Session("BusinessAreaID") = objRS("DefaultBusinessAreaID")
			   Session("CostCentreID") = objRS("DefaultCostCentre")
			   Session("BudgetID") = objRS("BudgetID")
			   Session("UserID") = objRS("UserID")
			   Session("UserTypeID") = 1
			   strName = objRS("FName") & " " & objRS("LName")
			   Session("UserName") = objRS("FName") & " " & objRS("LName")
					'objRS2.Open "SELECT Top 1 * FROM tblStaffingClassifications WHERE BudgetID = " & Session("BudgetID") & " AND StaffClassificationDesc = '" & strName & "'",objCon,3,1
					'	If Not objRS2.EOF Then
					'		Session("StaffingClassificationID") = objRS2("StaffingClassificationID")
					'		Session("EmployeeID") = objRS2("StaffingClassificationID")
					'	Else
					'		Session("StaffingClassificationID") = 0
					'	End If
					'objRS2.close
					
					'Set the Home page screens based on the user type
					If IsEmpty(Request.Form("UType")) Then
						Session("HomePage1") = "Reports/PLPieChart.asp"
						Session("HomePage2") = "Admin/CostCentreStatusHome.asp"
					Else
						If Request.Form("UType") = "Employee" Then
							Session("UType") = "Employee"
							Session("HomePage1") = "Reports/FBTChart.asp"
							Session("HomePage2") = "Admin/BenefitHome.asp"
							Session("HomePageTop") = "CC/HomeCC2.asp"
							response.redirect "indexCC2.asp"
						ElseIf Request.Form("UType") = "Manager" Then
							Session("UType") = "Manager"
							Session("HomePage1") = "FBT/FBTChartAll.asp"
							Session("HomePage2") = "Admin/CostCentreStatusHome.asp"
							Session("HomePageTop") = "CC/HomeCC2.asp"
							response.redirect "indexCC2.asp"
						ElseIf Request.Form("UType") = "FBT" Then
							Session("UType") = "FBT"
							Session("HomePage1") = "FBT/FBTChartAll.asp"
							Session("HomePage2") = "Admin/CostCentreStatusHome.asp"
							Session("HomePageTop") = "CC/HomeCC.asp"
							Session("BudgetID")=1
							Session("BusinessAreaID") = 1
							Session("VersionID") = 1
							response.redirect "index.asp"
							
						Else
							Session("UType") = "CreditCards"
							'Session("HomePage1") = "FBT/FBTChartAll.asp"
							'Session("HomePage2") = "Admin/CostCentreStatusHome.asp"
							Session("HomePageTop") = "CC/HomeCC.asp"
							response.redirect "indexCC2.asp"
						End If
					End If
					
					'These Session Variables must have values or errors will occur.
					Session("Level1ID") = 1
					Session("Level2ID") = 1
					Session("GFSCodeID") = 0
					Session("FundTypeID") = 0
					Session("FundSourceID") = 0
					Session("RecordID") = 0
					Session("Header") = "BERTFrameset.asp"
					Session("HeaderMenu") = "Header1.asp"	
					Session("UnitSaleItemID") = 1
					Session("TransferDocumentTypeID") = 0
					Session("InputSheetID") = 1

			   response.redirect "index.asp"
	        Else
				Session("BudgetID") = 1
    	         strMessage = "<FONT color=red>Username or password incorrect."
	             Session("UType") = "CreditCards"
				'Session("HomePage1") = "FBT/FBTChartAll.asp"
				'Session("HomePage2") = "Admin/CostCentreStatusHome.asp"
				'response.redirect "indexCC2.asp"
	        End If
	    
	    
	    objRS.Close
		
		If isnull(Session("ClientID")) or Session("ClientID") = "" Then
		    If strMessage = "" Then 
		        If isEmpty(Request.QueryString("WebUse")) Then
		        strMessage = "<FONT color=red>Username or Password is incorrect</FONT>"
		        End If
		    End If
		Else
		
	       
	        'Response.Write "Here:" & Session("ClientID")
	    End If
	    
	    'Response.Redirect "BudgetFrameset.asp"
	    
	End If
	
	If Not IsEmpty(Request.QueryString("UType")) Then
		If Request.QueryString("UType") = "Employee" Then
			strUserID = "8343022"
			
		End If
		
		If Request.QueryString("UType") = "Manager" Then
			strUserID = "8945455"
			
		End If
		
		strUType = Request.QueryString("UType")
	End If
	
%>

<html>
<head>

<script LANGUAGE="javascript">
	
	function Login()
{         
	var varSubmit = true						
    var varAlert="";	
          
   // if(isWhitespace(frm.username.value)==true)
   if(frm.username.value.length<1)
	{
	   varAlert += "Please enter a Username. \n \n";
       document.getElementById('username').style.backgroundColor="ff8080";       	   
	   varSubmit = false	   				
	}
	else document.getElementById('username').style.backgroundColor="ffffff";       
    
    //if(isWhitespace(frm.passwordbb.value)==true)
    //if(frm.passwordi.value.length<1)
	//{
	//   varAlert += "Please enter a Password. \n \n";
    //   document.getElementById('passwordi').style.backgroundColor="ff8080";       	   
	//   varSubmit = false	   				
	//}
	//else document.getElementById('passwordi').style.backgroundColor="ffffff";       
        	     		
	
	if(varSubmit == true)
	{	    
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}	
}
function NextScreen(){
    //alert('' + frm.ClientID.value + ''  + frm.UserID.value + '');
	//top.location="Verify.asp&dgj4ht=" + frm.ClientID.value + "&r83jnhh27s=" + frm.UserID.value + ""
	top.location="Verify.asp?dgj4ht=" + frm.ClientID.value + "&r83jnhh27s=" + frm.UserID.value + ""
}

function ChangeUser(){
    //alert('' + frm.ClientID.value + ''  + frm.UserID.value + '');
	//top.location="Verify.asp&dgj4ht=" + frm.ClientID.value + "&r83jnhh27s=" + frm.UserID.value + ""
	//top.location="Login.asp?dgj4ht=" + frm.ClientID.value + "&r83jnhh27s=" + frm.UserID.value + ""
	self.location="Login.asp?dgj4ht=" + frm.username.value + "&UType=" + frm.UType.value + ""
}


</script>
<title>Isidore</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="style/login_style.css" type="text/css" rel="stylesheet">
</head>
<body onLoad="document.login.email.focus()" bgcolor="#0072BC" topmargin="0" rightmargin="0" bottommargin="0" leftmargin="0">
<form action="login.asp?Action=Login" method="POST" id="frm" name="frm">
<div id="login_wrapper" align="center">

	<div id="login_box" style="width:456px; padding-top:205px; border:0px solid #ff0000;">
	
		<div id="login_top">&nbsp;</div>
		
		<div id="login_fields" style="height:136px; width:456px; text-align:left; background-image:url(images/login_box.gif); background-repeat:no-repeat; border:0px solid #ff0000;">
		
			<div id="login_fields_info" style="padding:20px 40px 30px 130px; font-family:Arial, Helvetica, sans-serif; font-size:11px">
				<form action="Index.asp?ClientID=1" method="POST" id="frm" name="frm" style="margin:0 0 0 0; style:0 0 0 0">
				<table border="0" cellpadding="2" cellspacing="0" width="100%">
				<tr>
					<td style="font-family:Arial, Helvetica, sans-serif; font-size:11px; font-weight:bold; width:30%;"></td>
					<td style="height:18px; width:100%;"><SELECT name="UType" id="UType" onChange="ChangeUser();">
					<%
						For x = 1 to 5
						
							If strUType = arrUType(x) Then
								strSelected = " SELECTED "
							Else
								strSelected = ""
							End If
							
							Response.Write "<option " & strSelected & " value=""" & arrUType(x) & """>" & arrUType(x) & "</option>"
						
						Next
					%>
					<!--<option value="Manager">Manager</option><option value="Employee">Employee</option><option value="CreditCards">CreditCards</option></select></td>-->
				</tr>
				<tr>
					<td style="font-family:Arial, Helvetica, sans-serif; font-size:11px; font-weight:bold; width:30%;">User ID:</td>
					<td style="height:18px; width:100%;"><input type="text" name="username" id="username" value="<%=strUserID%>" style="height:18px; width:100%;"></td>
				</tr>
				<tr>
					<td style="font-family:Arial, Helvetica, sans-serif; font-size:11px; font-weight:bold">Password:</td>
					<td><input type="password" name="passwordi" id="passwordi" value="" style="height:18px; width:100%;"></td>
				</tr>
				<tr>
					<td style="padding-top:10px; font-family:Arial, Helvetica, sans-serif; font-size:10px; color:#0072BC; font-weight:bold;"></td>
					<!--<td style="padding-top:10px; font-family:Arial, Helvetica, sans-serif; font-size:10px; color:#0072BC; font-weight:bold;">Forgot Password?</td>-->
					<td align="right"><div id="row_1_button_main_menu" style="height:22px; width:65px; margin-top:5px; border-right:1px solid #aebfd3; margin-right:11px;"><input name="main_menu_button" class="row_1_btn" type="button" onClick="Login();" value="Login" id="main_menu_button" /></div></td>
				</tr>
				<tr>
					
					<td align="right" colspan="2"><%=strMessage%></td>
				</tr>
				</table>
				</form>
			</div>
		
		</div>
		
		<div id="login_bottom">&nbsp;</div>
	
	</div>

</div>
</form>
</body>
</html>