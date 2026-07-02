<%@ Language=VBScript %>
<!-- #Include file=ADOVBS.inc -->
<%  
    'File:			Default.asp --- CAPS DPESIT WITH Manual Login
	'Written By:	MG
	'Written On:	April 2022
	'Edit History:	
	'Purpose:	    Application Login Screen.
	
	'''New for the COA Search tool --- 28th March 2025
	'''Check to see if the user has come from teh COA Search tool URL http://COASearch
	If Instr(1,Request.ServerVariables("All_Http"),"COASearch")>0 Then Response.Redirect "COASearch.asp"
	
	''''New Apr 2026 for DCCP redirect
	If Right(Request.ServerVariables("All_http"),10)="creditcard" Then Response.Redirect "ccportal/index.php?route=auth/login-agreement" 

	If Right(Request.ServerVariables("All_http"),11)="creditcard/" Then Response.Redirect "ccportal/index.php?route=auth/login-agreement" 

	If Instr(1,Request.ServerVariables("All_Http"),"creditcard/")>0 Then Response.Redirect "ccportal/index.php?route=auth/login-agreement"
	

	Response.Write "<div class=""content-body"" style=""position:absolute; top:10; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span><a href=""ccportal/index.php?route=auth/login-agreement"">Defence Credit Card Portal</a><br>If not redirected to the Defence Credit Card Portal, click above.</span></div></div></div>"


	Response.Expires = -1500
	
	Dim objCon
	Dim objRS
	Dim strMessage
    Dim strPassword
    Dim strUA 
    Dim strRegistered
    Dim strClientType
    Dim strName
	Dim arrUType(12,1)
	Dim strSelected
	Dim strLogin
	Dim strLogButton
	Dim strLogEnter
	Dim strLoginJava
	Dim x
	Dim strUTypeForm
	Dim strCAPSEmailFrom 
	Dim intError
	Dim strEmployeeID
	
	Dim strLastLoggedScreen
	
	Dim lngUserCount
	Dim strUsers
	Dim strU
	Dim intNewUser
	
	'arrUType(1) = "Manager"
	'arrUType(2) = "Employee"
	'arrUType(3) = "Admin"
	'arrUType(4) = "Other"
	
	intError = 0
	intNewUser=1
	
    'Get the User's browser to change some of the freeze pane screens and warn them about the best browser at the moment
    strUA = lcase(Request.ServerVariables("HTTP_USER_AGENT"))

	strLogin = Request.ServerVariables("Auth_User")
		
	'Cahnge above to belkow for UAT in DOD
	'strLogin = Request.Form("usernameconfirm")

    'By default the browser is set to all others (FF, Opera, Safari, etc..)
    Session("UBrowser") = "FF"
    'If instr(strUA,"msie") Then Session("UBrowser") = "IE"
    'MS IE Version 10 forces compatability so displays like all other browsers
    If instr(strUA,"msie 10") Then Session("UBrowser") = "FF"
   
	'If the browser is Internet Explorer 11 then use the horizontal menu (Bootstrap 3), otherwise vertical (Bootstrap 4)
	If instr(strUA,"trident") Then 
		Session("Menu") = "indexCC4.asp"
	Else
		Session("Menu") = "CAPSMenu.asp"
	End If
	
	Set objCon = Server.CreateObject("ADODB.Connection")
	Set objRS = Server.CreateObject("ADODB.Recordset")

	'Catch the Session Expired details
	If Request.QueryString("State") = "Expired" Then
	    strMessage = "<span style=""color:red; font-weight:bold;"">Your Session Expired. Please login again...</span>"
	End If
	
	'Set Database Connection
	Session("DBConnection") = "File Name=" & Server.MapPath("Database/CAPS.udl") & ";"
	
	objCon.Open Session("DBConnection")	    
			'response.write "<span style=""color:black;"">UserType = " & Request.Form("UType") & " Log: " & strLogin & "</span>"
		
		'Open a recordset to get a list of all the USer Types for the login drop-down
		'This is temporary for testing purposes
		objRS.Open "SELECT * FROM tblUserTypes WITH(NOLOCK) WHERE UserTypeID < 100",objCon,3,1
		
			If objRS.Eof Then
			
			Else
				Do Until objRS.EOF
				
					x = x + 1
					If x > 7 Then
						
					Else
						arrUType(x,0) = objRS("UserTypeID")
						arrUType(x,1) = objRS("UserTypeName")
					End If
					
				objRS.Movenext
				Loop
			
			End If
		objRS.Close
	
	strUTypeForm = cint(Request.Form("UType"))
	'strUTypeForm = Session("UserTypeID")

	'Get the UserID details if the user is from a network login as it isn't checked until they try to login when they won;t have a usertype yet, so will be logged in as the first entry in the list
	If isNull(strLogin) or strLogin = "" Then
	Else
		If Len(strLogin) >4 Then
			'If Left(strLogin,4) = "DRN\" Then strLogin = right(strLogin,len(strLogin)-4)
			
			'Above Replaced with dynamic removal of the Domain -- March 2022
			If InStr(1,strLogin,"\")>0 Then
				strLogin = Right(strLogin,Len(strLogin)-InStr(1,strLogin,"\"))
			End If
		End If
			
		objRS.Open "SELECT Top 1 * FROM qryUserLogon WITH(NOLOCK) WHERE UserLogon = '" & strLogin & "'",objCon,3,1
			If Not objRS.EOF Then
				strUTypeForm = objRS("UserTypeID")
				Session("UserTypeID") = objRS("UserTypeID")
			Else
			End If
		objRS.Close
	
	End If
	
	If Request.QueryString("Action") = "Login" Then
		
	    If strUTypeForm < 100 Then
		'If Request.Form("UType") = "Employee" OR Request.Form("UType") = "Manager" OR Request.Form("UType") = "Admin" Then
			strLogEnter = Request.Form("username")
			
			If Len(strLogEnter) >4 Then
				'If Left(strLogEnter,4) = "DRN\" Then strLogEnter = right(Request.Form("username"),len(Request.Form("username"))-4)
				
				'Above Replaced with dynamic removal of the Domain -- March 2022
				If InStr(1,strLogEnter,"\")>0 Then
					strLogEnter = Right(Request.Form("username"),Len(Request.Form("username"))-InStr(1,Request.Form("username"),"\"))
				End If
			
			End If
			
			'Added for DEV please check whether relevant for SIT/PROD as this uses the typed user
			'*********************
			If isNull(strLogEnter) or strLogEnter= "" Then strLogEnter = strLogin
			
		'response.write " strLogEnter=" & strLogEnter	& "| strLogin " & strLogin
			objRS.Open "SELECT Top 1 * FROM qryPORTALUserLogon WITH(NOLOCK) WHERE UserLogon = '" & strLogEnter & "'",objCon,3,1

	    Else
			objRS.Open "SELECT Top 1 * FROM qryPORTALUserLogon WITH(NOLOCK) WHERE EmailAddress = '" & Request.Form("username") & "'",objCon,3,1
		End If
		
	        If Not objRS.EOF Then			
			
				'First flag the user as having logged in to the Portal if they have not before
				'If IsNull(objRS("LoggedOntoPortal")) OR objRS("LoggedOntoPortal")="" Then
					'Update the User in the User table by calling the procedure for when they first login
				'	Call UpdateUserLogin(strLogEnter)
				'End If
				
    	       strMessage = "<FONT color=""green"">GREAT. </FONT><a href=""index.asp"" target=""_parent"">Logon Successful " & Request.Form("UType") & "</a>"
			   Session("Logon") = objRS("UserLogon")
			   
			   Session("EmployeeID") = objRS("EmployeeID")
			   		
			   'Session("BusinessAreaID") = objRS("DefaultBusinessAreaID")
			   'Session("CostCentreID") = objRS("DefaultCostCentre")
			   'Session("BudgetID") = objRS("BudgetID")
			   Session("UserID") = objRS("UserID")
			   
			   'Set the Global.asa values for the Users Logged in
			   If IsEmpty(Application("NamedUsers")) Then
					Application("NamedUsers") = Session("UserID")
				Else
					
					'''Add the userID to the application variable if they are not in there already
					lngUserCount = 0
					strUsers =  Split(Application("NamedUsers"),",")
					For each strU in strUsers
						If clng(Session("UserID")) = clng(strU) Then lngUserCount = 1
					Next
										
					If lngUserCount= 0 Then Application("NamedUsers") = Application("NamedUsers") & "," & Session("UserID")
			   End If
			   
			   'Set the UserID for the Authority page to the main user so they can only access their own authority details.
			   'This will never change, but the UserID can change if they login as someone else
				Session("UserIDAuthority") = objRS("UserID")
				'Session("UserIDAuthority") = objRS("EmployeeID")
				
			   Session("UserTypeID") = objRS("UserTypeID")
			   strName = objRS("FName") & " " & objRS("LName")
			   Session("UserName") = objRS("FName") & " " & objRS("LName")
			
				Session("BudgetID") = 1		
				Session("BusinessAreaID") = 1
				Session("CostCentreID") = 1
				Session("VersionID") = 1
			   
				Session("ServerPath") = GetSystemAdminDefault("ServerPath")
				If IsNull(Session("ServerPath")) or Session("ServerPath") = "" Then
					'Session("ServerPath") = "HTTP://vbmrsn05/CAPS/"
					intError = 0
					strCAPSEmailFrom = GetSystemAdminDefault("CAPSEmailAddressFrom")
					Response.Write "<div class=""container"" style=""position:relative; z-index:100; top:40px; left:40px;""><div class=""alert alert-danger"" role=""alert"" style=""position: absolute; top:40px; left:40px; z-index:100;"">Error! Server path not found. See System Admin: " & strCAPSEmailFrom & "</div></div>"
				
				End If
				
				'New update to make sure there is a UserTypeID
				If IsNull(strUTypeForm) or strUTypeForm = "" Then strUTypeForm = Session("UserTypeID")
				
				'Set the Home page screens based on the user type
				If IsEmpty(strUTypeForm) Then
					Session("HomePage1") = "Reports/PLPieChart.asp"
					Session("HomePage2") = "Admin/CostCentreStatusHome.asp"
				Else
					If strUTypeForm = 1 Then
					'If Request.Form("UType") = "Employee" Then
						Session("UType") = "Employee"
						Session("HomePage1") = "Reports/FBTChart.asp"
						Session("HomePage2") = "Admin/BenefitHome.asp"
						Session("HomePageTop") = "CC/HomeCC2.asp"
						'New page format -- Session("Menu") is the top page with includes rather than iframe
						Session("Menu") = Session("ServerPath") & "CC/HomeUser.asp" 'Server.MapPath("CC/HomeUser.asp")
						
						Session("BusinessArea") = "CAPS"
						
						Application.Lock
						Application("GeneralUsers") = Application("GeneralUsers") + 1
						Application.Unlock
						
						Response.Redirect Session("Menu")'"indexCC4.asp"
					ElseIf strUTypeForm = 2 Then
					'ElseIf Request.Form("UType") = "Manager" Then
						Session("UType") = "Manager"
						Session("HomePage1") = "FBT/FBTChartAll.asp"
						Session("HomePage2") = "Admin/CostCentreStatusHome.asp"
						Session("HomePageTop") = "CC/HomeCC2.asp"
						
						Session("BusinessArea") = "CAPS"
						
						Application.Lock
						Application("ManagerUsers") = Application("ManagerUsers") + 1
						Application.Unlock
						
						'Only redirect if there are no errors otherwise stay on this page
						If intError = 0 Then Response.Redirect Session("Menu")'"indexCC4.asp"
						
					ElseIf strUTypeForm = 4 Then
					'If Request.Form("UType") = "Employee" Then
						Session("UType") = "Employee"
						Session("HomePage1") = "Reports/FBTChart.asp"
						Session("HomePage2") = "Admin/BenefitHome.asp"
						Session("HomePageTop") = "CC/HomeCC2.asp"
						'New page format -- Session("Menu") is the top page with includes rather than iframe
						Session("Menu") = Session("ServerPath") & "CC/HomeUser.asp" 'Server.MapPath("CC/HomeUser.asp")
						
						Session("BusinessArea") = "Compliance"
						
						Application.Lock
						Application("GeneralUsers") = Application("GeneralUsers") + 1
						Application.Unlock
						
						'Only redirect if there are no errors otherwise stay on this page
						If intError = 0 Then Response.Redirect Session("Menu")'"indexCC4.asp"
						
					'PortalUser and PORTALAdmin
					ElseIf strUTypeForm = 5 OR strUTypeForm = 9 Then
						
						'Response.Write "Session(UserIDAuthority) =" & Session("UserIDAuthority") 
						strEmployeeID = objRS("EmployeeID")
						
						Session("UserIDAuthority") = objRS("UserID")

						'First flag the user as having logged in to the Portal if they have not before
						If IsNull(objRS("LoggedOntoPortal")) OR objRS("LoggedOntoPortal")="" Then
							'Update the User in the User table by calling the procedure for when they first login
							'Response.Write "llooogger=" & objRS("LoggedOntoPortal")
							Call UpdateUserLogin(strLogEnter,strEmployeeID)
							
							strLastLoggedScreen = "Portal/CAPSPortal.asp?1=" & objRS("TwoFAID") & ""
						Else
							'If the user has logged in before then redirect them to the page they were last on (from database)
							If IsNull(objRS("LastScreen")) or objRS("LastScreen") = "" Then
								strLastLoggedScreen = "Portal/CAPSPortal.asp?1=" & objRS("TwoFAID") & ""
							Else
								strLastLoggedScreen = "Portal/" & objRS("LastScreen") & "?1=" & objRS("TwoFAID") & ""
							End If

							'Record the user logging in to the PORTAL
							Call AddUserLoginRecord(strLogEnter)

						End If

						Session("UType") = "Portal"
						'Check to see if the user has Complete (Confirmed or Cancelled) their request and if so send them to the final screens (overwrite the last logged in)
						If Not IsNull(objRS("RolloutNewCardRequestedDate")) Then
							strLastLoggedScreen = "Portal/CAPSPortal5.asp"
						End If
						
						If Not IsNull(objRS("RolloutNewCardCancelledDate")) Then
							strLastLoggedScreen = "Portal/CAPSPortalCancel.asp"
						End If
						'Session("Menu") = Session("ServerPath") & "Portal/CAPSPortal.asp"
						'Session("ServerPath") = Session("ServerPath") & "Portal/"
						'Session("ServerPath") = "http://siepd85wsd0017.dpedev.protecteddev.mil.au/"
						'Response.write "<BR>2=" & Session("ServerPath")
						'Response.Redirect "Portal/CAPSPortal.asp?ApplicationEmployeeID=" & objRS("EmployeeID") & ""
						
						'Close objects before redirecting
						objRS.Close
						objCon.close

						Set objRS = Nothing
						Set objCon = Nothing

						Response.Redirect strLastLoggedScreen
						
						'Response.Redirect "Portal/" & strLastLoggedScreen & "?ApplicationEmployeeID=" & objRS("EmployeeID") & ""
						'Response.Redirect "Portal/CAPSPortal.asp?ApplicationEmployeeID=" & objRS("EmployeeID") & ""
						'Below if the Two Factor Authentication page if this is enabled and an email is to be sent
						'Response.Redirect "Portal/CAPSPortalFA.asp?ApplicationEmployeeID=" & objRS("EmployeeID") & ""
						
					Else
						Session("UType") = "Admin"
						'Session("HomePage1") = "FBT/FBTChartAll.asp"
						'Session("HomePage2") = "Admin/CostCentreStatusHome.asp"
						Session("HomePageTop") = "CC/HomeCC.asp"
						Session("Menu") = Session("ServerPath") & "CC/HomeAdmin.asp"
						
						'UPDATED 10/Nov/2020 ---- Changed to the new Credit Card Dashboard
						Session("Menu") = Session("ServerPath") & "CC/HomeCCAdmin.asp"
						
						Session("BusinessArea") = "CAPS"
						
						'Set the Application variable depending on the Admin User Type
						If Session("UserTypeID") = 10 Then
							Application.Lock
							Application("AdminUsers") = Application("AdminUsers") + 1
							Application.Unlock
						End If
						
						If Session("UserTypeID") = 11 Then
							Application.Lock
							Application("SuperAdminUsers") = Application("SuperAdminUsers") + 1
							Application.Unlock
						End If
						
						If Session("UserTypeID") = 99 Then
							Application.Lock
							Application("IsidoreAdminUsers") = Application("IsidoreAdminUsers") + 1
							Application.Unlock
						End If
						
						'Only redirect if there are no errors otherwise stay on this page
						If intError = 0 Then Response.Redirect Session("Menu")'"indexCC4.asp"
						
					End If
				End If
					
			   'Response.Redirect "index.asp"
	        Else
			
				'If there is no user logon record in qryPORTALUserLogin then add the user to the table tblUsers
				'from the CDMC table if they have a card
				'If IsNull(objRS("LoggedOntoPortal")) OR objRS("LoggedOntoPortal")="" Then
					'Update the User in the User table by calling the procedure for when they first login
					Call AddUserLogin(strLogEnter)
				'End If
						
				Session("BudgetID") = 1
				strMessage = "<FONT color=""red"">Access denied. " & strUTypeForm & ""
				Session("UType") = "CreditCards"

				'Do not log an error if the UserName is null/empty
				If IsNull(strLogEnter) Or strLogEnter = "" Then
				Else
					objCon.Execute "INSERT INTO tblErrorLog VALUES('CAPS/Default.asp','Login Failed:" & strLogEnter & " - No Active=Y in tblUsers',GetDate(),0)"
				End If

				'Set the variable to identify the user is new and should have been added to the table tblUser
				intNewUser=1
				
				'Session("HomePage1") = "FBT/FBTChartAll.asp"
				'Session("HomePage2") = "Admin/CostCentreStatusHome.asp"
				'response.redirect "indexCC2.asp"
	        End If
	    
	    
	    objRS.Close
		
		'Call the procedure to log the User is, if they have just been added as a new user
		If intNewUser=1 Then
			'This si the procedure which will log the user in and redirect them if they have a valid user record
			LoginUserPortal(strLogEnter)
		End If
		
		
		If isnull(strUTypeForm) or strUTypeForm = "" Then
		    If strMessage = "" Then 
		        If isEmpty(Request.QueryString("WebUse")) Then
		        strMessage = "<FONT color=""red"">Username or Password is incorrect</FONT>"
		        End If
		    End If
		Else
		
	       
	        'Response.Write "Here:" & Session("ClientID")
	    End If
	    
	    'Response.Redirect "BudgetFrameset.asp"
	    
	End If

	'If there is no User Network login then do not automate login
	If strLogin = "" or IsNull(strLogin) Then
		strLogButton = "<i class=""fa fa-sign-in""></i> Login"
		strLoginJava = ""
	Else
		If strMessage = "" or IsNull(strMessage) Then
			strLogButton = "<i class=""fa fa-spinner fa-spin""></i> Logging you in..."
			
			'Only set a delay in logging in if the user has logged off from using the application
			If Request.QueryString("Logoff") = "Y" Then
				strLoginJava = "onLoad=""setTimeout(Login, 5000);"""
			Else
				strLoginJava = "onLoad=""setTimeout(Login, 500);"""
			End If
		Else
			strLogButton = "<i class=""fa fa-sign-in""></i> Login"
			strLoginJava = ""
		End If
	End If
	
	''''New for the DCCP -- do not delay the logging in as per the statment above for CAPS
	If Instr(1,Request.ServerVariables("All_Http"),"creditcard/")>0 Then strLoginJava = "onLoad=""setTimeout(Login, 1);"""

	'Session("DPCProvider") = GetSystemAdminDefault("DPCProvider")

Public Function GetSystemAdminDefault(strParameterName)

Dim objRSFunc

Set objRSFunc = Server.CreateObject("ADODB.Recordset")

	'GetSystemAdmin = "See System Admin"
	'Exit Function
	objRSFunc.Open "SELECT * FROM tblCAPSSystemParameters WHERE ParameterName = '" & strParameterName & "'",objCon

		If objRSFunc.EOF Then
			GetSystemAdminDefault = "" '& strParameterName 
		Else
			GetSystemAdminDefault = objRSFunc(3)
		End If

	objRSFunc.Close
 
Set objRSFunc = Nothing
 
End Function


%>

<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
  <meta name="description" content="">
  <meta name="author" content="">
  <title>Defence Credit Card Applications</title>
  <link rel="icon" type="image/jpg" href="favicon.ico"/>
  <!-- Bootstrap core CSS-->
  <link href="vendor/bootstrap/css/bootstrap.min.css" rel="stylesheet">
  <!-- Custom fonts for this template-->
  <link href="vendor/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css">
  <!-- Custom styles for this template-->
  <link href="css/sb-admin.css" rel="stylesheet">
  
  <script LANGUAGE="javascript">
	
	function Login()
{         

	var varSubmit = true						
    var varAlert="";	
           
   // if(isWhitespace(frm.usernameconfirm.value)==true)
   if(frm.usernameconfirm.value.length<1)
	{
	   varAlert += "Please enter a Username. \n \n";
       document.getElementById('usernameconfirm').style.backgroundColor="ff8080";       	   
	   varSubmit = false	   				
	}
	else document.getElementById('usernameconfirm').style.backgroundColor="ffffff";       
   
    //if(isWhitespace(frm.passwordbb.value)==true)
    //if(frm.passwordbb.value.length<1)
	//{
	//   varAlert += "Please enter a Password. \n \n";
    //   document.getElementById('passwordbb').style.backgroundColor="ff8080";       	   
	//   varSubmit = false	   				
	//}
	//else document.getElementById('passwordbb').style.backgroundColor="ffffff";       
        	     		
	if(varSubmit == true)
	{	    
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}	
}

function ChangeUser(){
    //alert('' + frm.ClientID.value + ''  + frm.UserID.value + '');
	//top.location="Verify.asp&dgj4ht=" + frm.ClientID.value + "&r83jnhh27s=" + frm.UserID.value + ""
	//top.location="Login.asp?dgj4ht=" + frm.ClientID.value + "&r83jnhh27s=" + frm.UserID.value + ""
	self.location="Default.asp?dgj4ht=" + frm.username.value + "&UType=" + frm.UType.value + ""
}

</script>

<style>

.fullscreen_bg {
    position: fixed;
    top: 0;
    right: 0;
    bottom: 0;
    left: 0;
    background-size: cover;
    background-position: 50% 50%;
    background-image: url('img/VG_Cardtop.jpg');
    background-repeat:repeat;
  }
 .panel-default {
opacity: 1.9;
margin-top:30px;
}
.form-group.last { margin-bottom:0px; }

/* Style buttons */Fu
.buttonload {
  background-color: #4CAF50; /* Green background */
  border: none; /* Remove borders */
  color: white; /* White text */
  padding: 12px 16px; /* Some padding */
  font-size: 16px /* Set a font size */
}
  </style>
</head>

<body class="bg-dark" <%=strLoginJava%> >
<div id="fullscreen_bg" class="fullscreen_bg"/>
  <div class="container">
    <div class="card card-login mx-auto mt-5">
      <div class="card-header"><img src="images/defence_logo_dark.png" height="44px;" width="124px;"> &nbsp;Defence CAPS PORTAL Login</div>
      <div class="card-body">
        <form action="Default.asp?Action=Login" method="POST" id="frm" name="frm">
		 <div class="form-group">
            <label for="exampleInputEmail1">User Type</label>
			<SELECT class="form-control" READONLY name="UType" id="UType" >
					<%
						For x = 1 to 7
						
							'If cint(strUTypeForm) = cint(arrUType(x,0)) Then
							If cint(arrUType(x,0)) = cint(Session("UserTypeID")) Then
								strSelected = " SELECTED "
							Else
								strSelected = ""
							End If
							
							Response.Write "<option " & strSelected & " value=""" & arrUType(x,0) & """>" & arrUType(x,1) & "</option>"
						
						Next
					%>
					</select>
           
          </div>
          <div class="form-group">
            <label for="exampleInputEmail1">DRN Login</label>
            <input class="form-control" id="username" name="username" HIDDEN type="input" value="<%=strLogin%>" placeholder="Login">
			<input class="form-control" id="usernameconfirm" name="usernameconfirm" type="input" style="font-weight:bold;" DISABLED value="<%=strLogin%>" placeholder="Login">
          </div>
        <!--  <div class="form-group">
            <label for="exampleInputPassword1">Password (not validated)</label>
            <input class="form-control" id="passwordbb" name="passwordbb" type="password" placeholder="Password" onkeydown="if (event.keyCode == 13) document.getElementById('login').click()">
          </div>-->
		  <%=strMessage %>
          <!--<input type="button" name="login" id="login" onClick="Login()" value="Login"/>-->
          <a class="btn btn-primary btn-block" Style="background-color:#2394F2;" href="#" name="login" id="login" onClick="Login()"><%=strLogButton%></a>
  
        <div class="text-center">
          <a class="d-block small mt-3" href="Register.asp">No Login? Click here</a>
          <!--<a class="d-block small" href="ForgotPassword.asp">Forgot Password?</a>-->
        </div>
		</form>
      </div>
    </div>
  </div>
  </div>
  <!-- Bootstrap core JavaScript-->
  <script src="vendor/jquery/jquery.min.js"></script>
  <script src="vendor/bootstrap/js/bootstrap.bundle.min.js"></script>
  <!-- Core plugin JavaScript-->
  <script src="vendor/jquery-easing/jquery.easing.min.js"></script>
</body>

</html>


<%

Function Send2FAEmail()
'Procedure to send the user logging in a two factor identification email

Dim strSQL
Dim strEmailTo
Dim strTwoFactorID

	strSQL = "SELECT * FROM tblPORTALCards WHERE EmployeeID='" & strEmpID & "'"
	
	'adLockReadOnly	1
	'adOpenForwardOnly	0
	
	objRS.CursorType = adOpenStatic
	objRS.LockType = adLockReadOnly
	
	objRS.Open strSQL,objCon',3,1
		
		If objRS.EOF Then
			strEmailTo = objRS("Email")
			strTwoFactorID = objRS("TwoFactorID")
		Else
		
		
		End If
		
	objRS.Close
	
	'Call the procedure in PORTALFunctions.asp to send the email
	'Send_Email(strFrom,strTo,strSubject,strBody,strAttachment,strEmailType)

'''''NOT USED anymore send 2FA functions in CAPSPortalFA.asp

End Function

Sub UpdateUserLogin(strLogin,strEmpID)
'Procedure to update the records fro the user logging into the portal to note their successful logging in
Dim intRecord
Dim objCmd
Dim lngUserIDLocal

	'Make sure there is a UserId - there may not be one if the user does not have a login
	If IsNull(Session("UserIDAuthority")) or IsEmpty(Session("UserIDAuthority")) Then 
		lngUserIDLocal = 0
	Else
		lngUserIDLocal = Session("UserIDAuthority")
	End If
	
'Make sure the UerIDAuthority is numeric
	If IsNumeric(lngUserIDLocal)=Flase Then
		lngUserIDLocal = 0
	End If
	
    Set objCmd = Server.CreateObject("ADODB.Command")

	If IsNull(strLogin) OR strLogin = "" Then
		'Response.Write "<div class=""content-body"" style=""position:absolute; top:100px; width:80%; z-index:300"">" & _
		'	"<div class=""alert alert-danger"" role=""alert"">Error: UserLogon " & strLogin & " is empty. Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></div></div>"
			
		Response.Write "<div class=""content-body"" style=""position:absolute; top:100; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>Error: UserLogon " & strLogin & " is empty. Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></span></div></div></div>"


		Exit Sub
	End If
	
	'On Error Resume Next
	
	With objCmd

		.CommandType = 4
		.CommandText = "spPORTALUserLogin"

		.Parameters.Append objCmd.CreateParameter("UserLogon", adVarChar, adParamInput, 50)	
		.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarChar, adParamInput, 50)		
		.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("LoginIDOutput", adInteger, adParamOutput)
		
		.Parameters("UserLogon") = strLogin
		.Parameters("EmployeeID") = strEmpID
		.Parameters("UpdatedBy") = lngUserIDLocal'Session("UserIDAuthority")

		.ActiveConnection = objCon
		 
	End With
   
	objCmd.Execute        
  
	intRecord = objCmd.Parameters.Item("LoginIDOutput")
	
	If Err<>0 Then 
		'Resonse.Write "<div class=""content-body"" style=""position:absolute; top:100px; width:80%; z-index:300""><div class=""alert alert-success"" role=""alert"">Error " & Err.Number & ": " & Err.Descrition & " </div></div>"
		
		Response.Write "<div class=""content-body"" style=""position:absolute; top:100; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>Error " & Err.Number & ": " & Err.Descrition & " Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></span></div></div></div>"
	End If
	
	On Error Goto 0
	
	If intRecord = 0 Then
		'Response.Write "<div class=""content-body"" style=""position:absolute; top:100px; width:80%; z-index:300""><div class=""alert alert-danger"" role=""alert"">Error: User ID " & strLogin & " does not exist in the Rollout data. Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></div></div>"
		
		Response.Write "<div class=""content-body"" style=""position:absolute; top:100; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>Error: User ID " & strLogin & " does not exist in the Rollout data. Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></span></div></div></div>"
			
	Else
		'Response.Write "<div class=""content-body"" style=""position:absolute; top:100px; width:80%; z-index:300""><div class=""alert alert-success"" role=""alert"">" & intRecord & " Updated</div></div>"
		
		Response.Write "<div class=""content-body"" style=""position:absolute; top:100; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>" & strLogin & " added." & intRecord & " Updated. Please login again now.</a></span></div></div></div>"
	End If
		
End Sub


Sub AddUserLogin(strLogin)
'Procedure to add the user trying to login to the relevant tables so they can access the PORTAL
Dim intRecord
Dim objCmd
	
    Set objCmd = Server.CreateObject("ADODB.Command")

	If IsNull(strLogin) OR strLogin = "" Then
		'Response.Write "<div class=""content-body"" style=""position:absolute; top:100px; width:80%; z-index:300"">" & _
		'	"<div class=""alert alert-danger"" role=""alert"">Error: UserLogon " & strLogin & " is empty. Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></div></div>"
			
			Response.Write "<div class=""content-body"" style=""position:absolute; top:100; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>Error: UserLogon " & strLogin & " is empty. Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></span></div></div></div>"

			
		Exit Sub
	End If
	
	'On Error Resume Next
	
	With objCmd

		.CommandType = 4
		.CommandText = "spPORTALUserLoginInsert"

		.Parameters.Append objCmd.CreateParameter("UserLogon", adVarChar, adParamInput, 50)	
		.Parameters.Append objCmd.CreateParameter("LoginIDOutput", adInteger, adParamOutput)
		
		.Parameters("UserLogon") = strLogin

		.ActiveConnection = objCon
		 
	End With
   
	objCmd.Execute        
  
	intRecord = objCmd.Parameters.Item("LoginIDOutput")
	
	If Err<>0 Then 
		'Resonse.Write "<div class=""content-body"" style=""position:absolute; top:100px; width:80%; z-index:300""><div class=""alert alert-success"" role=""alert"">Error " & Err.Number & ": " & Err.Descrition & " </div></div>"
		
		Response.Write "<div class=""content-body"" style=""position:absolute; top:100; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>Error " & Err.Number & ": " & Err.Descrition & " Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></span></div></div></div>"
	End If
	
	On Error Goto 0
	
	If intRecord = 0 Then
		'Response.Write "<div class=""content-body"" style=""position:absolute; top:100px; width:80%; z-index:300""><div class=""alert alert-danger"" role=""alert"">Error: User ID " & strLogin & " does not exist in the Rollout data. Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></div></div>"
		
		Response.Write "<div class=""content-body"" style=""position:absolute; top:100; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>Error: User ID " & strLogin & " does not exist in the Rollout data. Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></span></div></div></div>"
			
	Else
		'Response.Write "<div class=""content-body"" style=""position:absolute; top:100px; width:80%; z-index:300""><div class=""alert alert-success"" role=""alert"">" & intRecord & " Updated</div></div>"
		
		Response.Write "<div class=""content-body"" style=""position:absolute; top:100; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>" & strLogin & " added." & intRecord & " Updated. Please login again now.</a></span></div></div></div>"
	End If
		
End Sub



Sub AddUserLoginRecord(strLogin)
'Procedure to record each user login
Dim intRecord
Dim objCmd
	
    Set objCmd = Server.CreateObject("ADODB.Command")

	If IsNull(strLogin) OR strLogin = "" Then
		'Response.Write "<div class=""content-body"" style=""position:absolute; top:100px; width:80%; z-index:300"">" & _
		'	"<div class=""alert alert-danger"" role=""alert"">Error: UserLogon " & strLogin & " is empty. Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></div></div>"
			
			Response.Write "<div class=""content-body"" style=""position:absolute; top:100; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>Error: UserLogon " & strLogin & " is empty. Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></span></div></div></div>"

			
		Exit Sub
	End If
	
	On Error Resume Next
	
	With objCmd

		.CommandType = 4
		.CommandText = "spPORTALAddUserLoginRecord"

		.Parameters.Append objCmd.CreateParameter("UserLogon", adVarChar, adParamInput, 50)	
		
		.Parameters("UserLogon") = strLogin

		.ActiveConnection = objCon
		 
	End With
   
	objCmd.Execute        
  
	'intRecord = objCmd.Parameters.Item("LoginIDOutput")
	
	If Err<>0 Then 
		'Resonse.Write "<div class=""content-body"" style=""position:absolute; top:100px; width:80%; z-index:300""><div class=""alert alert-success"" role=""alert"">Error " & Err.Number & ": " & Err.Descrition & " </div></div>"
		
		Response.Write "<div class=""content-body"" style=""position:absolute; top:100; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>Error Add User Login Record " & Err.Number & ": " & Err.Descrition & " Please contact <a href=""mailto:defence.creditcards@defence.gov.au"">CAPS System Administrators</a></span></div></div></div>"
	End If
	
	On Error Goto 0
	
		
End Sub


Public Sub LoginUserPortal(strLoginString)
'Procedure to login a user who has just been added to the User table, rather than making them click login again

Dim objRSLP

	Set objRSLP = Server.CreateObject("ADODB.Recordset")

	'Open the Portal Login query to validate the User and log them in/redirect them
	objRSLP.Open "SELECT Top 1 * FROM qryPORTALUserLogon WITH(NOLOCK) WHERE UserLogon = '" & strLoginString & "'",objCon,3,1

		If Not objRSLP.EOF Then		
			
			'Response.Write "Session(UserIDAuthority) =" & Session("UserIDAuthority") 
			strEmployeeID = objRSLP("EmployeeID")
			Session("ApplicationEmployeeID") = strEmployeeID
			
			
			Session("UserID") = objRSLP("UserID")
			Session("UserTypeID") = objRSLP("UserTypeID")
			strName = objRSLP("FName") & " " & objRSLP("LName")
			Session("UserName") = objRSLP("FName") & " " & objRSLP("LName")
			Session("Logon") = objRSLP("UserLogon")
			   
			Session("EmployeeID") = objRSLP("EmployeeID")
			Session("UserIDAuthority") = objRSLP("UserID")		
					
			Session("BudgetID") = 1		
			Session("BusinessAreaID") = 1
			Session("CostCentreID") = 1
			Session("VersionID") = 1
		   
			Session("ServerPath") = GetSystemAdminDefault("ServerPath")

			
			'First flag the user as having logged in to the Portal if they have not before
			If IsNull(objRSLP("LoggedOntoPortal")) OR objRSLP("LoggedOntoPortal")="" Then
				'Update the User in the User table by calling the procedure for when they first login
				'Response.Write "llooogger=" & objRSLP("LoggedOntoPortal")
				Call UpdateUserLogin(strLoginString,strEmployeeID)
				
				strLastLoggedScreen = "Portal/CAPSPortal.asp?1=" & objRSLP("TwoFAID") & ""
			Else
				'If the user has logged in before then redirect them to the page they were last on (from database)
				If IsNull(objRSLP("LastScreen")) or objRSLP("LastScreen") = "" Then
					strLastLoggedScreen = "Portal/CAPSPortal.asp?1=" & objRSLP("TwoFAID") & ""
				Else
					strLastLoggedScreen = "Portal/" & objRSLP("LastScreen") & "?1=" & objRSLP("TwoFAID") & ""
				End If
			End If

			Session("UType") = "Portal"
			'Check to see if the user has Complete (Confirmed or Cancelled) their request and if so send them to the final screens (overwrite the last logged in)
			If Not IsNull(objRSLP("RolloutNewCardRequestedDate")) Then
				strLastLoggedScreen = "Portal/CAPSPortal5.asp"
			End If
			
			If Not IsNull(objRSLP("RolloutNewCardCancelledDate")) Then
				strLastLoggedScreen = "Portal/CAPSPortalCancel.asp"
			End If
			'Session("Menu") = Session("ServerPath") & "Portal/CAPSPortal.asp"
			'Session("ServerPath") = Session("ServerPath") & "Portal/"
			'Session("ServerPath") = "http://siepd85wsd0017.dpedev.protecteddev.mil.au/"
			'Response.write "<BR>2=" & Session("ServerPath")
			'Response.Redirect "Portal/CAPSPortal.asp?ApplicationEmployeeID=" & objRSLP("EmployeeID") & ""
			
			'Close objects before redirecting
			objRSLP.Close
			objCon.close

			Set objRSLP = Nothing
			Set objCon = Nothing
			
			'Set the Global.asa values for the Users Logged in
			   If IsEmpty(Application("NamedUsers")) Then
					Application("NamedUsers") = Session("UserID")
				Else
				
					'''Add the userID to the application variable if they are not in there already
					lngUserCount = 0
					strUsers =  Split(Application("NamedUsers"),",")
					For each strU in strUsers
						If clng(Session("UserID")) = clng(strU) Then lngUserCount = 1
					Next
										
					If lngUserCount= 0 Then Application("NamedUsers") = Application("NamedUsers") & "," & Session("UserID")
			   End If

			Response.Redirect strLastLoggedScreen
			
			'Response.Redirect "Portal/" & strLastLoggedScreen & "?ApplicationEmployeeID=" & objRSLP("EmployeeID") & ""
			'Response.Redirect "Portal/CAPSPortal.asp?ApplicationEmployeeID=" & objRSLP("EmployeeID") & ""
			'Below if the Two Factor Authentication page if this is enabled and an email is to be sent
			'Response.Redirect "Portal/CAPSPortalFA.asp?ApplicationEmployeeID=" & objRSLP("EmployeeID") & ""
		
		Else
			'Log an Error as the user will have only arived here if they could not be added automatically as a user based on their network login (picked up automatically) and trying to match this against their email address (in sproc)
			objCon.Execute "INSERT INTO tblErrorLog VALUES('CAPS/Default.asp','Login Failed 2:" & strLogEnter & " - Unable to Auto Login from new add to User table',GetDate(),0)"
		End If

objRSLP.close

Set objRSLP = Nothing

End Sub


objCon.close

Set objRS = Nothing
Set objCon = Nothing

%>