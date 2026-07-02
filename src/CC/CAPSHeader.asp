<%@Language=VBScript CODEPAGE="65001"%>

<% 
Option Explicit
Dim objConHead
Dim objRSHead
'Dim objCmd
Dim strServerPath

Set objConHead = Server.CreateObject("ADODB.Connection")
Set objRSHead = Server.CreateObject("ADODB.Recordset")
'Set objCmd = Server.CreateObject("ADODB.Command")

If IsEmpty(Session("DBConnection")) Then 
	'Session("DBConnection") = "File Name=" & Server.MapPath("../Database/CAPS.udl") & ";"
	Response.Redirect("../Timeout.asp")
End IF

If IsEmpty(Session("DBConnectionHead")) Then
	'Set Database Connection
	Session("DBConnectionHead") = "File Name=" & Server.MapPath("../Database/CAPS.udl") & ";"
End If

objConHead.Open Session("DBConnectionHead")

strServerPath = Session("ServerPath")

	If Not IsEmpty(Request.QueryString("TopSearch")) Then
		'Response.Redirect Session("ServerPath") & "/CC/EmployeeHistory.asp?SearchAll=" + Request.QueryString("TopSearch")
		Response.Redirect Session("ServerPath") & "/CC/CDMCList.asp?Link=ED&SearchTerm=" + Request.QueryString("TopSearch")
		
		'Response.write "weew=" & Session("HomePageTop")
	End If
	
	'Session("HomePageTop") = "CC/HomeCC.asp"
	
	'If Not IsEmpty(Request.QueryString("HomePageTop")) Then
	'	Session("HomePageTop") = Request.QueryString("HomePageTop")
	'	'Response.write "weew=" & Session("HomePageTop")
	'End If

	'If Not IsEmpty(Request.QueryString("UType")) Then
'		Session("UType") = Request.QueryString("UType")
	'End If
	
	'If Session("UType") = "Employee" Then
	'	'Session("UType") = "Employee"
	'	Session("HomePageTop") = "CC/HomeCC2.asp"
	'	Session("HomePage1") = "MyCards.asp"
	'	Session("HomePage2") = "MyApplications.asp"
	'ElseIf Session("UType") = "Manager" Then
	'	'Session("UType") = "Manager"
	'	Session("HomePageTop") = "CC/HomeCC3.asp"
	'	Session("HomePage1") = "MyCards.asp"
	'	Session("HomePage2") = "MyApplications.asp"
	'Else
	'	'Session("UType") = "CreditCards"
	'	Session("HomePageTop") = "CC/HomeCC.asp"
	'	Session("HomePage1") = "CardTypeChart.asp"
	'	Session("HomePage2") = "ApplicationsHome.asp"
	'End If
					
	'If Not IsEmpty(Request.QueryString("HomeCC")) Then
	'	Session("HomeCC") = Request.QueryString("HomeCC")
	'End If
Dim arrActive(8)
Dim strActiveLink
Dim intHeaderX

	For intHeaderX = 1 to 8
		arrActive(intHeaderX) = ""
	Next
		
	'arrActive(1) = "active"
	
	If Not IsEmpty(Request.QueryString("Link")) Then
		'Response.write "</br>" & Request.Servervariables("PATH_INFO")
		
		'strActive = Right(Request.Servervariables("HTTP_REFERER"),2)
		strActiveLink = Request.QueryString("Link")
		
		'Response.write strActive
		
		If strActiveLink = "HM" Then arrActive(1) = "active"
		If strActiveLink = "AP" Then arrActive(2) = "active"
		If strActiveLink = "CD" Then arrActive(3) = "active"
		If strActiveLink = "AU" Then arrActive(4) = "active"
		If strActiveLink = "ED" Then arrActive(5) = "active"
		If strActiveLink = "AD" Then arrActive(6) = "active"
		If strActiveLink = "HP" Then arrActive(7) = "active"
		If strActiveLink = "RP" Then arrActive(8) = "active"
	Else
		'Make the home page link the default active
		arrActive(1) = "active"
	End If
	
	
 %>

<!DOCTYPE html>
<html>
  <head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta charset="utf-8" />
    <title>Defence Credit Card Application System (CAPS)</title>
	<link rel="icon" type="image/jpg" href="<%=Session("ServerPath")%>favicon.ico"/>
    <meta name="viewport" content="width=device-width, initial-scale=1" />

	<script src="<%=Session("ServerPath")%>js/jquery.js"></script>
    <script src="<%=Session("ServerPath")%>bootstrap/js/bootstrap.bundle.min.js"></script>
	 <link href="<%=Session("ServerPath")%>fontawesome/css/all.css" rel="stylesheet" />
    <link rel="stylesheet" href="<%=Session("ServerPath")%>css/myfi-bootstrap.css" />
	<link rel="stylesheet" href="<%=Session("ServerPath")%>introjs/introjs.min.css" />

<script LANGUAGE="javascript">

function TopEmployeeSearch() {
   
	alert(frm.TopSearch.value);
    //self.location = "<%=Session("ServerPath")%>/CC/EmployeeHistory.asp?SearchAll=" + frm.TopSearch.value;
      
}

function SelectApp(varAppID) {
	if(varAppID==undefined) {
		alert(varAppID);
	}
	{
	self.location = "EmployeeHistory.asp?Action=Search&ApplicationID=" + varAppID;
	}
}

function LoadSettings() {
  
  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ModalSettingsMessage").innerHTML = this.responseText;
    }
  };

  xhttp.open("GET", "../CC/AJAX/GetMySettings.asp", true);
  xhttp.send();

}

function SaveReadGDrive(varSet) {

  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ModalSettingsMessage").innerHTML = this.responseText;
    }
  };

  xhttp.open("GET", "../CC/AJAX/GetMySettings.asp?GDriveSet="+varSet, true);
  xhttp.send();

}

$(document).ready(function(){
//below does not remove the blurred background on hide (script at the bottom of this page)
//	$('#ModApp').modal('show');
});


</script>

<style>
.dropdown-submenu:hover > .dropdown-menu, .dropdown-submenu:focus > .dropdown-menu{
    display: flex;
    flex-direction: column;
    position: absolute !important;
    margin-top: -30px;
    left: 100%!important;

}


</style>


  </head>
  <body>
  
  <!-- Modal -->
<div class="loader" id="ModApp">
	<div class="wrap" style="position:relative; left:250px; top:350px; background-color:transparent;">
		<div class="spinner"></div>
		<span class="loading-message">Loading...</h6>
	</div>
</div>

  <!-- Start Settings Modal -->
	<div class="modal fade" id="ModalSettings" tabindex="-1" role="dialog" aria-labelledby="ModalDeleteCenterTitle" aria-hidden="true">
		<div class="modal-dialog modal-dialog-centered" role="document">
		  	<div class="modal-content">
				<div class="modal-header">
			  		<h5 class="modal-title" id="ModalSettingsLongTitle" style="font-weight:bold;">My Settings</h5>
			  		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				</div>
				<div class="modal-body">			  
				  	<div class="row">
						<div class="col-md-12 mb-3">			
						  		<div id="ModalSettingsMessage"></div>
						</div>
				  	</div>
				  	<div class="row">
					  	<div class="col-md-12 mb-3" style="text-align:right;">
						  	<button type="button" class="btn btn-secondary btn-sm" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
					  	</div>
				  	</div>
			  	</div>
			</div>	
		</div>
		<div class="modal-footer"></div>
	</div>
	<!-- End Settings Modal -->
	
	
    <header class="masthead">
      <div class="myfi-header">
        <div class="container">
          <div class="row">
            <div class="col-6 my-auto">
              <div class="logo-lockup">
                <img src="<%=Session("ServerPath")%>images/defence_logo_light.png" alt="Department of Defence" class="defence-logo" />
                <div class="app-info">
                  <div class="app-title">CAPS</div>
                  <div class="app-description d-none d-md-inline">
                    Credit Card Application System
                  </div>
                </div>
              </div>
            </div>
            <div class="col-6 my-auto text-right">
              <form class="header-search d-none d-lg-inline-block">
                <input class="form-control field-search" type="search" placeholder="Search" aria-label="Search" id="TopSearch" name="TopSearch" onChange="TopEmployeeSearch();" />
				<!-- <input class="form-control field-search" type="search" placeholder="Search" aria-label="Search" id="TopSearch" name="TopSearch" onChange="self.location='CC/EmployeeHistory.asp?SearchAll='"+frm.TopSearch.value />-->
              </form>
              <div class="header-actions">
                <a class="action-link" href="" & strServerPath & ""><i class="fa fa-home"></i></a>
                <!--<a class="action-link" href="#" data-container="body" data-toggle="popover" data-placement="bottom" data-content="Alerts for ><%=Session("UserName")%>" >
                  <span class="counter">5</span>
                  <i class="fa fa-bell"></i></a>-->
                <!--<a class="action-link" href="#" data-container="body" data-toggle="popover" data-placement="bottom" data-content="Messages for ><%=Session("UserName")%>">
                  <!--<span class="counter">16</span>--><!--<i class="fa fa-comment"></i></a>-->
               <!-- <a class="action-link" href=""><i class="fa fa-user"></i><span class="action-username d-none d-md-inline-block"><%=Session("UserName")%></span></a>-->
				     <div class="header-actions">
                  

                    <button class="action-link dropdown-toggle" id="userDropdown" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false" Title="<%="Session Time Left: " & Session.Timeout%>">
                      <i class="fa fa-user"></i><span class="action-username d-none d-md-inline-block"><%=Session("UserName")%></span>
                    </button>
                    <div class="dropdown-menu" aria-labelledby="userDropdown">
					<!--<ul class="navbar-nav">
					<li class="nav-item"><a class="dropdown-item" href="http://vbmrsn05/CAPS">Settings</a></li></ul>-->

					<%
						objRSHead.Open "SELECT * FROM qryModuleAccess WHERE UserID = " & Session("UserID") & "",objConHead
						
							Do Until objRSHead.EOF
								Response.Write  "<a class=""dropdown-item"" href=""" & objRSHead("ModuleURL") & "?ModuleID=" & objRSHead("Module") & "&Action=Login"">" & objRSHead("Module") & "</a>"
								objRSHead.Movenext
							Loop
							
						objRSHead.Close			
					
						Response.Write  "<li class=""dropdown-divider""></li>"
						
						If Session("UserTypeID") > 9 Then
							Response.Write  "<a class=""dropdown-item"" href=""#"" onClick=""LoadSettings()"" data-toggle=""modal"" data-target=""#ModalSettings""><i class=""fa fa-cogs""></i> My Settings</a>"
							'Response.Write "<a class=""dropdown-item"" href=""#"">" & strServerPath & "</a>"
						End If
					%>
					
                    </div>
                  </div>
					
					
				<a href="../Default.asp?Logoff=Y" style="color:white;"><i class="fa fa-power-off"></i></a>
				
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="myfi-nav-container">
        <div class="container">
          <div class="row">
            <div class="col-12">
              <nav class="navbar navbar-expand-md navbar-light">
                <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNavDropdown" aria-controls="navbarNavDropdown" aria-expanded="false" aria-label="Toggle navigation">
                  <span class="menu-text">Menu</span>
                  <i class="fa fa-bars"></i>
                </button>
                <div class="collapse navbar-collapse" id="navbarNavDropdown">
                  <ul class="navbar-nav">
				  
				  
<%
			'Display the menu options relating to personal screens - which is applicable to all users
			Response.write "<li class=""nav-item " & arrActive(1) & """><a class=""nav-link"" href=""" & Session("Menu") & "?Link=HM"">Home <span class=""sr-only"">(current)</span></a></li>"
			
			
			
			
			If Session("UType") = "Employee" or Session("UType") = "Manager" or Session("UType") = "Compliance" Then
					'Response.write "<li class=""nav-item active""><a class=""nav-link"" href=""HomeUser.asp"">Home <span class=""sr-only"">(current)</span></a></li>" & _
				'Response.write "<li class=""nav-item active""><a class=""nav-link"" href=""" & Session("Menu") & """>Home <span class=""sr-only"">(current)</span></a></li>" & _
				Response.write		"<li class=""nav-item""><a class=""nav-link"" href=""" & Session("ServerPath") & "CC/Cards.asp?UserView=User&Link=AP"" >My Cards</a></li>" & _
						"<li class=""nav-item""><a class=""nav-link"" href=""" & Session("ServerPath") & "CC/Applications.asp?Link=AP"" >My Applications</a></li>"
						
					'Response.Write "</ul></li>"
			End If
			
			
							 
				'Display the Administration pages menu if the user logged is a manager
			   If Session("UType") = "Admin" Then
			   
			   'add these back below for phase 2
			   	'"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/ApplicationsASFIN.asp?UserView=All&Link=AP"" >ASFIN Approvals</a><a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/Applications.asp?UserView=User&Link=AP"" >My Applications</a>" & _
						'"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/ApplicationsSubmit.asp?UserView=User&Link=AP"" >Current Application</a>" & _
						''"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/ApplicationsLimitSubmit.asp?UserView=User&Link=AP"" >Limit Change Application</a></div></li>" & _
			   
					Response.Write "<li class=""nav-item dropdown " & arrActive(2) & """><a class=""nav-link dropdown-toggle"" href=""#"" id=""navbarDropdownMenuLink"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false""> Applications</a>" & _
                        "<div class=""dropdown-menu"" aria-labelledby=""navbarDropdownMenuLink""><a class=""dropdown-item""  href=""" & Session("ServerPath") & "CC/Applications.asp?UserView=All&Link=AP"" >All Applications</a>" & _
                        "<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/LimitReductionSummary.asp?UserView=All&Link=AP"" >Limit Reduction Summary</a>" &  _		
						"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/CreditLimits.asp"" >Credit Limits</a>" &  _
						"<li class=""dropdown nav-item " & arrActive(3) & """ data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""#"" data-toggle=""dropdown""><span>Cards</span></a>" & _
						"<ul class=""dropdown-menu""><a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/Cards.asp?UserView=All&Link=CD"" >All Cards</a>" & _
						"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/Cards.asp?UserView=User&Link=CD"" >My Cards</a>" & _
						"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/AuditLog.asp?UserView=User&Link=CD"" >Card Audit History</a>" & _
						"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/ExpiringCards.asp?UserView=User&Link=CD"" >Cards Due To Expire</a>" & _
						"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "Admin/CardAccountTransfer.asp?UserView=User&Link=CD"" >Card Account Transfers</a></ul></li>" & _
						"<li class=""dropdown nav-item " & arrActive(4) & """ data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""#"" data-toggle=""dropdown""><span>Authority</span></a>" & _
						"<ul class=""dropdown-menu""><a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/Authority.asp?Link=AU"" >View/Edit Authority</a>" & _
						"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/Authority.asp?Link=AU"" >Approve Authority</a></ul></li>" & _
						"<li class=""dropdown nav-item " & arrActive(5) & """ data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""CC/CDMC.asp?Link=ED"" data-toggle=""dropdown""><span>Employee Details</span></a>" & _
						"<ul class=""dropdown-menu""><a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/EmployeeHistory.asp?Link=ED"" >Employee History</a>" & _
						"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/CDMCList.asp?Link=ED"" >Corporate Directory</a></ul></li>" 
               
					'Response.Write "<li class=""nav-item dropdown " & arrActive(6) & """><a class=""nav-link dropdown-toggle"" href=""#"" id=""navbarDropdownMenuLink"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">Administration</a>" & _
                    '      "<div class=""dropdown-menu"" aria-labelledby=""navbarDropdownMenuLink"">" & _
					'	  "<a class=""dropdown-item"" href=""" & Session("ServerPath") & "Admin/User.asp?Link=AD"" >Users</a>" & _
                    '      "<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/HomeAdmin.asp?Link=AD"" >Admin Menu</a>" & _
					'	  "<a class=""dropdown-item"" href=""" & Session("ServerPath") & "Admin/UploadCS2.asp?Link=AD"" >Upload CS From Diners File</a>" & _
					'	  "<a class=""dropdown-item"" href=""" & Session("ServerPath") & "Admin/UploadANZ.asp?Link=AD"" >Upload ANZ Cardlist File</a>" & _
					'	  "<a class=""dropdown-item"" href=""" & Session("ServerPath") & "Admin/UploadCDMC.asp?Link=AD"" >Upload CDMC File</a>" & _
					'	  "<a class=""dropdown-item"" href=""" & Session("ServerPath") & "Admin/UploadROMAN.asp?Link=AD"" >Upload ROMAN File</a>" & _
					'	  "<a class=""dropdown-item"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportCS.asp?Link=AD"" >Export CS To Diners File</a>" & _
					'	  "<a class=""dropdown-item"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportNA.asp?Link=AD"" >Export NA To Diners File</a>" & _
					'	  "<a class=""dropdown-item"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportANZ.asp?Link=AD"" >Export ANZ Cardlist File</a>" & _
					'	  "</div></li>" 
					
					Response.Write "<li class=""nav-item dropdown " & arrActive(6) & """><a id=""dropdownMenu1"" href=""#"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"" class=""nav-link dropdown-toggle"">Administration</a>" & _
                          "<ul aria-labelledby=""dropdownMenu1"" class=""dropdown-menu border-0 shadow"">" & _
						  "<li><a href=""" & Session("ServerPath") & "Admin/Users.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Users </a></li>" & _
						  "<li><a href=""" & Session("ServerPath") & "CC/HomeAdmin.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Daily Tasks </a></li>" & _
						  "<li><a href=""" & Session("ServerPath") & "CC/HomeCCAdmin.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">CC Team Dashboard </a></li>" & _
						  "<li><a href=""" & Session("ServerPath") & "Admin/SystemParameters.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">System Parameters </a></li>" & _
						  "<li><a href=""" & Session("ServerPath") & "Admin/EmailSend.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Send Emails </a></li>" & _
						  "<li><a href=""" & Session("ServerPath") & "Admin/EmailAdmin.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Email Admin</a></li>" & _
						  "<li class=""dropdown-divider""></li><li class=""dropdown-submenu"">" & _
                          "<a id=""dropdownMenu2"" href=""#"" role=""button"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"" class=""dropdown-item dropdown-toggle"" style=""padding-top:5px;"">Configuration</a>" & _
						  "<ul aria-labelledby=""dropdownMenu2"" class=""dropdown-menu border-0 shadow"">" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/SystemParameters.asp?Link=AD"" class=""dropdown-item"">System Parameters</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/EmailTemplate.asp?Link=AD"" class=""dropdown-item"">Email Template</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/EmailError.asp?Link=AD"" class=""dropdown-item"">Email Error Messages</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/EmailReplaceFields.asp?Link=AD"" class=""dropdown-item"">Email Message Replace</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSFileLoad.asp?Link=AD"" class=""dropdown-item"">File Load Admin</a></li>" & _
						  "<li class=""dropdown-submenu"">" & _
                          "<a id=""dropdownMenu3"" href=""#"" role=""button"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"" class=""dropdown-item dropdown-toggle"">NAB</a>" & _
						  "<ul aria-labelledby=""dropdownMenu3"" class=""dropdown-menu border-0 shadow"">" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/SystemParameters.asp?Link=AD"" class=""dropdown-item"">NAB System Parameters</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/EmailTemplate.asp?Link=AD"" class=""dropdown-item"">NAB Email Templates</a></li>" & _
						  "</ul></li>" & _
						  "</ul></li>" & _
						  "<li class=""dropdown-submenu"">" & _
                          "<a id=""dropdownMenu2"" href=""#"" role=""button"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"" class=""dropdown-item dropdown-toggle"" style=""padding-top:5px;"">Imports</a>" & _
						  "<ul aria-labelledby=""dropdownMenu2"" class=""dropdown-menu border-0 shadow"">" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadCSNAB.asp?CardType=DTC&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload CS From NAB File</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadCDMC.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload CDMC File</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadROMAN.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload ROMAN File</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/LoadXML.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload Applications</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadTraining.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload Training File</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadPM.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload ProMaster Data</a></li>" & _
						"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadCDMC_DCCP.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload CDMC File DCCP</a></li>" & _
							"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/DCCPSynch.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">DCCP Portal Synchronisation</a></li>" & _
						  "</ul></li>" & _
						  "<li class=""dropdown-submenu"">" & _
                          "<a id=""dropdownMenu2"" href=""#"" role=""button"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"" class=""dropdown-item dropdown-toggle"" style=""padding-top:5px;"">Exports</a>" & _
						  "<ul aria-labelledby=""dropdownMenu2"" class=""dropdown-menu border-0 shadow"">" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportCSNAB.asp?CardType=DTC&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"" >Export CM To NAB File</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportNANAB.asp?CardType=DPC&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Export NA To NAB</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportPM.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Export ProMaster Files</a></li>" & _
						"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/PGPFileLoader.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">PGP File Encrypter</a></li>" & _
						  "</ul></li>" & _
						   "<li class=""dropdown-submenu"">" & _
                          "<a id=""dropdownMenu2"" href=""#"" role=""button"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"" class=""dropdown-item dropdown-toggle"" style=""padding-top:5px;"">Detail</a>" & _
						  "<ul aria-labelledby=""dropdownMenu2"" class=""dropdown-menu border-0 shadow"">" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CSTransactionsToNAB.asp"" class=""dropdown-item"" style=""padding-top:5px;"">CM to NAB File</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/NATransactionsNAB.asp?CardType=NAB&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">NA to NAB File</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CSTransactionsNAB.asp"" class=""dropdown-item"" style=""padding-top:5px;"">CS from NAB File</a></li>" & _ 
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/TrainingTransactions.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Training File</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/PMReconciliation.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">ProMaster Reco</a></li>" & _
						  "</ul></li></ul>"
					'Tiffany - removed ANZ upload page - 389line
					' "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadANZ.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload ANZ Cardlist File</a></li>" & _
					'removed lines after 385
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadCS2DTCDPC.asp?CardType=DTC&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload CS From Diners File DTC</a></li>" & _
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadCS2DTCDPC.asp?CardType=DPC&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload CS From Diners File DPC - Branded</a></li>" & _
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadCS2DTCDPC.asp?CardType=DPCU&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Upload CS From Diners File DPC - Unbranded</a></li>" & _
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportNA.asp?CardType=DTC&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Export NA To Diners File - DTC</a></li>" & _
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportNA.asp?CardType=DPC&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Export NA To Diners File - DPC</a></li>" & _
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportCS.asp?CardType=DTC&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Export CS To Diners File</a></li>" & _
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportANZ.asp?Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">Export DPC To ANZ File</a></li>" & _
					'"<a class=""dropdown-item"" href=""" & Session("ServerPath") & "CC/UnactivatedCards.asp?UserView=User&Link=CD"" >Unactivated Cards</a></ul></li>" & _
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/NATransactionsDPC.asp?CardType=DPC&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">NA To Diners File - DPC</a></li>" & _
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CSTransactions.asp?CardType=DPC&Link=AD"" class=""dropdown-item"" style=""padding-top:5px;"">CS From Diners File - DPC</a></li>" & _
						  
					
					
					
					
					'Removed from above (line 378 - below CD To Diners File) as there will be one CS screen for DTC and DPC
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/ExportCSDTCDPC.asp?CardType=DPC&Link=AD"" class=""dropdown-item"">Export CS To Diners File DPC</a></li>" & _
					'(line 388) - CS To Detail screen
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CSTransactionsTo.asp?CardType=DPC&Link=AD"" class=""dropdown-item"">CS To Diners File - DPC</a></li>" & _
					
					'Removed from Administration Imports (Applications) Menu above
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/UploadApplications.asp?Link=AD"" class=""dropdown-item"">Upload Applications</a></li>" & _
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/LoadXML.asp?Link=AD"" class=""dropdown-item"">Upload Applications Server</a></li>" & _
					
					'Response.Write "<li class=""dropdown nav-item"" data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""#"" data-toggle=""dropdown""><span>Administration</span></a>" & _
					'				"<ul class=""dropdown-menu"">" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center""  href=""../Admin/User.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Users</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center""  href=""../Admin/AdminMenu.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Admin Menu</a></li>" & _
					'				"<li class=""dropdown dropdown-submenu"" data-menu=""dropdown-submenu""><a class=""dropdown-item align-items-center dropdown-toggle"" href=""#"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Imports</a>" & _
					'				"<ul class=""dropdown-menu"">" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center""  href=""../Admin/CAPSAdmin/UploadCS2.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Upload CS From Diners File</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center""  href=""../Admin/CAPSAdmin/UploadANZ.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Upload ANZ Cardlist File</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center""  href=""../Admin/CAPSAdmin/UploadCDMC.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Upload CDMC File</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center""  href=""../Admin/CAPSAdmin/UploadROMAN.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Upload ROMAN File</a></li>" & _
					'				"</ul></li>" & _
					'				"<li class=""dropdown dropdown-submenu"" data-menu=""dropdown-submenu""><a class=""dropdown-item align-items-center dropdown-toggle"" href=""#"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Exports</a>" & _
					'				"<ul class=""dropdown-menu"">" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center""  href=""../Admin/CAPSAdmin/ExportCS.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Export CS To Diners File</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center""  href=""../Admin/CAPSAdmin/ExportNA.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Export NA To Diners File</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center""  href=""../Admin/CAPSAdmin/ExportANZ.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Export ANZ Cardlist File</a></li>" & _
					'				"</ul></li>"
									
					'Response.write "<li class=""nav-item"" data-toggle=""tooltip"" data-placement=""right"" title=""Administration Menu"">" & _
					'		  "<a class=""nav-link nav-link-collapse collapsed"" data-toggle=""collapse"" href=""#collapseExamplePages"" data-parent=""#exampleAccordion"">" & _
					'			"<i class=""fa fa-fw fa-file""></i>" & _
					'			"<span class=""nav-link-text"">Administration</span></a>" & _
					'		  "<ul class=""sidenav-second-level collapse"" id=""collapseExamplePages"">" & _
					'			"<li><a  href=""CC/CSToDiners.asp?CSToDinersID=0"">CS To Diners File</a></li>" & _
					'			"<li><a  href=""CC/CSFromDiners.asp?CSFromDinersID=0"">CS From Diners File</a></li>" & _
					'			"<li><a  href=""CC/ANZUpdates.asp?ANZUpdateID=0"">ANZ Updates</a></li>" & _
					'			"<li><a  href=""../../Admin/User.asp"">Users</a></li>" & _
					'			"<li><a  href=""../../Admin/AdminMenu.asp"">Admin Menu</a></li>" & _
					'			"<li><a  href=""../../Admin/CAPSAdmin/UploadCDMC.asp"">Upload CDMC Data</a></li>" & _
					'			"<li><a  href=""../../Admin/CAPSAdmin/UploadCS.asp"">Upload CS From Diners File</a></li>" & _
					'		  "</ul></li>"

					'****---- APPLICATION Upload links removed from the menu but the pages might still be required for future testing ----*****
					'"<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/LoadApplication.asp?Link=AD"" class=""dropdown-item"">Upload Applications 2</a></li>" & _
					'	   "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/CAPSAdmin/LoadApplication3.asp?Link=AD"" class=""dropdown-item"">Upload Applications 3</a></li>" & _
			   End If
			   
			   'Write the Reports menu based on the User Type ID (4 = Compliance, 10, 11 and 99 = Admin)
			   If Session("UserTypeID") > 3 Then
			   
					Response.Write "<li class=""nav-item dropdown " & arrActive(8) & """><a id=""dropdownMenu1"" href=""#"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"" class=""nav-link dropdown-toggle"">Reports</a>" & _
                          "<ul aria-labelledby=""dropdownMenu1"" class=""dropdown-menu border-0 shadow"">" & _
						  "<li><a href=""" & Session("ServerPath") & "Reports/Training.asp?Link=RP"" class=""dropdown-item"">Training </a></li>" & _
						  "<li class=""dropdown-divider""></li><li><a href=""" & Session("ServerPath") & "Reports/CAPSSuperReporter.xlsm"" class=""dropdown-item"">CAPS Super Reporter </a></li>" & _
						  "<li class=""dropdown-divider""></li>" & _
						  "<li class=""dropdown-submenu"">" & _
                          "<a id=""dropdownMenu2"" href=""#"" role=""button"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"" class=""dropdown-item dropdown-toggle"">Compliance</a>" & _
						  "<ul aria-labelledby=""dropdownMenu2"" class=""dropdown-menu border-0 shadow"">" & _
						  "<li><a href=""" & Session("ServerPath") & "Admin/EmailSend.asp?Link=RP"" class=""dropdown-item"">Send Emails </a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Admin/EmailTemplate.asp?Link=RP"" class=""dropdown-item"">Email Template</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Reports/Training.asp?Link=RP"" class=""dropdown-item"">DTC Training</a></li>" & _
						  "<li><a tabindex=""-1"" href=""" & Session("ServerPath") & "Reports/Training.asp?Link=RP"" class=""dropdown-item"">DPC Training</a></li>" & _
						  "</ul></li>" & _
						  "<li><a href=""" & Session("ServerPath") & "CC/CancelCardsCheck.asp?Link=RP"" class=""dropdown-item"">Cancel Cards Check</a></li>" & _
						  "<li><a href=""" & Session("ServerPath") & "CC/ANZCancelCards.asp?Link=RP"" class=""dropdown-item"">Cancel ANZ Cards</a></li>" & _
						  "<li><a href=""" & Session("ServerPath") & "CC/CAPSCompareCDMC.asp?Link=RP"" class=""dropdown-item"">Card/CDMC Updates</a></li>"  & _
						  "<li><a href=""" & Session("ServerPath") & "Admin/CAPSAdmin/CheckLogs.asp?Link=RP"" class=""dropdown-item"">Check IIS Logs</a></li></ul>"
			   End If
			   
			   'Help menu
			   Response.write "<li class=""nav-item " & arrActive(7) & """><a class=""nav-link"" href=""" & Session("ServerPath") & "CC/HelpFile.html?Link=HP"" target=""_new"">Help <span class=""sr-only"">(current)</span></a></li>"
			   
			   'Write the Reports menu based on the User Type ID (4 = Compliance, 10, 11 and 99 = Admin)
			   If Session("UserTypeID") = 99 Then
					'Help menu
					Response.write "<li class=""nav-item " & arrActive(7) & """><a class=""nav-link"" href=""" & Session("ServerPath") & "Admin/ProjectDocs.asp?Link=HP""> Project Docs </a></li>"
			   
			   End If
			   %>
						   
						
					
                  </ul>
                </div>
              </nav>
            </div>
          </div>
        </div>
      </div>
    </header>
    <!-- BEGIN: Content-->
	
    <!-- END: Content-->
<%
'If there is a global message then write to the screen for all users
If Application("GlobalMessage") = "" Or IsNull(Application("GlobalMessage")) Then
Else
		Response.Write "<div class=""warning alert-warning"" role=""alert"" style=""text-align:center; font-weight:bold; height:40px; padding: 10px;"">" & Application("GlobalMessage") & "</div>"
End If

Set objConHead = Nothing
Set objRSHead = Nothing
%>
  <script>
	$(window).load(function() {
		$('#ModApp').hide();
	});
  </script>
