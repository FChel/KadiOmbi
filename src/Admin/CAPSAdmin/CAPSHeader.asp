<%@Language=VBScript CODEPAGE="65001"%>

<% 

'Dim objCon
'Dim objRS
'Dim objCmd


'Set objCon = Server.CreateObject("ADODB.Connection")
'Set objRS = Server.CreateObject("ADODB.Recordset")
'Set objCmd = Server.CreateObject("ADODB.Command")

'objCon.Open Session("DBConnection")


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
	
 %>

<!DOCTYPE html>
<html>
  <head>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta charset="utf-8" />
    <title>Defence Credit Card Application System (CAPS)</title>
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <link href="fontawesome/css/all.css" rel="stylesheet" />
    <link rel="stylesheet" href="css/myfi-bootstrap.css" />
  </head>
  <body>
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
                <input class="form-control field-search" type="search" placeholder="Search" aria-label="Search"/>
              </form>
              <div class="header-actions">
                <a class="action-link" href=""><i class="fa fa-home"></i></a>
                <a class="action-link" href="#" data-container="body" data-toggle="popover" data-placement="bottom" data-content="Vivamus sagittis lacus vel augue laoreet rutrum faucibus." >
                  <span class="counter">5</span>
                  <i class="fa fa-bell"></i></a>
                <a class="action-link" href="#" data-container="body" data-toggle="popover" data-placement="bottom" data-content="Vivamus sagittis lacus vel augue laoreet rutrum faucibus.">
                  <span class="counter">16</span><i class="fa fa-comment"></i></a>
                <a class="action-link" href=""><i class="fa fa-user"></i><span class="action-username d-none d-md-inline-block"><%=Session("UserName")%></span></a>
				
				<a href="../Default.asp" style="color:white;"><i class="fa fa-power-off"></i></a>
				
              </div>
            </div>
          </div>
        </div>
      </div>
      <div class="myfi-nav-container">
        <div class="container">
          <div class="row">
            <div class="col-12">
              <nav class="navbar navbar-expand-lg navbar-light">
                <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarNavDropdown" aria-controls="navbarNavDropdown" aria-expanded="false" aria-label="Toggle navigation">
                  <span class="menu-text">Menu</span>
                  <i class="fa fa-bars"></i>
                </button>
                <div class="collapse navbar-collapse" id="navbarNavDropdown">
                  <ul class="navbar-nav">
<%
			'Display the menu options relating to personal screens - which is applicable to all users
			'If Session("UType") = "Employee" or Session("UType") = "Admin" Then
					'Response.write "<li class=""nav-item active""><a class=""nav-link"" href=""HomeUser.asp"">Home <span class=""sr-only"">(current)</span></a></li>" & _
				Response.write "<li class=""nav-item active""><a class=""nav-link"" href=""" & Session("Menu") & """>Home <span class=""sr-only"">(current)</span></a></li>" & _
						"<li class=""nav-item""><a class=""nav-link"" href=""" & Session("ServerPath") & "CC/CardsEmployee.asp"" >My Cards</a></li>" & _
						"<li class=""nav-item""><a class=""nav-link"" href=""" & Session("ServerPath") & "CC/ApplicationsEmployeeHF.asp"" >My Applications</a></li>"
						
					'Response.Write "</ul></li>"
			'End If
							 
				'Display the Administration pages menu if the user logged is a manager
			   If Session("UType") = "Admin" Then
			   
					Response.Write "<li class=""nav-item dropdown""><a class=""nav-link dropdown-toggle"" href=""#"" id=""navbarDropdownMenuLink"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false""> Applications</a>" & _
                        "<div class=""dropdown-menu"" aria-labelledby=""navbarDropdownMenuLink""><a class=""dropdown-item""  href=""CC/Applications3.asp"" >All Applications</a><a class=""dropdown-item""href=""CC/Applications3_AG.asp"" >All Applications - New</a>" & _
                        "<a class=""dropdown-item"" href=""CC/Applications3GCFO.asp"" >ASFIN Approvals</a><a class=""dropdown-item"" href=""CC/ApplicationsSubmit2.asp"" >New Applications 2</a></div></li>" & _
						"<!-- Default drop down menu item" & _
						"<li class=""nav-item dropdown""><a class=""nav-link dropdown-toggle"" href=""#"" id=""navbarDropdownMenuLink"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false""> Dropdown link</a>" & _
                        "<div class=""dropdown-menu"" aria-labelledby=""navbarDropdownMenuLink""><a class=""dropdown-item"" href=""#"">Action</a><a class=""dropdown-item"" href=""#"">Another action</a>" & _
                        "<a class=""dropdown-item"" href=""#"">Something else here</a></div></li>-->" & _
						"<li class=""dropdown nav-item"" data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""#"" data-toggle=""dropdown""><span>Cards</span></a>" & _
						"<ul class=""dropdown-menu""><a class=""dropdown-item"" href=""CC/Cards3_2.asp"" >All Cards</a><a class=""dropdown-item"" href=""CC/Cards3_2.asp"" >My Cards</a>" & _
						"<a class=""dropdown-item"" href=""CC/Cards3.asp"" >All Cards 2</a></ul></li>" & _
						"<li class=""dropdown nav-item"" data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""#"" data-toggle=""dropdown""><span>Authority</span></a>" & _
						"<ul class=""dropdown-menu""><a class=""dropdown-item"" href=""CC/Authority.asp"" >View/Edit Authority</a>" & _
						"<a class=""dropdown-item"" href=""CC/CDMC.asp"" >Approve Authority</a></ul></li>" & _
						"<li class=""dropdown nav-item"" data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""#"" data-toggle=""dropdown""><span>Employee Details</span></a>" & _
						"<ul class=""dropdown-menu""><a class=""dropdown-item"" href=""CC/EmployeeHistory.asp"" >Employee History</a>" & _
						"<a class=""dropdown-item"" href=""CC/CDMC.asp"" >Corporate Directory</a></ul></li>" 
               
					Response.Write "<li class=""nav-item dropdown""><a class=""nav-link dropdown-toggle"" href=""#"" id=""navbarDropdownMenuLink"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">Dashboards</a>" & _
                          "<div class=""dropdown-menu"" aria-labelledby=""navbarDropdownMenuLink""><a class=""dropdown-item"" href=""CC/DashboardG1.asp"" >Dashboard Sample 1</a>" & _
						  "<a class=""dropdown-item"" href=""CC/DashboardSB2.asp"" >Dashboard Sample 2</a>" & _
                          "<a class=""dropdown-item"" href=""CC/DashboardFrest1.asp"" >Dashboard Sample 3</a>" & _
						  "<a class=""dropdown-item"" href=""CC/DashboardFrest2.asp"" >Dashboard Sample 4</a>" & _
						  "<a class=""dropdown-item"" href=""app-invoice-list.html"" >Invoice List</a>" & _
						  "</div></li>"
					
					Response.Write "<li class=""nav-item dropdown""><a class=""nav-link dropdown-toggle"" href=""#"" id=""navbarDropdownMenuLink"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">Administration</a>" & _
                          "<div class=""dropdown-menu"" aria-labelledby=""navbarDropdownMenuLink"">" & _
						  "<a class=""dropdown-item"" href=""../Admin/User.asp"" >Users</a>" & _
                          "<a class=""dropdown-item"" href=""../Admin/AdminMenu.asp"" >Admin Menu</a>" & _
						  "<a class=""dropdown-item"" href=""../Admin/CAPSAdmin/UploadCS2.asp"" >Upload CS From Diners File</a>" & _
						  "<a class=""dropdown-item"" href=""../Admin/CAPSAdmin/UploadANZ.asp"" >Upload ANZ Cardlist File</a>" & _
						  "<a class=""dropdown-item"" href=""../Admin/CAPSAdmin/UploadCDMC.asp"" >Upload CDMC File</a>" & _
						  "<a class=""dropdown-item"" href=""../Admin/CAPSAdmin/UploadROMAN.asp"" >Upload ROMAN File</a>" & _
						  "<a class=""dropdown-item"" href=""../Admin/CAPSAdmin/ExportCS.asp"" >Export CS To Diners File</a>" & _
						  "<a class=""dropdown-item"" href=""../Admin/CAPSAdmin/ExportNA.asp"" >Export NA To Diners File</a>" & _
						  "<a class=""dropdown-item"" href=""../Admin/CAPSAdmin/ExportANZ.asp"" >Export ANZ Cardlist File</a>" & _
						  "</div></li>"
						  
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

   
