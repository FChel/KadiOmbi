<%@Language=VBScript CODEPAGE="65001"%>

<% 

'Dim objCon
'Dim objRS
'Dim objCmd


'Set objCon = Server.CreateObject("ADODB.Connection")
'Set objRS = Server.CreateObject("ADODB.Recordset")
'Set objCmd = Server.CreateObject("ADODB.Command")

'objCon.Open Session("DBConnection")


	Session("HomePageTop") = "CC/HomeCC.asp"
	
	If Not IsEmpty(Request.QueryString("HomePageTop")) Then
		Session("HomePageTop") = Request.QueryString("HomePageTop")
		'Response.write "weew=" & Session("HomePageTop")
	End If

	If Not IsEmpty(Request.QueryString("UType")) Then
		Session("UType") = Request.QueryString("UType")
	End If
	
	If Session("UType") = "Employee" Then
		'Session("UType") = "Employee"
		Session("HomePageTop") = "CC/HomeCC2.asp"
		Session("HomePage1") = "MyCards.asp"
		Session("HomePage2") = "MyApplications.asp"
	ElseIf Session("UType") = "Manager" Then
		'Session("UType") = "Manager"
		Session("HomePageTop") = "CC/HomeCC3.asp"
		Session("HomePage1") = "MyCards.asp"
		Session("HomePage2") = "MyApplications.asp"
	Else
		'Session("UType") = "CreditCards"
		Session("HomePageTop") = "CC/HomeCC.asp"
		Session("HomePage1") = "CardTypeChart.asp"
		Session("HomePage2") = "ApplicationsHome.asp"
	End If
					
	If Not IsEmpty(Request.QueryString("HomeCC")) Then
		Session("HomeCC") = Request.QueryString("HomeCC")
	End If
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
                <img
                  src="images/defence_logo_light.png"
                  alt="Department of Defence"
                  class="defence-logo"
                />
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
                <input class="form-control field-search" type="search" placeholder="Searchx" aria-label="Search"/>
              </form>
              <div class="header-actions">
                <a class="action-link" href=""><i class="fa fa-home"></i></a>
                <a class="action-link" href="#" data-container="body" data-toggle="popover" data-placement="bottom" data-content="Vivamus sagittis lacus vel augue laoreet rutrum faucibus." >
                  <span class="counter">5</span>
                  <i class="fa fa-bell"></i></a>
                <a class="action-link" href="#" data-container="body" data-toggle="popover" data-placement="bottom" data-content="Vivamus sagittis lacus vel augue laoreet rutrum faucibus.">
                  <span class="counter">16</span><i class="fa fa-comment"></i></a>
                <a class="action-link" href=""><i class="fa fa-user"></i><span class="action-username d-none d-md-inline-block"><%=Session("UserName")%></span></a>
				
				<a href="Default.asp" style="color:white;"><i class="fa fa-power-off"></i></a>
				
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
						   
                    Response.write "<li class=""nav-item active""><a class=""nav-link"" href="""" & Session(""Menu"") & """">Home <span class=""sr-only"">(current)</span></a></li>" & _
						"<li class=""nav-item""><a class=""nav-link"" href=""CC/Cards3_2.asp"">My Cards</a></li>" & _
						"<li class=""nav-item""><a class=""nav-link"" href=""CC/ApplicationsEmployee.asp"" >My Applications</a></li>"
						'"<li class=""nav-item""><a class=""nav-link"" href=""CC/ApplicationsEmployee.asp"" Target=""activeframe"">My Applications</a></li>"
						
					'Response.Write "</ul></li>"
			'End If
							 
				'Display the Administration pages menu if the user logged is a manager
			   If Session("UType") = "Admin" Then
			   
					Response.Write "<li class=""nav-item dropdown""><a class=""nav-link dropdown-toggle"" href=""#"" id=""navbarDropdownMenuLink"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false""> Applications</a>" & _
                        "<div class=""dropdown-menu"" aria-labelledby=""navbarDropdownMenuLink""><a class=""dropdown-item""  href=""CC/Applications3.asp"" Target=""activeframe"">All Applications</a><a class=""dropdown-item""href=""CC/Applications3_AG.asp"" Target=""activeframe"">All Applications - New</a>" & _
                        "<a class=""dropdown-item"" href=""CC/Applications3GCFO.asp"" Target=""activeframe"">ASFIN Approvals</a><a class=""dropdown-item"" href=""CC/ApplicationsSubmit2.asp"" Target=""activeframe"">New Applications 2</a></div></li>" & _
						"<!-- Default drop down menu item" & _
						"<li class=""nav-item dropdown""><a class=""nav-link dropdown-toggle"" href=""#"" id=""navbarDropdownMenuLink"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false""> Dropdown link</a>" & _
                        "<div class=""dropdown-menu"" aria-labelledby=""navbarDropdownMenuLink""><a class=""dropdown-item"" href=""#"">Action</a><a class=""dropdown-item"" href=""#"">Another action</a>" & _
                        "<a class=""dropdown-item"" href=""#"">Something else here</a></div></li>-->" & _
						"<li class=""dropdown nav-item"" data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""#"" data-toggle=""dropdown""><span>Cards</span></a>" & _
						"<ul class=""dropdown-menu""><a class=""dropdown-item"" href=""CC/Cards3_2.asp"" Target=""activeframe"">All Cards</a><a class=""dropdown-item"" href=""CC/Cards3_2.asp"" Target=""activeframe"">My Cards</a>" & _
						"<a class=""dropdown-item"" href=""CC/Cards3.asp"" Target=""activeframe"">All Cards 2</a></ul></li>" & _
						"<li class=""dropdown nav-item"" data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""#"" data-toggle=""dropdown""><span>Authority</span></a>" & _
						"<ul class=""dropdown-menu""><a class=""dropdown-item"" href=""CC/Authority.asp"" Target=""activeframe"">View/Edit Authority</a>" & _
						"<a class=""dropdown-item"" href=""CC/Authority.asp"" Target=""activeframe"">Approve Authority</a></ul></li>" & _
						"<li class=""dropdown nav-item"" data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""#"" data-toggle=""dropdown""><span>Employee Details</span></a>" & _
						"<ul class=""dropdown-menu""><a class=""dropdown-item"" href=""CC/EmployeeHistory.asp"" Target=""activeframe"">Employee History</a>" & _
						"<a class=""dropdown-item"" href=""CC/CDMC.asp"" Target=""activeframe"">Corporate Directory</a></ul></li>" 
               
					Response.Write "<li class=""nav-item dropdown""><a class=""nav-link dropdown-toggle"" href=""#"" id=""navbarDropdownMenuLink"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">Dashboards</a>" & _
                          "<div class=""dropdown-menu"" aria-labelledby=""navbarDropdownMenuLink""><a class=""dropdown-item"" href=""CC/DashboardG1.asp"" Target=""activeframe"">Dashboard Sample 1</a>" & _
						  "<a class=""dropdown-item"" href=""CC/DashboardSB2.asp"" Target=""activeframe"">Dashboard Sample 2</a>" & _
                          "<a class=""dropdown-item"" href=""CC/DashboardFrest1.asp"" Target=""activeframe"">Dashboard Sample 3</a>" & _
						  "<a class=""dropdown-item"" href=""CC/DashboardFrest2.asp"" Target=""activeframe"">Dashboard Sample 4</a>" & _
						  "<a class=""dropdown-item"" href=""app-invoice-list.html"" Target=""activeframe"">Invoice List</a>" & _
						  "</div></li>"
					
					Response.Write "<li class=""nav-item dropdown""><a class=""nav-link dropdown-toggle"" href=""#"" id=""navbarDropdownMenuLink"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">Administration</a>" & _
                          "<div class=""dropdown-menu"" aria-labelledby=""navbarDropdownMenuLink"">" & _
						  "<a class=""dropdown-item"" href=""Admin/User.asp"" Target=""activeframe"">Users</a>" & _
                          "<a class=""dropdown-item"" href=""Admin/AdminMenu.asp"" Target=""activeframe"">Admin Menu</a>" & _
						  "<a class=""dropdown-item"" href=""Admin/CAPSAdmin/UploadCS2.asp"" Target=""activeframe"">Upload CS From Diners File</a>" & _
						  "<a class=""dropdown-item"" href=""Admin/CAPSAdmin/UploadANZ.asp"" Target=""activeframe"">Upload ANZ Cardlist File</a>" & _
						  "<a class=""dropdown-item"" href=""Admin/CAPSAdmin/UploadCDMC.asp"" Target=""activeframe"">Upload CDMC File</a>" & _
						  "<a class=""dropdown-item"" href=""Admin/CAPSAdmin/UploadROMAN.asp"" Target=""activeframe"">Upload ROMAN File</a>" & _
						  "<a class=""dropdown-item"" href=""Admin/CAPSAdmin/ExportCS.asp"" Target=""activeframe"">Export CS To Diners File</a>" & _
						  "<a class=""dropdown-item"" href=""Admin/CAPSAdmin/ExportNA.asp"" Target=""activeframe"">Export NA To Diners File</a>" & _
						  "<a class=""dropdown-item"" href=""Admin/CAPSAdmin/ExportANZ.asp"" Target=""activeframe"">Export ANZ Cardlist File</a>" & _
						  "</div></li>"
						  
					'Response.Write "<li class=""dropdown nav-item"" data-menu=""dropdown""><a class=""dropdown-toggle nav-link"" href=""#"" data-toggle=""dropdown""><span>Administration</span></a>" & _
					'				"<ul class=""dropdown-menu"">" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center"" Target=""activeframe"" href=""Admin/User.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Users</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center"" Target=""activeframe"" href=""Admin/AdminMenu.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Admin Menu</a></li>" & _
					'				"<li class=""dropdown dropdown-submenu"" data-menu=""dropdown-submenu""><a class=""dropdown-item align-items-center dropdown-toggle"" href=""#"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Imports</a>" & _
					'				"<ul class=""dropdown-menu"">" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center"" Target=""activeframe"" href=""Admin/CAPSAdmin/UploadCS2.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Upload CS From Diners File</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center"" Target=""activeframe"" href=""Admin/CAPSAdmin/UploadANZ.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Upload ANZ Cardlist File</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center"" Target=""activeframe"" href=""Admin/CAPSAdmin/UploadCDMC.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Upload CDMC File</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center"" Target=""activeframe"" href=""Admin/CAPSAdmin/UploadROMAN.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Upload ROMAN File</a></li>" & _
					'				"</ul></li>" & _
					'				"<li class=""dropdown dropdown-submenu"" data-menu=""dropdown-submenu""><a class=""dropdown-item align-items-center dropdown-toggle"" href=""#"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Exports</a>" & _
					'				"<ul class=""dropdown-menu"">" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center"" Target=""activeframe"" href=""Admin/CAPSAdmin/ExportCS.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Export CS To Diners File</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center"" Target=""activeframe"" href=""Admin/CAPSAdmin/ExportNA.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Export NA To Diners File</a></li>" & _
					'				"<li data-menu=""""><a class=""dropdown-item align-items-center"" Target=""activeframe"" href=""Admin/CAPSAdmin/ExportANZ.asp"" data-toggle=""dropdown""><i class=""bx bx-right-arrow-alt""></i>Export ANZ Cardlist File</a></li>" & _
					'				"</ul></li>"
									
					'Response.write "<li class=""nav-item"" data-toggle=""tooltip"" data-placement=""right"" title=""Administration Menu"">" & _
					'		  "<a class=""nav-link nav-link-collapse collapsed"" data-toggle=""collapse"" href=""#collapseExamplePages"" data-parent=""#exampleAccordion"">" & _
					'			"<i class=""fa fa-fw fa-file""></i>" & _
					'			"<span class=""nav-link-text"">Administration</span></a>" & _
					'		  "<ul class=""sidenav-second-level collapse"" id=""collapseExamplePages"">" & _
					'			"<li><a Target=""activeframe"" href=""CC/CSToDiners.asp?CSToDinersID=0"">CS To Diners File</a></li>" & _
					'			"<li><a Target=""activeframe"" href=""CC/CSFromDiners.asp?CSFromDinersID=0"">CS From Diners File</a></li>" & _
					'			"<li><a Target=""activeframe"" href=""CC/ANZUpdates.asp?ANZUpdateID=0"">ANZ Updates</a></li>" & _
					'			"<li><a Target=""activeframe"" href=""../Admin/User.asp"">Users</a></li>" & _
					'			"<li><a Target=""activeframe"" href=""../Admin/AdminMenu.asp"">Admin Menu</a></li>" & _
					'			"<li><a Target=""activeframe"" href=""../Admin/CAPSAdmin/UploadCDMC.asp"">Upload CDMC Data</a></li>" & _
					'			"<li><a Target=""activeframe"" href=""../Admin/CAPSAdmin/UploadCS.asp"">Upload CS From Diners File</a></li>" & _
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
	<div class="container">
    <div class="app-content content">
        <div class="content-overlay"></div>
        <div class="content-wrapper" style="padding-top:0.5rem; margin-top:1rem;">
          
            <div class="content-body" Style="height:800px;">
                <!-- Dashboard Analytics Start -->

	<!--Loading Wait Spinner-->
	<div class="modal fade bd-example-modal-lg modalWait" id="ModalWait" data-backdrop="static" data-keyboard="false" tabindex="-1">
    <div class="modal-dialog modal-sm">
        <div class="modal-content" style="width: 88px">
            <span style="color:black;" class="spinner-border spinner-border-lg"></span>
        </div>
    </div>
</div>
				<div id="wait" style="display: none;position: absolute;width: 350;height: 100;margin-left: 300;margin-top: 150;background-color: #FFFFFF; text-align: center; color:#333366; line-height:80px; vertical-align:middle; border: solid 1px #333366;">
               <img src="images/Load.gif" style="vertical-align:middle;" /> &nbsp;&nbsp;Please wait while loading...</div>
					
                <iframe height="95%" width="100%" src="<%=Session("HomePageTop")%>" name="activeframe" style="border:0;"></iframe>

                <!-- Dashboard Analytics end -->

            </div>
        </div>
    </div>
	</div>
    <!-- END: Content-->

    <footer class="footer myfi-footer" style="padding:10px;">
      <div class="container">
        <div class="row">
          <div class="col-md-3">
            <img
              src="images/defence_logo_dark.png"
              class="defence-logo"
              alt="Department of Defence"
            />
            <a class="btn btn-primary btn-myfi" href="#"
              ><i class="fa fa-home"></i> Back to MyFi Portal</a
            >
          </div>
          <div class="col-md-3">
            <h4 class="footer-title">Shortcuts</h4>
            <ul class="footer-nav">
              <li class="nav-item">
                <a class="nav-link" href="<%=Session("HomePageTop")%>" Target="activeframe">Dashboard</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="CC/Applications3.asp" Target="activeframe">Applications</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="CC/Cards3_2.asp?View=All" Target="activeframe">Cards</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="#">Reporting</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="CC/HomeCC2.asp" Target="activeframe">My details</a>
              </li>
            </ul>
          </div>
          <div class="col-md-3">
            <h4 class="footer-title">Getting started</h4>
            <ul class="footer-nav">
              <li class="nav-item">
                <a class="nav-link" href="#">Frequently Asked Questions</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="#">How to Guide</a>
              </li>
            </ul>
          </div>
          <div class="col-md-3">
            <h4 class="footer-title">Help</h4>
            <ul class="footer-nav">
              <li class="nav-item">
                <a class="nav-link" href="#">Frequently Asked Questions</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="#">How to Guide</a>
              </li>
              <li class="nav-item">
                <a class="nav-link" href="#">Contact</a>
              </li>
            </ul>
          </div>
        </div>
        <div class="row">
          <div class="col-12 text-center">
            <span class="copyright">© Department of Defence - Defence Finance Group 2020</span>
          </div>
        </div>
      </div>
    </footer>

   <!--<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>-->
	<script src="js/jquery.js"></script>
    <script src="bootstrap/js/bootstrap.bundle.min.js"></script>
    <script>
      jQuery(function ($) {
        $('[data-toggle="popover"]').popover();
      });
    </script>
  </body>
</html>
