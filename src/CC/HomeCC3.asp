<%@ Language=VBScript %>
<% Option Explicit
	
	Response.Expires = -1500

   'If IsEmpty(Session("Logon")) Then Response.Redirect("AccessDenied.asp")
   'If IsEmpty(Session("BudgetID")) Then Response.Redirect("Timeout.asp")
	If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp?State=Expired")

	'Session("CurrentPage") = "CC/HomeCC2.asp"
	
Dim objCon
Dim objRS
Dim objRS1
Dim objRS2
Dim strSelected
Dim arrStatusID(5)
Dim arrStatusName(5)
Dim arrStatusImg(5)
Dim arrCompanyID(2)
Dim arrCompanyName(2)
Dim x
Dim lngBudgetStatus
Dim lngBaseBudget
Dim lngBudget
Dim strColour 

Dim strManagerName
Dim intManagerID
Dim strManCStatus

Dim strApply, strGCFO, strDCC, strBank
Dim strApplyClass, strGCFOClass, strDCCClass, strBankClass
Dim strBankName

Dim strGCFODays, strDCCDays, strBankDays


arrStatusName(1) = "Open"
arrStatusName(2) = "Completed"
arrStatusName(3) = "Rejected"
arrStatusName(4) = "Approved"
arrStatusName(5) = "Closed"

arrStatusID(1) = 1
arrStatusID(2) = 2
arrStatusID(3) = 3
arrStatusID(4) = 4
arrStatusID(5) = 5

arrStatusImg(1) = "<IMG SRC='images/open.png'"
arrStatusImg(2) = "<IMG SRC='images/ready.gif'"
arrStatusImg(3) = "<IMG SRC='images/cross.png'" 
arrStatusImg(4) = "<IMG SRC='images/tick.png'"	
arrStatusImg(5) = "<IMG SRC='images/Closed.png'"
    			

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")
Set objRS2 = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

	
	If Not IsEmpty(Request.QueryString("ApplicationID")) Then
		Session("ApplicationID") = Request.QueryString("ApplicationID")
	End If
	
	If Not IsEmpty(Request.QueryString("HomeCC")) Then
		Session("HomeCC") = Request.QueryString("HomeCC")
	End If
	
	If Not IsEmpty(Request.QueryString("UType")) Then
		Session("UType") = Request.QueryString("UType")
	End If
	
	'Set the home page details based on the user type
'	If Session("UType") = "Employee" Then
'		'Session("UType") = "Employee"
'		Session("HomePage1") = "MyCards.asp"
'		Session("HomePage2") = "MyApplications.asp"
'	ElseIf Session("UType") = "Manager" Then
'		'Session("UType") = "Manager"
'		Session("HomePage1") = "CardTypeChart.asp"
'		Session("HomePage2") = "ApplicationsHome.asp"
'	Else
'		'Session("UType") = "CreditCards"
'		Session("HomePage1") = "CardTypeChart.asp"
'		Session("HomePage2") = "ApplicationsHome.asp"
'	End If
					
 %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">

 <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
  <meta name="description" content="">
  <meta name="author" content="">
  <title>Cards Home</title>
  <!-- Bootstrap core CSS-->
  <link href="../vendor/bootstrap/css/bootstrap.min.css" rel="stylesheet">
  <!-- Custom fonts for this template-->
  <link href="../vendor/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css">
  <!-- Custom styles for this template-->
  <link href="../css/sb-admin.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
  <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.0/css/bootstrap.min.css">

   <!-- Bootstrap -->
    <link href="../Temps/Gent/vendors/bootstrap/dist/css/bootstrap.min.css" rel="stylesheet">
    <!-- Font Awesome -->
    <link href="../Temps/Gent/vendors/font-awesome/css/font-awesome.min.css" rel="stylesheet">
    <!-- NProgress -->
    <link href="../Temps/Gent/vendors/nprogress/nprogress.css" rel="stylesheet">
    <!-- iCheck -->
    <link href="../Temps/Gent/vendors/iCheck/skins/flat/green.css" rel="stylesheet">
	
    <!-- bootstrap-progressbar -->
    <link href="../Temps/Gent/vendors/bootstrap-progressbar/css/bootstrap-progressbar-3.3.4.min.css" rel="stylesheet">
    <!-- JQVMap -->
    <link href="../Temps/Gent/vendors/jqvmap/dist/jqvmap.min.css" rel="stylesheet"/>
    <!-- bootstrap-daterangepicker -->
    <link href="../Temps/Gent/vendors/bootstrap-daterangepicker/daterangepicker.css" rel="stylesheet">

    <!-- Custom Theme Style -->
    <link href="../Temps/Gent/build/css/custom.min.css" rel="stylesheet">
	
</HEAD>
<BODY>
<FORM action="Home.asp" method="POST" id="frm" name="frm">

<div class="container-fluid" style="background:white;">
     
      <!-- page content -->
        <div style="background:white;" >
	 
	 <div class="col-md-3 col-sm-3 ">
              <div class="x_panel tile fixed_height_220 overflow_hidden" style="background:white;">
			<div class="x_content" style="background:white;">
                  <table class="" style="width:100%" style="background:white;">
                    <tr>
                      <th style="width:37%;">
                        <p>Applications</p>
                      </th>
                      <th>
                        <div class="col-lg-7 col-md-7 col-sm-7 ">
                          <p class="">Application Type</p>
                        </div>
                        <div class="col-lg-5 col-md-5 col-sm-5 ">
                          <p class="">Percentage</p>
                        </div>
                      </th>
                    </tr>
                    <tr>
                      <td>
                        <canvas class="canvasDoughnut" height="140" width="140" style="margin: 15px 10px 10px 0"></canvas>
                      </td>
                      <td>
                        <table class="tile_info">
                          <tr>
                            <td>
                              <p><i class="fa fa-square blue"></i>Defence Travel Card </p>
                            </td>
                            <td>30%</td>
                          </tr>
                          <tr>
                            <td>
                              <p><i class="fa fa-square green"></i>Defence Purchasing Card </p>
                            </td>
                            <td>10%</td>
                          </tr>
                          <tr>
                            <td>
                              <p><i class="fa fa-square purple"></i>Companion Master Card </p>
                            </td>
                            <td>20%</td>
                          </tr>
                          <tr>
                            <td>
                              <p><i class="fa fa-square aero"></i>CTS Account </p>
                            </td>
                            <td>15%</td>
                          </tr>
                          <tr>
                            <td>
                              <p><i class="fa fa-square red"></i>Limit Increases </p>
                            </td>
                            <td>30%</td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </div>
			</div>
		</div>	
		 <div class="col-md-2 col-sm-2 ">
              <div class="x_panel tile fixed_height_220 overflow_hidden">		
				<div class="sidebar-widget">
                        <h4>Application Approvals Rate</h4>
                        <canvas width="150" height="100" id="chart_gauge_01" class="" style="width: 160px; height: 100px;"></canvas>
                        <div class="goal-wrapper">
                          <span id="gauge-text" class="gauge-value pull-left">0</span>
                          <span class="gauge-value pull-left"> &nbsp;Applications</span>
                          <span id="goal-text" class="goal-value pull-right">100%</span>
                        </div>
                    </div>
			</div>
		</div>
		
		 <!-- Start to do list -->
                <div class="col-md-5 col-sm-5 ">
                  <div class="x_panel">
                    <div class="x_title">
                      <h2>To Do List <small>Application Approvals Awaiting Approval by <%=Session("Username") %></small></h2>
                      <ul class="nav navbar-right panel_toolbox">
                        <li><a class="collapse-link"><i class="fa fa-chevron-up"></i></a>
                        </li>
                        <li class="dropdown">
                          <a href="#" class="dropdown-toggle" data-toggle="dropdown" role="button" aria-expanded="false"><i class="fa fa-wrench"></i></a>
                          <div class="dropdown-menu" aria-labelledby="dropdownMenuButton">
                              <a class="dropdown-item" href="#">Show Outstanding Only</a>
                              <a class="dropdown-item" href="#">Show All</a>
                            </div>
                        </li>
                        <li><a class="close-link"><i class="fa fa-close"></i></a>
                        </li>
                      </ul>
                      <div class="clearfix"></div>
                    </div>
                    <div class="x_content">

                      <div class="">
                        <ul class="to_do">
                          <li>
                            <p>
                              <input type="checkbox" class="flat"> Glen Keyes DPC Application - 5 Days old</p>
                          </li>
                          <li>
                            <p>
                              <input type="checkbox" class="flat"> Johan Bankisnozer DPC Application - 10 Days old</p>
                          </li>
                          
                          <li>
                            <p>
                              <input type="checkbox" class="flat"> CTS Application for Army - <span style="color:red;">14 Days old</span></p>
                          </li>
                          
                        </ul>
                      </div>
                    </div>
                  </div>
                </div>
                <!-- End to do list -->
		
		<div class="tile_count">
            <div class="col-md-2 col-sm-4  tile_stats_count">
              <span class="count_top"><i class="fa fa-user"></i> Total Applications</span>
              <div class="count">2,500</div>
              <span class="count_bottom"><i class="green">4% </i> From last Week</span>
            </div>
            <div class="col-md-2 col-sm-4  tile_stats_count">
              <span class="count_top"><i class="fa fa-clock-o"></i> Average Processing Days</span>
              <div class="count">12.50</div>
              <span class="count_bottom"><i class="green"><i class="fa fa-sort-asc"></i>3% </i> From last Week</span>
            </div>
           
          </div>
		  
<HR>
<TABLE Align="Center" WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">	

	<TR>
		<TH Style="Width:40%;Height:25px;text-align:center;color:white;background-color:#86c5f9;"><a href="Cards3.asp?View=Employee" Style="color:white;background-color:#86c5f9; font-size:24px;"><i class="fa fa-credit-card"></i> My Cards</a></TH>
		<TH Style="Width:40%;Height:25px;text-align:center;background-color:#86c5f9;"><a href="ApplicationsEmployee.asp" Style="color:white;background-color:#86c5f9; font-size:24px;"><i class="fa fa-address-card"></i> My Applications</a></TH>
	</TR>
	<TR>
		<TD Colspan="2">&nbsp;</TD> 
	</TR>
	<tr>
		<TD ><iframe id="Iframe1" name="framecontent" src="<%=Session("HomePage1")%>" Width="100%" frameborder="0" height="500px"></iframe></TD>
		<TD ><iframe id="framecontent" name="framecontent" src="<%=Session("HomePage2")%>" Width="100%" frameborder="0" height="500px"></iframe></TD>
	</TR>
	
   <tr><th Style="Height:5px;"  colspan="2">&nbsp;</th></tr>
</TABLE>
</div>

 <!-- jQuery -->
    <script src="../Temps/Gent/vendors/jquery/dist/jquery.min.js"></script>
    <!-- Bootstrap -->
    <script src="../Temps/Gent/vendors/bootstrap/dist/js/bootstrap.bundle.min.js"></script>
    <!-- FastClick -->
    <script src="../Temps/Gent/vendors/fastclick/lib/fastclick.js"></script>
    <!-- NProgress -->
    <script src="../Temps/Gent/vendors/nprogress/nprogress.js"></script>
    <!-- Chart.js -->
    <script src="../Temps/Gent/vendors/Chart.js/dist/Chart.min.js"></script>
    <!-- gauge.js -->
    <script src="../Temps/Gent/vendors/gauge.js/dist/gauge.min.js"></script>
    <!-- bootstrap-progressbar -->
    <script src="../Temps/Gent/vendors/bootstrap-progressbar/bootstrap-progressbar.min.js"></script>
    <!-- iCheck -->
    <script src="../Temps/Gent/vendors/iCheck/icheck.min.js"></script>
    <!-- Skycons -->
    <script src="../Temps/Gent/vendors/skycons/skycons.js"></script>
    <!-- Flot -->
    <script src="../Temps/Gent/vendors/Flot/jquery.flot.js"></script>
    <script src="../Temps/Gent/vendors/Flot/jquery.flot.pie.js"></script>
    <script src="../Temps/Gent/vendors/Flot/jquery.flot.time.js"></script>
    <script src="../Temps/Gent/vendors/Flot/jquery.flot.stack.js"></script>
    <script src="../Temps/Gent/vendors/Flot/jquery.flot.resize.js"></script>
    <!-- Flot plugins -->
    <script src="../Temps/Gent/vendors/flot.orderbars/js/jquery.flot.orderBars.js"></script>
    <script src="../Temps/Gent/vendors/flot-spline/js/jquery.flot.spline.min.js"></script>
    <script src="../Temps/Gent/vendors/flot.curvedlines/curvedLines.js"></script>
    <!-- DateJS -->
    <script src="../Temps/Gent/vendors/DateJS/build/date.js"></script>
    
    <!-- bootstrap-daterangepicker -->
    <script src="../Temps/Gent/vendors/moment/min/moment.min.js"></script>
    <script src="../Temps/Gent/vendors/bootstrap-daterangepicker/daterangepicker.js"></script>

    <!-- Custom Theme Scripts -->
    <script src="../Temps/Gent/build/js/custom.min.js"></script>
	
</BODY>
</HTML>
<%




%>
