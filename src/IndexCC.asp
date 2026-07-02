<%@ Language=VBScript %>

<%  
    'File:			Default.asp
	'Written By:	Andrew Bull JMJ - Isidore www.isidore.com
	'Written On:	November 2007
	'Edit History:	
	'Purpose:	Application Start Screen.
	
	Response.Expires = -1500

   Session("DBConnection") = "File Name=" & Server.MapPath("Database/IsidoreGOL.udl") & ";"
    'Set Database Connection
    Select Case Request.QueryString("ClientID") 

        Case 1
		     Session("DBConnection") = "File Name=" & Server.MapPath("Database/IsidoreGOL.udl") & ";"
            Session("BAName") = "Business Area"
            Session("CCName") = "Cost Centre"
            Session("BANameShort") = "BA"
            Session("CCNameShort") = "CC"
            
                Session("Segment1") = "0"
                Session("Segment2") = "0"
                Session("Segment3") = "0"
                Session("Segment4") = "0"
                Session("Segment5") = "0"
                Session("Segment6") = "0"
                Session("Segment7") = "0"
                Session("Segment8") = "0"
                Session("Segment9") = "0"
                Session("Segment10") = "0" 
                Session("Level1ID") = 0  
                Session("ReportID") = 0
                Session("InputSheetID") = 1      

        Case 2
            Session("DBConnection") = "File Name=" & Server.MapPath("Database/IsidoreSUM.udl") & ";"
            Session("BAName") = "BA"
            Session("CCName") = "CC"
            Session("BANameShort") = "BA"
            Session("CCNameShort") = "CC"

                Session("Segment1") = "0"
                Session("Segment2") = "0"
                Session("Segment3") = "0"
                Session("Segment4") = "0"
                Session("Segment5") = "0"
                Session("Segment6") = "0"
                Session("Segment7") = "0"
                Session("Segment8") = "0"
                Session("Segment9") = "0"
                Session("Segment10") = "0" 
                Session("Level1ID") = 0  
                Session("ReportID") = 0
                Session("InputSheetID") = 1
        Case 3
            Session("DBConnection") = "File Name=" & Server.MapPath("Database/IsidoreMOF.udl") & ";"
            Session("BAName") = "BA"
            Session("CCName") = "CC"
            Session("BANameShort") = "BA"
            Session("CCNameShort") = "CC"
            
                Session("Segment1") = "0"
                Session("Segment2") = "0"
                Session("Segment3") = "0"
                Session("Segment4") = "0"
                Session("Segment5") = "0"
                Session("Segment6") = "0"
                Session("Segment7") = "0"
                Session("Segment8") = "0"
                Session("Segment9") = "0"
                Session("Segment10") = "0" 
                Session("Level1ID") = 0  
                Session("ReportID") = 0
                Session("InputSheetID") = 1         
     	Case else

		
    End Select

    If Session("Logon") = Null or IsEmpty(Session("Logon")) Then

     
    	
	    Dim objCon
	    Dim objRS
        
	    Set objCon = Server.CreateObject("ADODB.Connection")
	    Set objRS = Server.CreateObject("ADODB.Recordset")	    
	
        
	    objCon.Open Session("DBConnection")	   
    	 
	    'Get the default BudgetID from system Defaults table.
	    objRS.Open "SELECT * FROM tblSystemDefault WHERE CurrentDefault = 'Y'",objCon
	    	    
	        If Not objRS.EOF Then
	        
	            Session("BudgetID") = objRS("BudgetID")
    	        Session("RSServer") = objRS("RSServerIPAddress") 
    	        Session("RSFolder") = objRS("RSFolder")
		        'Session("ContingencyWarrantCC") = objRS("ContingencyWarrantCC")
    	        
    	    Else
    	    
    	        
    	    
	        End If
	        
	    objRS.close

    	    Session("Logon") = "IFMISZNZ\Sqlstart"
            'Session("UserTypeID") = 1
            'Session("Logon") = Request.ServerVariables("Auth_User")	
		    'First check whether the user has access.  If they don't then change their logon to READ ONLY and automatically log them in.
		    objRS.Open "SELECT Top 1 * FROM qryUserLogon WHERE UserLogon = '" & Session("Logon") & "' AND Active = 'Y'",objCon
	    	    
	            If objRS.EOF Then
	     	    
    	    		Session("Logon") = "READON"
			
	            End If
		
		    objRS.Close
            'Response.Write "SELECT Top 1 * FROM qryUserLogon WHERE UserLogon = '" & Session("Logon") & "' AND Active = 'Y' AND BudgetID = " & Session("BudgetID") & " Order By DefaultBusinessAreaID Asc"    
	        objRS.Open "SELECT Top 1 * FROM qryUserLogon WHERE UserLogon = '" & Session("Logon") & "' AND Active = 'Y' AND BudgetID = " & Session("BudgetID") & "",objCon
	    	
	            If Not objRS.EOF Then
	               
	                'Set all default variables
	                Session("UserID") = objRS("UserID")
	                Session("UserTypeID") = objRS("UserTypeID")
			        Session("UserType") = objRS("UserTypeName")
	                Session("FName") = objRS("FName")
	                Session("LName") = objRS("LName")
	                Session("BusinessAreaID") = objRS("DefaultBusinessAreaID")
	                Session("BusinessAreaName") = objRS("BusinessAreaName")
	                Session("Currency") = objRS("Currency")
	                Session("CostCentreID") = objRS("DefaultCostCentre")
	                Session("Currency") = objRS("Currency")
    	       	    Session("FullName") = Session("FName") & " " & Session("LName")
    	            Session("ModeID") = 2
    	            Session("FundID") = 1
    	            Session("YearID") = 1
    	            Session("PLReportID") = objRS("PLReportID")
		   	        Session("ProjectID") = 0
		            Session("GLCodeID") = 0
		            Session("Language") = objRS("Language")
                    Session("CCStatusID") = 0
                    Session("TargetID") = 0
    	        
    	            objCon.Execute "INSERT INTO tblLogons VALUES ('" & Session("UserID") & "',GetDate())"
    	           	    
	        Else
            
			    Response.Redirect "AccessDenied.asp"    	    
    	    
	        End If
	    
	    
	    objRS.Close			
		
		'Load System Settings
		
		objRS.Open "SELECT * FROM qrySystemDefault WHERE BudgetID = " & Session("BudgetID") & "",objCon
		
		If Not objRS.EOF Then
		
		    Session("BudgetID") = objRS("BudgetID")
		    Session("BudgetName") = objRS("BudgetName")
		    Session("VersionID") = objRS("DefaultVersionID")
		    Session("BaseBudgetVersionID") = objRS("BaseBudgetVersionID") 
		    Session("VersionTypeID")= objRS("VersionTypeID")
		    Session("FinacialYearID") = objRS("FinancialYearID")
		    Session("FinancialYear") = objRS("FinancialYearName")	
            Session("TargetID") = 0	   
		    
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
		
		Else
		
		    Response.Redirect "AccessDenied.asp"
		
		End If
		
		objRS.Close		
		
		'Get the Column Lock value of the Version
		
		objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & "",objCon
	
	        If not objRS.EOF Then
	            Session("ColumnLock") = objRS("ColumnLock")
			    Session("BaseVersionID") = objRS("BaseBudgetVersionID")
	        End If
	    
	    objRS.Close	

    End If

 
	
	
%>



<html lang="en" xmlns="http://www.w3.org/1999/xhtml">

<head runat="server">
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta name="description" content="" />
    <meta name="author" content="" />

    <title>Isidore</title>

<script type="text/javascript">

function WaitDiv()
        {
         document.getElementById('wait').style.display = 'block';

  setTimeout(fStop, 10000)
  //document.getElementById('wait').style.display = 'none';
            }

function fStop()
        {
        
  document.getElementById('wait').style.display = 'none';
            }
</script>

    <!-- Bootstrap Core CSS -->
    <link href="css/bootstrap.min.css" rel="stylesheet" />

    <!-- Custom Fonts -->
    <link href="font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css" />

    <!-- HTML5 Shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!-- WARNING: Respond.js doesn't work if you view the page via file:// -->
    <!--[if lt IE 9]>
        <script src="https://oss.maxcdn.com/libs/html5shiv/3.7.0/html5shiv.js"></script>
        <script src="https://oss.maxcdn.com/libs/respond.js/1.4.2/respond.min.js"></script>
    <![endif]-->
    <link href="css/Theme1.css" rel="stylesheet" />
    <link href="css/Common.css" rel="stylesheet" />
    <style type="text/css">
        .row {
            margin-left: -15px;
            margin-right: 0px;
        }

        .page-head {
            background: #FFF none repeat scroll 0px 0px;
            border-bottom: 1px solid #E9E9E9;
            box-shadow: 0px 0px 1px 0px rgba(0, 0, 0, 0.05);
            padding: 4px 10px;
            position: relative;
        }

        .form-horizontal.group-border-dashed .form-group {
            margin: 0;
            padding: 5px 0;
            border-bottom: 1px dashed #efefef;
        }

        .ErrorControl {
            background-color: #FBE3E4;
            border: solid 1px Red;
        }
    </style>
    <script src="js/jquery.js"></script>
    <script type="text/javascript" src="js/plugins/jquery.nanoscroller.js"></script>
    <script type="text/javascript" src="js/plugins/verticalmenu.js"></script>
    <script src="js/bootstrap.min.js"></script>
    <script type="text/javascript" src="js/plugins/jquery.flot.js"></script>
    <script type="text/javascript" src="js/plugins/jquery.flot.categories.js"></script>
    <asp:ContentPlaceHolder ID="head" runat="server">
    </asp:ContentPlaceHolder>
</head>
<body scroll="auto">

    <form id="form1" runat="server">
             
        <div id="head-nav" class="navbar navbar-default navbar-fixed-top">
            <div class="container-fluid">
                <div class="navbar-header">
                    <button type="button" data-toggle="collapse" data-target=".navbar-collapse" class="navbar-toggle"><span class="fa fa-gear"></span></button>
                    <a href="Index.asp" target="_parent" class="navbar-brand"><span>Isidore FBT</span></a>
                </div>
                <div class="navbar-collapse collapse">
                    <ul class="nav navbar-nav">
                        <li class="active"><a href="IndexBudget.asp?ClientID=1" target="_parent">Budget</a></li>
                        <li class="active"><a href="Index.asp?ClientID=1" target="_parent">FBT</a></li>
						<li class="active"><a href="IndexCC.asp?ClientID=1" target="_parent">Credit Cards</a></li>
						<li class="active"><a href="IndexCC2.asp?ClientID=1" target="_parent">Credit Cards 2</a></li>
                       <%
					   If Session("UType") = "Manager" Then
					   
							Response.write "<li class=""active""><a href=""ScrollingFrameset.asp?CurrentPage=Admin/AdminMenu.asp&HeaderMenu=Header2.asp"" target=""activeframe"">Admin</a></li>"
					   End If
					   
					   %>
                        
                    </ul>
                   
				   <ul>
				   
				   <li class="nav-item" data-toggle="tooltip" data-placement="right" title="Link">
          <a class="nav-link" href="#">
            <i class="fa fa-fw fa-link"></i>
            <span class="nav-link-text">Link</span>
          </a>
        </li>
      </ul>
      <ul class="navbar-nav sidenav-toggler">
        <li class="nav-item">
          <a class="nav-link text-center" id="sidenavToggler">
            <i class="fa fa-fw fa-angle-left"></i>
          </a>
        </li>
      </ul>
      <ul class="navbar-nav ml-auto">
        <li class="nav-item dropdown">
          <a class="nav-link dropdown-toggle mr-lg-2" id="messagesDropdown" href="#" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
            <i class="fa fa-fw fa-envelope"></i>
            <span class="d-lg-none">Messages
              <span class="badge badge-pill badge-primary">12 New</span>
            </span>
            <span class="indicator text-primary d-none d-lg-block">
              <i class="fa fa-fw fa-circle"></i>
            </span>
          </a>
          

			
                </div>
            </div>
        </div>
        <div id="cl-wrapper" class="fixed-menu">
            <!--Sidebar item function-->
            <!--Sidebar sub-item function-->
            <div data-position="right" data-step="1" data-intro="<strong>Fixed Sidebar</strong> <br/> It adjust to your needs." class="cl-sidebar">
                <div class="cl-toggle"><i class="fa fa-bars"></i></div>
                <div class="cl-navblock">
                    <div style="height: 94px;" class="menu-space nano nscroller has-scrollbar">
                        <div style="right: -17px;" tabindex="0" class="content">
                            <ul class="cl-vnavigation">
                                <li class="parent"><a href="Index.asp" target="_parent" ><i class="fa fa-home"></i><span>Summary Reports</span></a><ul class="sub-menu">
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=Reports/Dashboard.asp&HeaderMenu=Header2.asp" target="activeframe">Dashboard</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=FBT/CarParkingReport.asp&HeaderMenu=Header1.asp" target="activeframe"><%=Session("BANameShort") %> Car Parking Summary</a></li>
									<li><a href="ScrollingFrameset.asp?CurrentPage=FBT/ATOSummary.asp&HeaderMenu=Header2.asp" target="activeframe">ATO Payment Summary</a></li>
									<li><a href="ScrollingFrameset.asp?CurrentPage=FBT/CarParkingStaffReport.asp&HeaderMenu=Header1.asp" target="activeframe"><%=Session("BANameShort") %> Staff Benefit Report Summary</a></li>
                                </ul>
                                </li>                             
                                   
                                <li class="parent"><a href="#"><i class="fa fa-table"></i><span>FBT Individual Input</span></a><ul class="sub-menu">
									 <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=FBT/Loans.asp&HeaderMenu=Header1.asp" target="activeframe">Loans</a></li>
									 <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=FBT/CarParking.asp&HeaderMenu=Header1.asp" target="activeframe">Car Parking</a></li>
										<li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=FBT/Cars.asp&HeaderMenu=Header1.asp" target="activeframe">Car Owned/Leased</a></li>
									<li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=FBT/LAFHA.asp&HeaderMenu=Header1.asp" target="activeframe">Living Away From Home (LAFHA)</a></li>
									<li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=FBT/LAFHA.asp&HeaderMenu=Header1.asp" target="activeframe">Living In Allowance (LIA)</a></li>
									<li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=FBT/Property.asp&HeaderMenu=Header1.asp" target="activeframe">Property</a></li>
									<li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=FBT/CarParking.asp&HeaderMenu=Header1.asp" target="activeframe">Board</a></li>
									<li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=FBT/CarParking.asp&HeaderMenu=Header1.asp" target="activeframe">Other Residual</a></li>
									<li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Stalls.asp&HeaderMenu=Header1.asp" target="activeframe">Stalls</a></li>
									<li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Loans.asp&HeaderMenu=Header1.asp" target="activeframe">Loans 2</a></li>
                                    
                                </ul>
                                </li>
                                
                                <li class="parent"><a href="#"><i class="fa fa-envelope"></i><span>Workflow</span></a><ul class="sub-menu">
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Admin/CostCentreStatus.asp&HeaderMenu=Header2.asp&CCStatusID=0" target="activeframe">All</a></li>
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Admin/CostCentreStatus.asp&HeaderMenu=Header2.asp&CCStatusID=1" target="activeframe">Open</a></li>
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Admin/CostCentreStatus.asp&HeaderMenu=Header2.asp&CCStatusID=2" target="activeframe">Completed</a></li>
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Admin/CostCentreStatus.asp&HeaderMenu=Header2.asp&CCStatusID=3" target="activeframe">Rejected</a></li>
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Admin/CostCentreStatus.asp&HeaderMenu=Header2.asp&CCStatusID=4" target="activeframe">Approved</a></li>
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Admin/CostCentreStatus.asp&HeaderMenu=Header2.asp&CCStatusID=5" target="activeframe">Closed</a></li> 
                                </ul>
                                </li>
                                <li class="parent"><a href="#"><i class="fa fa-file"></i><span>Staff Reports</span></a><ul class="sub-menu">
				                    <li class="active" ><A HREF="ScrollingFrameset.asp?CurrentPage=Reports/StaffDataVarianceBusinessArea.asp&HeaderMenu=Header2.asp" target="activeframe">&nbsp;<%=Session("BAName") %> Employee Summary</A></li>
                                
								
                                </ul>
								
								<%
					   If Session("UType") = "Manager" Then
					   
							Response.write "<li class=""parent""><a href=""#""><i class=""fa fa-cog""></i><span>Admin</span></a><ul class=""sub-menu"">" & _
                                    "<li class=""active"" ><a href=""ScrollingFrameset.asp?CurrentPage=ProfitLoss/DataEntry6.asp&HeaderMenu=Header1.asp&CCStatusID=0"" target=""activeframe"">Staff</a></li>" & _
                                    "<li class=""active"" ><a href=""ScrollingFrameset.asp?CurrentPage=Admin/CostCentreStatus.asp&HeaderMenu=Header2.asp&CCStatusID=1"" target=""activeframe"">Users</a></li>"& _
                                    "<li class=""active"" ><a href=""ScrollingFrameset.asp?CurrentPage=Admin/CostCentreStatus.asp&HeaderMenu=Header2.asp&CCStatusID=2"" target=""activeframe"">Companies</a></li>"& _
									"</ul></li>"
					   End If
					   
					   %>
								
                            </ul>
                        </div>
                        <div style="display: none;" class="pane">
                            <div style="height: 20px; top: 0px;" class="slider"></div>
                        </div>
                    </div>
                    <div class="search-field collapse-button">
                        <input placeholder="Search..." class="form-control search" type="text" style="visibility: hidden;">
                        <button id="sidebar-collapse" class="btn btn-default" onclick="return false;"><i class="fa fa-angle-left"></i></button>
                    </div>
                </div>
            </div>
            <div id="pcont" class="container-fluid">
                   	<div id="wait" style="display: none;position: absolute;width: 350;height: 100;margin-left: 300;margin-top: 150;background-color: #FFFFFF; text-align: center; color:#333366; line-height:80px; vertical-align:middle; border: solid 1px #333366;">
                    <img src="images/Load.gif" style="vertical-align:middle;" /> &nbsp;&nbsp;Please wait while loading...</div>
                <iframe height="950" width="100%" src="Home.asp" name="activeframe" ></iframe>
            </div>
        </div>

		

        <script type="text/javascript">$(document).ready(function () {
    var url = window.location;
    $(".content li").removeClass("active");//this will remove the active class from
    $(".content li sub-menu").css("display", "none");
    $('.content li a').each(function () {
        if (this.href == url) {
            $(this).parent().addClass('active');
            $(this).parent().parent().css("display", "");
        }
    });

    //initialize the javascript
    App.init();
    App.dashboard();


});</script>
    </form>
    
</body>
</html>