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
            Session("BAName") = "Vote"
            Session("CCName") = "Sub Vote"
            Session("BANameShort") = "Vote"
            Session("CCNameShort") = "Sub Vote"
            
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
                Session("InputSheetID") = 0         
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
                    <a href="Index.asp" target="_parent" class="navbar-brand"><span>Isidore CBMS</span></a>
                </div>
                <div class="navbar-collapse collapse">
                    <ul class="nav navbar-nav">
                        <li class="active"><a href="Index.asp?ClientID=1" target="_parent">MOF</a></li>
                        <li class="active"><a href="Index.asp?ClientID=2" target="_parent">LGA</a></li>
						<li class="active"><a href="IndexCC4.asp?ClientID=2" target="_parent">CAPS</a></li>
 
                       
                        <li class="active"><a href="ScrollingFrameset.asp?CurrentPage=Admin/AdminMenu.asp&HeaderMenu=Header2.asp" target="activeframe">Admin</a></li>
                    </ul>
                    <ul class="nav navbar-nav navbar-right user-nav">
                        <li class="dropdown profile_menu"><a href="#" data-toggle="dropdown" class="dropdown-toggle">
                            <span>
                                <asp:Label ID="lblCurrentUser" runat="server"><%=Session("Logon") %></asp:Label></span><b class="caret"></b></a>
                            <ul class="dropdown-menu">
                                <li><a href="#">My Account</a></li>
                                <li><a href="#">Profile</a></li>
                                <li class="divider"></li>
                                <li><a href="" onclick="window.close()">Sign Out</a></li>
                            </ul>
                        </li>
                    </ul>
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
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/PLCostCentre1Quarterly.asp&HeaderMenu=Header1.asp" target="activeframe"><%=Session("CCName") %> Summary</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/PLBusinessArea1Quarterly.asp&HeaderMenu=Header2.asp" target="activeframe"><%=Session("BAName") %> Summary</a></li>
									<li><a href="ScrollingFrameset.asp?CurrentPage=Reports/PLBusinessArea1QuarterlyChart.asp&HeaderMenu=Header2.asp" target="activeframe"><%=Session("BAName") %> Summary Chart</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/PLBusinessArea1QuarterlyBaseBudget.asp&HeaderMenu=Header2.asp" target="activeframe"><%=Session("BAName") %> Original vs Revised Var</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/PLBusinessArea1QuarterlyCOFOG.asp&HeaderMenu=Header2.asp" onclick="WaitDiv();" target="activeframe"><%=Session("BAName") %> COFOG</a></li>                                    
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/PLCostCentre2QuarterlyByLineItem.asp&HeaderMenu=Header1.asp" target="activeframe"><%=Session("CCName") %> Summary By Line Item</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/PLBusinessArea2QuarterlyByLineItem.asp&HeaderMenu=Header2.asp" target="activeframe"><%=Session("BAName") %> Summary By Line Item</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=BudgetTransfers/VirementsSummary.asp&HeaderMenu=Header2.asp&TransferDocumentTypeID=1" target="activeframe">Virement Summary</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=BudgetTransfers/VirementsSummary.asp&HeaderMenu=Header2.asp&TransferDocumentTypeID=2" target="activeframe">Reallocation Summary</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/PLProject1Quarterly.asp&HeaderMenu=Header1.asp" target="activeframe">Project Summary</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=Strategy/StrategicPrioritySummary.asp&HeaderMenu=Header2.asp" target="activeframe">Strategic Priority Summary</a></li>
                                </ul>
                                </li>                             
                                    <li class="parent"><a href="#"><i class="fa fa-list-alt"></i><span>Expenditure Ceilings</span></a><ul class="sub-menu">
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=Indexation/MTFF.asp&HeaderMenu=Header3.asp" target="activeframe">Fiscal Framework</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=Appropriation/AppropriationControlPanel.asp&HeaderMenu=Header2.asp" target="activeframe">Ceiling Allocation</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/DataEntry2BCO.asp&HeaderMenu=Header1.asp&TransactionType=GEXP&Level1ID=0&CalculatedField=''" target="activeframe">Block Costings</a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=Indexation/BusinessAreaIndexation.asp&HeaderMenu=Header2.asp" target="activeframe"><%=Session("BAName") %> Ceilings </a></li>
                                    <li><a href="ScrollingFrameset.asp?CurrentPage=Indexation/CostCentreIndexation.asp&HeaderMenu=Header2.asp" target="activeframe"><%=Session("CCName") %> Ceilings </a></li> 
                                   
                                </ul>
                                </li>
                                <li class="parent"><a href="#"><i class="fa fa-table"></i><span>Budget Input</span></a><ul class="sub-menu">
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Strategy/VoteSetUp.asp&HeaderMenu=Header2.asp" target="activeframe">Vision & Mission</a></li>
									<li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Strategy/StrategicObjectives.asp&HeaderMenu=Header2.asp" target="activeframe">Objectives</a></li>
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Strategy/Strategies.asp&HeaderMenu=Header2.asp" target="activeframe">Strategies</a></li>
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Strategy/SmartTarget.asp&HeaderMenu=Header2.asp" target="activeframe">Activities</a></li>
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=Strategy/TargetPerformance.asp&HeaderMenu=Header2.asp" target="activeframe">Performance Indicators</a></li>
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/DataEntry2MOF.asp&HeaderMenu=Header1.asp&TransactionType=GEXP&Level1ID=0&CalculatedField=''" target="activeframe">MDA Input Sheet</a></li>
                                
					                <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/DataEntry6.asp&HeaderMenu=Header1.asp&TransactionType=ESTA" target="activeframe">Existing Staff</a></li>
					                <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/DataEntry6.asp&HeaderMenu=Header1.asp&TransactionType=NSTA" target="activeframe">New Staff</a></li> 
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=ProfitLoss/Bids.asp&HeaderMenu=Header2.asp" target="activeframe">Projects</a></li>
                                    
                                </ul>
                                </li>
                                <li class="parent"><a href="#"><i class="fa fa-map-marker"></i><span>Budget Management</span></a><ul class="sub-menu">
                                     <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=BudgetTransfers/TransferDocumentsList.asp?TransferDocumentTypeID=1&HeaderMenu=Header2.asp" target="activeframe">Reallocations</a></li>
                                    <li class="active" ><a href="ScrollingFrameset.asp?CurrentPage=BudgetTransfers/TransferDocumentsList.asp?TransferDocumentTypeID=2&HeaderMenu=Header2.asp" target="activeframe">Virements</a></li>
                                     <li><a href="ScrollingFrameset.asp?CurrentPage=Indexation/Forecasting.asp&HeaderMenu=Header2.asp" target="activeframe">Warrant Release </a></li> 
									<li><a href="ScrollingFrameset.asp?CurrentPage=Reports/BAFundsReleasedLine.asp&HeaderMenu=Header2.asp" target="activeframe">Funds Released Report</a></li> 
									<li><a href="ScrollingFrameset.asp?CurrentPage=Reports/BACashForecast.asp&HeaderMenu=Header2.asp" target="activeframe">Cash Forecast Report</a></li> 
									<li><a href="ScrollingFrameset.asp?CurrentPage=Reports/WarrantForecastReport.asp&HeaderMenu=Header2.asp" target="activeframe">Warrant Release Report</a></li> 
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
                                    <li class="active" ><A HREF="ScrollingFrameset.asp?CurrentPage=Reports/StaffDataCostCentre.asp&HeaderMenu=Header1.asp" target="activeframe">&nbsp;<%=Session("CCName") %> Staff FTE Report</A></li>
                                    <li class="active" ><A HREF="ScrollingFrameset.asp?CurrentPage=Reports/StaffDataBusinessArea.asp&HeaderMenu=Header2.asp" target="activeframe">&nbsp;<%=Session("BAName") %> Staff FTE Report</A></li>
				                    <li class="active" ><A HREF="ScrollingFrameset.asp?CurrentPage=Reports/StaffDataVarianceBusinessArea.asp&HeaderMenu=Header2.asp" target="activeframe">&nbsp;<%=Session("BAName") %> Variance Staffing Report</A></li>
				                    <li class="active" ><A HREF="ScrollingFrameset.asp?CurrentPage=Reports/StaffCostsVarianceBusinessArea.asp&HeaderMenu=Header2.asp" target="activeframe">&nbsp;Position Cost Report</A></li>
									<li class="active" ><A HREF="ScrollingFrameset.asp?CurrentPage=Reports/StaffCostsVarianceBusinessAreaChart.asp&HeaderMenu=Header2.asp" target="activeframe">&nbsp;Position Cost Report Chart</A></li>
									<li class="active" ><A HREF="ScrollingFrameset.asp?CurrentPage=Reports/StaffCostsVarianceBusinessAreaPie.asp&HeaderMenu=Header2.asp" target="activeframe">&nbsp;Position Cost Report Pie</A></li>
                                    <li class="active" ><A HREF="ScrollingFrameset.asp?CurrentPage=Reports/CostCentreStaffCostReport.asp&HeaderMenu=Header1.asp" target="activeframe">&nbsp;<%=Session("CCName") %> Staff Cost Report</A></li>
                                    <li class="active" ><A HREF="ScrollingFrameset.asp?CurrentPage=Admin/EstablishmentFundingSummary.asp&HeaderMenu=Header2.asp&Header=Header1.asp" target="activeframe">&nbsp;Establishments Report</A></li>
                                    
                                </ul>
                                </li>
                                    <li class="parent"><a href="#"><i class="fa fa-file"></i><span>Budget Papers</span></a><ul class="sub-menu">
                                   
                                    
                                </ul>
                                </li>
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