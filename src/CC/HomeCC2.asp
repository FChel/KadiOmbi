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
	If Session("UType") = "Employee" Then
		'Session("UType") = "Employee"
		Session("HomePage1") = "MyCards.asp"
		Session("HomePage2") = "MyApplications.asp"
	ElseIf Session("UType") = "Manager" Then
		'Session("UType") = "Manager"
		Session("HomePage1") = "CardTypeChart.asp"
		Session("HomePage2") = "ApplicationsHome.asp"
	Else
		'Session("UType") = "CreditCards"
		Session("HomePage1") = "CardTypeChart.asp"
		Session("HomePage2") = "ApplicationsHome.asp"
	End If
					
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

  
</HEAD>
<BODY>
<FORM action="Home.asp" method="POST" id="frm" name="frm">

<!-- Modal -->
<div class="modal fade" id="ModApp" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">Credit Application Declaration Form FRAME</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>I DECLARE that all the application details are correct.</label><br>
            <label for="name">First Name:</label>
            <input type="text" name="FirstName" id="FirstName" class="form-control input-md">
			<label for="name">Last Name:</label>
            <input type="text" name="LastName" id="LastName" class="form-control input-md">
			<label for="email">Email:</label>
            <input type="email" name="email" id="email" class="form-control input-md">
            <label for="phone">Phone:</label>
            <input type="text" name="phone" id="phone" class="form-control input-md">
			<label for="phone">Group:</label>
            
			<select class="form-control" id="Group">
			  <option>Please select your group...</option>
			  <option>Army</option>
			  <option>Navy</option>
			  <option>Air Force</option>
			  <option>CIOG</option>
			  <option>CFOG</option>
			  <option>DSTO</option>
			  <option>JOC</option>
			</select>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
        <button type="button" class="btn btn-primary">Save changes</button>
      </div>
    </div>
  </div>
</div>

<div class="container-fluid">
     
      
	  <%
	  
	  If Session("HomeCC") = "" or IsNull(Session("HomeCC")) Then Session("HomeCC") = "Default"
	  
	  If Session("HomeCC") = "Default" Then
      <!-- Icon Cards-->
  Response.Write "<div class=""container""><div class=""panel-group""><div class=""col-sm-3"">" & _
    "<div class=""panel panel-primary""><div class=""panel-heading"" class=""col-sm-6"">Frequently Asked Questions</div>" & _
      "<div class=""panel-body""><a href=""img/ApplyForACard.pdf"" targe=""_new"">Frequently Asked Questions</a> </div></div></div>" & _
   "<div class=""col-sm-3""><div class=""panel panel-primary""><div class=""panel-heading"" class=""col-sm-6"">Diners - DTC</div>" & _
      "<div class=""panel-body""><a href=""Diners.asp""><img src=""img/Diners2.png"" height=""20px"" width=""35px""> Contact Diners</a> </div></div></div>" & _
	"<div class=""col-sm-3""><div class=""panel panel-primary""><div class=""panel-heading"" class=""col-sm-6"">ANZ - DPC</div>" & _
      "<div class=""panel-body""><a href=""ANZ.asp""><img src=""img/ANZ2.png"" height=""20px"" width=""40px""> Contact ANZ</a> </div></div></div>" & _
	"<div class=""col-sm-3""><div class=""panel panel-primary"">" & _
      "<div class=""panel-heading"" class=""col-sm-6"">How To</div><div class=""panel-body""><a href=""HelpFile.html"">Help File</a></div></div></div></div></div>"
	  Else
	  
	  objRS.Open "SELECT * FROM qryApplications WHERE ApplicationID = " & Session("ApplicationID") & "",objCon
	  
		If objRS.eof Then
			strApply = ""
			strGCFO = ""
			strDCC = ""
			strBank = ""
		Else
		
			If objRS("DateReceived") = "" or IsNull(objRS("DateReceived")) Then
				strApply = "Awaiting"
				strApplyClass = "default"
			Else
				strApply = "<i class=""fa fa-check""></i> Complete " & FormatDateTime(objRS("DateReceived"),vbShortDate)
				strApplyClass = "success"
			End If
			
			If objRS("GCFOSigned") = "" or IsNull(objRS("GCFOSigned")) Then
				strGCFODays = DateDiff("d",objRS("DateReceived"),now()) & " Days"
				strGCFO = "Awaiting"
				strGCFOClass = "default"
			Else
				strGCFODays = DateDiff("d",objRS("DateReceived"),objRS("GCFOSignedDate")) & " Days"
				strGCFO = "<i class=""fa fa-check""></i> Complete " & FormatDateTime(objRS("GCFOSignedDate"),vbShortDate)
				strGCFOClass = "success"
			End If
			
			If objRS("DateReviewed") = "" or IsNull(objRS("DateReviewed")) Then
				strDCCDays = DateDiff("d",objRS("DateReviewed"),now()) & " Days"
				strDCC = "Awaiting"
				strDCCClass = "default"
			Else
				strDCCDays = DateDiff("d",objRS("GCFOSignedDate"),objRS("DateReviewed")) & " Days"
				strDCC = "<i class=""fa fa-check""></i> Complete " & FormatDateTime(objRS("DateReviewed"),vbShortDate)
				strDCCClass = "success"
			End If
			
			If objRS("BankResponseDate") = "" or IsNull(objRS("BankResponseDate")) Then
				strBankDays = DateDiff("d",objRS("BankResponseDate"),now()) & " Days"
				strBank = "Awaiting (" & DateDiff("d",objRS("DateReceived"),now()) & " Days)"
				strBankClass = "default"
			Else
				strBankDays = DateDiff("d",objRS("DateReviewed"),objRS("BankResponseDate")) & " Days"
				strBank = "<i class=""fa fa-check""></i> Complete " & FormatDateTime(objRS("BankResponseDate"),vbShortDate)
				strBankClass = "success"
			End If
			
			If objRS("CardTypeSub") = "" or IsNull(objRS("CardTypeSub")) Then
				strBankName = ""
			Else
				strBankName = objRS("CardTypeSub")
			End If
			
				
		End If
		
	  Response.Write objRS("CardTypeSub") & " ApplicationID: " & objRS("ApplicationID") & " <div class=""container""><div class=""panel-group""><div class=""col-sm-2"">" & _
    "<div class=""panel panel-" & strApplyClass & """><div class=""panel-heading"" class=""col-sm-2"">Application</div>" & _
      "<div class=""panel-body"">" & strApply & "</div></div></div><div class=""col-sm-1"">" & _
    "<div class=""panel panel-" & strApplyClass & """ style=""border: 0;""><div class=""panel-body"" class=""col-sm-1"" style=""text-align:center;""><button type=""button"" class=""btn btn-" & strApplyClass & """><i class=""fa fa-arrow-right"" Title=""" & strGCFODays & """></i></button> " & strGCFODays & "</div>" & _
      "<div class=""panel-body""></div></div></div>" & _
	  "<div class=""col-sm-2"">" & _
    "<div class=""panel panel-" & strGCFOClass & """><div class=""panel-heading"" class=""col-sm-2"">GCFO Approved</div>" & _
      "<div class=""panel-body"">" & strGCFO & "</div></div></div><div class=""col-sm-1"">" & _
    "<div class=""panel panel-" & strGCFOClass & """ style=""border: 0;""><div class=""panel-body"" class=""col-sm-1"" style=""text-align:center;""><button type=""button"" class=""btn btn-" & strGCFOClass & """><i class=""fa fa-arrow-right"" Title=""" & strDCCDays & """></i></button> " & strDCCDays & "</div>" & _
      "<div class=""panel-body""></div></div></div>" & _
	  "<div class=""col-sm-2"">" & _
    "<div class=""panel panel-" & strDCCClass & """><div class=""panel-heading"" class=""col-sm-2"">DCC Approved</div>" & _
      "<div class=""panel-body"">" & strDCC & "</div></div></div><div class=""col-sm-1"">" & _
    "<div class=""panel panel-" & strDCCClass & """ style=""border: 0;""><div class=""panel-body"" class=""col-sm-1"" style=""text-align:center;""><button type=""button"" class=""btn btn-" & strBankClass & """><i class=""fa fa-arrow-right"" Title=""" & strBankDays & """></i></button> " & strBankDays & "</div>" & _
      "<div class=""panel-body""></div></div></div>" & _
	  "<div class=""col-sm-2"">" & _
    "<div class=""panel panel-" & strBankClass & """><div class=""panel-heading"" class=""col-sm-2"">" & strBankName & " Received</div>" & _
      "<div class=""panel-body"">" & strBank & "</div></div></div>" & _
	  "</div></div>"
	  
	  
	  'Response.Write objRS("CardTypeSub") & " ApplicationID: " & objRS("ApplicationID") & " <div class=""container""><div class=""panel-group""><div class=""col-sm-2"">" & _
    '"<div class=""panel panel-" & strApplyClass & """><div class=""panel-heading"" class=""col-sm-2"">Application</div>" & _
     ' "<div class=""panel-body"">" & strApply & "</div></div></div><div class=""col-sm-1""><button type=""button"" class=""btn btn-" & strApplyClass & """><i class=""fa fa-arrow-right"" Title=""" & strGCFODays & """></i></button> " & strGCFODays & "</div>" & _
   '"<div class=""col-sm-2""><div class=""panel panel-" & strGCFOClass & """><div class=""panel-heading"" class=""col-sm-2"">GCFO Approved</div>" & _
    '  "<div class=""panel-body"">" & strGCFO & "</div></div></div><div class=""col-sm-1""><button type=""button"" class=""btn btn-" & strGCFOClass & """><i class=""fa fa-arrow-right"" Title=""" & strDCCDays & """></i></button> " & strDCCDays & "</div>" & _
	'"<div class=""col-sm-2""><div class=""panel panel-" & strDCCClass & """><div class=""panel-heading"" class=""col-sm-2"">DCC Approved</div>" & _
     ' "<div class=""panel-body"">" & strDCC & "</div></div></div><div class=""col-sm-1""><button type=""button"" class=""btn btn-" & strDCCClass & """><i class=""fa fa-arrow-right"" Title=""" & strBankDays & """></i></button> " & strBankDays & "</div>" & _
	'"<div class=""col-sm-2""><div class=""panel panel-" & strBankClass & """>" & _
     ' "<div class=""panel-heading"" class=""col-sm-2"">" & strBankName & " Received</div><div class=""panel-body""></button>" & strBank & "</div></div></div></div></div>"
	  
	  End If

%>
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
	
   <tr><th Style="Height:25px;"  colspan="2">&nbsp;</th></tr>
</TABLE>

</BODY>
</HTML>
<%




%>
