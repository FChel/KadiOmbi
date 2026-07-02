
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #Include file=CAPSFunctions.asp -->
<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp?State=Expired")


'Description:	Create and view applications
'Author:		MG
'Date:			January 2020

	Response.Expires = -1500	

Dim objCon
Dim objRS
Dim objCmd

Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL

Dim lngApplicationID
Dim strEmployeeID
Dim strEID
Dim strTitle
Dim strFirstName
Dim strLastName
Dim strAddress1
Dim strAddress2
Dim strAddress3
Dim strAddress4
Dim strSUburb
Dim strState
Dim strPostCode
Dim dteDateReceived
Dim strStatus
Dim strReviewedBy
Dim dteDateReduced
Dim lngCreditLimit
Dim lngCurrentRecord
Dim arrStatus(9)
Dim strOptions
Dim strCardNoFull
Dim strProcessStatus

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	

	arrStatus(1) = "Awaiting Export"
	arrStatus(2) = "Awaiting Export"
	arrStatus(3) = "Rejected"
	arrStatus(4) = "Awaiting Review"
	arrStatus(5) = "Deleted"
	arrStatus(6) = "On Hold"
	arrStatus(7) = "Submitted"
	arrStatus(8) = "Temp Hold"
	arrStatus(9) = "Rejected"

	If isNull(Session("ApplicationID")) Or Session("ApplicationID") = "" Then 
		Session("ApplicationID") = 0
	End If
	
	If Not IsEmpty(Request.QueryString("PageCombo")) Then
		Session("PageCombo") = Request.QueryString("PageCombo")
	End If
	
	If Not IsEmpty(Request.QueryString("UserView")) Then
		Session("UserView") = Request.QueryString("UserView")
	End If

	If Not IsEmpty(Request.QueryString("ProcessStatus")) Then
		Session("ProcessStatus") = Request.QueryString("ProcessStatus")
		
	End If
	
	If Not IsEmpty(Request.QueryString("Action")) Then
		If Request.QueryString("Action") = "Cancel" Then
			Call CancelApplication(Request.QueryString("ApplicationID"))
		End If
		
		If Request.QueryString("Action") = "Release" Then
			Call ReleaseApplication(Request.QueryString("ApplicationID"),Request.QueryString("EmployeeID"))
		End If
	End If
	
	If Not IsEmpty(Request.QueryString("SaveStatus")) Then
		'Response.write  Request.QueryString("SaveStatus") & " s=" & Request.QueryString("NewStat")
		Call UpdateStatus(Request.QueryString("SaveStatus"),Request.QueryString("NewStat"))
	End If
	
	If Not IsEmpty(Request.QueryString("ViewButton")) Then
		Session("ViewButton") = Request.QueryString("ViewButton")
		
		'Also clear the selected application Types -----CLEAR ApplicationType
		If Session("ViewButton") = "All" Then
			Session("ProcessStatus") = ""
		End If
		
	End If
  
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "SubmitApp" Then
			Call SubmitApplication()
		End If
	End If
	
	'Set Awaiting Review as the button default view
	If IsEmpty(Session("ViewButton")) Then Session("ViewButton") = "Review"
	
	Call LoadDetails()
  
  If IsNull(Session("CardType")) OR Session("CardType") = "" Then Session("CardType") = "DTC - Diners"
%>

<html>
<head>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

	<!--<link rel="stylesheet" type="text/css" href="../CAPSStyle.css">-->
	<!--<script src="../assets/node_modules/jquery/jquery-3.2.1.min.js"></script>-->
	  <!-- Custom fonts for this template-->
  <!--<link href="../vendor/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css">-->
<script LANGUAGE="javascript">

function SaveData(){
	var varSubmit = true
		frm.msgbox.value='Saving.......';
		frm.submit();
	//}
}

function CloseScreen() {

    //if(top.Header1.Header2.document.form1.SaveStatus.value=='S')
    //{var x=window.confirm("Changes have been made, do you wish to save these changes?")
    //    if (x){
    //        SaveData();
    //        self.location='index.asp';
    //    }
    //    else
    //        self.location='index.asp';}
    //    else
    { self.location = 'HomeCC2.asp'; }
}

function SaveStatus() {
	self.location = 'LimitReductionSummary.asp?SaveStatus='+frm.AppStatID.value+'&NewStat='+frm.NewStatus.options[frm.NewStatus.selectedIndex].value;
}


setTimeout( 'ShowTimeoutWarning();', 1080000 );

function ShowTimeoutWarning () {     
    window.alert( "********** Warning! **********' \n \n 'You will be automatically logged out in 2 minutes unless you change screens, Close or Save!" ); 
}


function OpenSs(cb) {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	//var e = document.getElementById("CardTypeSelect");
	//var result = e.options[e.selectedIndex].value;
	
	//document.getElementById('CardType').value=result;
	
	var id = cb.getAttribute('data-AppID');
	document.getElementById('AppStatID').value=id;
	
	var Userid = cb.getAttribute('data-AppName');
	document.getElementById('StatName').value=Userid;
	
	document.getElementById('CurrentStat').value=cb.innerHTML;
	
	document.getElementById('demoModalLabel').value='Status Change for ' + Userid + ':';
	
}

function LoadFiles() {

  //var id = cb.getAttribute('data-id');
  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ModalFileMessage").innerHTML = this.responseText;
    }
  };

  xhttp.open("GET", "AJAX/GetExcelFiles2.asp?FileStart=Applications", true);
  xhttp.send();
}

function ChangePage() {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("PageCombo");
	var result = e.options[e.selectedIndex].value;
	
	self.location = 'LimitReductionSummary.asp?PageCombo=' + result;
	//alert(result);
	//document.getElementById('CardType').value=result;
	
}

function Getit() {

	alert('asa')
	alert(document.getElementById("SearchInput").value);
	
}

function loadCSToDiners(varID,varCardNo) {

	//If(varCardNo)
	//const varCardNo= ['a'+ varCardNo];
	//alert(document.getElementById("CardNoMod").value)
	//varCardNo=document.getElementById("CardNoMod").value
	
	//alert(varID);
	//alert(varCardNo);
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
	document.getElementById("compareModalLabel").innerHTML = '<button type="button" class="btn btn-outline-secondary" title="Displaying CS To Diners. Click to View CS From Diners." onClick="loadCSFromDiners('+varID+',' + varCardNo +');">CS To Diners</button>'
	document.getElementById("CSToDinersDetail").innerHTML = '<img src="../images/Load.gif" style="vertical-align:middle;" /> '
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CSToDinersDetail").innerHTML = this.responseText;
    }
  };

  xhttp.open("GET", "AJAX/GetCSToDinersAudit2.asp?EmployeeID=" + varID + "&CardNo=" + varCardNo + "", true);
  xhttp.send();
}

function loadCSFromDiners(varID,varCardNo) {
//alert(varID);
//alert(varCardNo);
//varCardNo=document.getElementById("CardNoMod").value
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
  document.getElementById("compareModalLabel").innerHTML = '<button type="button" class="btn btn-outline-secondary" title="Displaying CS To Diners. Click to View CS From Diners." onClick="loadCSToDiners('+varID+',' + varCardNo +');">CS From Diners</button>'
  document.getElementById("CSToDinersDetail").innerHTML = '<img src="../images/Load.gif" style="vertical-align:middle;" /> '
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CSToDinersDetail").innerHTML = this.responseText;
    }
  };
	
  xhttp.open("GET", "AJAX/GetCSFromDinersAudit2.asp?EmployeeID=" + varID + "&CardNo=" + varCardNo + "", true);
  xhttp.send();
}

</script>

<style>

/* Bootstrap 4 text input with search icon */

.has-search .form-control {
    padding-left: 2.375rem;
}

.has-search .form-control-feedback {
    position: absolute;
    z-index: 2;
    display: block;
    width: 2.375rem;
    height: 2.375rem;
    line-height: 2.375rem;
    text-align: center;
    pointer-events: none;
    color: #aaa;
}

</style>
</head>
<body >
<main class="main py-3">
    <div class="container">
<!-- Modal -->
	<div class="loader" id="ModApp">
        <div class="wrap">
            <div class="spinner"></div>
            <span class="loading-message">Loading...</h6>
        </div>
    </div>

<form action="LimitReductionSummary.asp?Action=Search&Link=AP" method="POST" id="frm" name="frm">
				 
<!-- Modal -->
    <div class="modal fade" id="StatusModal" tabindex="-1" role="dialog" aria-labelledby="demoModalLabel" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="demoModalLabel">Status</h5>
            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
              <span aria-hidden="true">&times;</span>
            </button>
          </div>
          <div class="modal-body">
		  <div class="col-md-12">
			<table class="table table-hover">
			<tr><td style="font-weight:bold;">Changing Status for:</td><td><input id="StatName" name = "StatName" value="" style="border: 0;" class="form-control"/></td></tr>
           <%
		   
			For x = 1 to 9

				strOptions = strOptions & "<OPTION Value=""" & arrStatus(x) & """ ID=""StatusChange" & x & """ Name=""StatusChange" & x & """>" & arrStatus(x) & "</OPTION>"
				
			Next
		   
			Response.Write "<tr><td><label for""StatusChange""  style=""font-weight:bold;"">Change to:</label></td>" & _
					"<td><SELECT id=""NewStatus"" name=""NewStatus"" class=""form-control"">" & strOptions & "</SELECT></tr>"
			
		   %>
		   <tr><td><input id="AppStatID" name = "AppStatID" value="" HIDDEN /></td><td><input id="CurrentStat" name = "CurrentStat" value="" HIDDEN /></td></tr>
			</table>
		   </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
            <button type="button" class="btn btn-primary" onClick="SaveStatus();">Save changes</button>
          </div>
        </div>
      </div>
    </div>
	<!-- End Modal -->
	
<!-- Start Download Excel Modal -->
	<div class="modal fade" id="ModalExcel" tabindex="-1" role="dialog" aria-labelledby="ModalDeleteCenterTitle" aria-hidden="true">
		<div class="modal-dialog modal-dialog-centered" role="document">
		  	<div class="modal-content">
				<div class="modal-header">
			  		<h5 class="modal-title" id="ModalDeleteLongTitle" style="font-weight:bold;">Download Excel file</h5>
			  		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				</div>
				<div class="modal-body">			  
				  	<div class="row">
						<div class="col-md-12 mb-3">			
						  		<div id="ModalFileMessage">
								<span id="Progress" style="padding-left:20px; padding-bottom:20px;"><img src="../images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>
								</div>
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
	<!-- End Download Excel Modal -->
	
	<!-- CS To Diners Modal -->
<div class="modal fade" id="CSToDinersModal" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-large modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  CS To Diners - Displayed (click to View CS From Diners)
				  
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="CSToDinersDetail" height="100px">
               
			
               <img src="../images/Load.gif" style="vertical-align:middle;" />   
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
 <!-- End of CS To Diners Modal -->
	
<!-- End the first part of the Header Container -->
<div id='tbl-container'>

	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-8">
				<h2 class="text-left">Limit Reductions<% If Session("UserView") = "User" Then Response.write " for " & Session("UserName")%> <i class="fa help-tooltip fa-question-circle" data-toggle="tooltip" title="All Card Applications. Can be filtered and searched."></i></h2>
			</div>
			<div class="col-md-4 text-right">
				<button type="button" class="btn btn-outline-success" data-toggle="modal" data-target="#ModalExcel" onClick="LoadFiles();" title="Click to Export Current Card Search to Excel"><i class="fa fa-file-excel"></i> Export To Excel</button>&nbsp;&nbsp;
			
			</div>
			
		</div>

          <div class="row py-2">
            <div class="col-md-9">
              <%Call LoadViewButtons()%>
            </div>
			<div class="col-md-3">
				<div class="form-group has-search">
					<span class="fa fa-search form-control-feedback"></span>
				 <input type="text" class="form-control" type="search" id="SearchInput" name="SearchInput" placeholder="Search Applications by Keyword" onChange="frm.submit();" value="<%=Request.Form("SearchInput")%>"/>
				 </div>
			</div>
          </div>

      </section>
	  
	  
	 <section>
  

                 <%
        
				DisplayTableDetails()
        
				%>	
                
          </div>

      </section>
</div>


<!--</DIV>-->
</form>
</div>

</main>

	
	

    <!-- jQuery -->
    <script src="../js/jquery.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="../js/bootstrap.min.js"></script>

	
<!-- #Include file=CAPSFooter.asp -->

</body>
</html>
<%

Public Sub DisplayTableDetails()

Dim y, x
Dim strAction
Dim strStatus
Dim strAddress
Dim dteDateSubmitted
Dim dteLimitDateFrom
Dim dteLimitDateTo
Dim dteDateReduced
Dim strSearch
Dim lngTotalRecords
Dim lngStartingPage
Dim strPages
Dim strSort
Dim strOrderType
Dim strRecordMessage
Dim strWhere
Dim dteDateUpdated
Dim arrSort(11)
Dim strSortArrow

'Dim lngPage
Dim strPageCombo
Dim arrPagecombo(6)
Dim strPages2
Dim lngTotalPages
Dim strActive
Dim lngCurrentPage
Dim bolSkip

Dim arrNames
Dim strFNameSearch
Dim strLNameSearch
Dim strApplicationType
Dim strSearchDate
Dim strDateFrom
Dim strDateTo
Dim strSearchDay
Dim strSearchMonth
Dim strSearchYear

Dim strErrorCount
Dim strResolvedCount
Dim strViewAppType
Dim strViewAppTitle
Dim strApplicantName
Dim strProcessStatusSelected
Dim strLimitColour
Dim intLimitDateDiff

	'strSearch = Request.Form("SearchInput")
	
	strSearch = Replace(Request.Form("SearchInput"), "'", "''")
	
	If IsEmpty(Request.QueryString("SortType")) Then
		'strOrderType = "ASC"
	Else
		If Request.QueryString("SortType") = "ASC" Then
			strOrderType = "DESC"
			'Set the variable to be used in the sort order Fontawesone image
			strSortArrow = "-down"
		Else
			strOrderType = "ASC"
			'Set the variable to be used in the sort order Fontawesone image
			strSortArrow = "-up"
		End If
	End If
	
	If IsEmpty(Request.QueryString("Sort")) Then
		strSort = ""
	Else		
		strSort = " ORDER BY " & Request.QueryString("Sort") & " " & strOrderType
	End If
	
	'Build the TOP Statement
	If Session("PageCombo") = "" Or IsNull(Session("PageCombo")) Then
		Session("PageCombo") = 50
	End If
	
	'If there is no sort then sort by the most recent submitted
	If IsNull(strSort) Or strSort = "" Then strSort = " ORDER BY DateSubmitted DESC"
	
	If Session("ViewButton") = "Review" Then
		strWhere = " AND [Status] = 'Awaiting review' "
	ElseIf Session("ViewButton") = "OnHold" Then
		strWhere = " AND [Status] = 'On Hold' "
	ElseIf Session("ViewButton") = "AwaitingIssue" Then
		strWhere = " AND [Status] = 'Rejected' "
	ElseIf Session("ViewButton") = "AddedToNA" Then
		strWhere = " AND [Status] = 'Awaiting Export' "
	ElseIf Session("ViewButton") = "TempHold" Then
		strWhere = " AND [Status] = 'Temp Hold' "
		ElseIf Session("ViewButton") = "Reduced" Then
		strWhere = " AND [ProcessStatus] = 'Reduced' "
	Else
		'This catches ALL
		strWhere = ""
		
	End If

	'If an Application Type has been selected then add that to the WHERE statement
	If Not IsNull(Session("ProcessStatus")) Then
		If Session("ProcessStatus") = "" Then
		Else
			If Session("ProcessStatus") = "ERROR" Then
			
			Else
				strWhere = strWHERE & " AND [ProcessStatus] = '" & Session("ProcessStatus")  & "'"
			End If
		End If
		
	End If
	

	
If strSearch = "" OR ISNull(strSearch) Then
	'If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 1000 * FROM qryCAPSLimitReductionCheck WITH(NOLOCK) WHERE [ApplicationID] > 0 " & strWhere & strSort
		
		
	'Else
		'strSQL = "SELECT TOP 1000 * FROM qryCAPSLimitReductionCheck WITH(NOLOCK) " & strSort
	'End If
	
Else
	'If the user has entered the date lookup then process this, otherwise perform the EmployeeID and Name searches
	If UCASE(Left(strSearch,2)) = "D:" Then
	
		strSearchDate = Right(strSearch,Len(strSearch)-2)
		
		If Len(strSearchDate) > 5 Then
			'strSearchDate = MediumDate(strSearchDate)
			'Check for the Day in the string, which could be one or two characters (numbers)
			If Mid(strSearchDate,2,1) = "/" OR Mid(strSearchDate,2,1) = "-" Then
				strSearchDay = "0" & Mid(strSearchDate,2,1)
				'Get the Month value
				If Mid(strSearchDate,4,1) = "/" OR Mid(strSearchDate,4,1) = "-" Then
					strSearchMonth = Mid(strSearchDate,3,1)
				Else
					strSearchMonth = Mid(strSearchDate,3,2)
				End If
			Else
				strSearchDay = Left(strSearchDate,2)
				'Get the Month value
				If Mid(strSearchDate,5,1) = "/" OR Mid(strSearchDate,5,1) = "-" Then
					strSearchMonth = Mid(strSearchDate,4,1)
				Else
					strSearchMonth = Mid(strSearchDate,4,2)
				End If
			End If
			
			'Get the year value
			'response.write "mid=" & Mid(strSearchDate,Len(strSearchDate)-3,1) & "</br>"
			If Mid(strSearchDate,Len(strSearchDate)-2,1) = "/" OR Mid(strSearchDate,Len(strSearchDate)-2,1) = "-" Then
				strSearchYear = "20" & Right(strSearchDate,2)
			Else
				strSearchYear = Right(strSearchDate,4)
			End If
			
			'Original Search date string before splitting to Day, Month and Year.
			'strSearchDate = Left(strSearchDate,2) & "-" & MonthName(Mid(strSearchDate,4,2)) & "-" & Right(strSearchDate,4)
			
			strSearchDate = strSearchDay & "-" & MonthName(strSearchMonth) & "-" & strSearchYear
		End If
		'response.write strSearchDate
		If IsDate(strSearchDate) Then 'DateAdd("d", -1, strSearchDate)
		
			'strDateFrom = DateAdd("d", -1, strSearchDate)
			strDateFrom = strSearchDate
			strDateTo = DateAdd("d", +1, strSearchDate)
			'response.write strDateTo
			'response.write Left(strDateTo,2) & ","
			'strDateFrom = Left(strDateFrom,2) & "-" & MonthName(Mid(strDateFrom,4,2)) & "-" & Right(strDateFrom,4)
			'strDateTo = Left(strDateTo,2) & "-" & MonthName(Mid(strDateTo,4,2)) & "-" & Right(strDateTo,4)
			
			'strSQL = "SELECT * FROM qryCAPSLimitReductionCheck WITH(NOLOCK) WHERE (DateSubmitted > '" & strSearchDate & "' AND DateSubmitted < " & strSearchDate & ")" & strWhere & strSort
			strSQL = "SELECT TOP 1000 * FROM qryCAPSLimitReductionCheck WITH(NOLOCK) WHERE (DateSubmitted > '" & strDateFrom & "' AND DateSubmitted < '" & strDateTo & "')" & strWhere & strSort
			
			'strSQL = "SELECT * FROM qryCAPSLimitReductionCheck WITH(NOLOCK) WHERE (DateSubmitted > '" & DateAdd("d", -1, strSearchDate) & "' AND DateSubmitted < '" & DateAdd("d", +1, strSearchDate) & "')" & strWhere & strSort
			'response.write strsql
		Else
			strSQL = "SELECT TOP 1000 * FROM qryCAPSLimitReductionCheck WITH(NOLOCK) WHERE (DateSubmitted = '" & strSearchDate & "')" & strWhere & strSort
		End If
		
	Else
		If Session("UserView") = "All" Then
			'If the user has entered a search term with a space then assume this is a first and last name so search on that only
			If Instr(1,strSearch," ")>0 Then
				arrNames = Split(strSearch," ")
				strFNameSearch = arrNames(0)
				strLNameSearch = arrNames(1)
				
				'strWhere = " WHERE ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%')"
				strSQL = "SELECT TOP 1000 * FROM qryCAPSLimitReductionCheck WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%'))" & strWhere & strSort
			Else
				strSQL = "SELECT TOP 1000 * FROM qryCAPSLimitReductionCheck WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strWhere & strSort
			End If
		Else
			strSQL = "SELECT TOP 1000 * FROM qryCAPSLimitReductionCheck WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strSort
		End If
		
		
	
	'End of 
	End If
	

	
End If

'Get the Sort Order for the sort order fontawesome image for the selected sort field
'For x = 1 to 8

	Select Case Request.QueryString("Sort")
	
		Case "ApplicationID"
			arrSort(1) = strSortArrow
		Case "EmployeeID"
			arrSort(2) = strSortArrow
		Case "Surname"
			arrSort(3) = strSortArrow
		Case "CardType"
			arrSort(4) = strSortArrow
		Case "ApplicationType"
			arrSort(5) = strSortArrow
		Case "Status"
			arrSort(6) = strSortArrow
		Case "DateSubmitted"
			arrSort(7) = strSortArrow
		Case "DateReduced"
			arrSort(8) = strSortArrow
		Case "LimitDateFrom"
			arrSort(9) = strSortArrow
		Case "LimitDateTo"
			arrSort(10) = strSortArrow
	End Select
	
'Next

'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = strSearch
Else
	strRecordMessage = Session("UserName") 
End If

'Set the global query
Session("ExcelSearch") = strSQL

If not isEmpty(Session("ExcelSearch")) Then
	Session("ExcelSearch") = Replace(Session("ExcelSearch"),"Top 500","")
	Session("ExcelSearch") = Replace(Session("ExcelSearch"),"Top 1000","")
	
End If

	y = 0
	
	If IsEmpty(Request.QueryString("StartingPage")) Then
		lngStartingPage = 1
		lngCurrentPage = 1
	Else
		lngStartingPage = Request.QueryString("StartingPage")
		lngCurrentPage = Request.QueryString("StartingPage")
	End If
'response.write strSQL	
	objRS.PageSize = Session("PageCombo")
	objRS.CursorLocation = 3 ' adUseClient
	objRS.CacheSize = Session("PageCombo")
	objRS.Open strSQL,objCon',3,1
		
	'response.write objRS.PageCount
	lngTotalPages = objRS.PageCount
	
	'objRS.AbsolutePage = Session("PageCombo")
	'Response.Write "Page=" & lngStartingPage & " objRS.PageSize =" & Session("PageCombo")    
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Applications for " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	Else
		objRS.Movelast
		objRS.Movefirst
		lngTotalRecords = objRS.Recordcount
		
		'objRS.AbsolutePage = lngStartingPage
		
		'Set the Page combos here so can be transferred to other pages together
		'arrPagecombo(1) = "50"
		'arrPagecombo(2) = "100"
		'arrPagecombo(3) = "200"
		'arrPagecombo(4) = "500"
		'arrPagecombo(5) = "1000"
		'arrPagecombo(6) = "All"
		
		'Build the Page Combo for TOP statement
		'For x = 1 to 6
		
			'If Session("PageCombo") = arrPagecombo(x) Then
				'strSelected = " SELECTED "
			'Else
				'strSelected = ""
			'End If
			'strPageCombo = strPageCombo & "<option " & strSelected & " value=""" & arrPagecombo(x) & """>" & arrPagecombo(x) & "</option>"
		'Next
		
		'strPageCombo = "<SELECT ID=""PageCombo"" Name=""PageCombo"" onChange=""ChangePage();"">" & strPageCombo & "</select>"
		
		'If the PageCombo is not numeric (ALL or Null) then make it the total records for the recordset (which is set above)
		'If NOT IsNumeric(Session("PageCombo")) Then Session("PageCombo") = lngTotalRecords
	
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Displaying " & lngTotalRecords & " applications</span><span class=""panel-subheader"" style=""float:right;""></span></div></div>"
		
		
		'Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
        '          "<span class=""panel-subheader"">Displaying 50 of " & lngTotalRecords & " applications (" & lngStartingPage & " to " & lngStartingPage + 50 & ")</span></div></div>"

		'If the Application types selected are Limit Changes the display the Start Date for the Limit Change
		'If Right(Session("ProcessStatus"),12) = "Limit Change" Then
			'strProcessStatusSelected = "<th scope=""col"" style=""font-size:14px;""><a href=""LimitReductionSummary.asp?Sort=LimitDateFrom&Link=AP&SortType=" & strOrderType & """> Limit Start <i class=""fa fa-sort" & arrSort(9) & """></i></a></th>"
		'Else
			strProcessStatusSelected = ""
		'End If
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""LimitReductionSummary.asp?Sort=ApplicationID&Link=AP&SortType=" & strOrderType & """> App ID <i class=""fa fa-sort" & arrSort(1) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Date AE602 XML Form was LOADED into CAPS""><a href=""LimitReductionSummary.asp?Sort=DateSubmitted&Link=AP&SortType=" & strOrderType & """> Submitted <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
			"<th scope=""col""><a href=""LimitReductionSummary.asp?Sort=EmployeeID&Link=AP&SortType=" & strOrderType & """> EID <i class=""fa fa-sort" & arrSort(2) & """></i></a></th>" & _
			"<th scope=""col""><a href=""LimitReductionSummary.asp?Sort=Surname&Link=AP&SortType=" & strOrderType & """> Name <i class=""fa fa-sort" & arrSort(3) & """></i></a></th>" & _
			"<th scope=""col""><a href=""LimitReductionSummary.asp?Sort=ApplicationType&Link=AP&SortType=" & strOrderType & """> Limit Status <i class=""fa fa-sort" & arrSort(5) & """></i></a></th>" & _
			"<th scope=""col""><a href=""LimitReductionSummary.asp?Sort=LimitDateFrom&Link=AP&SortType=" & strOrderType & """> Limit Date From <i class=""fa fa-sort" & arrSort(9) & """></i></a></th>" & _
			"<th scope=""col""><a href=""LimitReductionSummary.asp?Sort=LimitDateTo&Link=AP&SortType=" & strOrderType & """> Limit Date To <i class=""fa fa-sort" & arrSort(10) & """></i></a></th>" & _
			"<th scope=""col"">Temp Limit</th>" & _
			"<th scope=""col"">Original Limit</th>" & _
			"<th scope=""col"">Current Limit</th>" & _
			"<th scope=""col""><a href=""LimitReductionSummary.asp?Sort=Status&Link=AP&SortType=" & strOrderType & """> Status  <i class=""fa fa-sort" & arrSort(6) & """></i></a> <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Card Process Status - the stage the application is at.""></i></th>" & _
			"<th scope=""col"">Reduction Status</th>" & _
			"<th scope=""col""><a href=""LimitReductionSummary.asp?Sort=DateReduced&Link=AP&SortType=" & strOrderType & """> Date Reduced <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
			strProcessStatusSelected & "<th scope=""col"">Action</th>" & _
			"</tr></thead><tbody class=""text-left"">"
				
				
	End If
    'response.write "If " & y & " <= " & cint(lngStartingPage)*cint(Session("PageCombo")) & "AND y >= " & (cint(lngStartingPage)*clng(Session("PageCombo")))-clng(Session("PageCombo")) & "Then"
	
	x = 0
	
    Do until objRS.EOF 
	'Do While Not (objRS.Eof )
	'Do While Not (objRS.Eof AND objRS.AbsolutePage <> lngCurrentPage )
	'Do While Not objRS.Eof AND x < Session("PageCombo")'objRS.Pagesize
	
		y = y + 1
		
		'Only write the first 50 records from the starting position
		'If y <= lngStartingPage + 50 AND y >= lngStartingPage - 50 Then
		'If y <= lngStartingPage + Session("PageCombo") AND y >= lngStartingPage - Session("PageCombo") Then
		'If y <= clng(lngStartingPage)*clng(Session("PageCombo")) AND y >= (clng(lngStartingPage)*clng(Session("PageCombo")))-clng(Session("PageCombo")) Then
		
			x = x + 1
			
			strViewAppTitle = ""
			
			'Get the Application Type
			If isNull(objRS("ApplicationTypeName")) Then
				strApplicationType = ""
			Else
				'If objRS("ApplicationTypeName") = "" Then objRS("ApplicationTypeName") = " "
				If Instr(objRS("ApplicationTypeName"),"Change") > 0 Then
					strApplicationType = "LimitChange"
				Else
					strApplicationType = "Normal"
				End If
			End If
			
			'Format the Date updated value
			If IsNull(objRS("DateUpdated")) OR objRS("DateUpdated")= "" Then
				dteDateUpdated = ""
			Else
				dteDateUpdated = FormatDateTime(objRS("DateUpdated"),vbShortDate)
			End If
			
			'Make sure the Diners Card Number is used not the mastercard for the Audit Log details
			If IsNull(objRS("CardTypeSubCard")) OR objRS("CardTypeSubCard")= "" Then
				strCardNoFull = objRS("CardNumber")
			Else
				If objRS("CardTypeSubCard")="Diners" Then
					strCardNoFull = objRS("CardNumber")
				Else
					strCardNoFull = objRS("Companion")
				End If
			End If
			
			strEID = objRS("EmployeeID")
			
			If IsNull(objRS("Surname")) Or objRS("Surname") = "" Then
				strApplicantName = ""
			Else
				strApplicantName = Left(objRS("FirstName") & " " & objRS("Surname"),15)
			End If
			
			
			'Select the Action and Status buttons/pills based on the application status
			Select Case objRS("Status")
			
			Case  "Received"
				strAction = "<button type=""button"" class=""btn btn-primary btn-sx"" onclick=""self.location='LimitReductionSummary.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> Release</button>"
				strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='LimitReductionSummary.asp?Action=Reject&Link=AP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			Case "Added To CS"

				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
			Case "Awaiting Export"

				'strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportNA.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View NA</button>"	
				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#CSToDinersModal"" HREF=""#"" onClick=""loadCSToDiners('" & strEID & "','" & strCardNoFull & "')""><i class=""fa fa-eye""></i> View Audit</button>"
				strStatus = "<span class=""badge badge-pill badge-info"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
			Case "Submitted"
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='LimitReductionSummary.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"

				'strStatus  = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='LimitReductionSummary.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted to GCFO</button>"
				strStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">Submitted to GCFO</span>"
			Case "Deleted"
				'strAction = "Deleted - " & objRS("DateUpdated")'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='LimitReductionSummary.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				strAction = "<Span style=""font-size:12px;"">Deleted - " & dteDateUpdated & "</span>"'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='LimitReductionSummary.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				'strStatus  = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='LimitReductionSummary.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Cancelled By Applicant</button>"
				strStatus = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">Deleted</span>"
			Case "Rejected"
				'strAction = "Deleted - " & objRS("DateUpdated")'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='LimitReductionSummary.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				strAction = "<Span style=""font-size:12px;"">Rejected - " & dteDateUpdated & "</span>"'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='LimitReductionSummary.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				'strStatus  = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='LimitReductionSummary.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Cancelled By Applicant</button>"
				strStatus = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">Rejected</span>"
			Case "ASFIN Approved"
				strAction = "<button type=""button"" title=""Approved by GCFO"" class=""btn btn-secondary btn-xs"" onclick=""self.location='LimitReductionSummary.asp?Link=AP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-check""></i>GCFO Approved</button>"
			
				'strStatus  = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='LimitReductionSummary.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Approved by GCFO</button>"
				strStatus = "<span class=""badge badge-pill badge-success"">Approved by ASFIN</span>"
			Case  "Awaiting Review"
				'strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='LimitReductionSummary.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "&EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Release</button>"
				'strAction = strAction & " <button type=""button"" class=""btn btn-outline-danger btn-xs"" onclick=""self.location='LimitReductionSummary.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-times""></i> Delete</button>"
				
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=AP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				
				'If the application is a Limit Change then open the Limit Change Submit screen, otherwise open the Normal submit screen
				If strApplicationType = "LimitChange" Then
					'strAction = strAction & "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsLimitSubmit.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
					strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#CSToDinersModal"" HREF=""#"" onClick=""loadCSToDiners('" & strEID & "','" & strCardNoFull & "')""><i class=""fa fa-eye""></i> View Audit</button>"
				Else
					strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsSubmit.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				End If
				
				''''SEP 2021 --- CHANGED from the below to the status line 2 below so that applications can have their status changed if they are Awaiting Review
				'strStatus = "<span class=""badge badge-pill badge-warning"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
				strStatus = "<span class=""badge badge-pill badge-warning "" data-toggle=""modal"" data-target=""#StatusModal"" data-AppID=""" & objRS("ApplicationID") & """  data-AppName=""" & objRS("FirstName") & " " & objRS("Surname") & " - " & objRS("CardType") & " " & objRS("CardTypeSub") & " Application"" onClick=""OpenSs(this);"">" & objRS("Status") & "</span>"
				
			Case  "Rejected"
				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportNA.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View NA</button>"	
				strStatus = "<span class=""badge badge-pill badge-info"">" & objRS("Status") & "</span>"
			Case "On Hold"
				strStatus = "<span class=""badge badge-pill badge-secondary "" data-toggle=""modal"" data-target=""#StatusModal"" data-AppID=""" & objRS("ApplicationID") & """  data-AppName=""" & objRS("FirstName") & " " & objRS("Surname") & " - " & objRS("CardType") & " " & objRS("CardTypeSub") & " Application"" onClick=""OpenSs(this);"">" & objRS("Status") & "</span>"
			
				
				'strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=AP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				
				'If the application is a Limit Change then open the Limit Change Submit screen, otherwise open the Normal submit screen
				If strApplicationType = "LimitChange" Then
				
					'strAction = strAction & "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsLimitSubmit.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
					strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#CSToDinersModal"" HREF=""#"" onClick=""loadCSToDiners('" & strEID & "','" & strCardNoFull & "')""><i class=""fa fa-eye""></i> View Audit</button>"
				Else
					strAction = strAction & "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsSubmit.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				End If
			Case "Temp Hold"
				strStatus = "<span class=""badge badge-pill badge-danger "" data-toggle=""modal"" data-target=""#StatusModal"" data-AppID=""" & objRS("ApplicationID") & """  data-AppName=""" & objRS("FirstName") & " " & objRS("Surname") & " - " & objRS("CardType") & " " & objRS("CardTypeSub") & " Application"" onClick=""OpenSs(this);"">" & objRS("Status") & "</span>"
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=AP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				
				'If the application is a Limit Change then open the Limit Change Submit screen, otherwise open the Normal submit screen
				If strApplicationType = "LimitChange" Then
				
					strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsLimitSubmit.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				Else
					strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsSubmit.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				End If
			Case  "Done"
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=AP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				strStatus = "<span class=""badge badge-pill badge-info"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
			Case Else
				'strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='LimitReductionSummary.asp?Action=Cancel&Link=AP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				'strAction = "Rejected"
				'strStatus  = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='LimitReductionSummary.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted</button>"
				'strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=AP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				'strStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
			End Select

		
			
			If IsNull(objRS("DateSubmitted")) Then
				dteDateSubmitted = ""
			Else
				dteDateSubmitted = FormatDateTime(objRS("DateSubmitted"),vbShortDate)
			End If
			
			If IsNull(objRS("LimitDateFrom")) Then
				dteLimitDateFrom = ""
			Else
				dteLimitDateFrom = FormatDateTime(objRS("LimitDateFrom"),vbShortDate)
			End If
			
			If IsNull(objRS("LimitDateTo")) Then
				dteLimitDateTo = ""
			Else
				dteLimitDateTo = FormatDateTime(objRS("LimitDateTo"),vbShortDate)
			End If
			
			If IsNull(objRS("DateReduced")) Then
				dteDateReduced = ""
			Else
				dteDateReduced = FormatDateTime(objRS("DateReduced"),vbShortDate)
			End If
			
			'If the Application types selected are Limit Changes the display the Start Date for the Limit Change
			If Right(Session("ProcessStatus"),12) = "Limit Change" Then
				'Get the Limit change date relative to today to determine the text colour
				intLimitDateDiff = DateDiff("d",objRS("LimitDateFrom"),Now())
				
				If intLimitDateDiff > 1 Then
					strLimitColour = "color:Green; font-weight:bold;"
				ElseIf intLimitDateDiff < 1 AND intLimitDateDiff > -5 Then
					strLimitColour = "color:red; font-weight:bold;"
				Else
					strLimitColour = ""
				End If	

			
				
				'strProcessStatusSelected = "<TD style=""font-size:13px; text-align:center; " & strLimitColour & """ Title=""" & intLimitDateDiff*-1 & " Days from today"">" & objRS("LimitDateFrom") & "</TD>"
			Else
			
				strProcessStatus = Check_Reduction(objRS("ProcessStatus"),objRS("CreditLimitOriginalVarChar"), objRS("TempLimit"), objRS("CardCreditLimit"),objRS("LimitDateFrom"),objRS("LimitDateTo"))
				If strProcessStatus = "Created" Then
					strProcessStatus = "<span class=""badge badge-pill badge-warning "">" & objRS("ProcessStatus") & "</span>"
				Elseif strProcessStatus = "Reduced" Then
					strProcessStatus = "<span class=""badge badge-pill badge-success "">" & objRS("ProcessStatus") & "</span>"
				Else
					strProcessStatus = "<span class=""badge badge-pill badge-danger "">" & strProcessStatus & "</span>"
				End If
			End If
			
			If Session("ProcessStatus") = "ERROR" Then
				If Check_Reduction(objRS("ProcessStatus"),objRS("CreditLimitOriginalVarChar"), objRS("TempLimit"), objRS("CardCreditLimit"),objRS("LimitDateFrom"),objRS("LimitDateTo")) = "Error" Then
					response.write "<TR><TD><a Target=""_self"" HREF=""ApplicationDetail.asp?Link=AP&ApplicationID=" & objRS("ApplicationID") & """>" & objRS("ApplicationID") & "</a></TD><TD style=""font-size:13px; text-align:center;"">" & dteDateSubmitted & "</TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
						"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""ApplicationDetail.asp?ApplicationID=" & objRS("ApplicationID") & """ title=""" & objRS("FirstName") & " " & objRS("Surname") & """>" & strApplicantName & "</a></TD>" & _
						"<TD style=""font-size:12px;"" title=""" & objRS("ApplicationType") & """><a Target=""_self"" HREF=""ApplicationDetail.asp?ApplicationID=" & objRS(0) & """>" & objRS("ApplicationTypeName") & "</a></TD><TD style=""font-size:12px;"">" & dteLimitDateFrom & "</TD><TD style=""font-size:12px;"">" & dteLimitDateTo & "</TD><TD style=""font-size:12px;"">" & formatnumber(objRS("TempLimit"),0) & "</TD><TD style=""font-size:12px;"">" & formatnumber(objRS("CreditLimitOriginalVarChar")/100,0) & "</TD><TD style=""font-size:12px;"">" & formatnumber(objRS("CardCreditLimit"),0) & "</TD><TD style=""font-size:12px;"">" & strStatus & "</TD><TD style=""font-size:12px;"">" & strProcessStatus & "</TD>" & _
						"<TD style=""font-size:13px; text-align:center;"">" & dteDateReduced & "</TD>" & strProcessStatusSelected & _
						"<TD>" & strAction & "</TD></TR>"
				
				End If
										
			Else
				response.write "<TR><TD><a Target=""_self"" HREF=""ApplicationDetail.asp?Link=AP&ApplicationID=" & objRS("ApplicationID") & """>" & objRS("ApplicationID") & "</a></TD><TD style=""font-size:13px; text-align:center;"">" & dteDateSubmitted & "</TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
						"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""ApplicationDetail.asp?ApplicationID=" & objRS("ApplicationID") & """ title=""" & objRS("FirstName") & " " & objRS("Surname") & """>" & strApplicantName & "</a></TD>" & _
						"<TD style=""font-size:12px;"" title=""" & objRS("ApplicationType") & """><a Target=""_self"" HREF=""ApplicationDetail.asp?ApplicationID=" & objRS(0) & """>" & objRS("ApplicationTypeName") & "</a></TD><TD style=""font-size:12px;"">" & dteLimitDateFrom & "</TD><TD style=""font-size:12px;"">" & dteLimitDateTo & "</TD><TD style=""font-size:12px;"">" & formatnumber(objRS("TempLimit"),0) & "</TD><TD style=""font-size:12px;"">" & formatnumber(objRS("CreditLimitOriginalVarChar")/100,0) & "</TD><TD style=""font-size:12px;"">" & formatnumber(objRS("CardCreditLimit"),0) & "</TD><TD style=""font-size:12px;"">" & strStatus & "</TD><TD style=""font-size:12px;"">" & strProcessStatus & "</TD>" & _
						"<TD style=""font-size:13px; text-align:center;"">" & dteDateReduced & "</TD>" & strProcessStatusSelected & _
						"<TD>" & strAction & "</TD></TR>"
				
			End If
			'response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""LimitReductionSummary.asp?ApplicationID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & strAction & "</a></TD>" & _
			'		"<TD><a Target=""_self"" HREF=""LimitReductionSummary.asp?ApplicationID=" & objRS(0) & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""LimitReductionSummary.asp?ApplicationID=" & objRS(0) & """>" & objRS(2) & "</a></TD>" & _
			'		"<TD style=""text-align:center;"">" & objRS(3) & "</TD><TD style=""text-align:center;"">" & objRS(4) & "</TD>" & _
			'		"<TD style=""text-align:center;"">" & objRS(5) & "</TD><TD style=""text-align:center;"">" & objRS(6) & "</TD>" & _
			'		"<TD style=""text-align:center;"">" & objRS(7) & "</TD><TD style=""text-align:center;"">" & objRS(10) & "</TD>" & _
			'		"<TD style=""text-align:center;"">" & strStatus & "</TD><TD style=""text-align:center;"">" & objRS(14) & "</TD><TD style=""text-align:center;"">" & objRS(15) & "</TD></TR>"
		
'		End If
			
		objRS.movenext
	Loop
	
	If y > 0 Then
		Response.Write "<TR><TH colspan=""10"">Total</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;"">" & x & "</TH></TR></tbody></table></div>"
				     
	End If
	
	'Create the number of pages
'	If IsNumeric(y) Then
'		If y > 1 Then
			
'			y = y / 50
			
'			For x = 1 to y
'				strPages = strPages & "<a href=""LimitReductionSummary.asp?Link=AP&StartingPage=" & (x * 50) & """> " & x & " </a>"
'			Next
			
'		End If
'	End If
	
'	If y > 0 Then
'		Response.Write "<TR><TH colspan=""9"" style=""text-align:center;""><a href=""LimitReductionSummary.asp?Link=AP&Previous&StartingPage=" & lngStartingPage -50 & """>Previous Page " & strPages & " <a href=""LimitReductionSummary.asp?Link=AP&Previous&StartingPage=" & lngStartingPage + 50 & """> Next Page</TH></TR>"
'	End If

	bolSkip = False
	
	For x = 1 to lngTotalPages
		
		'Determine which page number is active (displayed as active)
		If clng(x) = clng(lngCurrentPage) Then
			strActive = "active"
		Else
			strActive = ""
		End If
		
		If lngTotalPages > 20 Then
			If x > 1 Then
				'Add the Elipsis (...) to the start of the page numbers if there is more than 20 pages and the current place is beyond the first page
			'	If lngTotalPages > 20 Then
			'		strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""LimitReductionSummary.asp?StartingPage=" & lngTotalRecords - (clng(Session("PageCombo"))*20) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
			'	End If
			End If
		
			If x > 20 Then
				If bolSkip = False Then
					strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""LimitReductionSummary.asp?StartingPage=" & lngTotalPages & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
					bolSkip = True
				End If
			Else
				
			End If
		End If
		
		If bolSkip = True Then
		Else
			strPages2 = strPages2 & "<li class=""page-item " & strActive & """><a class=""page-link"" href=""LimitReductionSummary.asp?StartingPage=" & x & """>" & x & "</a></li>"
		End If
		
	Next
	
	'Write the Pagination objects for all pages based on the total records and the number records displayed on screen
	If lngTotalPages > 0 Then
		
		'Add the Elipsis (...) to the end of the page numbers if there is more than 20 pages
		'If x = 20 + cint(fix(lngPage)) Then
		'If x = 21 + lngPage Then
		'	strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""LimitReductionSummary.asp?StartingPage=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
		'End If
		
		'Add the Elipsis (...) to the start of the page numbers if there is more than 20 pages and the current place is beyond the first page
		'If x = 0 AND lngPage > 1 AND y > 20 Then
		'	strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""LimitReductionSummary.asp?StartingPage=" & lngTotalRecords - (clng(Session("PageCombo"))*20) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
		'End If
					
		'Response.Write "<div class=""container""><div class=""row""><div class=""col-12 text-center"">" & _
		'	"<nav aria-label=""Page navigation""><ul class=""pagination""><li class=""page-item"">" & _      
		'	"<a class=""page-link"" href=""LimitReductionSummary.asp?StartingPage=1"" aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
		'	strPages2 & _
		'	"<li class=""page-item"">" & _
		'	"<a class=""page-link"" href=""LimitReductionSummary.asp?StartingPage=" & lngTotalPages & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
		'	"</a></li></ul></nav></div></div></div>"

				'"<a class=""page-link"" href=""LimitReductionSummary.asp?StartingPage=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
	End If
	
	
objRS.Close

	'Response.Write strSQL

End Sub


Sub LoadDetails()

   'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM tblCAPSCDMC WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "'",objCon

		If Not objRS.EOF Then
		   
			'lngApplicationID = objRS("ApplicationID")
			strEmployeeID = objRS("EmployeeID")		
			strTitle = objRS("Title")
			strFirstName = objRS("FirstName")
			strLastName  = objRS("Surname")
			strAddress1 = objRS("Addressline1")
			strAddress2 = objRS("Addressline2")
			strAddress3 = objRS("Addressline3")
			'strAddress4 = objRS("Address4")
			strSuburb = objRS("Postaladdress_City")
			strState = objRS("Postaladdress_State")
			strPostCode = objRS("Postaladdress_PostCode")
			'dteDateReceived = objRS("DateReceived")
			'strStatus = objRS("Status")
			'strReviewedBy = objRS("ReviewedBy")
			'dteDateReduced = objRS("DateReduced")
			'If IsNull(objRS("CreditLimit")) or objRS("CreditLimit") = "" then
				lngCreditLimit = 30000
			'Else
			'	lngCreditLimit = objRS("CreditLimit") 
			'End If
		Else
			Session("ApplicationID") = 0
			lngApplicationID = 0'objRS("ApplicationID")
			strEmployeeID = ""
			strTitle = ""
			strFirstName = ""
			strLastName  = ""
			strAddress1 = ""
			strAddress2 = ""
			strAddress3 = ""
			strAddress4 = ""
			strSuburb = ""
			strState = ""
			strPostCode = ""
			dteDateReceived = ""
			strStatus = ""
			strReviewedBy = ""
			dteDateReduced = ""
			lngCreditLimit = 30000
	   End If

	objRS.Close
	
End Sub

Public Sub CancelApplication(lngApplicationID)

	strSQL = "UPDATE tblCAPSApplication SET Status = 'Deleted', DateUpdated = '" & Now() & "' WHERE ApplicationID = " & lngApplicationID & ""
	
	objCon.Execute strSQL
	
End Sub

Public Sub ReleaseApplication(lngApplicationID, lngEmpID)
Dim intRecord

	'strSQL = "UPDATE tblCAPSApplication SET Status = 'Awaiting Export', DateUpdated = '" & Now() & "' WHERE ApplicationID = " & lngApplicationID & ""
	
	'objCon.Execute strSQL

  	With objCmd

		.CommandType = 4
		.CommandText = "spCAPSReleaseApplication"

		.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarChar, adParamInput,10)
		.Parameters.Append objCmd.CreateParameter("ApplicationID", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("NAToDinersIDOutput", adInteger, adParamOutput)
		
		.Parameters("EmployeeID") = lngEmpID
		.Parameters("ApplicationID") = lngApplicationID
		.Parameters("UpdatedBy") = Session("UserID")

		.ActiveConnection = objCon
		 
	End With
	   
	objCmd.Execute        
  
	'Return the result of the Save Function.
	intRecord = objCmd.Parameters.Item("NAToDinersIDOutput") 
 
	If IsNull(intRecord) Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">NOT RELEASED! See System Admin. Error " & intRecord & "</div>"
	Else
		If intRecord = 0 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">NOT RELEASED! See System Admin. No Application Error " & intRecord & " A:" & lngApplicationID & "</div>"
		ElseIf intRecord = -1 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">NOT RELEASED! Applicant already has a card on the NA File. Error " & intRecord & "</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Application " & intRecord & " Released!  This will now appear on the NA File for Export</div>"
		End If
	End If
		
End Sub

Public Sub SubmitApplication()

Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spCDMCToApplication"

			.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("CardType", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("CardTypeSub", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("CreditLimit", adDouble, adParamInput) 
			.Parameters.Append objCmd.CreateParameter("CDMCToApplicationIDOutput", adInteger, adParamOutput)
			
			.Parameters("EmployeeID") = Session("EmployeeID")
			.Parameters("CardType") = Left(Request.Form("CardType"),3)
			.Parameters("CreditLimit") = Request.Form("CreditLimit")
			.Parameters("CardTypeSub") = Right(Request.Form("CardType"),Len(Request.Form("CardType"))-6)
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCToApplicationIDOutput") 
	 
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" /> Application " & intRecord & " submitted to your GCFO for approval!"

		strMessageColour = "Black"
		
End Sub


Public Sub UpdateStatus(lngApplicationID, strStatus)
Dim intRecord

  	With objCmd

		.CommandType = 4
		.CommandText = "spCAPSApplicationStatusUpdate"

		.Parameters.Append objCmd.CreateParameter("ApplicationID", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput,20)
		.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("ApplicationIDOutput", adInteger, adParamOutput)
		
		.Parameters("ApplicationID") = lngApplicationID
		.Parameters("Status") = strStatus
		.Parameters("UpdatedBy") = Session("UserID")

		.ActiveConnection = objCon
		 
	End With
	   
	objCmd.Execute        
  'response.write "EXEC spCAPSApplicationStatusUpdate(" & lngApplicationID & ",'" & strStatus & "'," & Session("UserID") & ",0)"
  
	'Return the result of the Save Function.
	intRecord = objCmd.Parameters.Item("ApplicationIDOutput") 
 
	If IsNull(intRecord) Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">STATUS NOT UPDATED! See System Admin. Error " & intRecord & "</div>"
	Else
		If intRecord = 0 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">STATUS NOT UPDATED See System Admin. No Application Error " & intRecord & " A:" & lngApplicationID & "</div>"
		ElseIf intRecord = -1 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">STATUS NOT UPDATED! Applicant already has the same status. Error " & intRecord & "</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Application " & intRecord & " Status Updated to " & strStatus & "!</div>"
		End If
	End If
		
End Sub


Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(7)
Dim strDropDown
Dim strAppText

If Session("ViewButton") = "Review" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "OnHold" Then
	arrButton(3) = "active"
ElseIf Session("ViewButton") = "AwaitingIssue" Then
	arrButton(4) = "active"
ElseIf Session("ViewButton") = "AddedToNA" Then
	arrButton(5) = "active"
ElseIf Session("ViewButton") = "TempHold" Then
	arrButton(6) = "active"
ElseIf Session("ViewButton") = "Reduced" Then
	arrButton(7) = "active"
	
Else
	'This catches ALL
	arrButton(1) = "active"
End If

'Make the Button text the selected Application Type if one is selected
If Session("ProcessStatus") = "" Or IsNull(Session("ProcessStatus")) then
	strAppText = "Limit Status"
Else
	strAppText = Left(Session("ProcessStatus"),10)
End If

'Create the Application Type selection/filter
strDropDown = "<div class=""dropdown""><button class=""btn btn-outline-primary dropdown-toggle"" type=""button"" id=""dropdownMenuButton"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">" & strAppText & "</button>" & _
		"<div class=""dropdown-menu"" aria-labelledby=""dropdownMenuButton""><a class=""dropdown-item"" href=""LimitReductionSummary.asp?ProcessStatus="">All</a>"
		
		objRS.Open "SELECT DISTINCT(ProcessStatus) FROM tblCAPSLimitChange",objCon

		Do Until objRS.EOF
		
			strDropDown = strDropDown & "<a class=""dropdown-item"" href=""LimitReductionSummary.asp?ProcessStatus=" & objRS("ProcessStatus") & """>" & objRS("ProcessStatus") & "</a>"
		
		objRS.Movenext
		Loop
		
		objRS.Close
		
		strDropDown = strDropDown & "<a class=""dropdown-item"" href=""LimitReductionSummary.asp?ProcessStatus=ERROR"">ERROR</a>"

		strDropDown = strDropDown & "</div></div>"
		
	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='LimitReductionSummary.asp?Link=AP&ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(7) & """ onClick=""self.location.href='LimitReductionSummary.asp?Link=AP&ViewButton=Reduced';""><i class=""fa fa-pause-circle""></i> Reduced</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='LimitReductionSummary.asp?Link=AP&ViewButton=OnHold';""><i class=""fa fa-pause-circle""></i> On Hold</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='LimitReductionSummary.asp?Link=AP&ViewButton=Review';""><i class=""fa fa-clock""></i> Awaiting Review</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(5) & """ onClick=""self.location.href='LimitReductionSummary.asp?Link=AP&ViewButton=AddedToNA';""><i class=""fa fa-file""></i> Awaiting Export</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='LimitReductionSummary.asp?Link=AP&ViewButton=AwaitingIssue';""><i class=""fa fa-truck""></i> Rejected</button>" & strDropDown & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(6) & """ onClick=""self.location.href='LimitReductionSummary.asp?Link=AP&ViewButton=TempHold';"" Title=""Click to view Applications on TEMP HOLD""><i class=""fa fa-hand-paper""></i> </button>" & _
				"</div>"

End Sub

Public Function Check_Reduction(Status,OLimit, TLimit, CLimit,DateFrom,DateTo)


Dim intDateDiffFrom
Dim intDateDiffTo

intDateDiffTo = DateDiff("d",DateTo,Now())
intDateDiffFrom = DateDiff("d",DateFrom,Now())


If Status = "Reduced" Then

	If cdbl(OLimit/100) <> cdbl(CLimit) Then
	
		IF  intDateDiffTo > 0 Then
			'Check_Reduction  = "Error A " & Status & ":" & intDateDiffFrom & ":" & intDateDiffTo
			Check_Reduction = "Error" 
		Else
			Check_Reduction = Status
		End If
		
	Else
	
		IF  intDateDiffTo > 0 Then
			Check_Reduction = Status
		Else
			'Check_Reduction = "Error B " & Status & ":" & intDateDiffFrom & ":" & intDateDiffTo
			Check_Reduction = "Error" 
		End If

	End If
	
Else

	If cdbl(OLimit/100) <> cdbl(CLimit) Then
	
		IF  intDateDiffFrom > 0 AND intDateDiffTo < 0 Then
			Check_Reduction  = Status
		Else
			'Check_Reduction = "Error C " & Status & ":" & intDateDiffFrom & ":" & intDateDiffTo
			Check_Reduction = "Error" 
		End If
		
	Else
	
		IF  intDateDiffFrom > 0 AND intDateDiffTo < 0 Then
			'Reduction not applied
			'Check_Reduction = "Error D " & Status & ":" & intDateDiffFrom & ":" & intDateDiffTo
			Check_Reduction = "Error" 
		Else
			Check_Reduction = Status
		End If

	End If


	
End If	

End Function


Set objRS = Nothing
Set objCon = Nothing
%>
