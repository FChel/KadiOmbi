
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #Include file=CAPSFunctions.asp -->
<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp?State=Expired")
If IsNull(Session("UserView")) Then

End If

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
Dim dteDateReviewed
Dim lngCreditLimit
Dim lngCurrentRecord
Dim arrStatus(9)
Dim strOptions

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	

	arrStatus(1) = "Added To NA"
	arrStatus(2) = "Awaiting Export"
	arrStatus(3) = "Awaiting Issue"
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

	If Not IsEmpty(Request.QueryString("ApplicationTypeName")) Then
		Session("ApplicationTypeNameView") = Request.QueryString("ApplicationTypeName")
		
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
			Session("ApplicationTypeNameView") = ""
		End If
		
	End If
  
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "SubmitApp" Then
			Call SubmitApplication()
		End If
	End If
	
	If Not IsEmpty(Request.QueryString("IgnoreZeroes")) Then
		Session("IgnoreZeroes") = Request.QueryString("IgnoreZeroes")
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
	self.location = 'CAPSCompareCDMC.asp?SaveStatus='+frm.AppStatID.value+'&NewStat='+frm.NewStatus.options[frm.NewStatus.selectedIndex].value;
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

  xhttp.open("GET", "AJAX/GetExcelFiles2.asp?FileStart=CAPSCompareCDMCForCSTo", true);
  xhttp.send();
}

function ChangePage() {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("PageCombo");
	var result = e.options[e.selectedIndex].value;
	
	self.location = 'CAPSCompareCDMC.asp?PageCombo=' + result;
	//alert(result);
	//document.getElementById('CardType').value=result;
	
}
function loadCSToDiners(varID,varCardNo) {

	//If(varCardNo)
	//const varCardNo= ['a'+ varCardNo];
	//alert(document.getElementById("CardNoMod").value)
	//alert(varID)
	//varCardNo=document.getElementById("CardNoMod").value
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

function loadCDMC(varID) {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CDMCDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "AJAX/GetCDMCDetails.asp?EmployeeID=" + varID + "", true);
  xhttp.send();
}

function loadCSFromDiners(varID,varCardNo) {

//alert(varCardNo)
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

<form action="CAPSCompareCDMC.asp?Action=Search&Link=RP" method="POST" id="frm" name="frm">
				 
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
 
 <!-- CDMC Modal -->
<div class="modal fade" id="CDMCModal" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-dialog-right modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  CDMC Detail
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="CDMCDetail">
               
				  
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
 <!-- End of CDMC To Diners Modal -->
 
 
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

	

<!-- End the first part of the Header Container -->
<div id='tbl-container'>

	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-8">
				<h2 class="text-left">Card Contact Details Different to CDMC <i class="fa help-tooltip fa-question-circle" data-toggle="tooltip" title="All Active Dines cards with contact details (Surname, Email, Address and Phone numbers) different to what is currently on the CDMC file"></i></h2>
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
	  
	  </form>
	  
	 <section class="table py-2">
  

                 <%
        
				DisplayTableDetails()
        
				%>	
                
    

      </section>
</div>


</DIV>
<!--</form>-->
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
Dim strAction2
Dim strStatus
Dim strAddress
Dim strSearch
Dim lngTotalRecords

Dim strSort
Dim strOrderType
Dim strRecordMessage
Dim strFields
Dim strWhere
Dim strQuery
Dim strTop
Dim dteDateUpdated
Dim arrSort(19)
Dim strSortArrow

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
Dim strCardHolderName
Dim strCardNoFull
Dim strOnCSFile

Dim strViewHeaders
Dim strViewFields

Dim strEmailShort
Dim strExcelTopSearch

Dim strZeroNumbers

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
	
	'If there is no sort then sort by the most recent submitted
	If IsNull(strSort) Or strSort = "" Then strSort = " ORDER BY DateUpdated DESC"
	
	'If Session("ViewButton") = "Review" Then
	'	strWhere = " AND [Status] = 'Awaiting review' "
	'ElseIf Session("ViewButton") = "OnHold" Then
	'	strWhere = " AND [Status] = 'On Hold' "
	'ElseIf Session("ViewButton") = "AwaitingIssue" Then
	'	strWhere = " AND [Status] = 'Awaiting issue' "
	'ElseIf Session("ViewButton") = "AddedToNA" Then
	'	strWhere = " AND [Status] = 'Added To NA' "
	'ElseIf Session("ViewButton") = "TempHold" Then
	'	strWhere = " AND [Status] = 'Temp Hold' "
	'Else
		'This catches ALL
		'strWhere = ""
		
	'End If
	
	'Build sort based on selection
	Select Case Request.QueryString("Sort")
	
		Case "CardID"
			arrSort(1) = strSortArrow
		Case "EmployeeID"
			arrSort(2) = strSortArrow
		Case "Surname"
			arrSort(3) = strSortArrow
		Case "CardType"
			arrSort(4) = strSortArrow
		Case "DateUpdated"
			arrSort(5) = strSortArrow
		Case "WorkPhone"
			arrSort(6) = strSortArrow
		Case "OutDinersWorkPhone"
			arrSort(7) = strSortArrow
		Case "MobilePhone"
			arrSort(8) = strSortArrow
		Case "OutDinersMobilePhone"
			arrSort(9) = strSortArrow
		Case "Address1"
			arrSort(10) = strSortArrow
		Case "OutDinersAddress1"
			arrSort(11) = strSortArrow
		Case "Address2"
			arrSort(12) = strSortArrow
		Case "OutDinersAddress2"
			arrSort(13) = strSortArrow
		Case "Suburb"
			arrSort(14) = strSortArrow
		Case "OutDinersSuburb"
			arrSort(15) = strSortArrow
		Case "Surname"
			arrSort(16) = strSortArrow
		Case "FormalLastName"
			arrSort(17) = strSortArrow
		Case "Email"
			arrSort(18) = strSortArrow
		Case "Email_Address"
			arrSort(19) = strSortArrow
	End Select
	
	'Set the FROM Statement and query used
	If Session("ViewButton") = "Phone" Then
		strQuery = " FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) "
		
		strViewHeaders = "<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=WorkPhone&Link=RP&SortType=" & strOrderType & """> CAPS Work Ph <i class=""fa fa-sort" & arrSort(6) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersWorkPhone&Link=RP&SortType=" & strOrderType & """> CDMC Work Ph <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Mobile Phone on Card (at Diners)""><a href=""CAPSCompareCDMC.asp?Sort=MobilePhone&Link=RP&SortType=" & strOrderType & """> CAPS Mobile Ph <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Mobile Phone on CDMC (after formatting and IF valid)""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersMobilePhone&Link=RP&SortType=" & strOrderType & """> CDMC Mobile Ph <i class=""fa fa-sort" & arrSort(9) & """></i></a></th>"
		
		'Set the view for ignore zero phone numbers button
		If Session("IgnoreZeroes") = "Yes" Then
			strZeroNumbers = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onClick=""self.location='CAPSCompareCDMC.asp?IgnoreZeroes=No'"" title=""Mobile Number zeroes ignored (where the mobile number in CAPS from Diners is blank, but 10 zeroes appear instead). Click to show zero numbers.""><i class=""fa fa-ban""></i> <i class=""fa fa-phone""></i> No Zeroes</button>"	
		Else
			strZeroNumbers = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onClick=""self.location='CAPSCompareCDMC.asp?IgnoreZeroes=Yes'"" title=""Mobile Number zeroes displayed (where the mobile number in CAPS from Diners is blank, but 10 zeroes appear instead). Click to hide zero numbers.""><i class=""fa fa-phone""></i> Show Zeroes</button>"	
		End If
			
	ElseIf Session("ViewButton") = "Address" Then
		strQuery = " FROM qryCAPSCompareCDMCCardsReport_Address WITH(NOLOCK) "
		
		strViewHeaders = "<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=Address1&Link=RP&SortType=" & strOrderType & """> CAPS Address 1 <i class=""fa fa-sort" & arrSort(10) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersAddress1&Link=RP&SortType=" & strOrderType & """> CDMC Address 1 <i class=""fa fa-sort" & arrSort(11) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Address 2 on Card (at Diners)""><a href=""CAPSCompareCDMC.asp?Sort=Address2&Link=RP&SortType=" & strOrderType & """> CAPS Address 2 <i class=""fa fa-sort" & arrSort(12) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Address 2 on CDMC (after formatting and IF valid)""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersAddress2&Link=RP&SortType=" & strOrderType & """> CDMC Address 2 <i class=""fa fa-sort" & arrSort(13) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Suburb, State and Postcode on Card (at Diners)""><a href=""CAPSCompareCDMC.asp?Sort=Suburb&Link=RP&SortType=" & strOrderType & """> CAPS Suburb, St, PC <i class=""fa fa-sort" & arrSort(14) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Suburb, State and Postcode on CDMC (after formatting and IF valid)""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersSuburb&Link=RP&SortType=" & strOrderType & """> CDMC Suburb, St, PC <i class=""fa fa-sort" & arrSort(15) & """></i></a></th>"
			
	ElseIf Session("ViewButton") = "Name" Then
		strQuery = " FROM qryCAPSCompareCDMCCardsReport_Name WITH(NOLOCK) "
		
		strViewHeaders = "<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=Surname&Link=RP&SortType=" & strOrderType & """> CAPS Surname <i class=""fa fa-sort" & arrSort(16) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=FormalLastName&Link=RP&SortType=" & strOrderType & """> CDMC Surname <i class=""fa fa-sort" & arrSort(17) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Email Address on Card (at Diners)""><a href=""CAPSCompareCDMC.asp?Sort=Email&Link=RP&SortType=" & strOrderType & """> CAPS EMail <i class=""fa fa-sort" & arrSort(18) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Email Address on CDMC (after formatting and IF valid)""><a href=""CAPSCompareCDMC.asp?Sort=Email_Address&Link=RP&SortType=" & strOrderType & """> CDMC EMail <i class=""fa fa-sort" & arrSort(19) & """></i></a></th>"
		
		
	ElseIf Session("ViewButton") = "OnCSFile" Then
		strQuery = " FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) "
	Else
		'This catches ALL
		strQuery = " FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) "
		
		strViewHeaders = "<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=WorkPhone&Link=RP&SortType=" & strOrderType & """> CAPS Work Ph <i class=""fa fa-sort" & arrSort(6) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersWorkPhone&Link=RP&SortType=" & strOrderType & """> CDMC Work Ph <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Mobile Phone on Card (at Diners)""><a href=""CAPSCompareCDMC.asp?Sort=MobilePhone&Link=RP&SortType=" & strOrderType & """> CAPS Mobile Ph <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Mobile Phone on CDMC (after formatting and IF valid)""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersMobilePhone&Link=RP&SortType=" & strOrderType & """> CDMC Mobile Ph <i class=""fa fa-sort" & arrSort(9) & """></i></a></th>"
			
	End If
	
	'Set the top value for all queries
	strTop = " TOP 500 "
	'strQuery = " FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) "
	strWhere = " WHERE [CardID] > 0 "
	strFields = " * "
	
If strSearch = "" OR ISNull(strSearch) Then
	'If Session("UserView") = "All" Then
		'strSQL = "SELECT TOP 100 * FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) WHERE [CardID] > 0 " & strWhere & strSort
		strWhere = " WHERE [CardID] > 0 "
	'Else
	'	strSQL = "SELECT TOP 1000 * FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' " & strSort
	'End If
	
Else
	'If the user has entered the date lookup then process this, otherwise perform the EmployeeID and Name searches
	If UCASE(Left(strSearch,2)) = "D:" Then
	
		strSearchDate = Right(strSearch,Len(strSearch)-2)
		
		If Len(strSearchDate) > 5 Then

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
			If Mid(strSearchDate,Len(strSearchDate)-2,1) = "/" OR Mid(strSearchDate,Len(strSearchDate)-2,1) = "-" Then
				strSearchYear = "20" & Right(strSearchDate,2)
			Else
				strSearchYear = Right(strSearchDate,4)
			End If
			
			strSearchDate = strSearchDay & "-" & MonthName(strSearchMonth) & "-" & strSearchYear
		End If
		'response.write strSearchDate
		If IsDate(strSearchDate) Then 'DateAdd("d", -1, strSearchDate)
		
			'strDateFrom = DateAdd("d", -1, strSearchDate)
			strDateFrom = strSearchDate
			strDateTo = DateAdd("d", +1, strSearchDate)

			'strSQL = "SELECT TOP 100 * FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) WHERE (DateUpdated > '" & strDateFrom & "' AND DateUpdated < '" & strDateTo & "')" & strWhere & strSort
			strWhere = "WHERE (DateUpdated > '" & strDateFrom & "' AND DateUpdated < '" & strDateTo & "')" 
		Else
			'strSQL = "SELECT TOP 100 * FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) WHERE (DateUpdated = '" & strSearchDate & "')" & strWhere & strSort
			strWhere = "WHERE (DateUpdated = '" & strSearchDate & "')"
		End If
		
	Else
		'If Session("UserView") = "All" Then
			'If the user has entered a search term with a space then assume this is a first and last name so search on that only
			If Instr(1,strSearch," ")>0 Then
				arrNames = Split(strSearch," ")
				strFNameSearch = arrNames(0)
				strLNameSearch = arrNames(1)
				
				'strWhere = " WHERE ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%')"
				'strSQL = "SELECT TOP 100 * FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%'))" & strWhere & strSort
				strWhere = "WHERE (EmployeeID Like '%" & strSearch & "%' OR ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%'))" 
			Else
				'strSQL = "SELECT TOP 100 * FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strWhere & strSort
				strWhere = "WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " 
			End If
		'Else
			'strSQL = "SELECT TOP 100 * FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strSort
		'	strWhere = "WHERE EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') "
		'End If
	
	'End of 
	End If
	

	
End If
	
'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = strSearch
Else
	strRecordMessage = Session("UserName") 
End If

	If Session("ViewButton") = "Phone" Then
		If Session("IgnoreZeroes") = "Yes" Then
			strWhere = strWhere & " AND MobilePhone<>'0000000000' "
		Else
		End If
	End If
	
	'Build the Select Statement ---- SQL
	strSQL = "SELECT " & strTop & strFields & strQuery & strWhere & strSort
	
	'Set the global query
	Session("ExcelSearch") = strSQL

	If Not isEmpty(Session("ExcelSearch")) Then
		'strExcelTopSearch = Instr(1,Session("ExcelSearch"),"TOP ")
		'If strExcelTopSearch > 0 Then
		'End If
		
		Session("ExcelSearch") = Replace(Session("ExcelSearch"),"TOP 100","")
		Session("ExcelSearch") = Replace(Session("ExcelSearch"),"TOP 500","")
		Session("ExcelSearch") = Replace(Session("ExcelSearch"),"TOP 1000","")
		
	End If

response.write strSQL	
	
	objRS.CursorLocation = 3 
	objRS.Open strSQL,objCon',3,1
		
	'Write a message in the list if there are no Cards
	If objRS.EOF Then
		Response.Write "<table class=""table table-compact text-left""><thead><TR><TH colspan=""10"" Style=""text-align:center;"">No Cards for " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR></thead></table>"
	Else
		'objRS.Movelast
		'objRS.Movefirst
		'lngTotalRecords = objRS.Recordcount
		
		
		'Set the FROM Statement and query used
		If Session("ViewButton") = "Phone" Then
			strViewHeaders = "<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=WorkPhone&Link=RP&SortType=" & strOrderType & """> CAPS Work Ph <i class=""fa fa-sort" & arrSort(6) & """></i></a></th>" & _
				"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersWorkPhone&Link=RP&SortType=" & strOrderType & """> CDMC Work Ph <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
				"<th scope=""col"" Title=""Mobile Phone on Card (at Diners)""><a href=""CAPSCompareCDMC.asp?Sort=MobilePhone&Link=RP&SortType=" & strOrderType & """> CAPS Mobile Ph <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
				"<th scope=""col"" Title=""Mobile Phone on CDMC (after formatting and IF valid)""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersMobilePhone&Link=RP&SortType=" & strOrderType & """> CDMC Mobile Ph <i class=""fa fa-sort" & arrSort(9) & """></i></a></th>"
			
		ElseIf Session("ViewButton") = "Address" Then
			
		ElseIf Session("ViewButton") = "Name" Then
			
		ElseIf Session("ViewButton") = "OnCSFile" Then
			
		Else
			'This catches ALL
			strViewHeaders = "<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=WorkPhone&Link=RP&SortType=" & strOrderType & """> CAPS Work Ph <i class=""fa fa-sort" & arrSort(6) & """></i></a></th>" & _
				"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersWorkPhone&Link=RP&SortType=" & strOrderType & """> CDMC Work Ph <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
				"<th scope=""col"" Title=""Mobile Phone on Card (at Diners)""><a href=""CAPSCompareCDMC.asp?Sort=MobilePhone&Link=RP&SortType=" & strOrderType & """> CAPS Mobile Ph <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
				"<th scope=""col"" Title=""Mobile Phone on CDMC (after formatting and IF valid)""><a href=""CAPSCompareCDMC.asp?Sort=OutDinersMobilePhone&Link=RP&SortType=" & strOrderType & """> CDMC Mobile Ph <i class=""fa fa-sort" & arrSort(9) & """></i></a></th>"
				
		End If
	
		
		'Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
        '          "<span class=""panel-subheader"">Displaying " & lngTotalRecords & " Cards</span><span class=""panel-subheader"" style=""float:right;""></span></div></div>"
		
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4>Diners Cards with Phone details different to CDMC &nbsp;&nbsp;&nbsp;&nbsp;" & strZeroNumbers & "</h4>" & _
                  "<span class=""panel-subheader""></span><span class=""panel-subheader"" style=""float:right;""></span></div></div>"
		
		Response.Write "<div class=""container""><div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=CardID&Link=RP&SortType=" & strOrderType & """> Card ID <i class=""fa fa-sort" & arrSort(1) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=EmployeeID&Link=RP&SortType=" & strOrderType & """> EID <i class=""fa fa-sort" & arrSort(2) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=Surname&Link=RP&SortType=" & strOrderType & """> Name <i class=""fa fa-sort" & arrSort(3) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=CardType&Link=RP&SortType=" & strOrderType & """> Card Type <i class=""fa fa-sort" & arrSort(4) & """></i></a></th>" & _
			strViewHeaders & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=DateUpdated&Link=RP&SortType=" & strOrderType & """> Date Updated <i class=""fa fa-sort" & arrSort(5) & """></i></a></th>" & _
			"<th scope=""col"" Title=""Is this card in today's CS File?"">On CS?</th>" & _
			"<th scope=""col"">View CDMC</th>" & _
			"<th scope=""col"">View CS</th>" & _
			"</tr></thead><tbody class=""text-left"">"
				
				
	End If
    
	x = 0
	
    Do until objRS.EOF 
	
			x = x + 1
			
			'Format the Date updated value
			If IsNull(objRS("DateUpdated")) OR objRS("DateUpdated")= "" Then
				dteDateUpdated = ""
			Else
				dteDateUpdated = FormatDateTime(objRS("DateUpdated"),vbShortDate)
			End If
			
			
			If IsNull(objRS("Surname")) Or objRS("Surname") = "" Then
				strCardHolderName = ""
			Else
				strCardHolderName = Left(objRS("FirstName") & " " & objRS("Surname"),15)
			End If
			

			strAddress = Trim(objRS("Address1")) & " " & Trim(objRS("Address2")) & " " & Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode"))
			
			If len(strAddress) > 15 Then strAddress = left(strAddress,15) & "..."
			
			If IsNull(objRS("DateUpdated")) Then
				dteDateUpdated = ""
			Else
				dteDateUpdated = FormatDateTime(objRS("DateUpdated"),vbShortDate)
			End If
			
			If IsNull(objRS("CardNumber")) Then
				strCardNoFull = ""
			Else
				strCardNoFull = objRS("CardNumber")
			End If
		
			strAction = ""
			
			strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#CDMCModal"" HREF=""#"" onClick=""loadCDMC(" & objRS("EmployeeID") & ")""><i class=""fa fa-eye""></i> CDMC</button>"
			'strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#CSToDinersModal"" HREF=""#"" onClick=""loadCSToDiners(" & objRS("EmployeeID") & ",'" & strCardNoFull & "')""><i class=""fa fa-eye""></i> View Audit</button>"	
			strAction2 = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#CSToDinersModal"" HREF=""#"" onClick=""loadCSToDiners(" & objRS("EmployeeID") & ",'" & strCardNoFull & "')""><i class=""fa fa-eye""></i> CS</button>"	
			
			If IsNull(objRS("CSToDinersID")) Then
				
				'strAction = strAction & "<i class=""fa fa-check""></i>"
				'strAction = strAction & " <button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CSTransactionsTo.asp?Link=AD&EmployeeID=" & objRS("EmployeeID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
				strOnCSFile = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;""><i class=""fa fa-times""></i></span>"
			Else
				'strAction = strAction & "Not on CS File"
				'strAction = strAction & "<i class=""fa fa-times""></i>"
				strOnCSFile = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;""><i class=""fa fa-check""></i></span>"
			End If
			
			
			'Set the FROM Statement and query used
			If Session("ViewButton") = "Phone" Then
				
				strViewFields = "<TD style=""font-size:11px; color:red;"" title=""Work phone on Card - " & objRS("WorkPhone") & " - (Diners current Work Phone)"">" & objRS("WorkPhone") & "</TD>" & _
							"<TD style=""font-size:11px; color:green;"" title=""Work phone on CDMC - " & objRS("OutDinersWorkPhone") & " - (CDMC Work Phone - formatted IF valid)"">" & objRS("OutDinersWorkPhone") & "</TD>" & _
							"<TD style=""font-size:11px; color:red;"" title=""Mobile phone on Card - " & objRS("MobilePhone") & " - (Diners current Mobile Phone)"">" & objRS("MobilePhone") & "</TD>" & _
							"<TD style=""font-size:11px; color:green;"" title=""Mobile phone on CDMC - " & objRS("OutDinersMobilePhone") & " - (CDMC Mobile Phone - formatted IF valid)"">" & objRS("OutDinersMobilePhone") & "</TD>"

			ElseIf Session("ViewButton") = "Address" Then
				
				strViewFields = "<TD style=""font-size:11px; color:red;"" title=""Address 1 on Card - " & objRS("Address1") & " - (Diners current Address1)"">" & objRS("Address1") & "</TD>" & _
							"<TD style=""font-size:11px; color:green;"" title=""Address 1 on CDMC - " & objRS("OutDinersAddress1") & " - (CDMC Address1 - formatted IF valid)"">" & objRS("OutDinersAddress1") & "</TD>" & _
							"<TD style=""font-size:11px; color:red;"" title=""Address 2 on Card - " & objRS("Address2") & " - (Diners current Address2)"">" & objRS("Address2") & "</TD>" & _
							"<TD style=""font-size:11px; color:green;"" title=""Address 2 on CDMC - " & objRS("OutDinersAddress2") & " - (CDMC Address2 - formatted IF valid)"">" & objRS("OutDinersAddress2") & "</TD>" & _
							"<TD style=""font-size:11px; color:red;"" title=""Suburb, State and PostCode on Card - " & objRS("Suburb") & " " & objRS("State") & " " & objRS("PostCode") & " - (Diners current Suburb, State and PostCode)"">" & objRS("Suburb") & " " & objRS("State") & " " & objRS("PostCode") & "</TD>" & _
							"<TD style=""font-size:11px; color:green;"" title=""Suburb, State and PostCode on CDMC - " & objRS("OutSuburb") & " " & objRS("OutState") & " " & objRS("OutPostCode") & " - (CDMC Suburb, State and PostCode - formatted IF valid)"">" & objRS("OutSuburb") & " " & objRS("OutState") & " " & objRS("OutPostCode") & "</TD>"

			ElseIf Session("ViewButton") = "Name" Then
				
				'Shorten the email if it is too long
				If Not IsNull(objRS("Email")) Then
					If Len(objRS("Email")) > 35 Then
						strEmailShort = "font-size:9px;"
					Else
						strEmailShort = "font-size:11px;"
					End If
				End If
				
				strViewFields = "<TD style=""font-size:11px; color:red;"" title=""Surname on Card - " & objRS("Surname") & " - (Diners current Surname)"">" & objRS("Surname") & "</TD>" & _
							"<TD style=""font-size:11px; color:green;"" title=""Surname on CDMC - " & objRS("FormalLastName") & " - (CDMC Formal Last Name - formatted IF valid)"">" & objRS("FormalLastName") & "</TD>" & _
							"<TD style=""" & strEmailShort & " color:red;"" title=""Email on Card - " & objRS("Email") & " - (Diners current Email)"">" & objRS("Email") & "</TD>" & _
							"<TD style=""" & strEmailShort & " color:green;"" title=""Email on CDMC - " & objRS("Email_Address") & " - (CDMC Email - formatted IF valid)"">" & objRS("Email_Address") & "</TD>"

							
			ElseIf Session("ViewButton") = "OnCSFile" Then
				
			Else
				'This catches ALL
				
				strViewFields = "<TD style=""font-size:11px; color:red;"" title=""Work phone on Card - " & objRS("WorkPhone") & " - (Diners current Work Phone)"">" & objRS("WorkPhone") & "</TD>" & _
							"<TD style=""font-size:11px; color:green;"" title=""Work phone on CDMC - " & objRS("OutDinersWorkPhone") & " - (CDMC Work Phone - formatted IF valid)"">" & objRS("OutDinersWorkPhone") & "</TD>" & _
							"<TD style=""font-size:11px; color:red;"" title=""Mobile phone on Card - " & objRS("MobilePhone") & " - (Diners current Mobile Phone)"">" & objRS("MobilePhone") & "</TD>" & _
							"<TD style=""font-size:11px; color:green;"" title=""Mobile phone on CDMC - " & objRS("OutDinersMobilePhone") & " - (CDMC Mobile Phone - formatted IF valid)"">" & objRS("OutDinersMobilePhone") & "</TD>"
			End If
	
		
			Response.Write "<TR><TD style=""font-size:11px;""><a Target=""_self"" HREF=""CardDetail.asp?Link=RP&CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:11px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """ title=""" & objRS("FirstName") & " " & objRS("Surname") & """>" & strCardHolderName & "</a></TD>" & _
					"<TD style=""font-size:11px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
					strViewFields & _
					"<TD style=""font-size:11px;"" title=""CDMC Update Received by CAPS - " & objRS("DateUpdated") & " - (These are the current CDMC details held by CAPS)"">" & dteDateUpdated & "</TD>" & _
					"<TD>" & strOnCSFile & "</TD>" & _
					"<TD>" & strAction & "</TD>" & _
					"<TD>" & strAction2 & "</TD></TR>"
		
		objRS.movenext
	Loop
	

		Response.Write "<TR><TH colspan=""10"">Total</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;"">" & x & "</TH></TR></tbody></table></div></div></div>"
				'"<TH colspan=""3"" style=""text-align:center;"">" & x & "</TH></TR></tbody></table></div></div></div>"
				     
	
	
	
objRS.Close

End Sub


Public Sub DisplayTableDetails_Old()

Dim y, x
Dim strAction
Dim strStatus
Dim strAddress
Dim dteDateUpdated
Dim dteDateReviewed
Dim strSearch
Dim lngTotalRecords
Dim lngStartingPage
Dim strPages
Dim strSort
Dim strOrderType
Dim strRecordMessage
Dim strWhere
Dim dteDateSubmitted
Dim arrSort(9)
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
Dim strApplicationTypeNameSelected
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
	If IsNull(strSort) Or strSort = "" Then strSort = " ORDER BY DateUpdated DESC"
	
	If Session("ViewButton") = "Review" Then
		strWhere = " AND [Status] = 'Awaiting review' "
	ElseIf Session("ViewButton") = "OnHold" Then
		strWhere = " AND [Status] = 'On Hold' "
	ElseIf Session("ViewButton") = "AwaitingIssue" Then
		strWhere = " AND [Status] = 'Awaiting issue' "
	ElseIf Session("ViewButton") = "AddedToNA" Then
		strWhere = " AND [Status] = 'Added To NA' "
	ElseIf Session("ViewButton") = "TempHold" Then
		strWhere = " AND [Status] = 'Temp Hold' "
	Else
		'This catches ALL
		strWhere = ""
		
	End If

	'If an Application Type has been selected then add that to the WHERE statement
	If Not IsNull(Session("ApplicationTypeNameView")) Then
		If Session("ApplicationTypeNameView") = "" Then
		Else
			strWhere = strWHERE & " AND [ApplicationTypeName] = '" & Session("ApplicationTypeNameView")  &"'"
		End If
		
	End If
	
If strSearch = "" OR ISNull(strSearch) Then
	'If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 1000 * FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) WHERE [ApplicationID] > 0 " & strWhere & strSort
	'Else
	'	strSQL = "SELECT TOP 1000 * FROM qryCAPSCompareCDMCCardsReport_Phone WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' " & strSort
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
			
			'strSQL = "SELECT * FROM qryCAPSApplicationsList WITH(NOLOCK) WHERE (DateUpdated > '" & strSearchDate & "' AND DateUpdated < " & strSearchDate & ")" & strWhere & strSort
			strSQL = "SELECT TOP 1000 * FROM qryCAPSApplicationsList WITH(NOLOCK) WHERE (DateUpdated > '" & strDateFrom & "' AND DateUpdated < '" & strDateTo & "')" & strWhere & strSort
			
			'strSQL = "SELECT * FROM qryCAPSApplicationsList WITH(NOLOCK) WHERE (DateUpdated > '" & DateAdd("d", -1, strSearchDate) & "' AND DateUpdated < '" & DateAdd("d", +1, strSearchDate) & "')" & strWhere & strSort
			'response.write strsql
		Else
			strSQL = "SELECT TOP 1000 * FROM qryCAPSApplicationsList WITH(NOLOCK) WHERE (DateUpdated = '" & strSearchDate & "')" & strWhere & strSort
		End If
		
	Else
		'If Session("UserView") = "All" Then
			'If the user has entered a search term with a space then assume this is a first and last name so search on that only
			If Instr(1,strSearch," ")>0 Then
				arrNames = Split(strSearch," ")
				strFNameSearch = arrNames(0)
				strLNameSearch = arrNames(1)
				
				'strWhere = " WHERE ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%')"
				strSQL = "SELECT TOP 1000 * FROM qryCAPSApplicationsList WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%'))" & strWhere & strSort
			Else
				strSQL = "SELECT TOP 1000 * FROM qryCAPSApplicationsList WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strWhere & strSort
			End If
		'Else
		'	strSQL = "SELECT TOP 1000 * FROM qryCAPSApplicationsList WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strSort
		'End If
	
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
		Case "DateUpdated"
			arrSort(7) = strSortArrow
		Case "DateReviewed"
			arrSort(8) = strSortArrow
		Case "LimitDateFrom"
			arrSort(9) = strSortArrow
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
		If Right(Session("ApplicationTypeNameView"),12) = "Limit Change" Then
			strApplicationTypeNameSelected = "<th scope=""col"" style=""font-size:14px;""><a href=""CAPSCompareCDMC.asp?Sort=LimitDateFrom&Link=RP&SortType=" & strOrderType & """> Limit Start <i class=""fa fa-sort" & arrSort(9) & """></i></a></th>"
		Else
			strApplicationTypeNameSelected = ""
		End If
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=ApplicationID&Link=RP&SortType=" & strOrderType & """> App ID <i class=""fa fa-sort" & arrSort(1) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=EmployeeID&Link=RP&SortType=" & strOrderType & """> EID <i class=""fa fa-sort" & arrSort(2) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=Surname&Link=RP&SortType=" & strOrderType & """> Name <i class=""fa fa-sort" & arrSort(3) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=CardType&Link=RP&SortType=" & strOrderType & """> Card Type <i class=""fa fa-sort" & arrSort(4) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=ApplicationType&Link=RP&SortType=" & strOrderType & """> App Type <i class=""fa fa-sort" & arrSort(5) & """></i></a></th>" & _
			"<th scope=""col"">Address</th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=Status&Link=RP&SortType=" & strOrderType & """> Status  <i class=""fa fa-sort" & arrSort(6) & """></i></a> <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Card Process Status - the stage the application is at.""></i></th>" & _
			"<th scope=""col"" Title=""Date AE602 XML Form was LOADED into CAPS""><a href=""CAPSCompareCDMC.asp?Sort=DateUpdated&Link=RP&SortType=" & strOrderType & """> Submitted <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSCompareCDMC.asp?Sort=DateReviewed&Link=RP&SortType=" & strOrderType & """> Reviewed <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
			strApplicationTypeNameSelected & "<th scope=""col"">Action</th>" & _
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
			
			'Get the errors and resolved count for the application for use in the display list View Appp Button
			If IsNull(objRS("Errors")) Or objRS("Errors") = "" Then
				strErrorCount = ""
				strViewAppType = "secondary"
			Else
				strErrorCount = objRS("Errors")
				
				strViewAppType = "danger"
				strViewAppTitle = strErrorCount & " Errors in the application"

			End If
			'Get the number of resolved errors
			If IsNull(objRS("Resolved")) Or objRS("Resolved") = "" Then
				strResolvedCount = ""
			Else
				strResolvedCount = objRS("Resolved")
				strViewAppTitle = strViewAppTitle & " : " & strResolvedCount & " Resolved"
			End If
			
			'If there have been errors then check to see if they are all resolved (number of errors = number of resolved)
			If IsNumeric(strErrorCount) = true Then
			
				If (cint(strErrorCount) - cint(strResolvedCount)) = 0 Then
					strViewAppType = "secondary"
				End If
			End If
			
			If IsNull(objRS("Surname")) Or objRS("Surname") = "" Then
				strApplicantName = ""
			Else
				strApplicantName = Left(objRS("FirstName") & " " & objRS("Surname"),15)
			End If
			
			
			'Select the Action and Status buttons/pills based on the application status
			Select Case objRS("Status")
			
			Case  "Received"
				strAction = "<button type=""button"" class=""btn btn-primary btn-sx"" onclick=""self.location='CAPSCompareCDMC.asp?Action=Release&Link=RP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> Release</button>"
				strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CAPSCompareCDMC.asp?Action=Reject&Link=RP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			Case "Added To CS"

				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
			Case "Added To NA"

				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportNA.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View NA</button>"	
				
				strStatus = "<span class=""badge badge-pill badge-info"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
			Case "Submitted"
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CAPSCompareCDMC.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"

				'strStatus  = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='CAPSCompareCDMC.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted to GCFO</button>"
				strStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">Submitted to GCFO</span>"
			Case "Deleted"
				'strAction = "Deleted - " & objRS("DateUpdated")'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='CAPSCompareCDMC.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				strAction = "<Span style=""font-size:12px;"">Deleted - " & dteDateUpdated & "</span>"'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='CAPSCompareCDMC.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				'strStatus  = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CAPSCompareCDMC.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Cancelled By Applicant</button>"
				strStatus = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">Deleted</span>"
			Case "Rejected"
				'strAction = "Deleted - " & objRS("DateUpdated")'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='CAPSCompareCDMC.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				strAction = "<Span style=""font-size:12px;"">Rejected - " & dteDateUpdated & "</span>"'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='CAPSCompareCDMC.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				'strStatus  = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CAPSCompareCDMC.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Cancelled By Applicant</button>"
				strStatus = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">Rejected</span>"
			Case "ASFIN Approved"
				strAction = "<button type=""button"" title=""Approved by GCFO"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CAPSCompareCDMC.asp?Link=RP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-check""></i>GCFO Approved</button>"
			
				'strStatus  = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='CAPSCompareCDMC.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Approved by GCFO</button>"
				strStatus = "<span class=""badge badge-pill badge-success"">Approved by ASFIN</span>"
			Case  "Awaiting Review"
				'strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='CAPSCompareCDMC.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "&EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Release</button>"
				'strAction = strAction & " <button type=""button"" class=""btn btn-outline-danger btn-xs"" onclick=""self.location='CAPSCompareCDMC.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-times""></i> Delete</button>"
				
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=RP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				
				'If the application is a Limit Change then open the Limit Change Submit screen, otherwise open the Normal submit screen
				If strApplicationType = "LimitChange" Then
					strAction = strAction & "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsLimitSubmit.asp?Action=Release&Link=RP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				Else
					strAction = strAction & "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsSubmit.asp?Action=Release&Link=RP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				End If
				
				''''SEP 2021 --- CHANGED from the below to the status line 2 below so that applications can have their status changed if they are Awaiting Review
				'strStatus = "<span class=""badge badge-pill badge-warning"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
				strStatus = "<span class=""badge badge-pill badge-warning "" data-toggle=""modal"" data-target=""#StatusModal"" data-AppID=""" & objRS("ApplicationID") & """  data-AppName=""" & objRS("FirstName") & " " & objRS("Surname") & " - " & objRS("CardType") & " " & objRS("CardTypeSub") & " Application"" onClick=""OpenSs(this);"">" & objRS("Status") & "</span>"
				
			Case  "Awaiting issue"
				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportNA.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View NA</button>"	
				strStatus = "<span class=""badge badge-pill badge-info"">" & objRS("Status") & "</span>"
			Case "On Hold"
				strStatus = "<span class=""badge badge-pill badge-secondary "" data-toggle=""modal"" data-target=""#StatusModal"" data-AppID=""" & objRS("ApplicationID") & """  data-AppName=""" & objRS("FirstName") & " " & objRS("Surname") & " - " & objRS("CardType") & " " & objRS("CardTypeSub") & " Application"" onClick=""OpenSs(this);"">" & objRS("Status") & "</span>"
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=RP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				
				'If the application is a Limit Change then open the Limit Change Submit screen, otherwise open the Normal submit screen
				If strApplicationType = "LimitChange" Then
				
					strAction = strAction & "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsLimitSubmit.asp?Action=Release&Link=RP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				Else
					strAction = strAction & "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsSubmit.asp?Action=Release&Link=RP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				End If
			Case "Temp Hold"
				strStatus = "<span class=""badge badge-pill badge-danger "" data-toggle=""modal"" data-target=""#StatusModal"" data-AppID=""" & objRS("ApplicationID") & """  data-AppName=""" & objRS("FirstName") & " " & objRS("Surname") & " - " & objRS("CardType") & " " & objRS("CardTypeSub") & " Application"" onClick=""OpenSs(this);"">" & objRS("Status") & "</span>"
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=RP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				
				'If the application is a Limit Change then open the Limit Change Submit screen, otherwise open the Normal submit screen
				If strApplicationType = "LimitChange" Then
				
					strAction = strAction & "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsLimitSubmit.asp?Action=Release&Link=RP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				Else
					strAction = strAction & "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsSubmit.asp?Action=Release&Link=RP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				End If
			Case  "Done"
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=RP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				strStatus = "<span class=""badge badge-pill badge-info"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
			Case Else
				'strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CAPSCompareCDMC.asp?Action=Cancel&Link=RP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				'strAction = "Rejected"
				'strStatus  = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='CAPSCompareCDMC.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted</button>"
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=RP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				strStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
			End Select

			strAddress = Trim(objRS("Address1")) & " " & Trim(objRS("Address2")) & " " & Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode"))
			
			If len(strAddress) > 15 Then strAddress = left(strAddress,15) & "..."
			
			If IsNull(objRS("DateUpdated")) Then
				dteDateUpdated = ""
			Else
				dteDateUpdated = FormatDateTime(objRS("DateUpdated"),vbShortDate)
			End If
			
			If IsNull(objRS("DateReviewed")) Then
				dteDateReviewed = ""
			Else
				dteDateReviewed = FormatDateTime(objRS("DateReviewed"),vbShortDate)
			End If
			
			'If the Application types selected are Limit Changes the display the Start Date for the Limit Change
			If Right(Session("ApplicationTypeNameView"),12) = "Limit Change" Then
				'Get the Limit change date relative to today to determine the text colour
				intLimitDateDiff = DateDiff("d",objRS("LimitDateFrom"),Now())
				
				If intLimitDateDiff > 1 Then
					strLimitColour = "color:Green; font-weight:bold;"
				ElseIf intLimitDateDiff < 1 AND intLimitDateDiff > -5 Then
					strLimitColour = "color:red; font-weight:bold;"
				Else
					strLimitColour = ""
				End If
				
				strApplicationTypeNameSelected = "<TD style=""font-size:13px; text-align:center; " & strLimitColour & """ Title=""" & intLimitDateDiff*-1 & " Days from today"">" & objRS("LimitDateFrom") & "</TD>"
			Else
				strApplicationTypeNameSelected = ""
			End If
		
			response.write "<TR><TD><a Target=""_self"" HREF=""ApplicationDetail.asp?Link=RP&ApplicationID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""ApplicationDetail.asp?ApplicationID=" & objRS(0) & """ title=""" & objRS("FirstName") & " " & objRS("Surname") & """>" & strApplicantName & "</a></TD><TD style=""font-size:12px;""><a Target=""_self"" HREF=""ApplicationDetail.asp?ApplicationID=" & objRS(0) & """>" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
					"<TD style=""font-size:12px;"" title=""" & objRS("ApplicationType") & """><a Target=""_self"" HREF=""ApplicationDetail.asp?ApplicationID=" & objRS(0) & """>" & objRS("ApplicationTypeName") & "</a></TD><TD style=""font-size:12px;"">" & strAddress & "</TD><TD style=""font-size:12px;"">" & strStatus & "</TD>" & _
					"<TD style=""font-size:13px; text-align:center;"">" & dteDateUpdated & "</TD><TD style=""font-size:13px; text-align:center;"">" & dteDateReviewed & "</TD>" & strApplicationTypeNameSelected & _
					"<TD>" & strAction & "</TD></TR>"
					
			'response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""CAPSCompareCDMC.asp?ApplicationID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & strAction & "</a></TD>" & _
			'		"<TD><a Target=""_self"" HREF=""CAPSCompareCDMC.asp?ApplicationID=" & objRS(0) & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""CAPSCompareCDMC.asp?ApplicationID=" & objRS(0) & """>" & objRS(2) & "</a></TD>" & _
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
'				strPages = strPages & "<a href=""CAPSCompareCDMC.asp?Link=RP&StartingPage=" & (x * 50) & """> " & x & " </a>"
'			Next
			
'		End If
'	End If
	
'	If y > 0 Then
'		Response.Write "<TR><TH colspan=""9"" style=""text-align:center;""><a href=""CAPSCompareCDMC.asp?Link=RP&Previous&StartingPage=" & lngStartingPage -50 & """>Previous Page " & strPages & " <a href=""CAPSCompareCDMC.asp?Link=RP&Previous&StartingPage=" & lngStartingPage + 50 & """> Next Page</TH></TR>"
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
			'		strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""CAPSCompareCDMC.asp?StartingPage=" & lngTotalRecords - (clng(Session("PageCombo"))*20) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
			'	End If
			End If
		
			If x > 20 Then
				If bolSkip = False Then
					strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""CAPSCompareCDMC.asp?StartingPage=" & lngTotalPages & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
					bolSkip = True
				End If
			Else
				
			End If
		End If
		
		If bolSkip = True Then
		Else
			strPages2 = strPages2 & "<li class=""page-item " & strActive & """><a class=""page-link"" href=""CAPSCompareCDMC.asp?StartingPage=" & x & """>" & x & "</a></li>"
		End If
		
	Next
	
	'Write the Pagination objects for all pages based on the total records and the number records displayed on screen
	If lngTotalPages > 0 Then
		
		'Add the Elipsis (...) to the end of the page numbers if there is more than 20 pages
		'If x = 20 + cint(fix(lngPage)) Then
		'If x = 21 + lngPage Then
		'	strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""CAPSCompareCDMC.asp?StartingPage=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
		'End If
		
		'Add the Elipsis (...) to the start of the page numbers if there is more than 20 pages and the current place is beyond the first page
		'If x = 0 AND lngPage > 1 AND y > 20 Then
		'	strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""CAPSCompareCDMC.asp?StartingPage=" & lngTotalRecords - (clng(Session("PageCombo"))*20) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
		'End If
					
		'Response.Write "<div class=""container""><div class=""row""><div class=""col-12 text-center"">" & _
		'	"<nav aria-label=""Page navigation""><ul class=""pagination""><li class=""page-item"">" & _      
		'	"<a class=""page-link"" href=""CAPSCompareCDMC.asp?StartingPage=1"" aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
		'	strPages2 & _
		'	"<li class=""page-item"">" & _
		'	"<a class=""page-link"" href=""CAPSCompareCDMC.asp?StartingPage=" & lngTotalPages & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
		'	"</a></li></ul></nav></div></div></div>"

				'"<a class=""page-link"" href=""CAPSCompareCDMC.asp?StartingPage=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
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
			'dteDateReviewed = objRS("DateReviewed")
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
			dteDateReviewed = ""
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
Dim arrButton(5)

If Session("ViewButton") = "Phone" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "Address" Then
	arrButton(3) = "active"
ElseIf Session("ViewButton") = "Name" Then
	arrButton(4) = "active"
ElseIf Session("ViewButton") = "OnCSFile" Then
	arrButton(5) = "active"
	
Else
	'This catches ALL
	arrButton(1) = "active"
End If

		
	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='CAPSCompareCDMC.asp?Link=RP&ViewButton=All';""><i class=""fa fa-wine-glass""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='CAPSCompareCDMC.asp?Link=RP&ViewButton=Phone';""><i class=""fa fa-phone""></i> Phone</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='CAPSCompareCDMC.asp?Link=RP&ViewButton=Address';""><i class=""fa fa-house-user""></i> Address</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='CAPSCompareCDMC.asp?Link=RP&ViewButton=Name';""><i class=""fa fa-address-card""></i> Name</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(5) & """ onClick=""self.location.href='CAPSCompareCDMC.asp?Link=RP&ViewButton=OnCSFile';""><i class=""fa fa-folder""></i> On CS File</button>" & _
				"</div>"

End Sub


Set objRS = Nothing
Set objCon = Nothing
%>
