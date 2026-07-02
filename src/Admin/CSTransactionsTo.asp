
<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #include file="../CC/CAPSFunctions.asp" -->
<!-- #Include file=../ADOVBS.inc -->
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
Dim objRS2
Dim objCmd

Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL

Dim lngCardID
Dim strEmployeeID

Dim strStatus
Dim strReviewedBy
Dim dteDateReviewed
Dim lngCreditLimit
Dim strCardType

Dim strHeaderText

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
	Set objRS2 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	

	If IsNull(Session("CardID")) OR Session("CardID") = "" Then Session("CardID")= 0

	If isNull(Session("CardID")) Or Session("CardID") = "" Then 
		Session("CardID") = 0
	End If
	
	If IsNull(Session("FileLoadID")) OR Session("FileLoadID") = "" Then Session("FileLoadID")= 0
	
	If Not IsEmpty(Request.QueryString("UserView")) Then
		Session("UserView") = Request.QueryString("UserView")
	End If

	If Not IsEmpty(Request.QueryString("CardID")) Then
		Session("CardID") = Request.QueryString("CardID")
	End If
	
	If Not IsEmpty(Request.QueryString("CardType")) Then
		strCardType = Request.QueryString("CardType")
	End If
	
	'Response.Write "Card Type = "
	'Response.Write strCardType
	
	If Not IsEmpty(Request.QueryString("FileLoadID")) Then
		Session("FileLoadID") = Request.QueryString("FileLoadID")
		'If the first record from the select list is selected then change the value to zero (0)
		If Session("FileLoadID") = "CS Transactions To Be Sent Today" Then Session("FileLoadID") = ""
		'If Session("FileLoadID") = "CS Transactions To Be Sent Today" Then Session("FileLoadID") = "'' or FileSeqNum Is Null "

	End If
	
	If Not IsEmpty(Request.QueryString("Action")) Then
		
	End If

	If Not IsEmpty(Request.QueryString("PageCombo")) Then
		Session("PageCombo") = Request.QueryString("PageCombo")
	End If
	
	If Not IsEmpty(Request.QueryString("ViewButton")) Then
		Session("ViewButton") = Request.QueryString("ViewButton")
	End If
  
	'Set the View Deleted search and display if user toggles via button
	If Not IsEmpty(Request.QueryString("ViewDeleted")) Then
		If Session("ViewDeleted") = "ViewDeleted" Then
			Session("ViewDeleted") = "NOTViewDeleted"
		Else
			Session("ViewDeleted") = "ViewDeleted"
		End If
	End If
	
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "SubmitApp" Then
			'Response.Write "CPID=" & Session("CarParkingID")
			'Session("CardID") = 0
			Call SubmitApplication()
		End If
	End If
	
	'If the Cancel/Remove has been clicked on the NA File Modal (within the file GetNAfile.asp) then flag the NA File record as removed
	If Request.QueryString("Action") = "CancelCS" Then

		Call RemoveCSRecord(Request.QueryString("CSToDinersID"),Request.QueryString("CSEID"),Request.QueryString("Status"))

	End If

	If Request.QueryString("Action") = "ExportCS" Then	
	
		Call ExportCSFile()
		'Call ExportCSFile(strCardType)
		
	End If
	
	'response.write "Session(FileLoadID)=" & Session("FileLoadID")
	IF Session("FileLoadID") = "0" OR Session("FileLoadID") = "" Then
		strHeaderText = "<span style=""font-weight:normal; font-style:italic; font-size:16px;"">CS File to Be Exported Today</span>"
	Else
		strHeaderText = Session("FileLoadID")
	End If
	
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

setTimeout( 'ShowTimeoutWarning();', 1080000 );

function ShowTimeoutWarning () {     
    window.alert( "********** Warning! **********' \n \n 'You will be automatically logged out in 2 minutes unless you change screens, Close or Save!" ); 
}


function OpenSs(cb) {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("CardTypeSelect");
	var result = e.options[e.selectedIndex].value;
	
	document.getElementById('CardType').value=result;
	
}

function ChangeBatch() {



	var e = document.getElementById("CSFileSelect");
	var result = e.options[e.selectedIndex].text;	
	//alert('CSTransactionsTo.asp?CardType=' + CardType + '&FileLoadID='+result);
	self.location='CSTransactionsTo.asp?FileLoadID='+result
	//self.location='CSTransactionsTo.asp?CardType=' + CardType + '&FileLoadID='+result
	

}

function ChangePage() {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("PageCombo");
	var result = e.options[e.selectedIndex].value;
	
	self.location = 'CSTransactionsTo.asp?PageCombo=' + result;
	//alert(result);
	//document.getElementById('CardType').value=result;
	
}

function loadDoc(varID) {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CSFromDinersDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetCSToDiners.asp?CSToDinersID=" + varID + "", true);
  xhttp.send();
}

function loadCard(varID) {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CSFromDinersDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetCSFromDinersCard.asp?CSFromDinersID=" + varID + "", true);
  xhttp.send();
}

$('#CardTypeSelect').change(function(){
    //alert($(this).val());
})
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

.ModText {
	border: 0px; 
	font-weight:bold;
	font-size: 13px;
	width: 100%;
}

.ModTextAudit {
	border: 0px; 
	font-weight:bold;
	font-size: 13px;
	width: 100%;
	background-color:#e6eeff;
}
	
.ModTextLabel {
	font-size: 13px;
}

</style>
</head>
<body >
<main class="main py-3">
    <div class="container">

	 <!-- Modal Compare -->
<div class="modal fade" id="CSFromDinersMod" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-large modal-dialog-centered modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  CS To Diners File Detail - <%=strCardType%>
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="CSFromDinersDetail">
               
				  
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
	  
<!-- Select Batch Number Modal -->
<div class="modal fade" id="ModalSelectBatch" tabindex="-1" role="dialog" aria-labelledby="ModalSelectBatch" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalApproveTitle" style="font-weight:bold;">CS File To/Export Batch Number</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <div class="col-md-4"><SELECT class="form-control" onChange="ChangeBatch();" name="CSFileSelect" id="CSFileSelect"><% Call LoadBatchList()%></Select></div><br><br>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
      </div>
    </div>
  </div>
</div>
<!-- End Select Batch Number Modal -->


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

			   
<!-- End the first part of the Header Container -->
<div id='tbl-container'>
  <form action="CSTransactionsTo.asp?Action=Search" method="POST" id="frm" name="frm">
	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-10">
				<!--<h4 class="text-left">CS File From Diners <%="File Load ID: " & strHeaderText%></h4>-->
				<h4 class="text-left" data-toggle="modal" data-target="#ModalSelectBatch"  Title="Click to Select a CS File to View">CS File To Diners - <%="File Load ID : " & strHeaderText%>&nbsp;&nbsp;- Date Loaded : <%=Get_Batch_Date(Session("FileLoadID"))%> <button type="button" class="btn btn-primary" onClick="#" data-toggle="modal" data-target="#ModalSelectBatch"  Title="Click to Select a CS File to View"><i class="fa fa-file"></i> Select File</button></h4>
			</div>
			<div class="col-md-2">
				<!--<button type="button" class="btn btn-primary" onClick='window.location="ApplicationsSubmit.asp"'><i class="fa fa-plus"></i> New Application</button>-->
			</div>
			
		</div>

          <div class="row py-2">
            <div class="col-md-9">
              <%Call LoadViewButtons()%>
            </div>
			<div class="col-md-3">
				<div class="form-group has-search">
					<span class="fa fa-search form-control-feedback" onClick="frm.submit();"></span>
				 <input type="text" class="form-control" type="search" id="SearchInput" name="SearchInput" placeholder="Search by Keyword"/>
				 </div>
			</div>
          </div>

      </section>
	  
	  
	 <section class="table py-2">
        <div class="container">
         
                 <%
        
				DisplayTableDetails()
        
				%>	
                
          </div>
        </div>
      </section>
</div>


<!--</DIV>-->
</form>
</div>

</main>
	
<!-- #Include file=../CC/CAPSFooter.asp -->

</body>
</html>
<%

Public Sub DisplayTableDetails()
Dim y
Dim strAction
Dim strStatus
Dim dteDateSubmitted
Dim dteDateReviewed
Dim strSearch
Dim strRecordMessage
Dim strCardNo
Dim strDaysColour
Dim strProcessStatus
Dim dteWarningDate
Dim strPages
Dim strSort
Dim strOrderType
Dim strPages2
Dim strActive
Dim lngPage
Dim strPageCombo
Dim arrPagecombo(6)
Dim strTop
Dim intWritten
Dim strWhere
Dim lngStartingRecord
Dim lngTotalRecords
Dim strFullName

	strSearch = Request.Form("SearchInput")
	
	If IsEmpty(Request.QueryString("SortType")) Then
		'strOrderType = "ASC"
	Else
		If Request.QueryString("SortType") = "ASC" Then
			strOrderType = "DESC"
		Else
			strOrderType = "ASC"
		End If
	End If
	
	If IsEmpty(Request.QueryString("Sort")) Then
		strSort = ""
	Else		
		strSort = " ORDER BY " & Request.QueryString("Sort") & " " & strOrderType
	End If
	
	If Session("ViewButton") = "Processed" Then
		strWhere = " AND CardType = " & Request.QueryString("CardType") & " AND [Status] = 'Processed' "
	ElseIf Session("ViewButton") = "NoChange" Then
		'strWhere = " AND [AuditLogID] IS NULL "
	ElseIf Session("ViewButton") = "Change" Then
		'strWhere = " AND [AuditLogID] IS NOT NULL "
	ElseIf Session("ViewButton") = "Cancelled" Then
		strWhere = " AND [CardStatus] ='01' "
	Else
		'This catches ALL
		strWhere = ""
	End If
	
	'Adjust the search based on whether the deleted CS records are to be displayed (or only those in the last 3 days)
	If Session("ViewDeleted") = "ViewDeleted" THEN
		'strWhere = strWhere & " AND [Status] <> 'Deleted' "
	Else
		strWhere = strWhere & " AND [Status] <> 'Deleted' "
	End If
	
	'Build the TOP Statement
'	If Session("PageCombo") = "" Or IsNull(Session("PageCombo")) Then
'		Session("PageCombo") = 50
'	End If
	
	'If IsNumeric(Session("PageCombo")) Then
	'	strTOP = "TOP " & Session("PageCombo")
	'Else
	'	strTOP = ""
	'End If
	
	'Response.Write strWhere
		
If strSearch = "" OR ISNull(strSearch) Then

	IF Session("FileLoadID") = "0" Then
		'strSQL = "SELECT " & strTOP & " * FROM qryCAPSCSToDinersList WITH(NOLOCK) WHERE CardType = '" & Request.QueryString("CardType") & "' AND (FileSeqNum = '' OR FileSeqNum IS NULL) AND [CSToDinersID] > 0 " & strWhere & strSort
		strSQL = "SELECT " & strTOP & " * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE (FileSeqNum = '' OR FileSeqNum IS NULL) AND Left(CardTypeSub,3) <> 'NAB' AND [CSToDinersID] > 0 " & strWhere & strSort
		
	Else
	
		'strSQL = "SELECT " & strTOP & " * FROM qryCAPSCSToDinersList WITH(NOLOCK) WHERE CardType = '" & Request.QueryString("CardType") & "' AND (FileSeqNum = '" & Session("FileLoadID") & "') AND [CSToDinersID] > 0 " & strWhere & strSort
		strSQL = "SELECT " & strTOP & " * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE (FileSeqNum = '" & Session("FileLoadID") & "') AND Left(CardTypeSub,3) <> 'NAB' AND [CSToDinersID] > 0 " & strWhere & strSort
		'strSQL = "SELECT " & strTOP & " * FROM qryCAPSCSFromDinersAuditLog WITH(NOLOCK) WHERE FileSeqNum = " & Session("FileLoadID") & " AND [CSFromDinersID] > 0 " & strWhere & strSort
	End If
Else
	If Session("FileLoadID") = 0 THEN
		'strSQL = "SELECT " & strTOP & " * FROM qryCAPSCSToDinersList WITH(NOLOCK) WHERE CardType = '" & Request.QueryString("CardType") & "' AND (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
		strSQL = "SELECT " & strTOP & " * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE Left(CardTypeSub,3) <> 'NAB' AND (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
	Else
		'strSQL = "SELECT " & strTOP & " * FROM qryCAPSCSToDinersList WITH(NOLOCK) WHERE CardType = '" & Request.QueryString("CardType") & "' AND FileSeqNum = '" & Session("FileLoadID") & "' AND (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
		strSQL = "SELECT " & strTOP & " * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE Left(CardTypeSub,3) <> 'NAB' AND FileSeqNum = '" & Session("FileLoadID") & "' AND (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
		
		'strSQL = "SELECT " & strTOP & " * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
		'strSQL = "SELECT " & strTOP & " * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE FileSeqNum = '" & Session("FileLoadID") & "' AND (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
		'strSQL = "SELECT " & strTOP & " * FROM qryCAPSCSFromDinersAuditLog WITH(NOLOCK) WHERE FileSeqNum = " & Session("FileLoadID") & " AND (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort

	End If
End If

'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = strSearch
Else
	'strRecordMessage = "for " & Session("UserName") 
End If
'response.write strSQL
objRS.Open strSQL,objCon,3,1

    y = 0
	
	If IsEmpty(Request.QueryString("StartingRecord")) Then
		lngStartingRecord = 0
	Else
		lngStartingRecord = Request.QueryString("StartingRecord")
	End If

	'Write a message in the list if there are no Unactivated Cards
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No CS To Diners records for " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	Else
		objRS.Movelast
		objRS.Movefirst
		lngTotalRecords = objRS.Recordcount
		
		'Set the Page combos here so can be transferred to other pages together
'		arrPagecombo(1) = "50"
'		arrPagecombo(2) = "100"
'		arrPagecombo(3) = "200"
'		arrPagecombo(4) = "500"
'		arrPagecombo(5) = "1000"
'		arrPagecombo(6) = "All"
		
		'Build the Page Combo for TOP statement
'		For x = 1 to 6
		
'			If Session("PageCombo") = arrPagecombo(x) Then
'				strSelected = " SELECTED "
'			Else
'				strSelected = ""
'			End If
'			strPageCombo = strPageCombo & "<option " & strSelected & " value=""" & arrPagecombo(x) & """>" & arrPagecombo(x) & "</option>"
'		Next
		
'		strPageCombo = "<SELECT ID=""PageCombo"" Name=""PageCombo"" onChange=""ChangePage();"">" & strPageCombo & "</select>"
		
		'If the PageCombo is not numeric (ALL or Null) then make it the total records for the recordset (which is set above)
'		If NOT IsNumeric(Session("PageCombo")) Then Session("PageCombo") = lngTotalRecords
	
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Displaying " & lngTotalRecords & " CS To Diners records </span></div></div>"
		
		'Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
        '          "<span class=""panel-subheader"">Displaying " & Session("PageCombo") & " of " & lngTotalRecords & " CS To Diners records (" & lngStartingRecord & " to " & lngStartingRecord + clng(Session("PageCombo")) & ")</span><span class=""panel-subheader"" style=""float:right;"">Number of records per page: " & strPageCombo  & "</span></div></div>"
		
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""CSTransactionsTo.asp?Sort=CSToDinersID&SortType=" & strOrderType & """> CS ID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""CSTransactionsTo.asp?Sort=EIDNo&SortType=" & strOrderType & """> EID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""CSTransactionsTo.asp?Sort=Surname&SortType=" & strOrderType & """> Name <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""CSTransactionsTo.asp?Sort=CardStatus&SortType=" & strOrderType & """> Card Status <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Card No.</th>" & _
			"<th scope=""col""><a href=""CSTransactionsTo.asp?Sort=Status&SortType=" & strOrderType & """> Status  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""CSTransactionsTo.asp?Sort=FileDateTime&SortType=" & strOrderType & """> File Date <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" title=""Card Type""> File Load ID </th>" & _
			"<th scope=""col""><a href=""CSTransactionsTo.asp?Sort=Notes&SortType=" & strOrderType & """> Change Details  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Action</th>" & _
			"</tr></thead><tbody class=""text-left"">"
					
	End If
	
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No CS From Diners records " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	End If
    
	x = 0
	
    Do until objRS.EOF 

'		y = y + 1
		
		'Only write the first 50 records from the starting position
'		If y <= lngStartingRecord + clng(Session("PageCombo")) AND y >= lngStartingRecord - clng(Session("PageCombo")) Then
		'If y <= lngStartingRecord + 50 AND y >= lngStartingRecord - 50 Then
		
			x = x + 1
			
			'Create the actions based on the Process Status of the card
'			Select Case objRS("ProcessStatus")
'			
'			Case  "Removed Unactivated"
'				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='CSTransactionsTo.asp?Action=UnRemove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-list""></i> Re-List</button>"
'			Case "Added to CS"
'
'				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportCS.asp?CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
'				
'			Case "Email Unactivated"
'				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CSTransactionsTo.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
'			
'			Case Else
'				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CSTransactionsTo.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
'				strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#EmailModal""><i class=""fa fa-minus-mail""></i> Email</button>"
'				strAction = strAction & "<button type=""button"" class=""btn btn-outline-info btn-xs"" onclick=""self.location='CSTransactionsTo.asp?Action=Remove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-cross""></i> Remove</button>"
'
'				'strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='CSTransactionsTo.asp?Action=Email&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-mail""></i> Email</button>"
'				'data-toggle="modal" data-target="#EmailModal"
'				
'			End Select
			
			'Create the Action button depending on the Status
			If trim(objRS("Status")) = "Awaiting Export" Then
				'strStatusDisplay = "<span class=""badge badge-pill badge-success"">" & strStatus & "</span>"
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CSTransactionsTo.asp?Action=CancelCS&CSToDinersID=" & objrs("CSToDinersID") & "&CSEID=" & objRS("EIDNo") & "&Status=Deleted'""; title=""Click to Remove from CS File""><i class=""fa fa-times""></i> Remove</button>"
				
			ElseIf objRS("Status") = "Deleted" Then
				'strStatusDisplay = "<span class=""badge badge-pill badge-danger"">" & strStatus & "</span>"
				strAction = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='CSTransactionsTo.asp?Action=CancelCS&CSToDinersID=" & objrs("CSToDinersID") & "&CSEID=" & objRS("EIDNo") & "&Status=Awaiting Export'""; title=""Click to Add to CS File""><i class=""fa fa-plus""></i> Add</button>"

			Else
				'strStatusDisplay = "<span class=""badge badge-pill badge-warning"">" & strStatus & "</span>"
				strAction = ""'<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CSTransactionsTo.asp?Action=CancelCS&CSToDinersID=" & objrs("CSToDinersID") & "&CSEID=" & objRS("EIDNo") & "&Status=Deleted'""; title=""Click to Remove from CS File""><i class=""fa fa-times""></i> Remove</button>"
			End If
			
			'Create the Status list badge based on the status field
			If IsNull(objRS("Status")) Then
				strStatus = ""
			Else
				If objRS("Status") = "Processed" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Processed</span>"
				ElseIf objRS("Status") = "Imported" Then
					strStatus = "<span class=""badge badge-pill badge-info"">Imported</span>"
				ElseIf objRS("Status") = "Awaiting Export" Then
					strStatus = "<span class=""badge badge-pill badge-warning badge-xs"" style=""font-size:12px;"">Awaiting Export</span>"
				ElseIf objRS("Status") = "Exported" Then
					strStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">Exported</span>"
				Else
					strStatus = objRS("Status")
				End If
			End If
			
			'Get the name from two fields and reduce if too long
			If IsNull(objRS("Surname")) Then
				strFullName = ""
			Else
				strFullName = trim(objRS("GivenNames")) & " " & trim(objRS("Surname"))
				
				If Len(strFullName) > 20 THEN
					strFullName = Left(strFullName,20)
				End If
			End If
			
			Response.write "<TR><TD ><a data-toggle=""modal"" data-target=""#CSFromDinersMod"" HREF=""#"" onClick=""loadDoc(" & objRS("CSToDinersID") & ")"">" & objRS(0) & "</a></TD><TD>" & objRS("EIDNo") & "</a></TD>" & _
					"<TD style=""font-size:12px;"" Title=""" & trim(objRS("GivenNames")) & " " & trim(objRS("Surname")) & """><a Target=""_self"" HREF=""../CC/CardDetail.asp?CardNo=" & objRS("CardNo") & """>" & strFullName & "</a></TD><TD><a Target=""_self"" HREF=""../CC/CardDetail.asp?CardNo=" & objRS("CardNo") & """>" & objRS("CardStatus") & "</a></TD>" & _
					"<TD style=""font-size:12px;"">" & MaskCard(objRS("CardNo")) & "</TD><TD >" & strStatus & "</TD>" & _
					"<TD style=""font-size:12px; background-color:#e6eeff;"">" & objRS("FileDateTime") & "</TD><TD style=""font-size:12px; background-color:#e6eeff;"">" & objRS("FileSeqNum") & "</TD><TD style=""font-size:12px; background-color:#e6eeff;"">" & objRS("Notes") & "</TD>" & _
					"<TD>" & strAction & "</TD></TR>"
					
'		End If
		
		objRS.movenext
	Loop

	''''****Variables *******'''''
'lngStartingRecord = The Number of the record (in order of display) starting from. So displays from that number (row) to the Total per page number (Session("PageCombo"))
'Session("PageCombo") = The Number of records to display per page, as selected by the user in the top drop-down
'y = The total number of records in the complete recordset, derived from counting as each record is processed above
'lngPage = The current page selected, mainly for the active flag to display this on screen (the number shaded as currently selected)
'lngTotalRecords  = The recordcount from the recordset when first opening it (movelast then movefirst)

'	If y > 0 Then
		Response.Write "<TR><TH colspan=""10"">Total</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;"">" & x & "</TH></TR>"
'	End If

	'Create the number of pages
'	If IsNumeric(y) Then
'		If y > 1 Then
			
			'Y is the total number of records. Session("PageCombo") is the Total records displayed on screen. Dividing Total records by the number displayed per page gets the number of pages for the bottom of the screen
'			y = y / clng(Session("PageCombo"))
			'y = y / 50
			
			'Determine the number of the page currently displayed
'			If lngStartingRecord = 0 Then
'				lngPage = 1
'			Else
'				If IsNumeric(lngStartingRecord) Then
'					lngPage = (lngStartingRecord/clng(Session("PageCombo")))'+1
					'lngPage = (lngStartingRecord/50)'+1
'				Else
'					lngPage = 1
'				End If
'			End If

'			For x = 0 to y
				
				'In looping through all pages, display only 20 pages, from the starting page (lngPage) to 20 more pages (lngPage + 20)
'				If x > 19 + lngPage OR x < lngPage - 20 Then
					
					'Add the Elipsis (...) to the end of the page numbers if there is more than 20 pages
'					If x = 20 + cint(fix(lngPage)) Then
					'If x = 21 + lngPage Then
'						strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""CSTransactionsTo.asp?StartingRecord=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
'					End If
					
					'Add the Elipsis (...) to the start of the page numbers if there is more than 20 pages and the current place is beyond the first page
'					If x = 0 AND lngPage > 1 AND y > 20 Then
'						strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""CSTransactionsTo.asp?StartingRecord=" & lngTotalRecords - (clng(Session("PageCombo"))*20) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
'					End If
'				Else
'					intWritten = intWritten + 1
'					If intWritten > 21 OR x < 1 then
'					Else
					'If x + 20 > lngPage Then
					'Determine which page number is active (displayed as active)
'					If clng(x) = clng(lngPage) Then
'						strActive = "active"
'					Else
'						strActive = ""
'					End If
				
					'strPages = strPages & "<a href=""CSTransactionsTo.asp?StartingRecord=" & (x * clng(Session("PageCombo"))) & """> " & x + 1 & " </a>"
'					strPages2 = strPages2 & "<li class=""page-item " & strActive & """><a class=""page-link"" href=""CSTransactionsTo.asp?StartingRecord=" & (x * clng(Session("PageCombo"))) & """>" & x & "</a></li>"
					
					'strPages = strPages & "<a href=""CSTransactionsTo.asp?StartingRecord=" & (x * 50) & """> " & x & " </a>"
					'strPages2 = strPages2 & "<li class=""page-item " & strActive & " -" & x & "|" & lngPage & """><a class=""page-link"" href=""CSTransactionsTo.asp?Previous&StartingRecord=" & lngStartingRecord -50 & """>" & x & "</a></li>"
'					End If
'				End If
'			Next
			
'		End If
'	End If
	
	'If y > 0 Then
	'	Response.Write "<TR><TH colspan=""9"" style=""text-align:center;""><a href=""CSTransactionsTo.asp?Previous&StartingRecord=" & lngStartingRecord -50 & """>Previous Page " & strPages & " <a href=""CSTransactionsTo.asp?Previous&StartingRecord=" & lngStartingRecord + 50 & """> Next Page</TH></TR>"
	'End If
	
	
	'Write the End of the table and divs for the above list, as the pagination (below) is in it's own container
	Response.Write "</tbody></table></div>"

	'Write the Pagination objects for all pages based on the total records and the number records displayed on screen
'	If y > 0 Then
		
'		Response.Write "<div class=""container""><div class=""row""><div class=""col-12 text-center"">" & _
'			"<nav aria-label=""Page navigation""><ul class=""pagination""><li class=""page-item"">" & _      
'			"<a class=""page-link"" href=""CSTransactionsTo.asp?StartingRecord=0"" aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
'			strPages2 & _
'			"<li class=""page-item"">" & _
'			"<a class=""page-link"" href=""CSTransactionsTo.asp?StartingRecord=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
'			"</a></li></ul></nav></div></div></div>"
			
			'"<a class=""page-link"" href=""CSTransactionsTo.asp?StartingRecord=" & lngStartingRecord - clng(Session("PageCombo")) & """ aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
'	End If
		
objRS.Close

End Sub

Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(5)
Dim strViewDeleted
Dim strDeletedText

If Session("ViewButton") = "Processed" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "NoChange" Then
	arrButton(3) = "active"
ElseIf Session("ViewButton") = "Change" Then
	arrButton(4) = "active"
ElseIf Session("ViewButton") = "Cancelled" Then
	arrButton(5) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

If Session("ViewDeleted") = "NOTViewDeleted" THEN
	strViewDeleted = "-outline"
	strDeletedText = "Deleted Hidden"
Else
	strViewDeleted = ""
	strDeletedText = "Deleted Displayed"
End If

	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='CSTransactionsTo.asp?ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='CSTransactionsTo.asp?ViewButton=NoChange';""><i class=""fa fa-thumbs-down""></i> View No Change</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='CSTransactionsTo.asp?ViewButton=Change';""><i class=""fa fa-thumbs-up""></i> View Changes</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(5) & """ onClick=""self.location.href='CSTransactionsTo.asp?ViewButton=Cancelled';""><i class=""fa fa-times""></i> View Cancels</button>" & _
				"&nbsp;&nbsp;<button type=""button"" class=""btn btn" & strViewDeleted & "-danger btn"" onClick=""self.location.href='CSTransactionsTo.asp?ViewDeleted=1';"" title=""Click to Show Deleted/Removed CS records""><i class=""fa fa-times""></i> " & strDeletedText & "</button>" & _
				"</div>"

				'"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='CSTransactionsTo.asp?ViewButton=Processed';""><i class=""fa fa-cogs""></i> View Processed</button>" & _
End Sub

Public Sub ExportCSFile()
'Public Sub ExportCSFile(strCardType)
'Procedure to Export the NA File including all of the records which are yet to be exported
Const fsoForWriting = 2
Dim strFilePath
Dim strFileNameStart
Dim strNextFileNumber
Dim intRecordCount
Dim strRecordCount
Dim strFileDateTime
Dim lngFileLoadID
Dim strFileDateTimeSec
Dim strError
Dim strWorkPhone
Dim strMobilePhone
Dim strHomePhone
Dim intCount
Dim strFilePathOnly

Dim objFSO
Dim strFileName

Dim strReportGroup
Dim strReportGroupSettings1
Dim strReportGroupSettings2
Dim strReportGroupSettingsDPC

Dim strSQL
Dim strCardTypeRS

'Get the Default file location for the server then add the filepath and name for the NA File
strFilePath = GetSystemAdmin("ServerFilePath")
'If strCardType = "DTC" Then
	strFileNameStart = GetSystemAdmin("CSFileStart")
'Else
'	strFileNameStart = GetSystemAdmin("CSFileStartDPC")
'End If

'If strCardType = "DTC" Then
	strNextFileNumber = GetSystemAdmin("CSFileNumberTo")
'Else
'	strNextFileNumber = GetSystemAdmin("CSFileNumberToDPC")
'End If

'Pad the number out to 6 digits
strNextFileNumber = PadDigits(strNextFileNumber,6)


'Compile the File name and path from the variables above
strFilePath = strFilePath & "\Admin\CAPSAdmin\Attachments\Diners\DinersTo\" & strFileNameStart & PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & ".txt"
'Get the file path only for the file transfer function
strFilePathOnly = strFilePath & "\Admin\CAPSAdmin\Attachments\Diners\DinersTo\"

'Set the filename to be used when moving the file in the procedure called at the end of this procedure
strFileName = strFileNameStart & PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & ".txt"


strFileDateTime = PadDigits(Right(Year(Now()),4),4) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2)
strFileDateTimeSec = PadDigits(Right(Year(Now()),4),4) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & PadDigits(Hour(Now()),2) & PadDigits(Minute(Now()),2) & PadDigits(Second(Now()),2)
'response.write 	"lngBatchNumber=" & strNextFileNumber & " strNextCSFileNumber= " &  strNextCSFileNumber

'Get the FileLoad details
'lngBatchNumber = GetFileLoadID("NAFile",strNextFileNumber,"")

'If IsNull(lngBatchNumber) OR lngBatchNumber = "" then
'	lngBatchNumber = strNextCSFileNumber
'End If

'Get the System Setting Report Groups to change for each record below
strReportGroupSettings1 = Trim(GetSystemAdmin("DefaultReportGroup"))
strReportGroupSettings2 = Trim(GetSystemAdmin("DefaultReportGroup2"))
strReportGroupSettingsDPC = Trim(GetSystemAdmin("DefaultReportGroupDPC"))

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

'Open the text file
Dim objTextStream

	objCon.Execute "UPDATE tblCAPSCSToDiners SEt Address1 = REPLACE(Address1,char(9),' ') WHERE charindex(char(9),Address1)>0"
	objCon.Execute "UPDATE tblCAPSCSToDiners SEt Address2 = REPLACE(Address2,char(9),' ') WHERE charindex(char(9),Address2)>0"
	objCon.Execute "UPDATE tblCAPSCSToDiners SEt Address3 = REPLACE(Address3,char(9),' ') WHERE charindex(char(9),Address3)>0"
	objCon.Execute "UPDATE tblCAPSCSToDiners SEt CardUpdateInd = 'E' WHERE CardUpdateInd Is Null"



	'Open a recordset of all of the NA File records yet to be exported
	strSQL = "SELECT * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE Left(CardTypeSub,3) <> 'NAB' AND (FileSeqNum Is Null OR FileSeqNum = '') AND ([Status] = 'Awaiting Export' Or [Status] = '')"
	objRS.Open strSQL,objCon
	'objRS.Open "SELECT * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE CardType = '" & strCardType & "' AND (FileSeqNum Is Null OR FileSeqNum = '') AND ([Status] = 'Awaiting Export' Or [Status] = '')",objCon
	
	'Response.Write "SELECT * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE CardType = '" & strCardType & "' AND (FileSeqNum Is Null OR FileSeqNum = '') AND ([Status] = 'Awaiting Export' Or [Status] = '')"

		If objRS.EOF Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No records to write to the CS File. See System Admin: " & strSQL & "</div>"
		Else
		
			'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
			'If strCardType = "DTC" Then
				lngFileLoadID = SaveFileLoadID ("CSToDiners",strFileNameStart & strFileDateTime & ".txt", strFilePath,-1,0,0,0,0,0,0,0,strFileDateTime,strNextFileNumber,"Exported",Session("UserID"),"N")
			'Else
			'	lngFileLoadID = SaveFileLoadID ("CSToDinersDPC",strFileNameStart & strFileDateTime & ".txt", strFilePath,-1,0,0,0,0,0,0,0,strFileDateTime,strNextFileNumber,"Exported",Session("UserID"),"N")
			'End If
	'lngFileLoadID = 0
			Set objTextStream = objFSO.OpenTextFile(strFilePath, fsoForWriting, True)
			objTextStream.WriteLine "H CS" & strFileDateTimeSec & strNextFileNumber & PadSpaceRight(" ",476)
			
			'format(Now(),"yyyymmddhhmmss")+Format(DLookup("Next_File_No","Next_File_No","File_Type = 'NA_Out'"),"000000");\
			
			'Write each record to the text file
			Do Until objRS.EOF
			
				intRecordCount = intRecordCount + 1
				'Display the contents of the text file
				strError = strError & CheckForNull(objRS("EIDNo"),"EIDNo",intRecordCount)
				strError = strError & CheckForNull(objRS("CardNo"),"Card No",intRecordCount)
				strError = strError & CheckForNull(objRS("CardUpdateInd"),"Card Update Ind",intRecordCount)
				strError = strError & CheckForNull(objRS("CardExpiryDate"),"Card Expiry Date",intRecordCount)
				strError = strError & CheckForNull(objRS("CardStatus"),"Card Status",intRecordCount)
				strError = strError & CheckForNull(objRS("Title"),"Title",intRecordCount)
				strError = strError & CheckForNull(objRS("Surname"),"Surname",intRecordCount)
				strError = strError & CheckForNull(objRS("GivenNames"),"Given Names",intRecordCount)
				strError = strError & CheckForNull(objRS("NameOnCard"),"Name on Card",intRecordCount)
				strError = strError & CheckForNull(objRS("Address1"),"Address 1",intRecordCount)
				strError = strError & CheckForNull(objRS("Address2"),"Address 2",intRecordCount)
				strError = strError & CheckForNull(objRS("Address3"),"Address 3",intRecordCount)
				strError = strError & CheckForNull(objRS("Suburb"),"Suburb",intRecordCount)
				strError = strError & CheckForNull(objRS("State"),"State",intRecordCount)
				strError = strError & CheckForNull(objRS("Postcode"),"Postcode",intRecordCount)
				strError = strError & CheckForNull(objRS("HomePhone"),"Home Phone",intRecordCount)
				strError = strError & CheckForNull(objRS("WorkPhone"),"Work Phone",intRecordCount)
				strError = strError & CheckForNull(objRS("MobilePhone"),"Mobile Phone",intRecordCount)
				strError = strError & CheckForNull(objRS("Email"),"Email",intRecordCount)
				strError = strError & CheckForNull(objRS("ReportGroup"),"Report Group",intRecordCount)
				strError = strError & CheckForNull(objRS("CreditLimit"),"Credit Limit",intRecordCount)				
				
				If IsNull(objRS("CSToDinersID")) = False Then
				
					'Set the phone numbers to all zeroes if they are empty
					If IsNull(objRS("HomePhone")) Or objRS("HomePhone") = "" Then
						strHomePhone = "000000000000"
					Else
						strHomePhone = objRS("HomePhone")
					End If
					
					If IsNull(objRS("MobilePhone")) Or objRS("MobilePhone") = "" Then
						'strMobilePhone = "000000000000"
						
						'CHANGE 18 OCT 2021 ------
						'Changed from above to below to remove the 2 leading zeroes from the number (file positions 322 and 323 in CSFileTo) as per Diners request
						strMobilePhone = "0000000000"
					Else
						strMobilePhone = objRS("MobilePhone")
					End If
					
					If IsNull(objRS("WorkPhone")) Or objRS("WorkPhone") = "" Then
						strWorkPhone = "000000000000"
					Else
						strWorkPhone = objRS("WorkPhone")
					End If
					
					'''NEW UPDATE 23rd Feb 2022 to try to fix records not being updated by Diners. This is a bit crazy, but as Diners do not send back some updates (they are not triggered in their system)
					'we are going to change the ReportGroup for each record to the opposite of what it currently is so a change is noted by diners.  This is what was used in the old CAPS for the same purpose
					'and has not yet been implemented in this new version of CAPS.
					
					If IsNull(objRS("CardType")) Or objRS("CardType") = "" Then
						strCardTypeRS = "DTC"
					Else
						strCardTypeRS = Trim(objRS("CardType"))
					End If
					
					'If the Card Type is a DTC then swap Report Groups, otherwise (DPC) make sure the Report GRoup is correct
					If strCardTypeRS = "DTC" Then
					
						If IsNull(objRS("ReportGroup")) Or objRS("ReportGroup") = "" Then
							strReportGroup = "00034776"
						Else
							strReportGroup = Trim(objRS("ReportGroup"))
							'If the Report Group is of one type then change to the other system Setting Default.  This will also capture any other Report Groups and make them only 1 of 2 possible defaults
							If strReportGroup = "00034776" THEN
								strReportGroup = strReportGroupSettings1
							Else
								strReportGroup = strReportGroupSettings2
							End If
						End If
					Else
						If IsNull(objRS("ReportGroup")) Or Trim(objRS("ReportGroup")) = "" Then
							strReportGroup = strReportGroupSettingsDPC
						Else
							strReportGroup = Trim(objRS("ReportGroup"))
						End If
					End If
					
				'UPDATE -- Changed phone number to pad space left instead of pad digits THEN replaced phone numbers with variables above for nulls/empties to have zeroes
					objTextStream.WriteLine "D " & PadSpaceLeft(objRS("EIDNo"),10) & PadDigits(objRS("CardNo"),19) & PadSpaceLeft(objRS("CardUpdateInd"),2) & PadSpaceLeft(objRS("CardExpiryDate"),8) & PadSpaceLeft(objRS("CardStatus"),2) & PadSpaceLeft(objRS("Title"),12) & PadSpaceLeft(objRS("Surname"),25) & PadSpaceLeft(objRS("GivenNames"),30) & PadSpaceLeft(objRS("NameOnCard"),26) & PadSpaceLeft(objRS("Address1"),40) & PadSpaceLeft(objRS("Address2"),40) & PadSpaceLeft(objRS("Address3"),40) & PadSpaceLeft(objRS("Suburb"),25) &  PadSpaceLeft(objRS("State"),4) & PadSpaceLeft(objRS("Postcode"),12) & PadSpaceRight(strHomePhone,12) &  PadSpaceRight(strWorkPhone,12) & PadSpaceRight(strMobilePhone,12) & PadSpaceLeft(objRS("Email"),70) & PadSpaceLeft(strReportGroup,8) & PadSpaceLeft(objRS("CreditLimit"),11) & PadSpaceLeft("",78)  
					'objTextStream.WriteLine "D " & PadSpaceLeft(objRS("EIDNo"),10) & PadDigits(objRS("CardNo"),19) & PadSpaceLeft(objRS("CardUpdateInd"),2) & PadSpaceLeft(objRS("CardExpiryDate"),8) & PadSpaceLeft(objRS("CardStatus"),2) & PadSpaceLeft(objRS("Title"),12) & PadSpaceLeft(objRS("Surname"),25) & PadSpaceLeft(objRS("GivenNames"),30) & PadSpaceLeft(objRS("NameOnCard"),26) & PadSpaceLeft(objRS("Address1"),40) & PadSpaceLeft(objRS("Address2"),40) & PadSpaceLeft(objRS("Address3"),40) & PadSpaceLeft(objRS("Suburb"),25) &  PadSpaceLeft(objRS("State"),4) & PadSpaceLeft(objRS("Postcode"),12) & PadSpaceRight(strHomePhone,12) &  PadSpaceRight(strWorkPhone,12) & PadSpaceRight(strMobilePhone,12) & PadSpaceLeft(objRS("Email"),70) & PadSpaceLeft(objRS("ReportGroup"),8) & PadSpaceLeft(objRS("CreditLimit"),11) & PadSpaceLeft("",78)  
					'objTextStream.WriteLine "D " & PadSpaceLeft(objRS("EIDNo"),10) & PadDigits(objRS("CardNo"),19) & PadSpaceLeft(objRS("CardUpdateInd"),2) & PadSpaceLeft(objRS("CardExpiryDate"),8) & PadSpaceLeft(objRS("CardStatus"),2) & PadSpaceLeft(objRS("Title"),12) & PadSpaceLeft(objRS("Surname"),25) & PadSpaceLeft(objRS("GivenNames"),30) & PadSpaceLeft(objRS("NameOnCard"),26) & PadSpaceLeft(objRS("Address1"),40) & PadSpaceLeft(objRS("Address2"),40) & PadSpaceLeft(objRS("Address3"),40) & PadSpaceLeft(objRS("Suburb"),25) &  PadSpaceLeft(objRS("State"),4) & PadSpaceLeft(objRS("Postcode"),12) & PadSpaceLeft(strHomePhone,12) &  PadSpaceLeft(strWorkPhone,12) & PadSpaceLeft(strMobilePhone,12) & PadSpaceLeft(objRS("Email"),70) & PadSpaceLeft(objRS("ReportGroup"),8) & PadSpaceLeft(objRS("CreditLimit"),11) & PadSpaceLeft("",78)                         
					'objTextStream.WriteLine "D " & PadSpaceLeft(objRS("EIDNo"),10) & PadDigits(objRS("CardNo"),19) & PadSpaceLeft(objRS("CardUpdateInd"),2) & PadSpaceLeft(objRS("CardExpiryDate"),8) & PadSpaceLeft(objRS("CardStatus"),2) & PadSpaceLeft(objRS("Title"),12) & PadSpaceLeft(objRS("Surname"),25) & PadSpaceLeft(objRS("GivenNames"),30) & PadSpaceLeft(objRS("NameOnCard"),26) & PadSpaceLeft(objRS("Address1"),40) & PadSpaceLeft(objRS("Address2"),40) & PadSpaceLeft(objRS("Address3"),40) & PadSpaceLeft(objRS("Suburb"),25) &  PadSpaceLeft(objRS("State"),4) & PadSpaceLeft(objRS("Postcode"),12) &  PadSpaceLeft(objRS("HomePhone"),12) &  PadSpaceLeft(objRS("WorkPhone"),12) & PadSpaceLeft(objRS("MobilePhone"),12) & PadSpaceLeft(objRS("Email"),70) & PadSpaceLeft(objRS("ReportGroup"),8) & PadSpaceLeft(objRS("CreditLimit"),11) & PadSpaceLeft("",78)                         
					'objTextStream.WriteLine "D " & PadSpaceLeft(objRS("EIDNo"),10) & PadDigits(objRS("CardNo"),19) & PadSpaceLeft(objRS("CardUpdateInd"),2) & PadSpaceLeft(objRS("CardExpiryDate"),8) & PadSpaceLeft(objRS("CardStatus"),2) & PadSpaceLeft(objRS("Title"),12) & PadSpaceLeft(objRS("Surname"),25) & PadSpaceLeft(objRS("GivenNames"),30) & PadSpaceLeft(objRS("NameOnCard"),26) & PadSpaceLeft(objRS("Address1"),40) & PadSpaceLeft(objRS("Address2"),40) & PadSpaceLeft(objRS("Address3"),40) & PadSpaceLeft(objRS("Suburb"),25) &  PadSpaceLeft(objRS("State"),4) & PadSpaceLeft(objRS("Postcode"),12) &  PadDigits(objRS("HomePhone"),12) &  PadDigits(objRS("WorkPhone"),12) & PadDigits(objRS("MobilePhone"),12) & PadSpaceLeft(objRS("Email"),70) & PadSpaceLeft(objRS("ReportGroup"),8) & PadSpaceLeft(objRS("CreditLimit"),11) & PadSpaceLeft("",78)                         
				
				strWorkPhone = ""
				strHomePhone = ""
				strMobilePhone = ""
				
				End If
				'If objRS("Status")= "Awaiting Export" Then ----Added this for when the file does not add all records due to a timeout when calling the Export process (Commnad object) below more than ~3000 times
				'intCount = intCount + 1
					'Call the procedure to update each record as exported once added to the CS File -- USE the File style Batch Number not FileLoadID
					'Call ExportCSRecord (objRS("CSToDinersID"),lngFileLoadID,intRecordCount)
					'Call ExportCSRecord (objRS("CSToDinersID"),strNextFileNumber,intCount)'intRecordCount)
					Call ExportCSRecord (objRS("CSToDinersID"),strNextFileNumber,intRecordCount)
				'End If
			objRS.Movenext
			Loop
			
			strRecordCount = PadDigits(intRecordCount,6)
			
			objTextStream.WriteLine "T " & strRecordCount & PadSpaceRight(" ",492)		
			
			
			'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			'strFileName = strFileNameStart & strFileDateTime & ".txt" ------Set at the top of the procedure now as it requires 2 digit year -- 4 digit year required within the file as header
			'If strCardType = "DTC" Then
				Call UpdateFileLoadSummary ("CSToDiners",strNextFileNumber, strFileName, lngFileLoadID)
			'Else
				
			'	Call UpdateFileLoadSummary ("CSToDinersDPC",strNextFileNumber, strFileName, lngFileLoadID)
			'End If
			
			'Call the procedure to update the System Parameter CSFileNumber. Increment the Number by 1.
			Call UpdateBatchNumber(strNextFileNumber)
			'Call UpdateBatchNumber(strNextFileNumber,strCardType)
			
			If strError = "" Then
				Response.Write "<div class=""alert alert-success"" role=""alert"">CS File " & "AUDC_INTOECS_DODCS_D" & strFileDateTime & ".txt" & " ADDED to the CS file export folder!</div>"
			Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">CS File " & "AUDC_INTOECS_DODCS_D" & strFileDateTime & ".txt" & " has errors : " & strError & "</div>"
			End If
				'Close the file and clean up
				objTextStream.Close
		End If

	objRS.Close
	
	'Call the procedure to move the file created to the G Drive
	Call MoveExportFiles(strFilePath, strFileName, strFilePathOnly)

	
Set objTextStream = Nothing
Set objFSO = Nothing

End Sub


Public Sub MoveExportFiles(strFilePathFrom, strFileName, strFilePathOnly)
'Procedure to move the exported files to the G Drive once the files have been produced.
Dim objNetwork
Dim strServer
Dim strUser
Dim strPass
'Dim strFileNameDefault
Dim strFileExtension
Dim objFSO
Dim objStartFolder
Dim strFilePathTo
Dim strFileNameNoExt

Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")

	'Get the System Parameter for the start of the Training File Location
	strServer = GetSystemAdmin("GDriveExportFilePath")

	'Get the System Parameter for the Service Account UserName and Password
	strUser = GetSystemAdmin("CAPSServiceAccountName")
	strPass = GetSystemAdmin("CAPSServiceAccountPassword")

	'Get the System Parameter for the fileName
	'strFileNameDefault = GetSystemAdmin("CSFromDinersFileName")
	
	'New Error check to make sure the file is found
	'Removed AB 080623
	'On Error Resume Next
			
	'Response.Write strServer			
	objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

	''New Error Capture for file not found --added 09/05/2023
	'Removed AB 080623
	'If Err.Number <> 0 Then
		'Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
				'"<span aria-hidden=""true"">&times;</span></button>" & _
				'"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
				'"<span>ERROR!!! CS To Export File NOT MOVED TO: """ & strServer & """ ! See System Admin " & Err.Number & " : " & Err.Description & " </BR></br> CS File Exported To: " & strFilePathFrom & "</span></div></div></div>"
		'Err.Clear
		'Exit Sub
	'End If
	'End remove
    'Removed AB 080623
	'On Error GoTo 0
			
			
		objStartFolder = strServer
	
	'strFilePathTo = strServer & strFileName
	strFilePathTo = strServer & "\" & strFileName

'Update to incliude trailing backslah after mapping drive (where trailing backslash causes an error)
	strServer = strServer & "\"

	'response.write "<Br>strFilePathFrom=" & strFilePathFrom
	'response.write "<Br>strFilePathTo=" & strFilePathTo
	
	'response.write "<Br>strServer=" & strServer & strFileNameNoExt & x & strFileExtension
	
	If objFSO.FileExists(strFilePathTo) Then
		
		strFileExtension = objFSO.GetExtensionName(strFilePathTo)
		strFileNameNoExt = Left(strFilePathTo,Len(strFilePathTo)-Len(strFileExtension)-1)
		
		For x = 1 to 10
			If objFSO.FileExists(strFilePathTo) Then
				strFilePathTo = strFileNameNoExt & x & "." & strFileExtension
			Else
				strFilePathTo = strFileNameNoExt & x + 1 & "." & strFileExtension
				'strFilePathTo = strServer & strFileNameNoExt & x & strFileExtension
				'strFilePathTo = strFilePathOnly & strFileName & x & strFileExtension
				
				'Move the file to the Loaded folder
				objFSO.MoveFile strFilePathFrom, strFilePathTo
				'objFSO.MoveFile strFilePathFrom, strServer & strFileNameNoExt & x & strFileExtension
				'objFSO.MoveFile strFilePathFrom & strFileNameNoExt & x & strFileExtension
				
				x = 10
			End If
		Next
	Else
		'Move the file to the Loaded folder
		'objFSO.MoveFile strFilePathOnly & strFileName,strServer & strFileName
		objFSO.MoveFile strFilePathFrom,strFilePathTo'strServer & strFileName
		
	End If
	
	'Remove the trailing backslash as the FSO object doesn;t like this on the new DPE server
	strServer = Left(strServer, len(strServer)- 1)			
	objNetwork.RemoveNetworkDrive strServer, True, False
		 
	Set objFSO = Nothing
	Set objNetwork = Nothing
	
	Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
				"<span aria-hidden=""true"">&times;</span></button>" & _
				"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
				"<span>SUCCESS!!! CS To Export File """ & strFilePathTo & """</span></div></div></div>"
				'"<span>SUCCESS!!! CS To Export File """ & filePath & """ Not Loaded! File path not Found. objExcelCon.Open. See System Admin " & Err.Number & " : " & Err.Description & "</span></div></div></div>"
				
End Sub

Public Sub ExportCSRecord(lngCSToDinersID,lngBatchNumber,x)
'Procedure to Change the Status of CS file records being exported and adds an Audit Log record
Dim intRecord

  	With objCmd

		.CommandTimeout = 900
		.CommandType = 4
		.CommandText = "spCAPSCSFileExportCard"
		
		'Only create the parameters the first time the procedure is created otherwise there will be an error
		If x = 1 Then
			.Parameters.Append objCmd.CreateParameter("CSToDinersID", adVarChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("BatchNumber", adVarChar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CSFileExportOutput", adInteger, adParamOutput)
		End If	
	
		.Parameters("CSToDinersID") = lngCSToDinersID
		.Parameters("BatchNumber") = lngBatchNumber
		.Parameters("UpdatedBy") = Session("UserID")
		
		.ActiveConnection = objCon
		 
	End With
   
	objCmd.Execute        
  
	'Return the result of the Save Function.
	intRecord = objCmd.Parameters.Item("CSFileExportOutput") 
 
End Sub


Public Sub UpdateBatchNumber(lngBatchNumber)
'Public Sub UpdateBatchNumber(lngBatchNumber,strCardType)
'Procedure to update the BatchNumber field in the System Parameters table with the next number
Dim strSQL

	'If the Batch Number is a number then update the System Parameter, otherwise post an error to the screen
	If IsNumeric(lngBatchNumber) Then
		lngBatchNumber = lngBatchNumber + 1
		
		'If strCardType = "DTC" Then
			strSQL = "UPDATE tblCAPSSystemParameters SET [ParameterValue] = '" & lngBatchNumber & "' WHERE [ParameterName] = 'CSFileNumberTo'"
		'Else
		'	strSQL = "UPDATE tblCAPSSystemParameters SET [ParameterValue] = '" & lngBatchNumber & "' WHERE [ParameterName] = 'CSFileNumberToDPC'"
		'End If
		
		objCon.Execute strSQL
	
	Else
		
		Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR! CS File Batch Number: " & lngBatchNumber & " is not a number. See System Admin.</div>"
		
	End If

End Sub

Public Sub RemoveCSRecord(lngCSToDinersID, strEmployeeID, strStatus)

Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCSFileRemoveCard"

			.Parameters.Append objCmd.CreateParameter("CSToDinersID", adVarChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("CSFileRemoveOutput", adInteger, adParamOutput)
			
			.Parameters("CSToDinersID") = lngCSToDinersID
			.Parameters("UpdatedBy") = Session("UserID")
			.Parameters("Status") = strStatus
			
			.ActiveConnection = objCon
			 
		End With
	   'response.write "spCAPSCSFileRemoveCard " & lngCSToDinersID & "," & Session("UserID") & "," & strStatus
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CSFileRemoveOutput") 
	 
		If intRecord = 0 Then
			If strStatus = "Deleted" Then
				Response.Write "<div class=""alert alert-danger"" role=""alert"">CS Record for " & strEmployeeID & " NOT Removed from CS File! An Error has occurred. See System Admin with CS File ID: " & lngCSToDinersID & " </div>"
			Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">CS Record for " & strEmployeeID & " NOT Added to the CS File! An Error has occurred. See System Admin with CS File ID: " & lngCSToDinersID & " </div>"
			End If
		Else
			If strStatus = "Deleted" Then
				Response.Write "<div class=""alert alert-success"" role=""alert"">CS Record for " & strEmployeeID & " REMOVED from the CS file!</div>"
			Else
				Response.Write "<div class=""alert alert-success"" role=""alert"">CS Record for " & strEmployeeID & " ADDED to the CS file!</div>"
			End If
		End If
		
	
End Sub

Public Sub LoadBatchList()
'Public Sub LoadBatchList(strCardType)
'Description:	Loads all Batch Numbers to a list for selecting and searching/filtering

	'If strCardType = "DTC" Then
		objRS.Open "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = 'CSToDiners' AND [Deleted] = 'N' ORDER By [FileSeqNum] DESC",objCon
	'Else
	'	objRS.Open "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = 'CSToDinersDPC' AND [Deleted] = 'N' ORDER By [FileSeqNum] DESC",objCon
	'End If
	'objRS.Open "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = 'CSToDiners' AND [Deleted] = 'N' ORDER By [FileSeqNum] DESC",objCon
  
	Response.write "<OPTION value=""0"">Select a Batch to View...</OPTION><OPTION value=""0"">CS Transactions To Be Sent Today</OPTION>"
	
		Do Until objRS.EOF 
			
			Response.write "<OPTION value=""" & objRS("FileLoadID") & """>" & objRS("FileSeqNum") & "</OPTION>"
			
			objRS.Movenext
			
		Loop
	
	objRS.Close
	
End Sub

Public Function Get_Batch_Date(strFileLoadID)
'Public Function Get_Batch_Date(strFileLoadID,strCardType)

	'If strCardType = "DTC" Then
		objRS2.Open "SELECT * FROM tblCAPSFileLoad WHERE FileType = 'CSToDiners' AND FileSeqNum = '" & strFileLoadID & "'",objCon
	'Else
	'	objRS2.Open "SELECT * FROM tblCAPSFileLoad WHERE FileType = 'CSToDinersDPC' AND FileSeqNum = '" & strFileLoadID & "'",objCon
	'End If
	
	If Not objRS2.EOF Then
	
		Get_Batch_Date = objRS2("DateLoaded")
	
	End If
	
	objRS2.Close

End Function



Public Function CheckForNull(strValue,strField,intRow)

	If IsNull(strValue) Then 
		CheckForNull = "</BR> Error in field " & strField & " is Null at row " & intRow & " : "
	Else
		CheckForNull = ""
	End If

End Function

Set objRS2 = Nothing
Set objRS = Nothing
Set objCon = Nothing
%>
