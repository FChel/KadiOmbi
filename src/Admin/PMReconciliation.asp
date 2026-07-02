
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
Dim objRS1
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
Dim strWherePM
Dim strAddAllButton

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
	Set objRS1 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	

	If IsNull(Session("CardID")) OR Session("CardID") = "" Then Session("CardID")= 0

	If isNull(Session("CardID")) Or Session("CardID") = "" Then 
		Session("CardID") = 0
	End If
	
	If IsNull(Session("FileLoadID")) OR Session("FileLoadID") = "" Then Session("FileLoadID")= 0
	If IsNull(Session("PMMissingReport")) OR Session("PMMissingReport") = "" Then  Session("PMMissingReport") = "Missing"
	
	If Not IsEmpty(Request.QueryString("UserView")) Then
		Session("UserView") = Request.QueryString("UserView")
	End If

	If Not IsEmpty(Request.QueryString("CardID")) Then
		Session("CardID") = Request.QueryString("CardID")
	End If
	
	If Not IsEmpty(Request.QueryString("FileLoadID")) Then
		Session("FileLoadID") = Request.QueryString("FileLoadID")
		'If the first record from the select list is selected then change the value to zero (0)
		If Session("FileLoadID") = "CS Transactions To Be Sent Today" Then Session("FileLoadID") = ""
		'If Session("FileLoadID") = "CS Transactions To Be Sent Today" Then Session("FileLoadID") = "'' or FileSeqNum Is Null "

	End If

	If Not IsEmpty(Request.QueryString("PageCombo")) Then
		Session("PageCombo") = Request.QueryString("PageCombo")
	End If
	
	If Not IsEmpty(Request.QueryString("ViewButton")) Then
		Session("ViewButton") = Request.QueryString("ViewButton")
	End If
  
	If Not IsEmpty(Request.QueryString("PendingViewButton")) Then
		Session("PendingViewButton") = Request.QueryString("PendingViewButton")
	End If
	
	If Not IsEmpty(Request.QueryString("PMMissingReport")) Then
		Session("PMMissingReport") = Request.QueryString("PMMissingReport")
		
		'Set the session variable for the Existing type (address, expiry etc..) t0 empty when the Missing is selected to avoid dataset issues
		Session("PMRecNameView") = ""
	End If

	If Not IsEmpty(Request.QueryString("PMRecNameView")) Then
		Session("PMRecNameView") = Request.QueryString("PMRecNameView")
	End If
	
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "AddToPM" Then
			Call AddToPM(Request.QueryString("CardID"),Request.QueryString("CardEID"),Request.QueryString("CardType"))
		End If
		
		If Request.QueryString("Action") = "Cancel" Then
			Call CancelFromPM(Request.QueryString("CardID"))
		End If
		
		If Request.QueryString("Action") = "CMSDetailsSave" Then
			Call CardCMSDetailsUpdate(Request.QueryString("CardID"),Request.QueryString("DefaultCompany"),Request.QueryString("CMSUser"),Request.QueryString("DefaultCostCentre"),Request.QueryString("PMLoadStatus"),Request.QueryString("ReportGroup"))
			
		End If
		
	End If
	
	'If the Cancel/Remove has been clicked on the NA File Modal (within the file GetNAfile.asp) then flag the NA File record as removed
	If Request.QueryString("Action") = "CancelCS" Then

		Call RemoveCSRecord(Request.QueryString("CSToDinersID"),Request.QueryString("CSEID"),Request.QueryString("Status"))

	End If

	If Request.QueryString("Action") = "ExportCS" Then	
	
		Call ExportCSFile()
		
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
	
	self.location='PMReconciliation.asp?FileLoadID='+result
}

function ChangePage() {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("PageCombo");
	var result = e.options[e.selectedIndex].value;
	
	self.location = 'PMReconciliation.asp?PageCombo=' + result;
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
  xhttp.open("GET", "../CC/AJAX/GetPMDetails.asp?CardID=" + varID + "", true);
  xhttp.send();

}

function loadDocCMS(varID, varCardID) {

	//document.getElementById("CSFromDinersDetail").innerHTML = '<div class="py-3"><span id="Progress" ><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Processing...</b></span></div>';
	
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CSFromDinersDetail").innerHTML = this.responseText;
    }
  };
  
  xhttp.open("GET", "../CC/AJAX/GetCMSAccount.asp?EmpID='" + varID + "'&CardID=" + varCardID + "", true);
  xhttp.send();

}

function addAllRecords() {
  var xhttp = new XMLHttpRequest();
  var strWherePM = document.getElementById("addall").value
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("addallresponse").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetAddAllCMSRecords.asp?Where=" + strWherePM + "", true);
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

function SaveCMSDetails() {
//Function to save the CSM details if updated in the CMS modal --GetCMSAccount.asp

document.getElementById("CSFromDinersDetailFoot").innerHTML = '<div class="py-3"><span id="Progress" ><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Processing...</b></span></div>';
	
	
	//If there is no list/table then the modal will write a text field (with T at the end of the ID) otherwise the option/select will exist, so check which value to get
	var e = document.getElementById("DefaultCompany");
		if(e){
			var DCompany = e.options[e.selectedIndex].value;
			} else {
			var DCompany = document.getElementById("DefaultCompanyT").value
		} 
	
	var f = document.getElementById("DefaultCostCentre");
		if(f){
			var DCostCentre = f.options[f.selectedIndex].value;
			} else {
			var DCostCentre = document.getElementById("DefaultCostCentreT").value
		} 
	
	
	self.location = "PMReconciliation.asp?Action=CMSDetailsSave&CardID="+document.getElementById("CardIDCMS").value+"&CMSUser="+document.getElementById("CMSUser").value+"&DefaultCompany="+DCompany+"&DefaultCostCentre="+DCostCentre+"&PMLoadStatus="+document.getElementById("PMLoadStatus").value+"&ReportGroup="+document.getElementById("ReportGroup").value;
}

function SaveNameChange() {

	self.location = "CardDetail.asp?Action=NameChange&NewNOC="+document.getElementById("NewNOC").value+"&AppStatus="+document.getElementById("NCAppNameStatus").value+"&NewFirstName="+document.getElementById("NewFirstName").value+"&NewTitle="+document.getElementById("NewTitle").value+"&NewSurname="+document.getElementById("NewSurname").value+"&NCCardType="+document.getElementById("NCCardType").value+"&NCAppID="+document.getElementById("NCAppID").value;
}

$('#CardTypeSelect').change(function(){
    alert($(this).val());
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
     <div class="modal-dialog modal modal-dialog-centered modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  CAPS In ProMaster Detail
                </h5>
				
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="CSFromDinersDetail">
               <div class="py-3"><span id="Progress" ><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Processing...</b></span></div>
		
				  
                
            </div>
				<div class="modal-body" id="CSFromDinersDetailFoot">
				</div>
			<div class="modal-footer">
				<button type="button" class="btn btn-primary" onClick="SaveCMSDetails();"><i class="fa fa-check"></i> Save changes</button>
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
               <img src="../images/Load.gif" style="vertical-align:middle;" /> &nbsp;&nbsp;Please wait while loading...</div>

			   
<!-- End the first part of the Header Container -->
<div id='tbl-container'>
  <form action="PMReconciliation.asp?Action=Search" method="POST" id="frm" name="frm">
	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row" >
			
			<div class="col-md-8">
				<!--<h4 class="text-left">CS File From Diners <%="File Load ID: " & Session("FileLoadID")%></h4>-->
				<!--<h4 class="text-left" data-toggle="modal" data-target="#ModalSelectBatch"  Title="Click to Select a CS File to View">CAPS Cards Not in ProMaster</h4>-->
				<h4 class="text-left" Title="Click to Select a CS File to View">CAPS Cards Not in ProMaster</h4>


			</div>
			<div class="col-md-4 float-right">
			<!--<%Call LoadPendingButtons()%>-->
				<!--<button type="button" class="btn btn-primary" onClick='window.location="ApplicationsSubmit.asp"'><i class="fa fa-plus"></i> New Application</button>-->
			</div>
			
		</div>

          <div class="row py-2">
            <div class="col-md-9">
              <%Call LoadViewButtons()%>
			<div id="addallresponse">
			
			</div>
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
Dim strPMStatus
Dim strDateIssued
Dim strTable
Dim strExtraHeaders
Dim strExtraFields
Dim strPMLoadDate
Dim strCardNoSearch
Dim strLinkString
Dim strAccountNumber
Dim strEmployeeIDVar
Dim strEmployeeName
Dim strSortSearch

	
	strSearch = Request.Form("SearchInput")
	strSortSearch = strSearch
	'Session("Search") = strSortSearch
	
	If IsEmpty(Request.QueryString("SortSearch")) THEN
		strSearch = Request.Form("SearchInput")
	Else
		strSortSearch = Request.QueryString("SortSearch")
	End If
	
	If IsEmpty(Request.QueryString("SortType")) Then
		strOrderType = "ASC"
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

	If Session("ViewButton") = "DTC" Then
		strWhere = " AND [CardType] = 'DTC' "
	ElseIf Session("ViewButton") = "DPC" Then
		strWhere = " AND [CardType] = 'DPC' "
	Else
		'This catches ALL
		strWhere = ""
	End If
	
	'Create the WHERE Statement based on the toggle button selected
	If Session("PendingViewButton") = "Pending" Then
		strWhere = strWhere & " AND [PMLoadStatus] <> 'Exported' "
	ElseIf Session("PendingViewButton") = "Exported" Then
		strWhere = strWhere & " AND [PMLoadStatus] = 'Exported' "
	Else
		'This catches ALL
		'strWhere = ""
	End If
	
	
	'Build the TOP Statement
	If Session("PageCombo") = "" Or IsNull(Session("PageCombo")) Then
		Session("PageCombo") = 50
	End If
	
strExtraHeaders = ""
strExtraFields = ""

'Set the table being queried based on the screen selection
If Session("PMMissingReport") = "" Or Isnull(Session("PMMissingReport")) Then
	strTable = "qryCAPSPMReconciliation"
Else
	If Session("PMMissingReport") = "Missing" Then
		strTable = "qryCAPSPMReconciliation"
	Else
		strTable = "qryCAPSPMReconciliationExisting"
		
		'Get there Where statement from the secondary Missing variable to determine what fields to compare
		If IsNull(Session("PMRecNameView")) Or Session("PMRecNameView") = "" THEN
		
			strWhere = strWhere & " AND (Address1 <> addr1)"
		Else
			If Session("PMRecNameView") = "Address1" THEN
				strWhere = strWhere & " AND (Address1 <> addr1)"
			ElseIf Session("PMRecNameView") = "Address2" THEN
				strWhere = strWhere & " AND (Address2 <> addr2)"
			ElseIf Session("PMRecNameView") = "Address3" THEN
				strWhere = strWhere & " AND (Address3 <> addr3)"
			ElseIf Session("PMRecNameView") = "Address4" THEN
				strWhere = strWhere & " AND (Address4 <> Addr4)"
			ElseIf Session("PMRecNameView") = "WorkPhone" THEN
				strWhere = strWhere & " AND (WorkPhone <> Work_Phone)"
			ElseIf Session("PMRecNameView") = "MobilePhone" THEN
				strWhere = strWhere & " AND (MobilePhone <> Mobile)"
			ElseIf Session("PMRecNameView") = "Expiry" THEN
				strWhere = strWhere & " AND (ExpiryCAPS <> ExpiryCMS)"
			ElseIf Session("PMRecNameView") = "Status" THEN
				strWhere = strWhere & " AND (Card_Status='C')"
			ElseIf Session("PMRecNameView") = "CreditLimit" THEN
				strWhere = strWhere & " AND (CreditLimit <> [monthly spend limit])"
			ElseIf Session("PMRecNameView") = "TransactionLimit" THEN
				strWhere = strWhere & " AND (TransactionLimit <> [transaction limit]) AND (CardTypeSub <> 'NAB Lodge')"
				Else
				strWhere = strWhere & " AND (Address1 <> addr1)"
			End If
			
			strExtraHeaders = "<th scope=""col"">CAPS " & Session("PMRecNameView") & "</th><th scope=""col"">CMS " & Session("PMRecNameView") & "</th>"
			'strWherePM = strWhere
			'response.write strSQL
		End If
		
	End If
End If

If strSort = "" OR ISNull(strSort) Then
	strSort = "Order By CardID Desc"
	'Session("strWherePM") = strWhere & strSort
End If

If strSortSearch <> ""  Then
	strSearch = strSortSearch
End If

If strSearch = "" OR ISNull(strSearch) Then
	strSQL = "SELECT " & strTOP & " * FROM " & strTable & " WITH(NOLOCK) WHERE CardID IS NOT NULL " & strWhere & strSort
Else
	Session("Search") = strSearch
	If UCase(Left(strSearch,2)) = "C:" Then
	
		strCardNoSearch = Right(strSearch,Len(strSearch)-2)
		
		strSQL = "SELECT " & strTOP & " * FROM " & strTable & " WITH(NOLOCK) WHERE [CardNumberShort] like '%" & strCardNoSearch & "%'" & strSort
		
	Else	
		If Session("PMRecNameView") = "Status" THEN
			strSQL = "SELECT " & strTOP & " * FROM " & strTable & " WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"  & strSort
		Else

		strSQL = "SELECT " & strTOP & " * FROM " & strTable & " WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort

		End If
	End If

End If

'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = strSearch
Else

End If


strWherePM = strWhere & strSort 

If Session("PMMissingReport") = "Existing" THEN
	If strSearch = "" OR ISNull(strSearch) THEN
		strAddAllButton = "<button type=""button"" id=""addall"" class=""btn btn-primary"" onclick=""addAllRecords()"" value=""" & strWherePM & """ title=""This button will update all the below listed cards to go through to CMS on the next updates export.""><i class=""fas fa-sync""></i> Update All to CMS</button>"
	Else
		strAddAllButton = ""
	End IF
End If


objRS.Open strSQL,objCon,3,1
    y = 0
	
	If IsEmpty(Request.QueryString("StartingRecord")) Then
		lngStartingRecord = 0
	Else
		lngStartingRecord = Request.QueryString("StartingRecord")
	End If

	'Write a message in the list if there are no Unactivated Cards
	If objRS.EOF Then
		If Session("PMRecNameView") = "" OR IsNull(Session("PMRecNameView")) THEN
				Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No CAPS Cards Missing from ProMaster records for " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
			Else
				Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No differences between CAPS and CMS for " & Session("PMRecNameView") & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
			End If
				'Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No CAPS Cards Missing from ProMaster records for " & strRecordMessage & "</TH>" & _
				'"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	Else
		objRS.Movelast
		objRS.Movefirst
		lngTotalRecords = objRS.Recordcount
		
		'Set the Page combos here so can be transferred to other pages together
		arrPagecombo(1) = "50"
		arrPagecombo(2) = "100"
		arrPagecombo(3) = "200"
		arrPagecombo(4) = "500"
		arrPagecombo(5) = "1000"
		arrPagecombo(6) = "All"
		
		'Build the Page Combo for TOP statement
		For x = 1 to 6
		
			If Session("PageCombo") = arrPagecombo(x) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End If
			strPageCombo = strPageCombo & "<option " & strSelected & " value=""" & arrPagecombo(x) & """>" & arrPagecombo(x) & "</option>"
		Next
		
		strPageCombo = "<SELECT ID=""PageCombo"" Name=""PageCombo"" onChange=""ChangePage();"">" & strPageCombo & "</select>"
		
		'If the PageCombo is not numeric (ALL or Null) then make it the total records for the recordset (which is set above)
		If NOT IsNumeric(Session("PageCombo")) Then Session("PageCombo") = lngTotalRecords
		
		Call LoadPendingButtons()
		Response.Write "<span>" & strAddAllButton & "</span></div>"

		
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
				  "<span class=""panel-subheader"">Displaying " & Session("PageCombo") & " of " & lngTotalRecords & " search results for: """ & strSearch & """</span><span class=""panel-subheader"" style=""float:right;"">Number of records per page: " & strPageCombo  & "</span></div></div>"
				  '"<span class=""panel-subheader"">Displaying " & Session("PageCombo") & " of " & lngTotalRecords & " CS To Diners records (" & lngStartingRecord & " to " & lngStartingRecord + clng(Session("PageCombo")) & ")</span><span class=""panel-subheader"" style=""float:right;"">Number of records per page: " & strPageCombo  & "</span></div></div>"
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""PMReconciliation.asp?SortSearch=" & strSortSearch & "&Sort=CardID&SortType=" & strOrderType & """> Card ID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""PMReconciliation.asp?SortSearch=" & strSortSearch & "&Sort=EIDNo&SortType=" & strOrderType & """> EID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""PMReconciliation.asp?SortSearch=" & strSortSearch & "&Sort=Surname&SortType=" & strOrderType & """> Name <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""PMReconciliation.asp?SortSearch=" & strSortSearch & "&Sort=CardType&SortType=" & strOrderType & """> Card Type <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Card No.</th>" & _
			"<th scope=""col""><a href=""PMReconciliation.asp?SortSearch=" & strSortSearch & "&Sort=Status&SortType=" & strOrderType & """> Status  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""PMReconciliation.asp?SortSearch=" & strSortSearch & "&Sort=DateIssued&SortType=" & strOrderType & """> Date Issued <i class=""fa fa-sort""></i></a></th>" & strExtraHeaders & _
			"<th scope=""col"" title=""Card Type""> PM Load Status </th>" & _
			"<th scope=""col"" style=""font-size:12px;""><a href=""PMReconciliation.asp?SortSearch=" & strSortSearch & "&Sort=PMLoadDate&SortType=" & strOrderType & """> PM Load Date  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" style=""font-size:12px;"">PM Account Ref</th>" & _
			"<th scope=""col"">Action</th>" & _
			"</tr></thead><tbody class=""text-left"">"
					
	End If
    
	x = 0
	
    Do until objRS.EOF 

		y = y + 1
		
		'Only write the first 50 records from the starting position
		If y <= lngStartingRecord + clng(Session("PageCombo")) AND y >= lngStartingRecord - clng(Session("PageCombo")) Then
		'If y <= lngStartingRecord + 50 AND y >= lngStartingRecord - 50 Then
		
			x = x + 1
			
			'Create the Action button depending on the Status
			If trim(objRS("Status")) = "Awaiting Export" Then
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='PMReconciliation.asp?Action=CancelCS&CSToDinersID=" & objrs("CSToDinersID") & "&CSEID=" & objRS("EmployeeID") & "&Status=Deleted'""; title=""Click to Remove from CS File""><i class=""fa fa-times""></i> Remove</button>"
				
			ElseIf objRS("Status") = "Deleted" Then
				strAction = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='PMReconciliation.asp?Action=CancelCS&CSToDinersID=" & objrs("CSToDinersID") & "&CSEID=" & objRS("EmployeeID") & "&Status=Awaiting Export'""; title=""Click to Add to CS File""><i class=""fa fa-plus""></i> Add</button>"

			Else
				strAction = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='PMReconciliation.asp?Action=AddToPM&CardID=" & objRS("CardID") & "&CardEID=" & objRS("EmployeeID") & "&CardType=" & objRS("CardType") & "" & objRS("CardTypeSub") & "'""; title=""Click to Add to ProMaster Export File for today""><i class=""fa fa-plus""></i> Add</button>"
			End If
			
			'Add the Cancel Action to all records
			strAction = strAction & " <a href=""PMReconciliation.asp?Action=Cancel&CardID=" & objRS("CardID") & """ title=""Click to CANCEL ProMaster Account. This will add the Card Account to the ProMaster File for CANCELLATION!""><span class=""badge badge-pill-xs badge-danger badge-xs"" style=""padding:10px;"">X</span></a>"
			
			'Create the Status list badge based on the status field
			If IsNull(objRS("Status")) Then
				strStatus = "Err"
			Else
				If objRS("Status") = "" OR objRS("Status") = "XS" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
				ElseIf objRS("Status") = "VX" Then
					strStatus = "<span class=""badge badge-pill badge-info"">Cancelled</span>"
				ElseIf objRS("Status") = "Pending update acc" Then
					strStatus = "<span class=""badge badge-pill badge-secondary"">Imported</span>"
				Else
					strStatus = objRS("Status")
				End If
			End If
			
			If IsNull(objRS("PMLoadStatus")) Then
				strPMStatus = ""
			Else
				strPMStatus = objRS("PMLoadStatus")
			End IF
			
			'Get the ProMaster Status
			If IsNull(objRS("PMLoadStatus")) Then
				strPMStatus = ""
			Else
				If objRS("PMLoadStatus") = "Pending" Then
					strPMStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:10px;"" title=""" & objRS("PMLoadStatus") & """>Pending</span>"
				ElseIf objRS("PMLoadStatus") = "Pending New Bill" Then
					strPMStatus = "<span class=""badge badge-pill badge-info"" style=""font-size:10px;"" title=""" & objRS("PMLoadStatus") & """>Pending New Bill</span>"
				ElseIf objRS("PMLoadStatus") = "Pending update" or objRS("PMLoadStatus") = "Pending Update" Then
					strPMStatus = "<span class=""badge badge-pill badge-secondary"" style=""font-size:10px;"" title=""" & objRS("PMLoadStatus") & """>Pending update acc</span>"
				ElseIf objRS("PMLoadStatus") = "New Card" Then
					strPMStatus = "<span class=""badge badge-pill badge-info"" style=""font-size:10px;"" title=""" & objRS("PMLoadStatus") & """>New Card</span>"
				ElseIf Left(strPMStatus,11)="Pending New" Then
					strPMStatus = "<span class=""badge badge-pill badge-info"" style=""font-size:10px;"" title=""" & objRS("PMLoadStatus") & """>New Card To Acct</span>"	
				ElseIf objRS("PMLoadStatus") = "Check Required" Then
					strPMStatus = "<span class=""badge badge-pill badge-danger"" style=""font-size:10px;"" title=""" & objRS("PMLoadStatus") & """>Check Required</span>"
				Else
					strPMStatus = objRS("PMLoadStatus")
				End If
			End If
			
			'Get the Date Issued and format
			If IsNull(objRS("PMLoadStatus")) Then
				strDateIssued = ""
			Else
				IF IsDate(objRS("DateIssued")) Then
				strDateIssued = FormatDateTime(objRS("DateIssued"),vbShortDate)
				Else
				strDateIssued = ""
				End If
			End If
			
			'Get the PM Load Date Issued and format
			If IsNull(objRS("PMLoadDate")) Then
				strPMLoadDate = ""
			Else
				strPMLoadDate = FormatDateTime(objRS("PMLoadDate"),vbShortDate)
			End If
			
			
			'Get there Where statement from the secondary Missing variable to determine what fields to compare
			If IsNull(Session("PMRecNameView")) Or Session("PMRecNameView") = "" THEN
			
			Else
				If Session("PMRecNameView") = "Address1" THEN
					strExtraFields = "<TD style=""font-size:12px;"">" & CheckString(objRS("Address1")) & "</TD><TD style=""font-size:12px;"">" & CheckString(objRS("Addr1")) & "</TD>"
				ElseIf Session("PMRecNameView") = "Address2" THEN
					strExtraFields = "<TD style=""font-size:12px;"">" & CheckString(objRS("Address2")) & "</TD><TD style=""font-size:12px;"">" & CheckString(objRS("Addr2")) & "</TD>"
				ElseIf Session("PMRecNameView") = "Address3" THEN
					strExtraFields = "<TD style=""font-size:12px;"">" & CheckString(objRS("Address3")) & "</TD><TD style=""font-size:12px;"">" & CheckString(objRS("Addr3")) & "</TD>"
				ElseIf Session("PMRecNameView") = "Address4" THEN
					strExtraFields = "<TD style=""font-size:12px;"">" & CheckString(objRS("Address4")) & "</TD><TD style=""font-size:12px;"">" & CheckString(objRS("Addr4")) & "</TD>"				
				ElseIf Session("PMRecNameView") = "WorkPhone" THEN
					strExtraFields = "<TD style=""font-size:12px;"">" & CheckString(objRS("WorkPhone")) & "</TD><TD style=""font-size:12px;"">" & CheckString(objRS("Work_Phone")) & "</TD>"
				ElseIf Session("PMRecNameView") = "MobilePhone" THEN
					strExtraFields = "<TD style=""font-size:12px;"">" & CheckString(objRS("MobilePhone")) & "</TD><TD style=""font-size:12px;"">" & CheckString(objRS("Mobile")) & "</TD>"
				ElseIf Session("PMRecNameView") = "Expiry" THEN
					strExtraFields = "<TD style=""font-size:12px;"">" & CheckString(objRS("ExpiryCAPS")) & "</TD><TD style=""font-size:12px;"">" & CheckString(objRS("ExpiryCMS")) & "</TD>"
				ElseIf Session("PMRecNameView") = "Status" THEN
					strExtraFields = "<TD style=""font-size:12px;"">" & CheckString(objRS("Status")) & "</TD><TD style=""font-size:12px;"">" & CheckString(objRS("Card_Status")) & "</TD>"
				ElseIf Session("PMRecNameView") = "CreditLimit" THEN
					strExtraFields = "<TD style=""font-size:12px;"">" & CheckNumber(objRS("CreditLimit")) & "</TD><TD style=""font-size:12px;"">" & CheckNumber(objRS("monthly spend limit")) & "</TD>"
				ElseIf Session("PMRecNameView") = "TransactionLimit" THEN
					strExtraFields = "<TD style=""font-size:12px;"">" & CheckNumber(objRS("TransactionLimit")) & "</TD><TD style=""font-size:12px;"">" & CheckNumber(objRS("transaction limit")) & "</TD>"				
				Else
					strExtraFields = "<TD style=""font-size:12px;"">" & objRS("Address1") & "</TD><TD style=""font-size:12px;"">" & objRS("Addr1") & "</TD>"
				End If
				strExtraHeaders = "<th scope=""col"">CAPS " & Session("PMRecNameView") & "</th><th scope=""col"">CMS " & Session("PMRecNameView") & "</th>"
				
				
			End If
		
			'Build the link to open the Modal with extra detail based on the view type Existing or Missing
			If Session("PMMissingReport") = "Missing" Then
				If IsNull(objRS("EmployeeID")) or Trim(objRS("EmployeeID")) = "" Then
					strEmployeeIDVar = "0"
				Else
					strEmployeeIDVar = objRS("EmployeeID")
				End If
				
				strLinkString = "<a data-toggle=""modal"" data-target=""#CSFromDinersMod"" HREF=""#"" onClick=""loadDocCMS(" & strEmployeeIDVar & "," & objRS("CardID") & ")"">" & objRS(0) & "</a>"
			Else
				strLinkString = "<a data-toggle=""modal"" data-target=""#CSFromDinersMod"" HREF=""#"" onClick=""loadDoc(" & objRS("CardID") & ")"">" & objRS(0) & "</a>"
			End If
			
			'GEt the Accunt Number
			If IsNull(objRS("AccountNumber")) or objRS("AccountNumber") = "" Then
				strAccountNumber = ""
			Else
				strAccountNumber = "Account Number: " & MaskCard(objRS("AccountNumber"))
			End If
			
			If IsNull(objRS("Surname")) THEN
				strEmployeeName = ""
			Else
				strEmployeeName = trim(objRS("FirstName")) & " " & trim(objRS("Surname"))
				If Len(strEmployeeName) > 14 THEN
					strEmployeeName = Left(strEmployeeName,14) &  "..."
				End If
			End If
			
			response.write "<TR><TD style=""font-size:12px;"">" & strLinkString & "</TD><TD style=""font-size:12px;"">" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""../CC/CardDetail.asp?CardNo=" & objRS("CardNumber") & """>" & strEmployeeName & "</a></TD><TD><a Target=""_self"" HREF=""../CC/CardDetail.asp?CardNo=" & objRS("CardNumber") & """>" & objRS("CardTypeSub") & "</a></TD>" & _
					"<TD style=""font-size:12px;"" title=""" & strAccountNumber & """ >" & MaskCard(objRS("CardNumber")) & "</TD><TD >" & strStatus & "</TD>" & _
					"<TD style=""font-size:12px;"">" & strDateIssued & "</TD>" & strExtraFields & "<TD style=""font-size:12px; background-color:#e6eeff;"">" & strPMStatus & "</TD><TD style=""font-size:12px; background-color:#e6eeff;"">" & strPMLoadDate & "</TD><TD style=""font-size:12px; background-color:#99b9ff;"">" & objRS("True_Account_Ref") & "</TD>" & _
					"<TD>" & strAction & "</TD></TR>"
			

		End If
		
		objRS.movenext
	Loop

	''''****Variables *******'''''
'lngStartingRecord = The Number of the record (in order of display) starting from. So displays from that number (row) to the Total per page number (Session("PageCombo"))
'Session("PageCombo") = The Number of records to display per page, as selected by the user in the top drop-down
'y = The total number of records in the complete recordset, derived from counting as each record is processed above
'lngPage = The current page selected, mainly for the active flag to display this on screen (the number shaded as currently selected)
'lngTotalRecords  = The recordcount from the recordset when first opening it (movelast then movefirst)

	If y > 0 Then
		Response.Write "<TR><TH colspan=""10"">Total</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;"">" & x & "</TH></TR>"
	End If

	'Create the number of pages
	If IsNumeric(y) Then
		If y > 1 Then
			
			'Y is the total number of records. Session("PageCombo") is the Total records displayed on screen. Dividing Total records by the number displayed per page gets the number of pages for the bottom of the screen
			y = y / clng(Session("PageCombo"))
			'y = y / 50
			
			'Determine the number of the page currently displayed
			If lngStartingRecord = 0 Then
				lngPage = 1
			Else
				If IsNumeric(lngStartingRecord) Then
					lngPage = (lngStartingRecord/clng(Session("PageCombo")))'+1
					'lngPage = (lngStartingRecord/50)'+1
				Else
					lngPage = 1
				End If
			End If

			For x = 0 to y
				
				'In looping through all pages, display only 20 pages, from the starting page (lngPage) to 20 more pages (lngPage + 20)
				If x > 19 + lngPage OR x < lngPage - 20 Then
					
					'Add the Elipsis (...) to the end of the page numbers if there is more than 20 pages
					If x = 20 + cint(fix(lngPage)) Then
					'If x = 21 + lngPage Then
						strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""PMReconciliation.asp?StartingRecord=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
					End If
					
					'Add the Elipsis (...) to the start of the page numbers if there is more than 20 pages and the current place is beyond the first page
					If x = 0 AND lngPage > 1 AND y > 20 Then
						strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""PMReconciliation.asp?StartingRecord=" & lngTotalRecords - (clng(Session("PageCombo"))*20) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
					End If
				Else
					intWritten = intWritten + 1
					If intWritten > 21 OR x < 1 then
					Else
					'If x + 20 > lngPage Then
					'Determine which page number is active (displayed as active)
					If clng(x) = clng(lngPage) Then
						strActive = "active"
					Else
						strActive = ""
					End If
				
					strPages2 = strPages2 & "<li class=""page-item " & strActive & """><a class=""page-link"" href=""PMReconciliation.asp?StartingRecord=" & (x * clng(Session("PageCombo"))) & """>" & x & "</a></li>"
					End If
				End If
			Next
			
		End If
	End If	
	
	'Write the End of the table and divs for the above list, as the pagination (below) is in it's own container
	Response.Write "</tbody></table></div>"

	'Write the Pagination objects for all pages based on the total records and the number records displayed on screen
	If y > 0 Then
		
		Response.Write "<div class=""container""><div class=""row""><div class=""col-12 text-center"">" & _
			"<nav aria-label=""Page navigation""><ul class=""pagination""><li class=""page-item"">" & _      
			"<a class=""page-link"" href=""PMReconciliation.asp?StartingRecord=0"" aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
			strPages2 & _
			"<li class=""page-item"">" & _
			"<a class=""page-link"" href=""PMReconciliation.asp?StartingRecord=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
			"</a></li></ul></nav></div></div></div>"
			
	End If

objRS.Close

End Sub

Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(4)
Dim arrButton2(2)
Dim strMissText
Dim arrChecks(10)
Dim x
Dim strDropDown

arrChecks(1) = "Address1"
arrChecks(2) = "Address2"
arrChecks(3) = "Address3"
arrChecks(4) = "Address4"
arrChecks(5) = "WorkPhone"
arrChecks(6) = "MobilePhone"
arrChecks(7) = "Expiry"
arrChecks(8) = "Status"
arrChecks(9) = "CreditLimit"
arrChecks(10) = "TransactionLimit"

If Session("ViewButton") = "DTC" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "DPC" Then
	arrButton(3) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

If Session("PMMissingReport") = "Missing" Then
	arrButton2(1) = "active"
Else
	'This catches ALL
	arrButton2(2) = "active"
End If

	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='PMReconciliation.asp?ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>"& _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='PMReconciliation.asp?ViewButton=DTC';""><i class=""fa fa-plane""></i> View DTC</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='PMReconciliation.asp?ViewButton=DPC';""><i class=""fa fa-dollar-sign""></i> View DPC</button>" & _
				"</div>"

'Create the Missing button with drop-downs				
'Make the Button text the selected Application Type if one is selected
If Session("PMRecNameView") = "" Or IsNull(Session("PMRecNameView")) Then
	strMissText = "Address1"
Else
	strMissText = Session("PMRecNameView")
End If

'Create the Application Type selection/filter
strDropDown = "<div class=""dropdown""><button class=""btn btn-outline-secondary dropdown-toggle"" type=""button"" id=""dropdownMenuButton"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">" & strMissText & "</button>" & _
				"<div class=""dropdown-menu"" aria-labelledby=""dropdownMenuButton"">"
		'"<div class=""dropdown-menu"" aria-labelledby=""dropdownMenuButton""><a class=""dropdown-item"" href=""Applications.asp?ApplicationTypeName="">All</a>"
		
		For x = 1 to 10
		
			strDropDown = strDropDown & "<a class=""dropdown-item"" href=""PMReconciliation.asp?PMMissingReport=Existing&PMRecNameView=" & arrChecks(x) & """>" & arrChecks(x) & "</a>"
		
		Next

		strDropDown = strDropDown & "</div></div>"
		
		
	Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-secondary " & arrButton2(1) & """ onClick=""self.location.href='PMReconciliation.asp?PMMissingReport=Missing';""><i class=""fa fa-clipboard""></i> View Missing</button>" & _
				"<button type=""button"" class=""btn btn-outline-secondary " & arrButton2(2) & """ onClick=""self.location.href='PMReconciliation.asp?PMMissingReport=Existing';""><i class=""fa fa-clipboard-check""></i> View Existing</button>" & _
				strDropDown & "</div>"
				
End Sub


Public Sub LoadPendingButtons()
'Load the the Status Selector buttons depending on what has been clicked
Dim strStatusButton


'Get the Status View button depending on what has been selected
	If Session("PendingViewButton") = "Pending" Then
		strStatusButton ="<button type=""button"" class=""btn btn-outline-success active float-right"" onClick=""self.location.href='PMReconciliation.asp?PendingViewButton=Exported';"" title=""Click to view Exported Cards Only (Sent to CMS but not in CMS)""><i class=""fa fa-check""></i> View Pending</button>"
	ElseIf Session("PendingViewButton") = "Exported" Then
		strStatusButton ="<button type=""button"" class=""btn btn-outline-danger active float-right"" onClick=""self.location.href='PMReconciliation.asp?PendingViewButton=All';"" title=""Click to View ALL Status Cards""><i class=""fa fa-times""></i> View Exported</button>"
	Else
		strStatusButton ="<button type=""button"" class=""btn btn-outline-secondary active float-right"" onClick=""self.location.href='PMReconciliation.asp?PendingViewButton=Pending';"" title=""Click to view Pending Cards Only (Cards to be sent to CMS today)""><i class=""fa fa-asterisk""></i> View All Statuses</button>"
	End If

	Response.Write 	"<div class=""d-flex justify-content-end""><div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				strStatusButton & _
				"</div>"
	
End Sub

Public Function GetPMLOadStatus(lngCardID)
'Procedure to load ProMaster details from the CAPS database (rather than a live link to ProMaster - which is in procedure LoadCMSDetails() )
Dim strSQL
Dim strCardNumber
Dim strNABCardNumber
Dim strAccountNumber
Dim strCardTypeSub
Dim strCMSExpiryDate
Dim strCheckCard

	strCheckCard = "N"
	
		'First Get the Card Number and Account Number
		strSQL = "SELECT * FROM tblCAPSCard WITH(NOLOCK) WHERE [CardID] = " & lngCardID & ""
		
		objRS1.Open strSQL,objCon
		
			If objRS1.EOF Then
				strCardNumber = ""
				strAccountNumber = ""
				strCardTypeSub = ""
				strNABCardNumber = ""
			Else
				strCardNumber = objRS1("CardNumberShort")
				strAccountNumber = objRS1("AccountNumber")
				strCardTypeSub = objRS1("CardTypeSub")
				strNABCardNumber = objRS1("CardNumber")
				
				'Remove the leading zeroes from the Account Number
				If Len(strAccountNumber) > 5 Then
					If Left(strAccountNumber,5) = "00000" Then
						strAccountNumber = Right(strAccountNumber,Len(strAccountNumber)-5)
					End If
				End IF
			End If
		
		objRS1.Close
		
		'If there is no Card Then exit.
		If strCardNumber = "" and strNABCardNumber = "" Then
			GetPMLOadStatus = ""
			Exit Function
		End If
		
		'If the card is a mastercard then check that the card exists rather than the Account, otherwise check the account.
		If strCardTypeSub="Mastercard" Then
			strSQL = "SELECT * FROM qryCAPSProMasterAccountsUserDecode WITH(NOLOCK)  WHERE [CardAccountNumber] = '" & strAccountNumber & "'"
			'strSQL = "SELECT * FROM qryCAPSProMasterAccountsUserDecode WITH(NOLOCK)  WHERE [CardAccountNumber] = '" & strCardNumber & "'"
		Else
			strSQL = "SELECT * FROM qryCAPSProMasterAccountsUserDecode WITH(NOLOCK)  WHERE [CardAccountNumber] = '" & strAccountNumber & "'"
		End If
		
		'response.write strSQL
		
		objRS1.Open strSQL,objCon
		
			If objRS1.EOF Then
			'If there is no card Account number in ProMaster for the Account Number in CAPS
				If strCardTypeSub = "Diners" THEN
					GetPMLOadStatus = "Pending"
					'GetPMLOadStatus = "Pending New Bill"
				End If
				
				If strCardTypeSub = "Mastercard" THEN
					GetPMLOadStatus = "Pending update acc"
					'GetPMLOadStatus = "Pending New Compani"
				End If
				
				If strCardTypeSub = "ANZ" THEN
					GetPMLOadStatus = "Pending"
				End If
				
				If Left(strCardTypeSub, 3) = "NAB" THEN
					GetPMLOadStatus = "Pending"
				End If
				
			Else
			'If the CAPS card Account number already exists in ProMaster
			
				If strCardTypeSub = "Diners" THEN
					GetPMLOadStatus = "Pending"
					'GetPMLOadStatus = "Pending update acc"
					
					'Set the local variable to Check to see if the Account also exists or not
					strCheckCard = "Y"
				End If
				
				If strCardTypeSub = "Mastercard" THEN
					'GetPMLOadStatus = "Pending update acc"
					GetPMLOadStatus = "Pending New Compani"
				End If
				
				If strCardTypeSub = "ANZ" THEN
					''If the CARD is already in PM then update otherwise new card on same billing (Pending)
					If strCardNumber = objRS1("CardAccountNumber") Then
						GetPMLOadStatus = "Pending update acc"
					Else
						GetPMLOadStatus = "Pending"
					End If
				
				If Left(strCardTypeSub, 3) = "NAB" THEN
					GetPMLOadStatus = "Pending update acc"
				End If
					
				End If
				
			End If
			
		objRS1.Close
	
	
		'Do a final check for Diners cards if the card exists to make sure the Card number exists
		If strCheckCard = "Y" Then
			strSQL = "SELECT * FROM qryCAPSProMasterAccountsUserDecode WITH(NOLOCK)  WHERE [CardAccountNumber] = '" & strCardNumber & "'"
			
			objRS1.Open strSQL,objCon
			
				If objRS1.EOF Then
			
				Else
				'If the CAPS card Account number already exists in ProMaster
				
					If strCardTypeSub = "Diners" THEN
						'GetPMLOadStatus = "Pending"
						GetPMLOadStatus = "Pending update acc"
					End If
				End If
			
			objRS1.Close
		End If
	
End Function




Public Sub AddToPM(lngCardID,lngEID,strCardType)
'Procedure to add the Card to the relevant ProMaster file for import in ProMaster
Dim intRecord
Dim strSQL
Dim strPMUpdateStatus

	'Get the relevant PMUpload Status
	strPMUpdateStatus = GetPMLOadStatus(lngCardID)
	
'Response.write "lngCardID="& lngCardID & " lngEID=" & lngEID

	If strPMUpdateStatus = "" THEN
		
		'Exit the procedure as there is no card found in tblCapsCard, which is unlikely
		Response.Write "<div class=""alert alert-danger"" role=""alert"">" & strCardType & " Card : " & lngCardID & " for Employee: " & lngEID & " has NOT been updated. No CARD found in tblCAPSCard. See system admin.</div>"
		
		Exit Sub
	End If
	
	
	strSQL = "UPDATE tblCAPSCard SET [PMLoadStatus] = '" & strPMUpdateStatus & "',PMLoadDate=GetDate() WHERE [CardID] = " & lngCardID & ""
		
	objCon.Execute strSQL

	Response.Write "<div class=""alert alert-success"" role=""alert"">" & strCardType & " Card : " & lngCardID & " for Employee: " & lngEID & " has been added to the ProMaster export file for today (if Afternoon Tasks have not been run).</div>"
		

'''''EXIT the Sub as the procedure below is not run due to it being for applications
Exit Sub
  	With objCmd

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

Public Sub CancelFromPM(lngCardID)
'Procedure to flag the Card Account for CANCEL in ProMaster file, after import in ProMaster

Dim strSQL

	If IsNull(lngCardID) OR lngCardID = "" THEN
	
		'Exit the procedure as there is no card found in tblCapsCard, which is unlikely
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Card ID: " & lngCardID & " has NOT been updated. No CARD found in tblCAPSCard. See system admin.</div>"
		
		Exit Sub
	End If
	
	
	strSQL = "UPDATE tblCAPSCard SET [PMLoadStatus] = 'Cancel',PMLoadDate=GetDate() WHERE [CardID] = " & lngCardID & ""
		
	objCon.Execute strSQL

	Response.Write "<div class=""alert alert-success"" role=""alert"">Card ID : " & lngCardID & " has been added to the ProMaster export file for today for CANCEL (if Afternoon Tasks have not been run).</div>"
		
End Sub

Public Sub UpdateBatchNumber(lngBatchNumber)
'Procedure to update the BatchNumber field in the System Parameters table with the next number
Dim strSQL

	'If the Batch Number is a number then update the System Parameter, otherwise post an error to the screen
	If IsNumeric(lngBatchNumber) Then
		lngBatchNumber = lngBatchNumber + 1
		
		strSQL = "UPDATE tblCAPSSystemParameters SET [ParameterValue] = '" & lngBatchNumber & "' WHERE [ParameterName] = 'CSFileNumberTo'"
		
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
'Description:	Loads all Batch Numbers to a list for selecting and searching/filtering


	objRS.Open "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = 'CSToDiners' AND [Deleted] = 'N' ORDER By [FileSeqNum] DESC",objCon
	'objRS.Open "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = 'CSToDiners' AND [Deleted] = 'N' ORDER By [FileSeqNum] DESC",objCon
  
	Response.write "<OPTION value=""0"">Select a Batch to View...</OPTION><OPTION value=""0"">CS Transactions To Be Sent Today</OPTION>"
	
		Do Until objRS.EOF 
			
			Response.write "<OPTION value=""" & objRS("FileLoadID") & """>" & objRS("FileSeqNum") & "</OPTION>"
			
			objRS.Movenext
			
		Loop
	
	objRS.Close
	
End Sub


Public Sub CardCMSDetailsUpdate(intCardID,strDefaultCompany,strCMSUser,strDefaultCostCentre,strPMLoadStatus,strReportGroup)
'Procedure to run a stored procedure which updates the CMS related detail for a Card
Dim intRecord

	With objCmd
	
		.CommandType = 4
		.CommandText = "spCAPSCardCMSDetailsUpdate"
		
		.Parameters.Append objCmd.CreateParameter("CardID", adInteger,adParamInput)
		.Parameters.Append objCmd.CreateParameter("CMSUser", adVarchar, adParamInput,50)
		.Parameters.Append objCmd.CreateParameter("DefaultCompany", adVarchar, adParamInput,20)
		.Parameters.Append objCmd.CreateParameter("DefaultCostCentre", adVarchar, adParamInput,20)
		.Parameters.Append objCmd.CreateParameter("PMLoadStatus", adVarchar, adParamInput,50)
		.Parameters.Append objCmd.CreateParameter("ReportGroup", adVarchar, adParamInput,50)
		.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("CardCMSIDOutput", adInteger, adParamOutput)
		
		.Parameters("CardID") = intCardID'Session("CardID") 'The Application currently viewed.
		.Parameters("CMSUser") = Trim(strCMSUser) 'The new name on card
		.Parameters("DefaultCompany") = Trim(strDefaultCompany) 'The new name on card
		.Parameters("DefaultCostCentre") = Trim(strDefaultCostCentre) 'The new name on card
		.Parameters("PMLoadStatus") = Trim(strPMLoadStatus) 'The new name on card --NO LONGER USED
		.Parameters("ReportGroup") = Trim(strReportGroup)
		.Parameters("UpdatedBy") = Session("UserID")
		
		.ActiveConnection = objCon
				
	End With
				
	objCmd.Execute        

	'Return the result of the Save Function.
	 intRecord = objCmd.Parameters.Item("CardCMSIDOutput")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Card details NOT updated for CARD ID: " & intCardID & " !</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">CMS Details updated for CARD ID: " & intCardID & "!</div>"
	End If
	
	
End Sub


Public Function CheckForNull(strValue,strField,intRow)

	If IsNull(strValue) Then 
		CheckForNull = "</BR> Error in field " & strField & " is Null at row " & intRow & " : "
	Else
		CheckForNull = ""
	End If

End Function

Set objRS1 = Nothing
Set objRS = Nothing
Set objCon = Nothing

%>
