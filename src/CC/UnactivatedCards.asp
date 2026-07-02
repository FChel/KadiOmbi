
<!-- #Include file=CAPSHeader.asp -->
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
Dim objCmd

Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL

Dim lngCardID
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

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	

    Session("CurrentPage") = "CC/ApplicationsEmployeeHF.asp"

	If IsNull(Session("CardID")) OR Session("CardID") = "" Then Session("CardID")= 0

	If isNull(Session("CardID")) Or Session("CardID") = "" Then 
		Session("CardID") = 0
	End If
	
	If Not IsEmpty(Request.QueryString("UserView")) Then
		Session("UserView") = Request.QueryString("UserView")
	End If

	If Not IsEmpty(Request.QueryString("CardID")) Then
		Session("CardID") = Request.QueryString("CardID")
	End If
	
	If Not IsEmpty(Request.QueryString("BatchID")) Then
		Session("BatchID") = Request.QueryString("BatchID")
	End If
	
	If Not IsEmpty(Request.QueryString("Action")) Then
		If Request.QueryString("Action") = "Cancel" Then
			Call CancelCard()
		End If
		
		If Request.QueryString("Action") = "Email" Then
			Call EmailCard()
		End If
		
		If Request.QueryString("Action") = "Remove" Then
			Call RemoveCard()
		End If
		
		'Call the procedure to add a new Batch
		If Request.QueryString("Action") = "CreateBatch" Then
			Call CreateBatch()
		End If
		
		'Call the procedure to Cancel the selected Batch
		If Request.QueryString("Action") = "CancelBatch" Then
			Call CancelBatch("Cards Cancelled")
		End If
		
		'Call the procedure to Update to Email sent the selected Batch
		If Request.QueryString("Action") = "EmailBatch" Then
			Call CancelBatch("Email Sent")
		End If
		
	End If

	If Not IsEmpty(Request.QueryString("ViewButton")) Then
		Session("ViewButton") = Request.QueryString("ViewButton")
	End If
  
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "SubmitApp" Then
			'Response.Write "CPID=" & Session("CarParkingID")
			'Session("CardID") = 0
			Call SubmitApplication()
		End If
	End If
	
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

function loadEmail(cb) {


//var id = cb.getAttribute('data-id');
var name = cb.getAttribute('data-EmployeeName');

	document.getElementById("EmployeeEmailName").innerHTML = 'Send Unactivated Card Email to - ' + name + '?';
	//document.getElementById("ModalDelete").style.display = "block";
	document.getElementById("EmployeeEmailID").value = id;

}

function CreateBatch() {


//var id = cb.getAttribute('data-id');
//var name = cb.getAttribute('data-EmployeeName');

	//document.getElementById("BatchName").innerHTML = 'Create New Batch - ' + name + '?';
	//document.getElementById("ModalDelete").style.display = "block";
	//document.getElementById("EmployeeEmailID").value = id;
	
	 var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("BatchName").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetUnactivatedBatch.asp", true);
  xhttp.send();
  
}

function CreateBatch1(cb) {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	//var e = document.getElementById("CardTypeSelect");
	//var result = e.options[e.selectedIndex].value;
	
	//Removed as no longer in screen
	//var SelWarn = document.getElementById("SelWarning");
	//var result1 = SelWarn.options[SelWarn.selectedIndex].value;
	
	var SelEm = document.getElementById("SelEmail");
	var result2 = SelEm.options[SelEm.selectedIndex].value;
	
	//Removed as no longer in screen
	//var Days = document.getElementById("DaysEmail").value;
	
	//document.getElementById('CardType').value=result;
	
	self.location='UnactivatedCards.asp?Action=CreateBatch&Emails='+result2
	//self.location='UnactivatedCards.asp?Action=CreateBatch&Warning=' +result1+'&Emails='+result2+'Days='+Days
	
}

function OpenExcelReport() {
	var strExcel = document.getElementById('WhereClause').value;
	
	window.open('ExcelExport.asp?tbl=qryCAPSUnactivatedCardsBatch&W=' + strExcel + '')
	//window.open('../CC/ExcelExport.asp?tbl=qryCAPSTrainingReportExport&W=' + strExcel + '&Top=100')
	//window.open('../CC/ExcelExport.asp?tbl=qryCAPSTrainingReport&Top=100')

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

$('#CardTypeSelect').change(function(){
    alert($(this).val());
})

jQuery(document).ready(function($) {
    $(".clickable-row").click(function() {
        window.location = $(this).data("href");
    });
});

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

<!-- Modal -->
<div class="modal fade" id="EmailModal" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle"> Email Cardholder</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
       
      </div>
	  <div id="EmployeeEmailName">
	  
	  </div>
      <div class="modal-footer">
        <div class="row">
			<div class="col-md-12 mb-3" style="text-align:right;">
				<input type="hidden" id="EmployeeEmailID"></input>
				<!--<button class="btn btn-primary btn-sm" onClick="self.location='Training.asp?Action=UpdateTraining'"><i class="fa fa-check"></i> Save</button>-->
				<button class="btn btn-primary btn-sm" onClick="UpdateTraining(this)"><i class="fa fa-check"></i> Save</button>
				<button type="button" class="btn btn-secondary btn-sm" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
			</div>
		</div>
      </div>
	 
    </div>
  </div>
</div>


<!-- Modal -->
<div class="modal fade" id="BatchModal" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle"> Create New Batch</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
       
      </div>
	  <div id="BatchName">
	  
	  </div>
      <div class="modal-footer">
        <div class="row">
			<div class="col-md-12 mb-3" style="text-align:right;">
				<input type="hidden" id="BatchNameID"></input>
				<!--<button class="btn btn-primary btn-sm" onClick="self.location='UnactivatedCards.asp?Action=CreateBatch'"><i class="fa fa-check"></i> Save</button>-->
				<button class="btn btn-primary btn-sm" onClick="CreateBatch1(this)"><i class="fa fa-check"></i> Save</button>
				<button type="button" class="btn btn-secondary btn-sm" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
			</div>
		</div>
      </div>
	 
    </div>
  </div>
</div>

	   
<!-- End the first part of the Header Container -->
<div id='tbl-container'>
  <form action="UnactivatedCards.asp?Action=Search" method="POST" id="frm" name="frm">
	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-6">
				<h4 class="text-left">Unactivated Cards <%="As at: " & FormatDateTime(now(),2)%></h4>
				
			</div>
			<div class="col-md-6" style="text-align:right;">
				
				<%Call LoadAdminButtons()%>
				<button type="button" class="btn btn-primary" data-toggle="modal" data-target="#BatchModal" HREF="#" onClick="CreateBatch();"><i class="fa fa-beer"></i> Create New Batch</button>
				<!--<button type="button" class="btn btn-primary" data-toggle="modal" data-target="#BatchModal" HREF="#" onClick="CreateBatch();" onClick='self.location="ApplicationsSubmit.asp"'><i class="fa fa-beer"></i> Create New Batch</button> -->
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
	  
	  <div class="row py-2">
			<div class="col-md-3">
				<div class="panel panel-shadow mb-3">

					<div class="panel-header">
						<h4>Batches</h4>
							<%Call LoadBatch() %>
					  </div>
				</div>
				
				
			</div>
			

            <div class="col-md-9">
				 <section class="table py-2">
					<div class="container">
					 
							 <%
					
							DisplayTableDetails()
					
							%>	
							</tbody>
						  </table>
						</div>
					  </div>
					</div>
				  </section>

			</div>
        </div>
		
	 
</div>


<!--</DIV>-->
</form>
</div>

</main>
	
<!-- #Include file=CAPSFooter.asp -->

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
Dim strWhere
Dim lngStartingRecord
Dim lngTotalRecords
Dim strBatches
Dim strCountSQL
Dim strBatchName
Dim strName
Dim strActiveFlag
Dim strRowStyle
Dim strRecordBatchDetails
Dim strBatchStatusText
Dim strUnactivatedStatusSummary
Dim lngActivationFlagSummaryY
Dim lngActivationFlagSummaryN
Dim lngStatusSummaryActive
Dim lngStatusSummaryCancelled
Dim strBatchDate
Dim strBatchIDSearch

	strSearch = Request.Form("SearchInput")

	If Not IsNull(strSearch) Then
		strSearch = Trim(strSearch)
	End If
	
	'If Not IsEmpty(Request.QueryString("Action")) Then
	'	If Request.QueryString("Action") = "Search" Then
	'		strSearch = Trim(Request.Form("SearchInput"))
	'	End If
	'End If
	
	
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
		strSort = " ORDER By [DateIssued]"
	Else		
		strSort = " ORDER BY " & Request.QueryString("Sort") & " " & strOrderType
	End If
	
	If Session("ViewButton") = "Emailed" Then
		strWhere = " AND [ProcessStatus] = 'Email Unactivated' "
	ElseIf Session("ViewButton") = "Removed" Then
		strWhere = " AND [ProcessStatus] = 'Removed Unactivated' "
	ElseIf Session("ViewButton") = "AddedToCS" Then
		strWhere = " AND [ProcessStatus] = 'Added To CS' "
	ElseIf Session("ViewButton") = "InABatch" Then
		strWhere = " AND [UnactivatedStatus] IS NOT NULL "
	Else
		'This catches ALL
		strWhere = ""
	End If
	
	'Add the Unactivated Date difference (Cards Issued greater than 45 days ago) as a default
	strWhere = strWhere & " AND  (DATEDIFF(day, DateIssued, GETDATE()) > 45)"
	
If strSearch = "" OR ISNull(strSearch) Then
	'If Session("UserView") = "All" Then
		'strSQL = "SELECT TOP 100 * FROM qryCAPSCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND Status = '00' AND ([CardTypeSub] = 'Diners' OR [CardTypeSub] = 'Mastercard')" & strWhere & strSort
	'Else
	'	strSQL = "SELECT top 100 * FROM qryCAPSCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND EmployeeID = '" & Session("EmployeeID") & "'"
	'End If
	
	strSQL = "SELECT TOP 100 * FROM qryCAPSUnactivatedCardsBatch WITH(NOLOCK) WHERE [ActivationFlag] = 'N' " & strWhere & strSort
	'strSQL = "SELECT TOP 100 * FROM qryCAPSUnactivatedCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' " & strWhere & strSort
	
Else
	'If Session("UserView") = "All" Then
		'strSQL = "SELECT TOP 100 * FROM qryCAPSCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND Status = '00' AND ([CardTypeSub] = 'Diners' OR [CardTypeSub] = 'Mastercard') AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
	'Else
	'	strSQL = "SELECT top 100 * FROM qryCAPSCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
	'End If
	
	strSQL = "SELECT TOP 100 * FROM qryCAPSUnactivatedCardsBatch WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
	'strSQL = "SELECT TOP 100 * FROM qryCAPSUnactivatedCards WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
	strWhere = "AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
End If

	'Overwrite the Other searches if there is a Batch selected
	'Determine whether there has been a Batch Selected to display
	If IsEmpty(Session("BatchID")) Or Session("BatchID") = "" Then
		strBatches = ""
		strBatchIDSearch = ""
		strCountSQL = " FROM qryCAPSUnactivatedCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' " & strWhere
		
		strBatchName = ""
	Else		
		strBatches = " BatchID= " & Session("BatchID") & " " & strWhere
		
		'Get the Where statement for the Batch ID only as the above has the DateIssued Date Diff for cards over 45 days which is for the new batches only
		strBatchIDSearch = " BatchID= " & Session("BatchID") & ""
		
		strSQL = "SELECT * FROM qryCAPSUnactivatedCardsBatch WITH(NOLOCK) WHERE " & strBatches & strSort
		strCountSQL = "FROM qryCAPSUnactivatedCardsBatch WITH(NOLOCK) WHERE " & strBatchIDSearch'strBatches
		strBatchName = " BatchID = " & Session("BatchID") 
	End If
	
	
'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = strSearch
Else
	'strRecordMessage = "for " & Session("UserName") 
'	strRecordMessage = "No Batch Selected"
End If
'response.write strCountSQL

	'select COUNT(*) as CountEIDs,UnactivatedStatus,ActivationFlag,[Status] FROM qryCAPSUnactivatedCardsBatch WITH(NOLOCK) WHERE BatchID= 30 GROUP BY UnactivatedStatus,ActivationFlag,[Status]

	'If there is no batch selected then do not display summary details
	If IsEmpty(Session("BatchID")) Or Session("BatchID") = "" Then
	
	Else
	
	lngActivationFlagSummaryY = 0
	lngActivationFlagSummaryN = 0
	lngStatusSummaryActive = 0
	lngStatusSummaryCancelled = 0
		
	objRS.Open "SELECT COUNT(*) as CountEIDs,UnactivatedStatus,ActivationFlag,[Status],CONVERT(VARCHAR, [DateCreated], 106) as BatchDate " & strCountSQL & " GROUP BY UnactivatedStatus,ActivationFlag,[Status],CONVERT(VARCHAR, [DateCreated], 106)",objCon,3,1
	'Response.Write "SELECT COUNT(*) as CountEIDs,UnactivatedStatus,ActivationFlag,[Status],CONVERT(VARCHAR, [DateCreated], 106) as BatchDate " & strCountSQL & " GROUP BY UnactivatedStatus,ActivationFlag,[Status],CONVERT(VARCHAR, [DateCreated], 106)"
	
		Do Until objRS.EOF
		
			lngTotalRecords = lngTotalRecords + objRS("CountEIDs")
			
			'Count the cards by Activation Flag
			If Not IsNull(objRS("ActivationFlag")) then
				If objRS("ActivationFlag") = "Y" Then
					lngActivationFlagSummaryY = lngActivationFlagSummaryY + objRS("CountEIDs")
				Else
					lngActivationFlagSummaryN = lngActivationFlagSummaryN + objRS("CountEIDs")
				End If
			End If
			
			'Count the cards by Card Status
			If Not IsNull(objRS("Status")) Then
				If objRS("Status") = "00" Then
					lngStatusSummaryActive = lngStatusSummaryActive + objRS("CountEIDs")
				Else
					lngStatusSummaryCancelled = lngStatusSummaryCancelled + objRS("CountEIDs")
				End If
			End If
			
			strUnactivatedStatusSummary = objRS("UnactivatedStatus")
			strBatchDate = objRS("BatchDate")
			
		objRS.Movenext
		Loop
		
	objRS.Close

	'If IsNumeric(lngTotalRecords) Then lngTotalRecords = FormatNumber(lngTotalRecords,0)
	'If IsNumeric(lngActivationFlagSummaryY) Then lngActivationFlagSummaryY = FormatNumber(lngActivationFlagSummaryY,0)
	'If IsNumeric(lngActivationFlagSummaryN) Then lngActivationFlagSummaryN = FormatNumber(lngActivationFlagSummaryN,0)
	'If IsNumeric(lngStatusSummaryActive) Then lngStatusSummaryActive = FormatNumber(lngStatusSummaryActive,0)
	'If IsNumeric(lngStatusSummaryCancelled) Then lngStatusSummaryCancelled = FormatNumber(lngStatusSummaryCancelled,0)
	
	'END Batch Check for summary details
	End If
	
'	objRS.Open "SELECT Count(*) AS [CountEIDs] " & strCountSQL,objCon,3,1
	'objRS.Open "SELECT Count(*) AS [CountEIDs] FROM qryCAPSUnactivatedCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' " & strWhere & " " & strBatches,objCon,3,1
	'objRS.Open "Select Count(*) AS [CountEIDs] FROM qryCAPSCards WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND Status = '00' AND ([CardTypeSub] = 'Diners' OR [CardTypeSub] = 'Mastercard') " & strWhere,objCon,3,1
	
'		If not objRS.EOF Then
		
'			lngTotalRecords = objRS("CountEIDs")
'			If IsNumeric(lngTotalRecords) Then lngTotalRecords = FormatNumber(lngTotalRecords,0)
'		End If
		
'	objRS.Close

	'response.write strsql
objRS.Open strSQL,objCon,3,1

    y = 0
	
	If IsEmpty(Request.QueryString("StartingRecord")) Then
		lngStartingRecord = 0
	Else
		lngStartingRecord = Request.QueryString("StartingRecord")
	End If
	
	'Write a message in the list if there are no Unactivated Cards
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Unactivated Cards for " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
		
		'If there is no batch selected then do not display summary details
		If IsEmpty(Session("BatchID")) Or Session("BatchID") = "" Then
			Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Search all Batches for <span style=""color:red; font-weight:bold;"">" & strSearch & "</span></div></div>"
		Else
			Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Search Batch <span style=""color:red; font-weight:bold;"">" & Session("BatchID") & "</span> for <span style=""color:red; font-weight:bold;"">" & strSearch & "</span></div></div>"
		End If
	Else
		objRS.Movelast
		objRS.Movefirst
		'lngTotalRecords = objRS.Recordcount
		
		If NOT objRS.EOF Then
			If IsNull(objRS("UnactivatedStatus")) or objRS("UnactivatedStatus") = "" Then
			Else
			
				If objRS("UnactivatedStatus") = "Generated" Then
					strAction = "<button type=""button"" class=""btn btn-info btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=EmailBatch&BatchID=" & objrs("BatchID") & "'"";><i class=""fa fa-envelope""></i> Email Batch</button>"
				End If
				
				If objRS("UnactivatedStatus") = "Email Sent" Then
					strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=CancelBatch&BatchID=" & objrs("BatchID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel Batch</button>"
				End If
			End If
		End If
		
		'If there is no batch selected then do not display summary details
		If IsEmpty(Session("BatchID")) Or Session("BatchID") = "" Then
			Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Search all Batches for <span style=""color:red; font-weight:bold;"">" & strSearch & "</span></div></div>"
		Else
		
		'Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
        '          "<span class=""panel-subheader"">Displaying all of " & lngTotalRecords & " unactivated cards in " & strBatches & "</span> " & strAction & "</br>" & _ 
		'		  "<span class=""panel-subheader"">Active Cards: " & lngActivationFlagSummaryY & " InActive Cards: " & lngActivationFlagSummaryN & "</br>" & _
		'		  "Cancelled Cards: " & lngStatusSummaryCancelled & " Active Cards Cards: " & lngStatusSummaryActive & "</span></div></div>"
		
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Displaying all of " & lngTotalRecords & " unactivated cards in " & strBatches & "</span> " & strAction & "</br>" & _ 
				  "</span></div></div>"
		
		Response.Write "<div class=""panel panel-light mb-1""><div class=""panel-header"" style=""padding-top:5px;"">" & _
				  "<span class=""panel-subheader"" style=""padding-top:5px;""><table><tr><th colspan=""5"" style=""border:0px;""><span style=""color:red; font-weight:bold;"">" & strBatchDate & "</span> Batch Summary</th></tr>" & _
				  "<tr><th style=""padding:2px;"" Title=""Cards which have been Activated with Diners (Activation Flag in table below)"">Actived Cards:</th><td style=""padding:2px; text-align:right;"">" & lngActivationFlagSummaryY & "</td><td style=""width:20px; border-right:1px solid black;""></td><th style=""padding:2px;"" title=""Cards which have been cancelled (Status in table below)"">Cancelled Cards:</th><td style=""padding:2px; text-align:right;"">" & lngStatusSummaryCancelled & "</td></tr>" & _
				  "<tr><th style=""padding:2px;"" Title=""Cards which have not yet been Activated with Diners (Activation Flag in table below)"">Not Actived Cards:</th><td style=""padding:2px; text-align:right;"">" & lngActivationFlagSummaryN & "</td><td style=""width:20px; border-right:1px solid black;""></td><th style=""padding:2px;"" title=""Cards which are current (not cancelled) Status=00 (Status in table below)"">Active Cards:</th><td style=""padding:2px; text-align:right;"">" & lngStatusSummaryActive & "</td></tr>" & _
				  "<tr><th style=""padding:2px; border-top:1px solid black;"">Activation Flag Total:</th><td style=""padding:2px; text-align:right; border-top:1px solid black; font-weight:bold;"">" & lngActivationFlagSummaryY + lngActivationFlagSummaryN & "</td><td style=""width:20px; border-right:1px solid black; border-top:1px solid black;""></td><th style=""padding:2px;  border-top:1px solid black;"">Card Status Total:</th><td style=""padding:2px; text-align:right; border-top:1px solid black; font-weight:bold;"">" & lngStatusSummaryCancelled + lngStatusSummaryActive & "</td></tr></table></span></div></div>"
				  
				  
		'Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
        '          "<span class=""panel-subheader"">Displaying all of " & lngTotalRecords & " unactivated cards in " & strBatches & "</span> " & strAction & "</br>" & _ 
		'		  "<span class=""panel-subheader"" style=""padding-top:10px; padding-left:60px;""><table><tr><th colspan=""5""><span style=""color:red; font-weight:bold;"">" & strBatchDate & "</span> Batch Summary</th></tr>" & _
		'		  "<tr><th style=""padding:2px;"">Active Cards:</th><td style=""padding:2px; text-align:right;"">" & lngActivationFlagSummaryY & "</td><td style=""width:20px;""></td><th style=""padding:2px;"">Cancelled Cards:</th><td style=""padding:2px; text-align:right;"">" & lngStatusSummaryCancelled & "</td></tr>" & _
		'		  "<tr><th style=""padding:2px;"">InActive Cards:</th><td style=""padding:2px; text-align:right;"">" & lngActivationFlagSummaryN & "</td><td style=""width:20px;""></td><th style=""padding:2px;"">Active Cards:</th><td style=""padding:2px; text-align:right;"">" & lngStatusSummaryActive & "</td></tr>" & _
		'		  "<tr><th style=""padding:2px;"">Activation Flag Total:</th><td style=""padding:2px; text-align:right;"">" & lngActivationFlagSummaryY + lngActivationFlagSummaryN & "</td><td style=""width:20px;""></td><th style=""padding:2px;"">Card Status Total:</th><td style=""padding:2px; text-align:right;"">" & lngStatusSummaryCancelled + lngStatusSummaryActive & "</td></tr></table></span></div></div>"
				  
				  
				'"<span class=""panel-subheader"">Displaying all of " & lngTotalRecords & " unactivated cards in " & strBatches & "</span> " & strAction & "</div></div>"
		
		'End batch check		
		End If
		
		strAction = ""
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col"" Style=""font-size:14px;""><a href=""UnactivatedCards.asp?Sort=ApplicationID&SortType=" & strOrderType & """> Card ID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" Style=""font-size:14px;""><a href=""UnactivatedCards.asp?Sort=EmployeeID&SortType=" & strOrderType & """> EID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" Style=""font-size:14px;""><a href=""UnactivatedCards.asp?Sort=Surname&SortType=" & strOrderType & """> Name <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" Style=""font-size:14px;""><a href=""UnactivatedCards.asp?Sort=CardType&SortType=" & strOrderType & """> Card Type <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" Style=""font-size:14px;"">Card No.</th>" & _
			"<th scope=""col"" Style=""font-size:14px;""><a href=""UnactivatedCards.asp?Sort=Status&SortType=" & strOrderType & """> Status  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" Style=""font-size:14px;""><a href=""UnactivatedCards.asp?Sort=DateIssued&SortType=" & strOrderType & """> Issued <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" title=""Days inactive since issued (up to today)"" Style=""font-size:14px;""> Days </th>" & _
			"<th scope=""col"" Style=""font-size:14px;""><a href=""UnactivatedCards.asp?Sort=ActivationFlag&SortType=" & strOrderType & """> Activation Flag <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" Style=""font-size:14px;""><a href=""UnactivatedCards.asp?Sort=ProcessStatus&SortType=" & strOrderType & """> Process Status  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" style=""text-align:center; font-size:14px;"">Action</th>" & _
			"</tr></thead><tbody class=""text-left"">"
			
			
			'Removed from the last column after Process Status
			'"<th scope=""col"" style=""text-align:center;"">Action</th>" & _
	End If
	
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Unactivated Cards " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	End If
    	
    Do until objRS.EOF 

		y = y + 1
		
		'Only write the first 50 records from the starting position
		'If y <= lngStartingRecord + 50 AND y >= lngStartingRecord - 50 Then
		
			x = x + 1
			
			'Create the actions based on the Process Status of the card
			'Select Case objRS("ProcessStatus")
			Select Case objRS("UnactivatedStatus")
			
			Case  "Removed Unactivated"
				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=UnRemove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-list""></i> Re-List</button>"
			Case "Added to CS"

				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportCS.asp?CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
				
			Case "Email Unactivated"
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
			Case "Email Sent"
				strAction = "<button type=""button"" class=""btn btn-outline-info btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=Remove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Remove</button>"
			Case Else
			'	strAction = "<button type=""button"" class=""btn btn-outline-danger btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
			'	strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#EmailModal"" data-EmployeeName=""" & objRS("FirstName") & " " & objRS("Surname") & """ data-EmployeeID=""" & objrs("EmployeeID") & """ onClick=""loadEmail(this);""><i class=""fa fa-envelope""></i> Email</button>"
			'	strAction = strAction & "<button type=""button"" class=""btn btn-outline-info btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=Remove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-times""></i> Remove</button>"

				'strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='UnactivatedCards.asp?Action=Email&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-mail""></i> Email</button>"
				'data-toggle="modal" data-target="#EmailModal"
				
			End Select
			
			'Create the Status list badge based on the status field
			If IsNull(objRS("Status")) Then
				strStatus = ""
			Else
				If objRS("Status") = "00" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
				ElseIf objRS("Status") = "01" OR objRS("Status") = "02" Then
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				Else
					strStatus = ""
				End If
			End If
			
			If IsNull(objRS("DateIssued")) Then
				dteDateSubmitted = ""
			Else
				dteDateSubmitted = FormatDateTime(objRS("DateIssued"),vbShortDate)
			End If
			
			If IsNull(objRS("DateIssued")) Then
				dteDateReviewed = ""
			Else
				dteDateReviewed = DateDiff("d",objRS("DateIssued"),now())
				If dteDateReviewed > 80 Then
					strDaysColour = "Style=""color:red; font-weight:bold;"""
				ElseIf dteDateReviewed > 45 Then
					strDaysColour = "Style=""color:orange; font-weight:bold;"""
				Else
					strDaysColour = "Style=""color:black"""
				End If
			End If
			
			'Format the Card number so it is masked depending on the card type
			If IsNull(objRS("CardNumber")) Then
				strCardNo = ""
			Else
				strCardNo = objRS("CardNumber")
				If mid(strCardNo,5,1)=0 Then 
					strCardNo = mid(strCardNo,6,2) & "****" & right(strCardNo,4)
				Else
					strCardNo = mid(strCardNo,4,2) & "****" & right(strCardNo,4)
				End If
				'If len(strCardNo)>8 Then strCardNo = mid(strCardNo,6,2) & "****" & right(strCardNo,4)
			End If
			
			'If IsNull(objRS("ProcessStatus")) or objRS("ProcessStatus") = "" Then
			'	strProcessStatus = ""
			'Else
			'	If objRS("ProcessStatus") = "Email Unactivated" Then
			'		strProcessStatus = "Emailed"
			'	Elseif objRS("ProcessStatus") = "Removed Unactivated" Then
			'		strProcessStatus = "Removed"
			'	Elseif objRS("ProcessStatus") = "Added to CS" Then
			'		strProcessStatus = "Added To CS"
			'	End If
			'End If
			
			'Changed from the above to the Status of the card if added to the Unactivated table
			If IsNull(objRS("UnactivatedStatus")) or objRS("UnactivatedStatus") = "" Then
				strProcessStatus = ""
				strBatchStatusText = "Future"
			Else
				If objRS("UnactivatedStatus") = "Email Unactivated" Then
					strProcessStatus = "Emailed"
					strBatchStatusText = "Future"
				Elseif objRS("UnactivatedStatus") = "Removed Unactivated" Then
					strProcessStatus = "Removed"
					strBatchStatusText = "Past"
				Elseif objRS("UnactivatedStatus") = "Added to CS" Then
					strProcessStatus = "Added To CS"
					strBatchStatusText = "Future"
				Else
					strProcessStatus = objRS("UnactivatedStatus")
					If strProcessStatus = "Cards Cancelled" Then
						strBatchStatusText = "Past"
					Else
						strBatchStatusText = "Future"
					End If
				End If
			End If
			
			dteWarningDate = "title=" & objRS("Warning") & ""
			
			'Get the Full name and reduce to 15 characters if more than 15
			If IsNull(objRS("Surname")) Or objRS("Surname") = "" Then
				strName = ""
			Else
				strName = objRS("FirstName") & " " & objRS("Surname")
				If Len(strName) > 15 Then
					strName = Left(strName,15)
				End If
			End If
			
			'Set the Batch Details for the record
			strRecordBatchDetails = " - | BatchID: " & objRS("BatchID") & " | DateCreated: " & objRS("DateCreated") & " | Status: " & objRS("UnactivatedStatus") & ""
			
			'Get the ActivationFlag details for each record
			If IsNull(objRS("ActivationFlag")) or objRS("ActivationFlag")="" Then
				strActiveFlag = ""
				strRowStyle = ""
			Else
				strActiveFlag = objRS("ActivationFlag")
				
				If strActiveFlag = "Y" Then
					strActiveFlag = "<span class=""badge badge-pill-xs badge-success"">Activated</span>"
					
					'Determine the grammar depending on whether the batch has been sent yet or not
					If strBatchStatusText = "Future" Then
						strRowStyle = "background-color:#e6ffe6;"" title=""Card has been activated so it WILL NOT be cancelled in this batch " & strRecordBatchDetails & " "
					Else
						strRowStyle = "background-color:#e6ffe6;"" title=""Card has been activated so it WAS NOT cancelled in this batch " & strRecordBatchDetails & " "
					End If
				ElseIf strActiveFlag = "N" Then
					strActiveFlag = "<span class=""badge badge-pill-sm badge-danger"">Not Activated</span>"
					
					'Determine the grammar depending on whether the batch has been sent yet or not
					If strBatchStatusText = "Future" Then
						strRowStyle = " "" title=""Card has NOT been activated so it WILL be cancelled in this batch " & strRecordBatchDetails & " "
					Else
						strRowStyle = " "" title=""Card was NOT activated so it WAS cancelled in this batch " & strRecordBatchDetails & " "
					End If
				Else
					strActiveFlag = ""
					strRowStyle = ""
				End If
				
			End If
			
			response.write "<TR Style=""font-size:13px; " & strRowStyle & """><TD Style=""font-size:13px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD Style=""font-size:13px;"">" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & strName & "</a></TD><TD Style=""font-size:14px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
					"<TD Style=""font-size:14px;"">" & strCardNo & "</TD><TD >" & strStatus & "</TD><TD Style=""font-size:14px;"">" & dteDateSubmitted & "</TD>" & _
					"<TD " & strDaysColour & " Style=""font-size:14px;"">" & dteDateReviewed & "</TD><TD>" & strActiveFlag & "</TD><TD " & dteWarningDate & " Style=""font-size:14px;"">" & strProcessStatus & "</TD>" & _
					"<TD>" & strAction & "</TD></TR>"
					
			'response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & strAction & "</a></TD>" & _
			'		"<TD><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?CardID=" & objRS(0) & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?CardID=" & objRS(0) & """>" & objRS(2) & "</a></TD>" & _
			'		"<TD style=""text-align:center;"">" & objRS(3) & "</TD><TD style=""text-align:center;"">" & objRS(4) & "</TD>" & _
			'		"<TD style=""text-align:center;"">" & objRS(5) & "</TD><TD style=""text-align:center;"">" & objRS(6) & "</TD>" & _
			'		"<TD style=""text-align:center;"">" & objRS(7) & "</TD><TD style=""text-align:center;"">" & objRS(10) & "</TD>" & _
			'		"<TD style=""text-align:center;"">" & strStatus & "</TD><TD style=""text-align:center;"">" & objRS(14) & "</TD><TD style=""text-align:center;"">" & objRS(15) & "</TD></TR>"
				
				'y = y + 1
		'End If
		
		objRS.movenext
	Loop
	
	
	'Add the row totals in the bottom of the table and a hidden input field with the WHERE Clause for the Excel Export button (which loads before this, hence the hidden input for javascript!)
	If x > 0 Then
		strWhere = strSQL
		'Make sure there is WHERE in the where statement
		If Instr(1,strWhere,"WHERE")>0 Then
			'strWhere = Replace(strWhere,"AND","WHERE")
			'If Trim(Left(strWhere,4)) = "AND" Then
				strWhere = " WHERE " & Right(strWhere,len(strWhere)-Instr(1,strWhere,"WHERE")-4)
			'End If
		End If
		
	End If
	
	If y > 0 Then
		Response.Write "<TR><TH colspan=""7"">Total <input type=""HIDDEN"" id=""WhereClause"" name=""WhereClause"" value=""" & strWhere & """ ></TH>" & _
				"<TH colspan=""2"" style=""text-align:center;"">" & x & "</TH></TR>"
	End If
	
	
objRS.Close

End Sub


Sub LoadDetails()

   'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM tblCAPSCDMC WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "'",objCon

		If Not objRS.EOF Then
		   
			'lngCardID = objRS("CardID")
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
			Session("CardID") = 0
			lngCardID = 0'objRS("CardID")
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

Public Sub CancelCard()

Dim strCardNo
Dim strCardType
Dim intCancelID

	'Get the Card Details for the CardID actioned
	objRS.Open "SELECT [CardType],[CardTypeSub],[CardNumber] FROM tblCAPSCard WITH(NOLOCK) WHERE CardID = '" & Request.QueryString("CardID") & "'",objCon

		If objRS.EOF Then
			strCardNo = ""
			strCardType = ""
		Else
			strCardNo = objRS("CardNumber")
			strCardType = objRS("CardType") & " " & objRS("CardTypeSub")
		End If
		
	objRS.Close
	response.write " can=" & Request.QueryString("CardEID") & "," & Request.QueryString("CardID") & "," & strCardNo
	'Call the function to add the card to the CS To Diners table
	intCancelID = CancelCardToCS(0,Request.QueryString("CardEID"),Request.QueryString("CardID"),strCardNo,"Cancelled in Unactivated Cards Admin Screen")
	
	'Check for errors returned from the Stored Procedure
	If intCancelID = -1 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Card Not Cancelled OR added to the CS To Diners File as " & Request.QueryString("CardEID") & " is ALREADY ON THE CS FILE!</div>"
		Exit Sub
	End If
	
	strSQL = "UPDATE tblCAPSCard SET ProcessStatus = 'Added to CS', Warning = '" & Left(now(),20) & "' WHERE CardID = " & Session("CardID") & ""
	
	objCon.Execute strSQL
	
	'Call the procedure to save a message for the card
	Call SaveMessage("Card Cancelled from Unactivated List")
	
	'Call the function to save the Audit Log record
	Call SaveAuditLog(0,"Unactivated Card","Card Cancelled",Request.QueryString("CardEID"),strCardType,strCardNo,Session("UserID"),"","","","Card Cancelled in Unactivated Card screen",Request.QueryString("CardID"),0,0,0,"UnactivatedCards.asp")
	
End Sub

Public Sub RemoveCard()
Dim strCardNo
Dim strCardType

	strSQL = "UPDATE tblCAPSCard SET ProcessStatus = 'Removed Unactivated', Warning = '" & Left(now(),20) & "' WHERE CardID = " & Session("CardID") & ""
	
	objCon.Execute strSQL
	
	'Remove from the Unactivated table
	strSQL = "DELETE FROM tblCAPSUnactivated WHERE CardID = " & Session("CardID") & ""
	
	objCon.Execute strSQL
	
	'Call the procedure to save a message for the card
	Call SaveMessage("Card Removed from Unactivated List")
	
	'Get the Card Details for the CardID actioned
	objRS.Open "SELECT [CardType],[CardTypeSub],[CardNumber] FROM tblCAPSCard WITH(NOLOCK) WHERE CardID = '" & Request.QueryString("CardID") & "'",objCon

		If Not objRS.EOF Then
			strCardNo = ""
			strCardType = ""
		Else
			strCardNo = objRS("CardNumber")
			strCardType = objRS("CardType") & " " & objRS("CardTypeSub")
		End If
		
	objRS.Close
	
	'Call the function to save the Audit Log record
	Call SaveAuditLog(0,"Unactivated Card","Card Cancelled",Request.QueryString("CardEID"),strCardType,strCardNo,Session("UserID"),"","","","Card Removed from list in Unactivated Card screen",Request.QueryString("CardID"),0,0,0,"UnactivatedCards.asp")
	
	Response.Write "<div class=""alert alert-success"" role=""alert"">Card " & strCardNo & " Removed from Unactivated list!</div>"
	
End Sub


Public Sub EmailCard()
Dim strCardNo
Dim strCardType

	strSQL = "UPDATE tblCAPSCard SET ProcessStatus = 'Email Unactivated', Warning = '" & Left(now(),20) & "' WHERE CardID = " & Session("CardID") & ""
	
	objCon.Execute strSQL
	
	'Call the procedure to save a message for the card
	Call SaveMessage("Card Emailed from Unactivated List")
	
	'Get the Card Details for the CardID actioned
	objRS.Open "SELECT [CardType],[CardTypeSub],[CardNumber] FROM tblCAPSCard WITH(NOLOCK) WHERE CardID = '" & Request.QueryString("CardID") & "'",objCon

		If Not objRS.EOF Then
			strCardNo = ""
			strCardType = ""
		Else
			strCardNo = objRS("CardNumber")
			strCardType = objRS("CardType") & " " & objRS("CardTypeSub")
		End If
		
	objRS.Close
	
	'Call the function to save the Audit Log record
	Call SaveAuditLog(0,"Unactivated Card","Card Cancelled",Request.QueryString("CardEID"),strCardType,strCardNo,Session("UserID"),"","","","Card Emailed in Unactivated Card screen",Request.QueryString("CardID"),0,0,0)
	
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
			.Parameters.Append objCmd.CreateParameter("CDMCToCardIDOutput", adInteger, adParamOutput)
			
			.Parameters("EmployeeID") = Session("EmployeeID")
			.Parameters("CardType") = Left(Request.Form("CardType"),3)
			.Parameters("CreditLimit") = Request.Form("CreditLimit")
			.Parameters("CardTypeSub") = Right(Request.Form("CardType"),Len(Request.Form("CardType"))-6)
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCToCardIDOutput") 
	 
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" /> Application " & intRecord & " submitted to your GCFO for approval!"

		strMessageColour = "Black"
		
End Sub

Public Sub SaveMessage(strMessage)

Dim lngMessageID
Dim lngAdminID
Dim intRecord
Dim strMessageTitle

If Session("MessageID") = "" or IsNull(Session("MessageID")) Then
	lngMessageID = 0
Else
	lngMessageID = Session("MessageID")
End If

If Session("AdminID") = "" or IsNull(Session("AdminID")) Then
	lngAdminID = 0
Else
	lngAdminID = Session("AdminID")
End If

'Set the message title to Application
strMessageTitle = "Application"

	'Makes sure that there is content in the message or do not save
	If strMessage = "" Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Message detail to save...NOT SAVED!</div>"
	Else
		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSMessageSave"

			.Parameters.Append objCmd.CreateParameter("MessageID", adInteger)
			.Parameters.Append objCmd.CreateParameter("MessageFrom", adInteger)
			.Parameters.Append objCmd.CreateParameter("MessageTo", adInteger)
			.Parameters.Append objCmd.CreateParameter("MessageTitle", adVarChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("MessageDetail", adVarChar, adParamInput, 500)
			.Parameters.Append objCmd.CreateParameter("MessageDate", adInteger)
			.Parameters.Append objCmd.CreateParameter("MessageStatus", adVarChar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("MessageRead", adChar, adParamInput, 1)
			.Parameters.Append objCmd.CreateParameter("MessageThreadID", adInteger)
			.Parameters.Append objCmd.CreateParameter("Object", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("ObjectID", adInteger)
			.Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput, 1)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
			.Parameters.Append objCmd.CreateParameter("MessageIDOutput", adInteger, adParamOutput)
			
			.Parameters("MessageID") = lngMessageID
			.Parameters("MessageFrom") = Session("UserID")
			.Parameters("MessageTo") = lngAdminID
			.Parameters("MessageTitle") = strMessageTitle
			.Parameters("MessageDetail") = strMessage
			.Parameters("MessageDate") = now()
			.Parameters("MessageStatus") = "Created"'Session("MessageStatus")
			.Parameters("MessageRead") = "N"'Session("MessageRead")
			.Parameters("MessageThreadID") = 0
			.Parameters("Object") = "Card"
			.Parameters("ObjectID") = Session("CardID")
			.Parameters("Active") = "Y"
			.Parameters("UpdatedBy") = Session("UserID")
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("MessageIDOutput") 
	 
		Response.Write "<div class=""alert alert-success"" role=""alert"">Message " & intRecord & " Saved!</div>"
	End If

End Sub


Public Sub CreateBatch()
'Procedure to create a new batch of Unactivated Cards
Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCreateNewBatch"

			.Parameters.Append objCmd.CreateParameter("CardID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmailDetailID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Days", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Warning", adChar, adParamInput, 1) 
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("NewBatchOutputID", adInteger, adParamOutput)
			
			.Parameters("CardID") = 0'Session("CardID")
			.Parameters("EmailDetailID") = Request.QueryString("Emails")
			.Parameters("Days") = 45'Request.QueryString("Days")
			.Parameters("Warning") = "Y"'Request.QueryString("Warning")
			.Parameters("UpdatedBy") = Session("UserID")
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("NewBatchOutputID") 
	 
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" /> Application " & intRecord & " submitted to your GCFO for approval!"

		strMessageColour = "Black"
		
End Sub


Public Sub CancelBatch(strStatus)
'Procedure to cancel an existing batch
' TO BE UPDATED to send emails in the future
Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCancelBatch"

			.Parameters.Append objCmd.CreateParameter("BatchID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("BatchOutputID", adInteger, adParamOutput)
			
			.Parameters("BatchID") = Session("BatchID")
			.Parameters("Status") = strStatus
			.Parameters("UpdatedBy") = Session("UserID")
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("BatchOutputID") 
	 
		If intRecord < 1 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Batch Not Cancelled! See System Admin</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Batch " & Session("BatchID") & " " & strStatus & " !</div>"
		End If
		
End Sub


Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(5)

If Session("ViewButton") = "Emailed" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "Removed" Then
	arrButton(3) = "active"
ElseIf Session("ViewButton") = "AddedToCS" Then
	arrButton(4) = "active"
ElseIf Session("ViewButton") = "InABatch" Then
	arrButton(5) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='UnactivatedCards.asp?ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='UnactivatedCards.asp?ViewButton=Emailed';""><i class=""fa fa-envelope""></i> View Emailed</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='UnactivatedCards.asp?ViewButton=Removed';""><i class=""fa fa-times""></i> View Removed</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='UnactivatedCards.asp?ViewButton=AddedToCS';""><i class=""fa fa-arrow-circle-down""></i> View Added To CS</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(5) & """ onClick=""self.location.href='UnactivatedCards.asp?ViewButton=InABatch';""><i class=""fa fa-beer""></i> In a Batch</button>" & _
				"</div>"

End Sub

Public Sub LoadBatch()
'Load the Campaign details for emails already sent
Dim strStatusDisplay
Dim strEmailSubject
Dim lngBatchCount
Dim strCreatedDate
Dim strClear
Dim strBGColor
Dim strTitle

	'Allow the user to clear the selceted batch if one is selected
	If IsNull(Session("BatchID")) Or Session("BatchID") = ""  Then
		strClear = ""
	Else
		strClear = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onClick=""self.location.href='UnactivatedCards.asp?BatchID=';""><i class=""fa fa-times""></i> Clear Batch</button>"
	End If
	
	'Description:	Loads CDMC Details onto the page called from
	objRS.Open "SELECT * FROM qryCAPSUnactivatedBatch WITH(NOLOCK) ORDER BY [DateCreated] DESC" ,objCon

		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"" style=""font-weight:bold;"">Unactivated Batches</span>" & strClear & "</div></div>"
		
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact table-hover text-left""><thead>" & _
			"<tr><th style=""font-weight:bold; font-size:12px;"">Date Created</th><th style=""font-weight:bold; font-size:12px;"">Count</th>" & _
			"<th style=""font-weight:bold; font-size:12px;"">Status</th></tr></thead><tbody>"
		
		
		Do Until objRS.EOF
			
			'Determine the pill display based on the Status
			If IsNull(objRS("Status")) or objRS("Status") = "" Then
				strStatusDisplay = "<span class=""badge badge-info"">None</span>"
			Else
				If objRS("Status") = "Generated" Then
					strStatusDisplay = "<span class=""badge badge-secondary"">" & objRS("Status") & "</span>"
				ElseIf objRS("Status") = "Cards Cancelled" Then
					strStatusDisplay = "<span class=""badge badge-danger"">" & objRS("Status") & "</span>"
				Else
					strStatusDisplay = "<span class=""badge badge-success"">" & objRS("Status") & "</span>"
				End If
			End If
			
			If IsNull(objRS("DateCreated")) Then
				strCreatedDate = ""
			Else
				strCreatedDate = FormatDateTime(objRS("DateCreated"),vbShortDate)
			End If
			
			'Get the selected Batch for background colouring
			If IsNull(Session("BatchID")) or Session("BatchID") = "" Then
				strBGColor = ""
			Else
				If cstr(Session("BatchID")) = cstr(objRS("BatchID")) Then
					strBGColor = "table-active"
				Else
					strBGColor = ""
				End If
				
			End If
			
			strTitle = " Title=""Batch ID = " & objRS("BatchID") & """ "
			
			lngBatchCount = objRS("CardCount")
			
			Response.Write "<tr class=""clickable-row " & strBGColor & """ data-href='UnactivatedCards.asp?BatchID=" & objRS("BatchID") & "' data-target='_blank' " & strBGColor & " " & strTitle & "><td style=""font-size:12px;"">" & strCreatedDate & "</td><td style=""font-size:12px;"">" & lngBatchCount & "</td>" & _
				"<td style=""font-size:16px;"">" & strStatusDisplay & "</td></TR>"
			
			
		objRS.Movenext
		Loop
		
		'Else
		'	Response.write "No campaigns"
	   'End If

	objRS.Close
	
	Response.write "</tbody></table></div></div>"
	
End Sub

Public Sub LoadAdminButtons()
'Procedure to write the Excel export button bsaed on the user type of the person logge din

If Session("UserTypeID") > 10 then

	Response.write "<button type=""button"" class=""btn btn-outline-success"" onClick=""OpenExcelReport();""><i class=""fa fa-file-excel""></i> Export To Excel</button>"
	
End If

End Sub

Set objRS = Nothing
Set objCon = Nothing
%>
