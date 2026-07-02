<!-- #Include file=../../CC/CAPSHeader.asp -->
<!-- #Include file=../../ADOVBS.inc -->
<!-- #Include file=../../CC/CAPSFunctions.asp -->
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
Dim objCmd1

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
Dim strNextNAFileNumber
Dim strCardType

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")
	Set objCmd1 = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	

	If IsNull(Session("CardID")) OR Session("CardID") = "" Then Session("CardID")= 0

	If isNull(Session("CardID")) Or Session("CardID") = "" Then 
		Session("CardID") = 0
	End If
	
	If Not IsEmpty(Request.QueryString("UserView")) Then
		Session("UserView") = Request.QueryString("UserView")
	End If
	
	If Not IsEmpty(Request.QueryString("CardType")) Then
		strCardType = Request.QueryString("CardType")
	Else
		strCardType = Session("CardType")
	End If
	
	'Response.Write "<BR>Card Type " & strCardType & "<BR>"

	If Not IsEmpty(Request.QueryString("CardID")) Then
		Session("CardID") = Request.QueryString("CardID")
	End If
	
	If Not IsEmpty(Request.QueryString("FileLoadID")) Then
		Session("FileLoadID") = Request.QueryString("FileLoadID")
	End If
	
	If Not IsEmpty(Request.QueryString("Action")) Then
		
	End If

	If Not IsEmpty(Request.QueryString("PageCombo")) Then
		Session("PageCombo") = Request.QueryString("PageCombo")
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
		
		'If the Export NA File button has been clicked then call the procedure to export the NA File
		If Request.QueryString("Action") = "ExportNA" Then
			Call ExportNAFile(strCardType)
		End If
	End If
	
	'If the Cancel/Remove has been clicked on the NA File Modal (within the file GetNAfile.asp) then flag the NA File record as removed
	If Request.QueryString("Action") = "CancelNA" Then

		Call RemoveNARecord(Request.QueryString("NAToDinersID"),Request.QueryString("NAEID"),Request.QueryString("Status"))

	End If	

	'Get the next NA File Number for display on screen
	If Not IsEmpty(Request.QueryString("BatchNumber")) Then
		strNextNAFileNumber = Request.QueryString("BatchNumber")
		If IsNumeric(strNextNAFileNumber) Then strNextNAFileNumber = PadDigits(strNextNAFileNumber,6)
	Else
		strNextNAFileNumber = GetSystemAdmin("NAFileNumber")
		If IsNumeric(strNextNAFileNumber) Then strNextNAFileNumber = PadDigits(strNextNAFileNumber,6)
	
	End If	
	
	'strNextNAFileNumber = Session("FileLoadID")
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

function ConfirmExport(cb) {
	
	var id = cb.getAttribute('data-NANumber');
	document.getElementById('NAFileExport').value=id;

}

function ChangeBatch() {

	var e = document.getElementById("NAFileSelect");
	var CType = document.getElementById("CardTypeMod").value;
	var result = e.options[e.selectedIndex].text;

	self.location='NATransactions.asp?FileLoadID='+result +'&CardType='+ CType + '&BatchNumber='+result
}

function ChangePage() {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("PageCombo");
	var result = e.options[e.selectedIndex].value;
	
	self.location = 'NATransactions.asp?PageCombo=' + result;
	//alert(result);
	//document.getElementById('CardType').value=result;
	
}

function loadDoc(varID) {

  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("NAFileDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../../CC/AJAX/GetNAToDinersAudit.asp?NAToDinersID=" + varID + "", true);
  xhttp.send();

}

function loadCard(varID) {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("NAFileDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetNAFromDinersCard.asp?NAFromDinersID=" + varID + "", true);
  xhttp.send();
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

<!-- Modal Detail Compare -->
<div class="modal fade" id="NAFileMod" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-large modal-dialog-centered modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  <%=strCardType%> - NA To Diners File Detail
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="NAFileDetail">
               
				  
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
<!-- END Modal Detail Compare -->

<!-- Approve Modal -->
<div class="modal fade" id="ModalApprove" tabindex="-1" role="dialog" aria-labelledby="ModalApprove" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalApproveTitle" style="font-weight:bold;">NA File Export Confirmation</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <span><%=strCardType%> - Export NA File Number: <input type="text" name="NAFileExport" id="NAFileExport" style="border:0; font-weight:bold; width:100px; text-align:right;" value="<%=strNextNAFileNumber%>" >?</span><br><br>
      </div>
      <div class="modal-footer">
		<button type="button" class="btn btn-primary" onClick='window.location="NATransactions.asp?CardType=<%=strCardType%>&Action=ExportNA"'><i class="fa fa-check"></i> Yes</button>
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
        
      </div>
    </div>
  </div>
</div>
<!-- End Approve Modal -->

<!-- Select Batch Number Modal -->
<div class="modal fade" id="ModalSelectBatch" tabindex="-1" role="dialog" aria-labelledby="ModalSelectBatch" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalApproveTitle" style="font-weight:bold;">NA File Export Batch Number</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
	  <div class="col-md-4"><INPUT Type="hidden" id="CardTypeMod" name="CardTypeMod" value="<%=strCardType%>" ></div>
        <div class="col-md-4"><SELECT class="form-control" onChange="ChangeBatch();" name="NAFileSelect" id="NAFileSelect"><% Call LoadBatchList()%></Select></div><br><br>
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
  <form action="NATransactions.asp?CardType=<%=strCardType%>&Action=Search" method="POST" id="frm" name="frm">
	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-9">
				<h4 class="text-left" data-toggle="modal" data-target="#ModalSelectBatch"  Title="Click to Select an NA File to View"><%=strCardType%> - NA File To Diners <%="File Load ID: " & strNextNAFileNumber%> </h4>
			</div>
			<div class="col-md-3 float-right">
				<button type="button" class="btn btn-primary" onclick="window.open('NAToDinersExportExcel.asp?CardType=<%=strCardType%>')"><i class="fa fa-file"></i> Export NA File to Excel</button>
				<!--<button type="button" class="btn btn-primary float-right" onClick='window.location="NATransactions.asp?Action=ExportNA"'><i class="fa fa-file"></i> Export NA File</button>-->
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
	
<!-- #Include file=../../CC/CAPSFooter.asp -->

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
Dim strBatchNumber
Dim strWhere
Dim lngStartingRecord
Dim lngTotalRecords
Dim strUpdateDate

	strSearch = Request.Form("SearchInput")
	
	'strSearch = Request.Form("SearchInput") & " " & Request.Form("BatchNumber")
	
	'If a Batch Number has been selected, then only show the records in that batch
	If IsEmpty(Request.QueryString("BatchNumber")) Then
		strBatchNumber = 0
	Else
		strWhere = " AND [BatchNumber] = " & Request.QueryString("BatchNumber")
		strWhere = " AND [FileSeqNum] = '" & Request.QueryString("BatchNumber") & "'"
	End If
	
	'The below overwrites the above.  Remove the above once all tested....
	If IsEmpty(Request.QueryString("FileLoadID")) Then
		strBatchNumber = 0
	Else
		'strWhere = " AND [BatchNumber] = " & Request.QueryString("BatchNumber")
		strWhere = " AND [FileSeqNum] = '" & Request.QueryString("FileLoadID") & "'"
	End If
	
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
		strSort = " ORDER BY [FileSeqNum] " & strOrderType
	Else		
		strSort = " ORDER BY " & Request.QueryString("Sort") & " " & strOrderType
	End If
	
	If Session("ViewButton") = "WaitingExport" Then
		strWhere = strWhere & " AND [Status] = 'Added To NA' "
	ElseIf Session("ViewButton") = "NoResponse" Then
		strWhere = strWhere & " AND [AuditLogID] IS NULL "
	ElseIf Session("ViewButton") = "Response" Then
		strWhere = strWhere & " AND [AuditLogID] IS NOT NULL "
	Else
		'This catches ALL
		'strWhere = ""
	End If
	
	'Build the TOP Statement
	If Session("PageCombo") = "" Or IsNull(Session("PageCombo")) Then
		Session("PageCombo") = 50
	End If
	
	'If IsNumeric(Session("PageCombo")) Then
	'	strTOP = "TOP " & Session("PageCombo")
	'Else
	'	strTOP = ""
	'End If
	
	'Update to make sure that the complete recordset is not returned when no filters are selected as it takes a long time to load
	If strWhere ="" OR IsNull(strWhere) Then strTop = "Top 100"
	
If strSearch = "" OR ISNull(strSearch) Then
	strSQL = "SELECT " & strTOP & " * FROM qryCAPSNAToDinersDetail WITH(NOLOCK) WHERE CardType = '" & strCardType & "' AND [NAToDinersID] > 0 " & strWhere & strSort
Else
	strSQL = "SELECT " & strTOP & " * FROM qryCAPSNAToDinersDetail WITH(NOLOCK) WHERE CardType = '" & strCardType & "' AND (EmployeeID Like '%" & strSearch & "%' OR FormalFirstName Like '%" & strSearch & "%' OR FormalLastName Like '%" & strSearch & "%')" & strWhere & strSort
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
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No NA File records for " & strRecordMessage & " " & strWhere & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
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
	
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Displaying " & Session("PageCombo") & " of " & lngTotalRecords & " NA File to Diners records (" & lngStartingRecord & " to " & lngStartingRecord + clng(Session("PageCombo")) & ")</span><span class=""panel-subheader"" style=""float:right;"">Number of records per page: " & strPageCombo  & "</span></div></div>"
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""NATransactions.asp?CardType=" & strCardType & "&Sort=NAToDinersID&SortType=" & strOrderType & """> NA ID<i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""> Record Type <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""NATransactions.asp?CardType=" & strCardType & "&Sort=Status&SortType=" & strOrderType & """> Status <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Record Text</th>" & _
			"<th scope=""col""><a href=""NATransactions.asp?CardType=" & strCardType & "&Sort=BatchNumber&SortType=" & strOrderType & """> Batch No. <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""NATransactions.asp?CardType=" & strCardType & "&Sort=DateUpdated&SortType=" & strOrderType & """> Date Updated <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""NATransactions.asp?CardType=" & strCardType & "&Sort=UpdatedBy&SortType=" & strOrderType & """> Updated By </th>" & _
			"<th scope=""col""><a href=""NATransactions.asp?CardType=" & strCardType & "&Sort=EmployeeID&SortType=" & strOrderType & """> Employee ID  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""NATransactions.asp?CardType=" & strCardType & "&Sort=ApplicationID&SortType=" & strOrderType & """> Application ID  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Applicant Name</th>" & _
			"<th scope=""col"">Action</th>" & _
			"</tr></thead><tbody class=""text-left"">"
					
	End If
	
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No NA File to Diners records " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	End If
    
	x = 0
	
    Do until objRS.EOF 

		y = y + 1
		
		'Only write the first 50 records from the starting position
		If y <= lngStartingRecord + clng(Session("PageCombo")) AND y >= lngStartingRecord - clng(Session("PageCombo")) Then
		'If y <= lngStartingRecord + 50 AND y >= lngStartingRecord - 50 Then
		
			x = x + 1
			
			'Create the Status list badge and Action button based on the status field
			If IsNull(objRS("Status")) Then
				strStatus = ""
			Else
				If objRS("Status") = "Exported" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Exported</span>"
					
					strAction = ""'
				ElseIf objRS("Status") = "Added To NA" Then
					strStatus = "<span class=""badge badge-pill badge-warning"">Added To NA</span>"
					
					strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='NATransactions.asp?CardType=" & strCardType & "&Action=CancelNA&NAToDinersID=" & objrs("NAToDinersID") & "&NAEID=" & objRS("EmployeeID") & "&Status=Deleted'""; title=""Click to Remove from NA File""><i class=""fa fa-times""></i> Remove</button>"
					
				ElseIf objRS("Status") = "Deleted" Then
					strStatus = "<span class=""badge badge-pill badge-danger"">Deleted</span>"
					
					strAction = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='NATransactions.asp?CardType=" & strCardType & "&Action=CancelNA&NAToDinersID=" & objrs("NAToDinersID") & "&NAEID=" & objRS("EmployeeID") & "&Status=Added To NA'""; title=""Click to Add to NA File""><i class=""fa fa-plus""></i> Add</button>"
				Else
					strStatus = objRS("Status")
					
					strAction = ""'
				End If
			End If
			
			'Format the Date updated for short display
			If IsNull(objRS("DateUpdated")) Then
				strUpdateDate = ""
			Else
				If IsDate(objRS("DateUpdated")) Then
					strUpdateDate = FormatDateTime(objRS("DateUpdated"), vbShortDate)
				Else
					strUpdateDate = objRS("DateUpdated")
				End If
			End If
			
			
			Response.Write "<TR><TD ><a data-toggle=""modal"" data-target=""#NAFileMod"" HREF=""#"" onClick=""loadDoc(" & objRS(0) & ")"">" & objRS(0) & "</a></TD><TD>" & objRS("RecordType") & "</a></TD>" & _
					"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""../../CC/ApplicationDetail.asp?ApplicationID=" & objRS("ApplicationID") & """>" & trim(objRS("RecordText")) & "</a></TD><TD>" & strStatus & "</TD>" & _
					"<TD style=""text-align:center;"">" & objRS("FileSeqNum") & "</TD><TD title=""" & objRS("DateUpdated") & """>" & strUpdateDate & "</TD><TD>" & objRS("UpdatedByName") & "</TD><TD >" & objRS("EmployeeID") & "</TD>" & _
					"<TD style=""font-size:12px; background-color:#e6eeff;"">" & objRS("ApplicationID") & "</TD><TD style=""font-size:12px; background-color:#e6eeff;"">" & objRS("ApplicantName") & "</TD>" & _
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
						strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""NATransactions.asp?CardType=" & strCardType & "&StartingRecord=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
					End If
					
					'Add the Elipsis (...) to the start of the page numbers if there is more than 20 pages and the current place is beyond the first page
					If x = 0 AND lngPage > 1 AND y > 20 Then
						strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""NATransactions.asp?CardType=" & strCardType & "&StartingRecord=" & lngTotalRecords - (clng(Session("PageCombo"))*20) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
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
				
					'strPages = strPages & "<a href=""NATransactions.asp?StartingRecord=" & (x * clng(Session("PageCombo"))) & """> " & x + 1 & " </a>"
					strPages2 = strPages2 & "<li class=""page-item " & strActive & """><a class=""page-link"" href=""NATransactions.asp?CardType=" & strCardType & "&StartingRecord=" & (x * clng(Session("PageCombo"))) & """>" & x & "</a></li>"
					
					'strPages = strPages & "<a href=""NATransactions.asp?StartingRecord=" & (x * 50) & """> " & x & " </a>"
					'strPages2 = strPages2 & "<li class=""page-item " & strActive & " -" & x & "|" & lngPage & """><a class=""page-link"" href=""NATransactions.asp?Previous&StartingRecord=" & lngStartingRecord -50 & """>" & x & "</a></li>"
					End If
				End If
			Next
			
		End If
	End If
	
	'If y > 0 Then
	'	Response.Write "<TR><TH colspan=""9"" style=""text-align:center;""><a href=""NATransactions.asp?Previous&StartingRecord=" & lngStartingRecord -50 & """>Previous Page " & strPages & " <a href=""NATransactions.asp?Previous&StartingRecord=" & lngStartingRecord + 50 & """> Next Page</TH></TR>"
	'End If
	
	
	'Write the End of the table and divs for the above list, as the pagination (below) is in it's own container
	Response.Write "</tbody></table></div>"

	'Write the Pagination objects for all pages based on the total records and the number records displayed on screen
	If y > 0 Then
		
		Response.Write "<div class=""container""><div class=""row""><div class=""col-12 text-center"">" & _
			"<nav aria-label=""Page navigation""><ul class=""pagination""><li class=""page-item"">" & _      
			"<a class=""page-link"" href=""NATransactions.asp?CardType=" & strCardType & "&StartingRecord=0"" aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
			strPages2 & _
			"<li class=""page-item"">" & _
			"<a class=""page-link"" href=""NATransactions.asp?CardType=" & strCardType & "&StartingRecord=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
			"</a></li></ul></nav></div></div></div>"
			
			'"<a class=""page-link"" href=""NATransactions.asp?StartingRecord=" & lngStartingRecord - clng(Session("PageCombo")) & """ aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
	End If
		
objRS.Close

End Sub

Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(4)

If Session("ViewButton") = "WaitingExport" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "NoResponse" Then
	arrButton(3) = "active"
ElseIf Session("ViewButton") = "Response" Then
	arrButton(4) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='NATransactions.asp?CardType=" & strCardType & "&ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='NATransactions.asp?CardType=" & strCardType & "&ViewButton=WaitingExport';""><i class=""fa fa-cogs""></i> View Awaiting Export</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='NATransactions.asp?CardType=" & strCardType & "&ViewButton=NoResponse';""><i class=""fa fa-thumbs-down""></i> View No Response</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='NATransactions.asp?CardType=" & strCardType & "&ViewButton=Response';""><i class=""fa fa-thumbs-up""></i> View Response</button>" & _
				"</div>"

End Sub


Public Sub RemoveNARecord(lngNAToDinersID, strEmployeeID, strStatus)
'Procedure to add and remove records from the NA File before it is exported
Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSNAFileRemoveCard"

			.Parameters.Append objCmd.CreateParameter("NAToDinersID", adVarChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("NAFileRemoveOutput", adInteger, adParamOutput)
			
			.Parameters("NAToDinersID") = lngNAToDinersID
			.Parameters("UpdatedBy") = Session("UserID")
			.Parameters("Status") = strStatus
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("NAFileRemoveOutput") 
	 
		If intRecord = 0 Then
			If strStatus = "Deleted" Then
				Response.Write "<div class=""alert alert-danger"" role=""alert"">Application for " & strEmployeeID & " NOT Removed from NA File! An Error has occurred. See System Admin with NA File ID: " & lngNAToDinersID & " </div>"
			Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">Application for " & strEmployeeID & " NOT Added to the NA File! An Error has occurred. See System Admin with NA File ID: " & lngNAToDinersID & " </div>"
			End If
		Else
			If strStatus = "Deleted" Then
				Response.Write "<div class=""alert alert-success"" role=""alert"">Application for " & strEmployeeID & " REMOVED from the NA file!</div>"
			Else
				Response.Write "<div class=""alert alert-success"" role=""alert"">Application for " & strEmployeeID & " ADDED to the NA file!</div>"
			End If
		End If
		
	
End Sub

Public Sub ExportNAFile(strCardType)
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
Dim strFilePathOnly

Dim objFSO
Dim strFileName


'Get the Default file location for the server then add the filepath and name for the NA File
strFilePath = GetSystemAdmin("ServerFilePath")

If strCardType = "DTC" Then
	strFileNameStart = GetSystemAdmin("NAFileStart")
	strNextFileNumber = GetSystemAdmin("NAFileNumber")
Else
	strFileNameStart = GetSystemAdmin("NAFileStartDPC")
	strNextFileNumber = GetSystemAdmin("NAFileNumberDPC")
End If

'Pad the number out to 6 digits
strNextFileNumber = PadDigits(strNextFileNumber,6)
strNextFileNumber = PadDigits(strNextFileNumber,6)

'Compile the File name and path from the variables above
strFilePath = strFilePath & "\Admin\CAPSAdmin\Attachments\Diners\DinersTo\" & strFileNameStart & PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & ".txt"

'Set the filename to be used when moving the file in the procedure called at the end of this procedure
strFileName = strFileNameStart & PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & ".txt"

'Get the file path only for the file transfer function
strFilePathOnly = strFilePath & "\Admin\CAPSAdmin\Attachments\Diners\DinersTo\"

strFileDateTime = PadDigits(Right(Year(Now()),4),4) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2)
strFileDateTimeSec = PadDigits(Right(Year(Now()),4),4) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & Hour(Now()) & Minute(Now()) & Second(Now())
'response.write 	"lngBatchNumber=" & strNextFileNumber & " strNextNAFileNumber= " &  strNextNAFileNumber

'Get the FileLoad details
'lngBatchNumber = GetFileLoadID("NAFile",strNextFileNumber,"")

'If IsNull(lngBatchNumber) OR lngBatchNumber = "" then
'	lngBatchNumber = strNextNAFileNumber
'End If



Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

'Open the text file
Dim objTextStream

	'Open a recordset of all of the NA File records yet to be exported
	'Response.Write "SELECT * FROM tblCAPSNAToDiners WITH(NOLOCK) WHERE CardType = '" & strCardType & "' AND BatchNumber = 0 AND Status = 'Added To NA'"
	''Updated April 2023 for DPC change --added Card Type
	objRS.Open "SELECT * FROM tblCAPSNAToDiners WITH(NOLOCK) WHERE CardType = '" & strCardType & "' AND BatchNumber = 0 AND Status = 'Added To NA'",objCon

		If objRS.EOF Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No records to write to the NA File.</div>"
		Else
			Set objTextStream = objFSO.OpenTextFile(strFilePath, fsoForWriting, True)
			'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
			lngFileLoadID = SaveFileLoadID ("NAFile",strFileNameStart & strFileDateTime & ".txt", strFilePath,-1,0,0,0,0,0,0,0,strFileDateTime,strNextFileNumber,"Exported",Session("UserID"),"N")
			
			objTextStream.WriteLine "H NA" & strFileDateTimeSec & strNextFileNumber
			'objTextStream.WriteLine "1 H NA" & strFileDateTimeSec & strNextFileNumber
			
			'format(Now(),"yyyymmddhhmmss")+Format(DLookup("Next_File_No","Next_File_No","File_Type = 'NA_Out'"),"000000");\
			
			'Write each record to the text file
			Do Until objRS.EOF
			
				intRecordCount = intRecordCount + 1
				'Display the contents of the text file
				If IsNull(objRS("RecordText")) = False Then
					objTextStream.WriteLine objRS("RecordText")
				End If
				'Call the procedure to update each record as exported once added to the NA File -- USE the File style Batch Number not FileLoadID
				Call ExportNARecord (objRS("NAToDinersID"),lngFileLoadID,intRecordCount,strCardType)
				'Call ExportNARecord (objRS("NAToDinersID"),strNextFileNumber,intRecordCount)
				
			objRS.Movenext
			Loop
			
			strRecordCount = PadDigits(intRecordCount,6)
			
			objTextStream.WriteLine "T" & strRecordCount
			'objTextStream.WriteLine "3 T" & strRecordCount			
			
			'Call the procedure to update the System Parameter NAFileNumber. Increment the Number by 1.
			
			Call EmailNARecord(PadDigits(strNextFileNumber,6))
			Call UpdateBatchNumber(strNextFileNumber)
		
			'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			'strFileName = strFileNameStart & strFileDateTime & ".txt" ------Set at the top of the procedure now as it requires 2 digit year -- 4 digit year required within the file as header
			Call UpdateFileLoadSummary ("NAFile",strNextFileNumber, strFileName, lngFileLoadID)
			
			Response.Write "<div class=""alert alert-success"" role=""alert"">NA File " & "AUDC_INTOECS_DODNA_D" & strFileDateTime & ".txt" & " ADDED to the NA file export folder!</div>"
			objTextStream.Close
		End If

	objRS.Close


	'Call the procedure to move the file created to the G Drive
	Call MoveExportFiles(strFilePath, strFileName, strFilePathOnly)
	
'Close the file and clean up
Set objTextStream = Nothing
Set objFSO = Nothing



End Sub

Public Sub ExportNARecord(lngNAToDinersID,lngBatchNumber,x,strCardType)
'Procedure to Change the Status of NA file records being exported and adds an Audit Log record
Dim intRecord

  	With objCmd

		.CommandType = 4
		.CommandText = "spCAPSNAFileExportCard"
		
		'Only create the parameters the first time the procedure is created otherwise there will be an error
		If x = 1 Then
			.Parameters.Append objCmd.CreateParameter("NAToDinersID", adVarChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("BatchNumber", adVarChar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("CardType", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("NAFileExportOutput", adInteger, adParamOutput)
		End If	
	
		.Parameters("NAToDinersID") = lngNAToDinersID
		.Parameters("BatchNumber") = lngBatchNumber
		.Parameters("CardType") = strCardType
		.Parameters("UpdatedBy") = Session("UserID")
		
		.ActiveConnection = objCon
		 
	End With
   
	objCmd.Execute        
  
	'Return the result of the Save Function.
	intRecord = objCmd.Parameters.Item("NAFileExportOutput") 
 
End Sub

Public Sub EmailNARecord(strFileSeqNum)
'Procedure to Change the Status of NA file records being exported and adds an Audit Log record
Dim intRecord

  	With objCmd1

		.CommandType = 4
		.CommandText = "spCAPSPostNARecordEmails"		
	
			.Parameters.Append objCmd1.CreateParameter("FileSeqNum", adVarChar, adParamInput,10)
			.Parameters.Append objCmd1.CreateParameter("UpdatedBy", adInteger, adParamInput)	
	
			.Parameters("FileSeqNum") = strFileSeqNum
			.Parameters("UpdatedBy") = Session("UserID")
		
		.ActiveConnection = objCon
		 
	End With
   
	objCmd1.Execute        
   
End Sub


Public Sub UpdateBatchNumber(lngBatchNumber)
'Procedure to update the BatchNumber field in the System Parameters table with the next number
Dim strSQL

	'If the Batch Number is a number then update the System Parameter, otherwise post an error to the screen
	If IsNumeric(lngBatchNumber) Then
		lngBatchNumber = lngBatchNumber + 1
		
		strSQL = "UPDATE tblCAPSSystemParameters SET [ParameterValue] = '" & lngBatchNumber & "' WHERE [ParameterName] = 'NAFileNumber'"

		objCon.Execute strSQL
	
	Else
		
		Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR! NA File Batch Number: " & lngBatchNumber & " is not a number. See System Admin.</div>"
		
	End If

End Sub

Public Sub LoadBatchList()
'Description:	Loads all Batch Numbers to a list for selecting and searching/filtering


	objRS.Open "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = 'NAFile' AND [Deleted] = 'N' ORDER By [FileSeqNum] DESC",objCon
  
	Response.write "<OPTION value=""0"">Select a Batch to View...</OPTION>"
	
		Do Until objRS.EOF 
			
			Response.write "<OPTION value=""" & objRS("FileLoadID") & """>" & objRS("FileSeqNum") & "</OPTION>"
			
			objRS.Movenext
			
		Loop
	
	objRS.Close
	
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

'Response.Write "<BR>strServer=" & strServer
'Response.Write "<BR>strUser=" & strUser

	'Get the System Parameter for the fileName
	'strFileNameDefault = GetSystemAdmin("CSFromDinersFileName")

'TEMP DELETE BELOOWWWW**********	
'strServer = "\\d85groupdata.dpe.protected.mil.au\groupdata_rus\CFO\CFO\CMS Admin\CAPS\Export Files"
		
	objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

		objStartFolder = strServer
	
	'strFilePathTo = strServer & strFileName
	strFilePathTo = strServer & "\" & strFileName
	
'Update to incliude trailing backslah after mapping drive (where trailing backslash causes an error)
	strServer = strServer & "\"

	'response.write "<Br>strFilePathFrom=" & strFilePathFrom
	'response.write "<Br>strFilePathTo=" & strFilePathTo
	
	'response.write "<Br>strServer=" & strServer & strFileNameNoExt & x & strFileExtension
	
	'strFilePathFrom = "D:\Apps\CAPS\AspNew\Admin\CAPSAdmin\Attachments\TestMG.txt"
	'strFilePathTo = "D:\Apps\CAPS\AspNew\Admin\CAPSAdmin\Attachments\TestMG001.txt"
	
	'Set the file extenstion to text if it doesn't have one
	If IsNull(strFileExtension) or strFileExtension = "" Then strFileExtension = "txt"
		
	'Set the first FileTo Name to check if it exists below
	'strFilePathTo = strFileNameNoExt & x & "." & strFileExtension
		
	If objFSO.FileExists(strFilePathTo) Then
		
		strFileExtension = objFSO.GetExtensionName(strFileName)
		strFileNameNoExt = Left(strFilePathTo,Len(strFilePathTo)-Len(strFileExtension)-1)
		
		For x = 1 to 10
			If objFSO.FileExists(strFilePathTo) Then
				strFilePathTo = strFileNameNoExt & x & "." & strFileExtension
				'strFilePathTo = strServer & strFileNameNoExt & x & "." & strFileExtension
				'Response.Write "<BR>strFilePathTo NO=" & strFilePathTo
			Else
				strFilePathTo = strFileNameNoExt & x & "." & strFileExtension
				'strFilePathTo = strServer & strFileNameNoExt & x & strFileExtension
				'strFilePathFrom = strFilePathOnly & strFileName & x & strFileExtension
				
				'Move the file to the Loaded folder
				objFSO.MoveFile strFilePathFrom, strFilePathTo
				'objFSO.MoveFile strFilePathFrom, strServer & strFileNameNoExt & x & strFileExtension
				'objFSO.MoveFile strFilePathFrom & strFileName & x & strFileExtension
				'Response.Write "<BR>strFilePathTo YES=" & strFilePathTo
				x = 10
			End If
		Next
	Else

	'Response.Write "<BR>strFilePathFrom=" & strFilePathFrom
	'Response.Write "<BR>strFilePathTo=" & strFilePathTo

		'Move the file to the Loaded folder
		'objFSO.MoveFile strFilePathOnly & strFileName,strServer & strFileName
		objFSO.MoveFile strFilePathFrom,strFilePathTo'strServer & strFileName
		
		Response.Write "<div class=""alert alert-success"" role=""alert"">NA File exported FROM: " & strFilePathFrom & "</div>"
		Response.Write "<div class=""alert alert-success"" role=""alert"">NA File exported TO: " & strFilePathTo & "</div>"
	
	End If

	'Remove the trailing backslash as the FSO object doesn;t like this on the new DPE server
	strServer = Left(strServer, len(strServer)- 1)
	objNetwork.RemoveNetworkDrive strServer, True, False
		 
	Set objFSO = Nothing
	Set objNetwork = Nothing
		
End Sub

Set objRS = Nothing
Set objCon = Nothing
%>
