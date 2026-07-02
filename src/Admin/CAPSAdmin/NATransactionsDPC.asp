
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

	If Not IsEmpty(Request.QueryString("CardID")) Then
		Session("CardID") = Request.QueryString("CardID")
	End If
	
	If Not IsEmpty(Request.QueryString("CardType")) Then
		strCardType = Request.QueryString("CardType")
		Session("CardTypeNA") = strCardType
	End If
	
	If Session("CardTypeNA") = "" or IsNull(Session("CardTypeNA")) Then  Session("CardTypeNA") = Session("CardType")
	
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
			Call ExportNAFile
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
	var result = e.options[e.selectedIndex].text;
	
	self.location='NATransactionsDPC.asp?FileLoadID='+result + '&BatchNumber='+result
}

function ChangePage() {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("PageCombo");
	var result = e.options[e.selectedIndex].value;
	
	self.location = 'NATransactionsDPC.asp?PageCombo=' + result;
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
  xhttp.open("GET", "../../CC/AJAX/GetNAToDinersAudit.asp?NAtoDinersID=" + varID + "", true);
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
                  <%=Session("CardTypeNA")%> - NA To Diners File Detail
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
        <span><%=Session("CardTypeNA")%> - Export 55 NA File Number: <input type="text" name="NAFileExport" id="NAFileExport" style="border:0; font-weight:bold; width:100px; text-align:right;" value="<%=strNextNAFileNumber%>" >?</span><br><br>
      </div>
      <div class="modal-footer">
		<button type="button" class="btn btn-primary" onClick='window.location="NATransactionsDPC.asp?CardType=<%=Session("CardTypeNA")%>&Action=ExportNA"'><i class="fa fa-check"></i> Yes</button>
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
  <form action="NATransactionsDPC.asp?CardType=<%=Session("CardTypeNA")%>&Action=Search" method="POST" id="frm" name="frm">
	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-9">
				<h4 class="text-left" data-toggle="modal" data-target="#ModalSelectBatch"  Title="Click to Select an NA File to View"><%=Session("CardTypeNA")%> - NA File To Diners <%="File Load ID: " & strNextNAFileNumber%> </h4>
			</div>
			<div class="col-md-3 float-right">
				<button type="button" class="btn btn-primary" onclick="window.open('NAToDinersExportExcel.asp?CardType=<%=Session("CardTypeNA")%>')"><i class="fa fa-file"></i> Export NA File to Excel</button>
				<!--<button type="button" class="btn btn-primary float-right" onClick='window.location="NATransactionsDPC.asp?Action=ExportNA"'><i class="fa fa-file"></i> Export NA File</button>-->
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
		'strWhere = " AND [FileSeqNum] = '" & strNextNAFileNumber & "'" 
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
	
	'Update to make sure that the complete recordset is not returned when no filters are selected as it takes a long time to load
	If strWhere ="" OR IsNull(strWhere) Then strTop = "Top 50"
	
If strSearch = "" OR ISNull(strSearch) Then
	strSQL = "SELECT " & strTOP & " * FROM qryCAPSNAToDinersDetail WITH(NOLOCK) WHERE [NAToDinersID] > 0 " & strWhere & strSort
Else
	strSQL = "SELECT " & strTOP & " * FROM qryCAPSNAToDinersDetail WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
End If

'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = strSearch
Else
	'strRecordMessage = "for " & Session("UserName") 
End If
response.write strSQL

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
		
		
	
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Displaying " & lngTotalRecords & " of " & lngTotalRecords & " NA File to Diners records </div></div>"
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""NATransactionsDPC.asp?Sort=NAToDinersID&SortType=" & strOrderType & """> NA ID<i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""> Record Type <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""NATransactionsDPC.asp?Sort=Status&SortType=" & strOrderType & """> Status <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Record Text</th>" & _
			"<th scope=""col""><a href=""NATransactionsDPC.asp?Sort=BatchNumber&SortType=" & strOrderType & """> Batch No. <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""NATransactionsDPC.asp?Sort=DateUpdated&SortType=" & strOrderType & """> Date Updated <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""NATransactionsDPC.asp?Sort=UpdatedBy&SortType=" & strOrderType & """> Updated By </th>" & _
			"<th scope=""col""><a href=""NATransactionsDPC.asp?Sort=EmployeeID&SortType=" & strOrderType & """> Employee ID  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""NATransactionsDPC.asp?Sort=ApplicationID&SortType=" & strOrderType & """> Application ID  <i class=""fa fa-sort""></i></a></th>" & _
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
					
					strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='NATransactionsDPC.asp?Action=CancelNA&NAToDinersID=" & objrs("NAToDinersID") & "&NAEID=" & objRS("EmployeeID") & "&Status=Deleted'""; title=""Click to Remove from NA File""><i class=""fa fa-times""></i> Remove</button>"
					
				ElseIf objRS("Status") = "Deleted" Then
					strStatus = "<span class=""badge badge-pill badge-danger"">Deleted</span>"
					
					strAction = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='NATransactionsDPC.asp?Action=CancelNA&NAToDinersID=" & objrs("NAToDinersID") & "&NAEID=" & objRS("EmployeeID") & "&Status=Added To NA'""; title=""Click to Add to NA File""><i class=""fa fa-plus""></i> Add</button>"
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
					
	
		objRS.movenext
	Loop

	
	'Write the End of the table and divs for the above list, as the pagination (below) is in it's own container
	Response.Write "</tbody></table></div>"

	
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
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='NATransactionsDPC.asp?ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='NATransactionsDPC.asp?ViewButton=WaitingExport';""><i class=""fa fa-cogs""></i> View Awaiting Export</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='NATransactionsDPC.asp?ViewButton=NoResponse';""><i class=""fa fa-thumbs-down""></i> View No Response</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='NATransactionsDPC.asp?ViewButton=Response';""><i class=""fa fa-thumbs-up""></i> View Response</button>" & _
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

Public Sub ExportNAFile()
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

If Session("CardTypeNA") = "DTC" Then
	strFileNameStart = GetSystemAdmin("NAFileStart")
	strNextFileNumber = GetSystemAdmin("NAFileNumber")
Else
	strFileNameStart = GetSystemAdmin("NAFileStartDPC")
	strNextFileNumber = GetSystemAdmin("NAFileNumberDPC")
End If

'Pad the number out to 6 digits
strNextFileNumber = PadDigits(strNextFileNumber,6)
strNextFileNumber = PadDigits(strNextFileNumber,6)

'---UPDATE APRIL 2023 for the new DPC change --- File for Diners cannot have the date and time at the end
'Compile the File name and path from the variables above
strFilePath = strFilePath & "\Admin\CAPSAdmin\Attachments\Diners\DinersTo\" & strFileNameStart & ".txt"
'strFilePath = strFilePath & "\Admin\CAPSAdmin\Attachments\Diners\DinersTo\" & strFileNameStart & PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & ".txt"

'Set the filename to be used when moving the file in the procedure called at the end of this procedure
'****UPDATE April 2023 ---- Removed the Date stamp at the end of the file name, as per Diners requirements for the DPC NA file ------ ******
strFileName = strFileNameStart 
'strFileName = strFileNameStart & PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & ".txt"

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
	''Updated April 2023 for DPC change --added Card Type
	objRS.Open "SELECT * FROM tblCAPSNAToDiners WITH(NOLOCK) WHERE CardType = '" & Session("CardTypeNA") & "' AND BatchNumber = 0 AND Status = 'Added To NA'",objCon

		If objRS.EOF Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No records to write to the NA File.</div>"
		Else
			Set objTextStream = objFSO.OpenTextFile(strFilePath, fsoForWriting, True)
			'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
			lngFileLoadID = SaveFileLoadID ("NAFileDPC",strFileNameStart & strFileDateTime & ".txt", strFilePath,-1,0,0,0,0,0,0,0,strFileDateTime,strNextFileNumber,"Exported",Session("UserID"),"N")
			
			objTextStream.WriteLine PadSpaceLeft("H NA" & strFileDateTimeSec & strNextFileNumber,428)
			'------Original Line below before DPC Diners changes Feb 2023
			'objTextStream.WriteLine "H NA" & strFileDateTimeSec & strNextFileNumber
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
				Call ExportNARecord (objRS("NAToDinersID"),lngFileLoadID,intRecordCount,Session("CardTypeNA"))
				'Call ExportNARecord (objRS("NAToDinersID"),strNextFileNumber,intRecordCount)
				
			objRS.Movenext
			Loop
			
			strRecordCount = PadDigits(intRecordCount,6)
			
			objTextStream.WriteLine PadSpaceLeft("T " & strRecordCount,428)
			'------Original Line below before DPC Diners changes Feb 2023
			'objTextStream.WriteLine "T" & strRecordCount
			'objTextStream.WriteLine "3 T" & strRecordCount			
			
			'Call the procedure to update the System Parameter NAFileNumber. Increment the Number by 1.
			
			Call EmailNARecord(PadDigits(strNextFileNumber,6))
			Call UpdateBatchNumber(strNextFileNumber)
		
			'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			'strFileName = strFileNameStart & strFileDateTime & ".txt" ------Set at the top of the procedure now as it requires 2 digit year -- 4 digit year required within the file as header
			Call UpdateFileLoadSummary ("NAFileDPC",strNextFileNumber, strFileName, lngFileLoadID)
			
			Response.Write "<div class=""alert alert-success"" role=""alert"">NA File " & strFileName & ".txt" & " ADDED to the NA file export folder!</div>"
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
		
		strSQL = "UPDATE tblCAPSSystemParameters SET [ParameterValue] = '" & lngBatchNumber & "' WHERE [ParameterName] = 'NAFileNumberDPC'"

		objCon.Execute strSQL
	
	Else
		
		Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR! NA File Batch Number: " & lngBatchNumber & " is not a number. See System Admin.</div>"
		
	End If

End Sub

Public Sub LoadBatchList()
'Description:	Loads all Batch Numbers to a list for selecting and searching/filtering


	objRS.Open "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = 'NAFileDPC' AND [Deleted] = 'N' ORDER By [FileSeqNum] DESC",objCon
  
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
	
	If objFSO.FileExists(strFilePathTo) Then
		
		strFileExtension = objFSO.GetExtensionName(strFileName)
		strFileNameNoExt = Left(strFilePathTo,Len(strFilePathTo)-Len(strFileExtension))
		'strFileNameNoExt = Left(strFilePathTo,Len(strFilePathTo)-Len(strFileExtension)-1)
		
		'Set the file extenstion to text if it doesn't have one
		If IsNull(strFileExtension) or strFileExtension = "" Then strFileExtension = "txt"
		
		'Set the first FileTo Name to check if it exists below
		'strFilePathTo = strFileNameNoExt & x & "." & strFileExtension
		
		'response.write "<BR>strFileName=" & strFileName
		For x = 1 to 10
			If objFSO.FileExists(strFilePathTo) Then
				strFilePathTo = strFileNameNoExt & x + 1 & "." & strFileExtension
				'strFilePathTo = strServer & strFileNameNoExt & x & "." & strFileExtension
				Response.Write "<BR>strFilePathTo NO=" & strFilePathTo
			Else
				Response.Write "<BR>strFilePathTo before NO=" & strFilePathTo
				strFilePathTo = strFileNameNoExt & x & "." & strFileExtension
				'strFilePathTo = strServer & strFileNameNoExt & x & strFileExtension
				'strFilePathFrom = strFilePathOnly & strFileName & x & strFileExtension
				response.write "<BR>strFilePathTo=" & strFilePathTo
				response.write "<BR>strFileExtension=" & strFileExtension
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
	Response.Write "<div class=""alert alert-success"" role=""alert"">NA File exported FROM: " & strFilePathFrom & "</div>"
	Response.Write "<div class=""alert alert-success"" role=""alert"">NA File exported TO: " & strFilePathTo & "</div>"
	
		'Move the file to the Loaded folder
		'objFSO.MoveFile strFilePathOnly & strFileName,strServer & strFileName
		If Right(strFilePathTo,4)<>".txt" Then strFilePathTo = strFilePathTo & ".txt"
		
		objFSO.MoveFile strFilePathFrom,strFilePathTo'strServer & strFileName
		
		
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
