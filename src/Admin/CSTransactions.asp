
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
	
	Response.Write "Card Type = " & strCardType
	
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

function ChangeBatch(CardType) {

	var e = document.getElementById("CSFileSelect");
	var result = e.options[e.selectedIndex].text;
	
	self.location='CSTransactions.asp?CardType=' + CardType + '&FileLoadID='+result
}

function OpenExcelReport() {
	var strExcel = document.getElementById('WhereClause').value;

	window.open('../CC/ExcelExport.asp?tbl=qryCAPSCSFromDinersAuditLog&W=' + strExcel + '')
	//window.open('../CC/ExcelExport.asp?tbl=qryCAPSTrainingReportExport&W=' + strExcel + '&Top=100')
	//window.open('../CC/ExcelExport.asp?tbl=qryCAPSTrainingReport&Top=100')

}

function ChangePage(CardType) {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	
	var e = document.getElementById("PageCombo");
	var result = e.options[e.selectedIndex].value;
	
	self.location = 'CSTransactions.asp?CardType=' + CardType + '&PageCombo=' + result;
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
  xhttp.open("GET", "../CC/AJAX/GetCSFromDinersAudit.asp?CSFromDinersID=" + varID + "", true);
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
     <div class="modal-dialog modal-large modal-dialog-centered modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  CS From Diners File Detail
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
        <div class="col-md-4"><SELECT class="form-control" onChange="ChangeBatch('<%=strCardType%>');" name="CSFileSelect" id="CSFileSelect"><% Call LoadBatchList()%></Select></div><br><br>
      </div>
      <div class="modal-footer">
		<button type="button" class="btn btn-primary" onClick="self.location='CSTransactions.asp?CardType=" & strCardType & "&FileLoadID=0'"><i class="fa fa-times"></i> Clear</button>
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
  <form action="CSTransactions.asp?CardType=<%=strCardType%>&Action=Search" method="POST" id="frm" name="frm">
	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-10">
				<!--<h4 class="text-left">CS File From Diners <%="File Load ID: " & Session("FileLoadID")%></h4>-->
				<h4 class="text-left" data-toggle="modal" data-target="#ModalSelectBatch"  Title="Click to Select a CS File to View"><%=strCardType%> - CS File From Diners <%="File Load ID: " & Session("FileLoadID")%>&nbsp;&nbsp;- Date Loaded : <%=Get_Batch_Date(Session("FileLoadID"))%></h4>
			</div>
			<div class="col-md-2">
				<button type="button" class="btn btn-outline-success" onClick="OpenExcelReport();"><i class="fa fa-file-excel"></i> Export To Excel</button>
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
Dim strWhere2 
Dim strFileSeqNum
Dim strFileDateTime

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

	'Set the second part of the Where statement
	If Session("ViewButton") = "Processed" Then
		strWhere = " AND CardTypeCard = '" & strCardType & "' AND [Status] = 'Processed' "
	ElseIf Session("ViewButton") = "NoChange" Then
		strWhere = " AND CardTypeCard = '" & strCardType & "' AND [AuditLogID] IS NULL "
	ElseIf Session("ViewButton") = "Change" Then
		strWhere = " AND CardTypeCard = '" & strCardType & "' AND [AuditLogID] IS NOT NULL "
	Else
		'This catches ALL
		strWhere = ""
	End If
	
	'Set the first part of the Where statement based on the File Load selected and add the second part of the where
	If Session("FileLoadID") = "0" Then
		If strSearch = "" OR ISNull(strSearch) Then
			If strWhere = "" Then
				strWhere = " WHERE CardTypeCard = '" & strCardType & "' AND [CSFromDinersID] = 0 "
			Else
				strWhere = " WHERE CardTypeCard = '" & strCardType & "' AND " & Right(strWhere,Len(strWhere)-4) & " AND (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
			End If
		Else
			strWhere = " WHERE CardTypeCard = '" & strCardType & "' AND (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
		End If
	Else
		If strSearch = "" OR ISNull(strSearch) Then
			strWhere = " WHERE CardTypeCard = '" & strCardType & "' AND FileSeqNum = '" & Session("FileLoadID") & "'" & strWhere
		Else
			strWhere = " WHERE CardTypeCard = '" & strCardType & "' AND FileSeqNum = '" & Session("FileLoadID") & "'" & strWhere & " AND (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
		End If
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
	
If strSearch = "" OR ISNull(strSearch) Then

	strSQL = "SELECT " & strTOP & " * FROM qryCAPSCSFromDinersAuditLog WITH(NOLOCK) " & strWhere & strSort
	'strSQL = "SELECT " & strTOP & " * FROM qryCAPSCSFromDinersAuditLog WITH(NOLOCK) WHERE FileSeqNum = '" & Session("FileLoadID") & "' AND [CSFromDinersID] > 0 " & strWhere & strSort
Else
	strSQL = "SELECT " & strTOP & " * FROM qryCAPSCSFromDinersAuditLog WITH(NOLOCK) " & strWhere & strSort
	'strSQL = "SELECT " & strTOP & " * FROM qryCAPSCSFromDinersAuditLog WITH(NOLOCK) WHERE FileSeqNum = '" & Session("FileLoadID") & "' AND (EIDNo Like '%" & strSearch & "%' OR GivenNames Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
End If

Response.Write strSQL

'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = strSearch
Else
	'strRecordMessage = "for " & Session("UserName") 
End If

'Build the Where statement used for Excel (which just contains the existing WHERE with the additionals not in the dynamic variable)
strWhere2 = "WHERE FileSeqNum = '" & Session("FileLoadID") & "' " & strWhere

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
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No CS From Diners records for " & strRecordMessage & "</TH>" & _
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
		
		strPageCombo = "<SELECT ID=""PageCombo"" Name=""PageCombo"" onChange=""ChangePage('" & strCardType & "');"">" & strPageCombo & "</select>"
		
		'If the PageCombo is not numeric (ALL or Null) then make it the total records for the recordset (which is set above)
		If NOT IsNumeric(Session("PageCombo")) Then Session("PageCombo") = lngTotalRecords
	
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Displaying " & Session("PageCombo") & " of " & lngTotalRecords & " CS From Diners records (" & lngStartingRecord & " to " & lngStartingRecord + clng(Session("PageCombo")) & ")</span><span class=""panel-subheader"" style=""float:right;"">Number of records per page: " & strPageCombo  & "</span></div></div>"
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""CSTransactions.asp?CardType=" & strCardType & "&Sort=CSFromDinersID&SortType=" & strOrderType & """> CS ID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""CSTransactions.asp?CardType=" & strCardType & "&Sort=EIDNo&SortType=" & strOrderType & """> EID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""CSTransactions.asp?CardType=" & strCardType & "&Sort=Surname&SortType=" & strOrderType & """> Name <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""CSTransactions.asp?CardType=" & strCardType & "&Sort=CardType&SortType=" & strOrderType & """> Card Status <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Card No.</th>" & _
			"<th scope=""col""><a href=""CSTransactions.asp?CardType=" & strCardType & "&Sort=Status&SortType=" & strOrderType & """> Status  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""CSTransactions.asp?CardType=" & strCardType & "&Sort=Type&SortType=" & strOrderType & """> Change Type <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" title=""Card Type""> Card Type Sub </th>" & _
			"<th scope=""col""><a href=""CSTransactions.asp?CardType=" & strCardType & "&Sort=ChangeDetails&SortType=" & strOrderType & """> Change Details  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">File Date</th>" & _
			"</tr></thead><tbody class=""text-left"">"
					
	End If
	
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No CS From Diners records " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	End If
    
	x = 0
	
    Do until objRS.EOF 

		y = y + 1
		
		'Only write the first 50 records from the starting position
		If y <= lngStartingRecord + clng(Session("PageCombo")) AND y >= lngStartingRecord - clng(Session("PageCombo")) Then
		'If y <= lngStartingRecord + 50 AND y >= lngStartingRecord - 50 Then
		
			x = x + 1
			
			'Create the actions based on the Process Status of the card
'			Select Case objRS("ProcessStatus")
'			
'			Case  "Removed Unactivated"
'				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='CSTransactions.asp?Action=UnRemove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-list""></i> Re-List</button>"
'			Case "Added to CS"
'
'				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportCS.asp?CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
'				
'			Case "Email Unactivated"
'				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CSTransactions.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
'			
'			Case Else
'				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='CSTransactions.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
'				strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#EmailModal""><i class=""fa fa-minus-mail""></i> Email</button>"
'				strAction = strAction & "<button type=""button"" class=""btn btn-outline-info btn-xs"" onclick=""self.location='CSTransactions.asp?Action=Remove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-cross""></i> Remove</button>"
'
'				'strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='CSTransactions.asp?Action=Email&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-mail""></i> Email</button>"
'				'data-toggle="modal" data-target="#EmailModal"
'				
'			End Select
			
			'Create the Status list badge based on the status field
			If IsNull(objRS("Status")) Then
				strStatus = ""
			Else
				If objRS("Status") = "Processed" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Processed</span>"
				ElseIf objRS("Status") = "Imported" Then
					strStatus = "<span class=""badge badge-pill badge-warning"">Imported</span>"
				Else
					strStatus = objRS("Status")
				End If
			End If
			
			If IsNull(objRS("FileDateTime")) or objRS("FileDateTime") = "" Then			
				strFileDateTime = ""
			Else
				If Len(objRS("FileDateTime"))>10 Then
					strFileDateTime = Mid(objRS("FileDateTime"),7,2) & "/" & Mid(objRS("FileDateTime"),5,2) & "/" & Left(objRS("FileDateTime"),4)
				Else
					strFileDateTime = ""
				End If
			End If
			
			
			Response.Write "<TR><TD ><a data-toggle=""modal"" data-target=""#CSFromDinersMod"" HREF=""#"" onClick=""loadDoc(" & objRS(0) & ")"">" & objRS(0) & "</a></TD><TD>" & objRS("EIDNo") & "</a></TD>" & _
					"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""../CC/CardDetail.asp?CardID=" & objRS(0) & """>" & trim(objRS("GivenNames")) & " " & trim(objRS("Surname")) & "</a></TD><TD><a Target=""_self"" HREF=""../CC/CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("CardStatus") & "</a></TD>" & _
					"<TD >" & MaskCard(objRS("CardNumber")) & "</TD><TD >" & strStatus & "</TD>" & _
					"<TD style=""font-size:12px; background-color:#e6eeff;"">" & objRS("Type") & "</TD><TD style=""font-size:12px; background-color:#e6eeff;"">" & objRS("SubType") & "</TD><TD style=""font-size:12px; background-color:#e6eeff;"">" & objRS("ChangeDetails") & "</TD>" & _
					"<TD style=""font-size:12px;"">" & strFileDateTime & "</TD></TR>"
					
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
		Response.Write "<TR><TH colspan=""10"">Total <input type=""HIDDEN"" id=""WhereClause"" name=""WhereClause"" value=""" & strWhere2 & """ ></TH>" & _
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
						strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""CSTransactions.asp?CardType=" & strCardType & "&StartingRecord=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
					End If
					
					'Add the Elipsis (...) to the start of the page numbers if there is more than 20 pages and the current place is beyond the first page
					If x = 0 AND lngPage > 1 AND y > 20 Then
						strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""CSTransactions.asp?CardType=" & strCardType & "&StartingRecord=" & lngTotalRecords - (clng(Session("PageCombo"))*20) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
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
				
					'strPages = strPages & "<a href=""CSTransactions.asp?StartingRecord=" & (x * clng(Session("PageCombo"))) & """> " & x + 1 & " </a>"
					strPages2 = strPages2 & "<li class=""page-item " & strActive & """><a class=""page-link"" href=""CSTransactions.asp?CardType=" & strCardType & "&StartingRecord=" & (x * clng(Session("PageCombo"))) & """>" & x & "</a></li>"
					
					'strPages = strPages & "<a href=""CSTransactions.asp?StartingRecord=" & (x * 50) & """> " & x & " </a>"
					'strPages2 = strPages2 & "<li class=""page-item " & strActive & " -" & x & "|" & lngPage & """><a class=""page-link"" href=""CSTransactions.asp?Previous&StartingRecord=" & lngStartingRecord -50 & """>" & x & "</a></li>"
					End If
				End If
			Next
			
		End If
	End If
	
	'If y > 0 Then
	'	Response.Write "<TR><TH colspan=""9"" style=""text-align:center;""><a href=""CSTransactions.asp?Previous&StartingRecord=" & lngStartingRecord -50 & """>Previous Page " & strPages & " <a href=""CSTransactions.asp?Previous&StartingRecord=" & lngStartingRecord + 50 & """> Next Page</TH></TR>"
	'End If
	
	
	'Write the End of the table and divs for the above list, as the pagination (below) is in it's own container
	Response.Write "</tbody></table></div>"

	'Write the Pagination objects for all pages based on the total records and the number records displayed on screen
	If y > 0 Then
		
		Response.Write "<div class=""container""><div class=""row""><div class=""col-12 text-center"">" & _
			"<nav aria-label=""Page navigation""><ul class=""pagination""><li class=""page-item"">" & _      
			"<a class=""page-link"" href=""CSTransactions.asp?CardType=" & strCardType & "&StartingRecord=0"" aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
			strPages2 & _
			"<li class=""page-item"">" & _
			"<a class=""page-link"" href=""CSTransactions.asp?CardType=" & strCardType & "&StartingRecord=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
			"</a></li></ul></nav></div></div></div>"
			
			'"<a class=""page-link"" href=""CSTransactions.asp?StartingRecord=" & lngStartingRecord - clng(Session("PageCombo")) & """ aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
	End If
		
objRS.Close

End Sub

Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(4)

If Session("ViewButton") = "Processed" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "NoChange" Then
	arrButton(3) = "active"
ElseIf Session("ViewButton") = "Change" Then
	arrButton(4) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='CSTransactions.asp?CardType=" & strCardType & "&ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='CSTransactions.asp?CardType=" & strCardType & "&ViewButton=Processed';""><i class=""fa fa-cogs""></i> View Processed</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='CSTransactions.asp?CardType=" & strCardType & "&ViewButton=NoChange';""><i class=""fa fa-thumbs-down""></i> View No Change</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='CSTransactions.asp?CardType=" & strCardType & "&ViewButton=Change';""><i class=""fa fa-thumbs-up""></i> View Changes</button>" & _
				"</div>"

End Sub

Public Sub ExportCSFile()
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

'Get the Default file location for the server then add the filepath and name for the NA File
strFilePath = GetSystemAdmin("ServerFilePath")
strFileNameStart = GetSystemAdmin("CSFileStart")
strNextFileNumber = GetSystemAdmin("CSFileNumberTo")

'Pad the number out to 6 digits
strNextFileNumber = PadDigits(strNextFileNumber,6)
strNextFileNumber = PadDigits(strNextFileNumber,6)

'Compile the File name and path from the variables above
strFilePath = strFilePath & "\Admin\CAPSAdmin\Attachments\Diners\DinersTo\" & strFileNameStart & PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & ".txt"

strFileDateTime = PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2)
strFileDateTimeSec = PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & Hour(Now()) & Minute(Now()) & Second(Now())
'response.write 	"lngBatchNumber=" & strNextFileNumber & " strNextCSFileNumber= " &  strNextCSFileNumber

'Get the FileLoad details
'lngBatchNumber = GetFileLoadID("NAFile",strNextFileNumber,"")

'If IsNull(lngBatchNumber) OR lngBatchNumber = "" then
'	lngBatchNumber = strNextCSFileNumber
'End If

Dim objFSO
Dim strFileName
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

'Open the text file
Dim objTextStream


	'Open a recordset of all of the NA File records yet to be exported
	
	objRS.Open "SELECT * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE FileSeqNum Is Null AND Status = 'Awaiting Export'",objCon

		If objRS.EOF Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No records to write to the CS File.</div>"
		Else
		
			'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
			lngFileLoadID = SaveFileLoadID ("CSToDiners",strFileNameStart & strFileDateTime & ".txt", strFilePath,-1,0,0,0,0,0,0,0,strFileDateTime,strNextFileNumber,"Exported",Session("UserID"),"N")
			Set objTextStream = objFSO.OpenTextFile(strFilePath, fsoForWriting, True)
			objTextStream.WriteLine "H CS" & strFileDateTimeSec & strNextFileNumber
			
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
					objTextStream.WriteLine "D " & PadSpaceLeft(objRS("EIDNo"),10) & PadDigits(objRS("CardNo"),19) & PadSpaceLeft(objRS("CardUpdateInd"),2) & PadSpaceLeft(objRS("CardExpiryDate"),8) & PadSpaceLeft(objRS("CardStatus"),2) & PadSpaceLeft(objRS("Title"),12) & PadDigitsLeft(objRS("Surname"),25) & PadSpaceLeft(objRS("GivenNames"),30) & PadSpaceLeft(objRS("NameOnCard"),26) & PadSpaceLeft(objRS("Address1"),40) & PadSpaceLeft(objRS("Address2"),40) & PadSpaceLeft(objRS("Address3"),40) & PadSpaceLeft(objRS("Suburb"),25) &  PadSpaceLeft(objRS("State"),4) & PadSpaceLeft(objRS("Postcode"),12) &  PadSpaceLeft(objRS("HomePhone"),12) &  PadSpaceLeft(objRS("WorkPhone"),12) & PadSpaceLeft(objRS("MobilePhone"),12) & PadSpaceLeft(objRS("Email"),70) & PadSpaceLeft(objRS("ReportGroup"),8) &  PadSpaceLeft(objRS("CreditLimit"),11)                         
				End If
				'Call the procedure to update each record as exported once added to the CS File -- USE the File style Batch Number not FileLoadID
				'Call ExportCSRecord (objRS("CSToDinersID"),lngFileLoadID,intRecordCount)
				Call ExportCSRecord (objRS("CSToDinersID"),strNextFileNumber,intRecordCount)
				
			objRS.Movenext
			Loop
			
			strRecordCount = PadDigits(intRecordCount,6)
			
			objTextStream.WriteLine "T" & strRecordCount		
			
			
			'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			strFileName = strFileNameStart & strFileDateTime & ".txt"
			
			Call UpdateFileLoadSummary ("CSToDiners",strNextFileNumber, strFileName, lngFileLoadID)
			'Call the procedure to update the System Parameter CSFileNumber. Increment the Number by 1.
			Call UpdateBatchNumber(strNextFileNumber)
			
			If strError = "" Then
				Response.Write "<div class=""alert alert-success"" role=""alert"">CS File " & "AUDC_INTOECS_DODNA_D" & strFileDateTime & ".txt" & " ADDED to the CS file export folder!</div>"
			Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">CS File " & "AUDC_INTOECS_DODNA_D" & strFileDateTime & ".txt" & " has errors : " & strError & "</div>"
			End If
				'Close the file and clean up
				objTextStream.Close
		End If

	objRS.Close
	


Set objTextStream = Nothing
Set objFSO = Nothing



End Sub


Public Sub ExportCSRecord(lngCSToDinersID,lngBatchNumber,x)
'Procedure to Change the Status of CS file records being exported and adds an Audit Log record
Dim intRecord

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

Public Sub RemoveCSRecord(lngNAToDinersID, strEmployeeID, strStatus)
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
				Response.Write "<div class=""alert alert-danger"" role=""alert"">Application for " & strEmployeeID & " NOT Removed from CS File! An Error has occurred. See System Admin with CS File ID: " & lngNAToDinersID & " </div>"
			Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">Application for " & strEmployeeID & " NOT Added to the CS File! An Error has occurred. See System Admin with CS File ID: " & lngNAToDinersID & " </div>"
			End If
		Else
			If strStatus = "Deleted" Then
				Response.Write "<div class=""alert alert-success"" role=""alert"">Application for " & strEmployeeID & " REMOVED from the CS file!</div>"
			Else
				Response.Write "<div class=""alert alert-success"" role=""alert"">Application for " & strEmployeeID & " ADDED to the CS file!</div>"
			End If
		End If
		
	
End Sub

Public Sub LoadBatchList()
'Description:	Loads all Batch Numbers to a list for selecting and searching/filtering


	objRS.Open "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = 'CSFromDiners' AND [Deleted] = 'N' ORDER By [FileSeqNum] DESC",objCon
	'objRS.Open "SELECT * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE [FileType] = 'CSToDiners' AND [Deleted] = 'N' ORDER By [FileSeqNum] DESC",objCon
  
	Response.write "<OPTION value=""0"">Select a Batch to View...</OPTION>"
	
		Do Until objRS.EOF 
			
			Response.write "<OPTION value=""" & objRS("FileLoadID") & """>" & objRS("FileSeqNum") & "</OPTION>"
			
			objRS.Movenext
			
		Loop
	
	objRS.Close
	
End Sub

Public Function Get_Batch_Date(strFileLoadID)

	objRS2.Open "SELECT * FROM tblCAPSFileLoad WHERE FileType = 'CSFromDiners' AND FileSeqNum = '" & strFileLoadID & "'",objCon

	If Not objRS2.EOF Then
	
		Get_Batch_Date = objRS2("DateLoaded")
	
	End If
	
	objRS2.Close

End Function


Public Function CheckForNull(strValue,strField,intRow)

	If IsNull(strValue) Then 
		CheckForNull = " Error in field " & strField & " is Null at row " & intRow & " : "
	Else
		CheckForNull = ""
	End If

End Function

Set objRS2 = Nothing
Set objRS = Nothing
Set objCon = Nothing
%>
