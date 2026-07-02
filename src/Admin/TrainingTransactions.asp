
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

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

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

function ChangePage() {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("PageCombo");
	var result = e.options[e.selectedIndex].value;
	
	self.location = 'TrainingTransactions.asp?PageCombo=' + result;
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
  <form action="TrainingTransactions.asp?Action=Search" method="POST" id="frm" name="frm">
	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-10">
				<h4 class="text-left">Training from CAMPUS <%="File Load ID: " & Session("FileLoadID")%></h4>
			</div>
			<div class="col-md-2">
				<!--<button type="button" class="btn btn-primary" onClick='window.location="ApplicationsSubmit.asp"'><i class="fa fa-plus"></i> New Application</button>-->
			</div>
			
		</div>

          <div class="row py-2">
            <div class="col-md-9">
              <%'Call LoadViewButtons()%>
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
Dim dteDateUpdated
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
		strWhere = " AND [Loaded] = 'Y' "
	ElseIf Session("ViewButton") = "NoChange" Then
		strWhere = " AND [AuditLogID] IS NULL "
	ElseIf Session("ViewButton") = "Change" Then
		strWhere = " AND [AuditLogID] IS NOT NULL "
	Else
		'This catches ALL
		strWhere = ""
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

	strSQL = "SELECT " & strTOP & " * FROM qryCAPSTraining WITH(NOLOCK) WHERE [TrainingID] > 0 " & strWhere & strSort
Else
	strSQL = "SELECT " & strTOP & " * FROM qryCAPSTraining WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR LastName Like '%" & strSearch & "%')" & strWhere & strSort
End If

'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = strSearch
Else
	'strRecordMessage = "for " & Session("UserName") 
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
		
		strPageCombo = "<SELECT ID=""PageCombo"" Name=""PageCombo"" onChange=""ChangePage();"">" & strPageCombo & "</select>"
		
		'If the PageCombo is not numeric (ALL or Null) then make it the total records for the recordset (which is set above)
		If NOT IsNumeric(Session("PageCombo")) Then Session("PageCombo") = lngTotalRecords
	
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Displaying " & Session("PageCombo") & " of " & lngTotalRecords & " Training records (" & lngStartingRecord & " to " & lngStartingRecord + clng(Session("PageCombo")) & ")</span><span class=""panel-subheader"" style=""float:right;"">Number of records per page: " & strPageCombo  & "</span></div></div>"
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""TrainingTransactions.asp?Sort=CSFromDinersID&SortType=" & strOrderType & """> Training ID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""TrainingTransactions.asp?Sort=EmployeeID&SortType=" & strOrderType & """> Employee ID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""TrainingTransactions.asp?Sort=Surname&SortType=" & strOrderType & """> Name <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""TrainingTransactions.asp?Sort=email&SortType=" & strOrderType & """> Email <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""TrainingTransactions.asp?Sort=CompletionDate&SortType=" & strOrderType & """>Completion Date <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" title=""Card Type""> Date Updated </th>" & _
			"<th scope=""col""><a href=""TrainingTransactions.asp?Sort=ChangeDetails&SortType=" & strOrderType & """> Updated By  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" title=""Card Type""> Updated </th>" & _
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

		y = y + 1
		
		'Only write the first 50 records from the starting position
		If y <= lngStartingRecord + clng(Session("PageCombo")) AND y >= lngStartingRecord - clng(Session("PageCombo")) Then
		'If y <= lngStartingRecord + 50 AND y >= lngStartingRecord - 50 Then
		
			x = x + 1
			
			'Create the actions based on the Process Status of the card
'			Select Case objRS("ProcessStatus")
'			
'			Case  "Removed Unactivated"
'				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='TrainingTransactions.asp?Action=UnRemove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-list""></i> Re-List</button>"
'			Case "Added to CS"
'
'				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportCS.asp?CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
'				
'			Case "Email Unactivated"
'				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='TrainingTransactions.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
'			
'			Case Else
'				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='TrainingTransactions.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
'				strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#EmailModal""><i class=""fa fa-minus-mail""></i> Email</button>"
'				strAction = strAction & "<button type=""button"" class=""btn btn-outline-info btn-xs"" onclick=""self.location='TrainingTransactions.asp?Action=Remove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-cross""></i> Remove</button>"
'
'				'strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='TrainingTransactions.asp?Action=Email&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-mail""></i> Email</button>"
'				'data-toggle="modal" data-target="#EmailModal"
'				
'			End Select
			
			'Create the Loaded list badge based on the Loaded field
			If IsNull(objRS("Loaded")) Then
				strStatus = ""
			Else
				If objRS("Loaded") = "Processed" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Processed</span>"
				ElseIf objRS("Loaded") = "Imported" Then
					strStatus = "<span class=""badge badge-pill badge-warning"">Imported</span>"
				Else
					strStatus = objRS("Loaded")
				End If
			End If
			
			If IsNull(objRS("DateUpdated")) or objRS("DateUpdated") = "" Then
				dteDateUpdated = ""
			Else
				dteDateUpdated = FormatDateTime(objRS("DateUpdated"),vbShortDate)
			End If
			
			response.write "<TR><TD ><a data-toggle=""modal"" data-target=""#CSFromDinersMod"" HREF=""#"" onClick=""loadDoc(" & objRS("TrainingID") & ")"">" & objRS("TrainingID") & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:13px;""><a Target=""_self"" HREF=""../CC/CardDetail.asp?CardID=" & objRS(0) & """>" & trim(objRS("FirstName")) & " " & trim(objRS("LastName")) & "</a></TD>" & _
					"<TD style=""font-size:12px;"">" & objRS("Email") & "</TD><TD >" & objRS("CompletionDate") & "</TD>" & _
					"<TD >" & dteDateUpdated & "</TD><TD style=""font-size:12px;"">" & objRS("UpdatedByName") & "</TD><TD style=""font-size:12px; background-color:#e6eeff;"">" & objRS("UpdatedBy") & "</TD>" & _
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
						strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""TrainingTransactions.asp?StartingRecord=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
					End If
					
					'Add the Elipsis (...) to the start of the page numbers if there is more than 20 pages and the current place is beyond the first page
					If x = 0 AND lngPage > 1 AND y > 20 Then
						strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""TrainingTransactions.asp?StartingRecord=" & lngTotalRecords - (clng(Session("PageCombo"))*20) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
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
				
					'strPages = strPages & "<a href=""TrainingTransactions.asp?StartingRecord=" & (x * clng(Session("PageCombo"))) & """> " & x + 1 & " </a>"
					strPages2 = strPages2 & "<li class=""page-item " & strActive & """><a class=""page-link"" href=""TrainingTransactions.asp?StartingRecord=" & (x * clng(Session("PageCombo"))) & """>" & x & "</a></li>"
					
					'strPages = strPages & "<a href=""TrainingTransactions.asp?StartingRecord=" & (x * 50) & """> " & x & " </a>"
					'strPages2 = strPages2 & "<li class=""page-item " & strActive & " -" & x & "|" & lngPage & """><a class=""page-link"" href=""TrainingTransactions.asp?Previous&StartingRecord=" & lngStartingRecord -50 & """>" & x & "</a></li>"
					End If
				End If
			Next
			
		End If
	End If
	
	'If y > 0 Then
	'	Response.Write "<TR><TH colspan=""9"" style=""text-align:center;""><a href=""TrainingTransactions.asp?Previous&StartingRecord=" & lngStartingRecord -50 & """>Previous Page " & strPages & " <a href=""TrainingTransactions.asp?Previous&StartingRecord=" & lngStartingRecord + 50 & """> Next Page</TH></TR>"
	'End If
	
	
	'Write the End of the table and divs for the above list, as the pagination (below) is in it's own container
	Response.Write "</tbody></table></div>"

	'Write the Pagination objects for all pages based on the total records and the number records displayed on screen
	If y > 0 Then
		
		Response.Write "<div class=""container""><div class=""row""><div class=""col-12 text-center"">" & _
			"<nav aria-label=""Page navigation""><ul class=""pagination""><li class=""page-item"">" & _      
			"<a class=""page-link"" href=""TrainingTransactions.asp?StartingRecord=0"" aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
			strPages2 & _
			"<li class=""page-item"">" & _
			"<a class=""page-link"" href=""TrainingTransactions.asp?StartingRecord=" & lngTotalRecords - clng(Session("PageCombo")) & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
			"</a></li></ul></nav></div></div></div>"
			
			'"<a class=""page-link"" href=""TrainingTransactions.asp?StartingRecord=" & lngStartingRecord - clng(Session("PageCombo")) & """ aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
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
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='TrainingTransactions.asp?ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='TrainingTransactions.asp?ViewButton=Processed';""><i class=""fa fa-cogs""></i> View Processed</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='TrainingTransactions.asp?ViewButton=NoChange';""><i class=""fa fa-thumbs-down""></i> View No Change</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='TrainingTransactions.asp?ViewButton=Change';""><i class=""fa fa-thumbs-up""></i> View Changes</button>" & _
				"</div>"

End Sub

Set objRS = Nothing
Set objCon = Nothing
%>
