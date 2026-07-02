
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
	

		
		
		If Request.QueryString("Action") = "Email" Then
			Call EmailCard()
			
		End If
		
		

	
	'Get the Number of records to display in the table if it has changed
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

setTimeout( 'ShowTimeoutWarning();', 1080000 );

function ShowTimeoutWarning () {     
    window.alert( "********** Warning! **********' \n \n 'You will be automatically logged out in 2 minutes unless you change screens, Close or Save!" ); 
}

function OpenExcelReport() {
	var strExcel = document.getElementById('WhereClause').value;
	
	//window.open('ExcelExport.asp?tbl=qryCAPSCardsExpiry')
	window.open('ExcelExport.asp?tbl=qryCAPSCardsExpiry&W=' + strExcel + '')
	//window.open('../CC/ExcelExport.asp?tbl=qryCAPSTrainingReportExport&W=' + strExcel + '&Top=100')
	//window.open('../CC/ExcelExport.asp?tbl=qryCAPSTrainingReport&Top=100')

}

function ChangePage() {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("PageCombo");
	var result = e.options[e.selectedIndex].value;
	
	self.location = 'ExpiringCards.asp?PageCombo=' + result;
	//alert(result);
	//document.getElementById('CardType').value=result;
	
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
</script>
<script js>
$(function(){
    $('#myFormSubmit').click(function(e){
      e.preventDefault();
      $('#formResults').text($('#frm').serialize());
      
      $.post('ApplicationsEmployeeHF.asp?Action=SubmitApp', 
         $('#frm').serialize(), 
      /*   function(data, status, xhr){
           // do something here with response;
         });
      */
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


			   
<!-- End the first part of the Header Container -->
<div id='tbl-container'>
  <form action="ExpiringCards.asp?Action=Search" method="POST" id="frm" name="frm">
	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-10">
				<h4 class="text-left">Expiring Cards <%="As at: " & FormatDateTime(now(),2)%></h4>
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
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </section>
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
Dim strEmailSent
Dim dteWarningDate
Dim strPages
Dim strSort
Dim strOrderType
Dim strWhere
Dim lngStartingRecord
Dim lngTotalRecords

Dim strPageCombo
Dim arrPagecombo(6)
Dim strTop

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
	
	If Session("ViewButton") = "Emailed" Then
		strWhere = " AND [EmailSent] = 'Email Expiring' "
	ElseIf Session("ViewButton") = "Removed" Then
		strWhere = " AND [EmailSent] = 'Removed Expiring' "
	ElseIf Session("ViewButton") = "AddedToCS" Then
		strWhere = " AND [EmailSent] = 'Added To CS' "
	Else
		'This catches ALL
		strWhere = ""
	End If
	
	'Build the TOP Statement
	If Session("PageCombo") = "" Or IsNull(Session("PageCombo")) Then
		Session("PageCombo") = 50
	End If
	
	If IsNumeric(Session("PageCombo")) Then
		strTOP = " TOP " & Session("PageCombo")
	Else
		strTOP = ""
	End If
	
	'Determine the Daye to Expire being displayed
	strWhere = strWhere & " AND ([DaysToExpiry] > -30 AND [DaysToExpiry] < 120) "
	
If strSearch = "" OR ISNull(strSearch) Then
	'If Session("UserView") = "All" Then
		strSQL = "SELECT " & strTOP & " * FROM qryCAPSCardExpiryEmails WITH(NOLOCK) WHERE (Status='00' OR Status ='') " & strWhere & strSort
	'Else
	'	strSQL = "SELECT top 100 * FROM qryCAPSCardExpiryEmails WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND EmployeeID = '" & Session("EmployeeID") & "'"
	'End If
	
Else
	'If Session("UserView") = "All" Then
		strSQL = "SELECT " & strTOP & " * FROM qryCAPSCardExpiryEmails WITH(NOLOCK) WHERE (Status='00' OR Status ='') AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
	'Else
	'	strSQL = "SELECT top 100 * FROM qryCAPSCardExpiryEmails WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
	'End If
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
	
	'Write a message in the list if there are no Expiring Cards
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Expiring Cards for " & strRecordMessage & "</TH>" & _
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
                  "<span class=""panel-subheader"">Displaying " & Session("PageCombo") & " of " & lngTotalRecords & " expiring cards (" & lngStartingRecord & " to " & lngStartingRecord + clng(Session("PageCombo")) & ")</span><span class=""panel-subheader"" style=""float:right;"">Number of records per page: " & strPageCombo  & "</span></div></div>"
		
		
		
		'Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
         '         "<span class=""panel-subheader"">Displaying 50 of " & lngTotalRecords & " expiring cards (" & lngStartingRecord & " to " & lngStartingRecord + 50 & ")</span></div></div>"
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=ApplicationID&SortType=" & strOrderType & """> Card ID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=EmployeeID&SortType=" & strOrderType & """> EID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=Surname&SortType=" & strOrderType & """> Name <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=CardType&SortType=" & strOrderType & """> Card Type <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Card No.</th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=Status&SortType=" & strOrderType & """> Card Status  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=Expiry&SortType=" & strOrderType & """> Expiry Date <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" title=""Days until Card Expires (up to today)""> Days </th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=EmailSent&SortType=" & strOrderType & """> Email Sent  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Address Status</th>" & _
			"<th scope=""col"">Action</th>" & _
			"</tr></thead><tbody class=""text-left"">"
					
	End If
	
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Expiring Cards " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	End If
    
	x = 0 
	y = 0 
	
    Do until objRS.EOF 

		y = y + 1
		
		'Only write the first 50 records from the starting position
		'If y <= lngStartingRecord + 50 AND y >= lngStartingRecord - 50 Then
		
			x = x + 1	
			

'			
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
			
			If IsNull(objRS("Expiry")) Then
				dteDateSubmitted = ""
			Else
				dteDateSubmitted = FormatDateTime(objRS("Expiry"),vbShortDate)
			End If
			
			If IsNull(objRS("Expiry")) Then
				dteDateReviewed = ""
			Else
				dteDateReviewed = DateDiff("d",now(),objRS("Expiry"))
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
			
			'If IsNull(objRS("EmailSent")) or objRS("EmailSent") = "" Then
				'strEmailSent = ""
			'Else
				'If objRS("EmailSent") = "Email Expiring" Then
					'strEmailSent = "Emailed"
				'Elseif objRS("EmailSent") = "Removed Expiring" Then
					'strEmailSent = "Removed"
				'Elseif objRS("EmailSent") = "Added to CS" Then
					'strEmailSent = "Added To CS"
				'End If
					strEmailSent = objRS("EmailSentDate")
			'End If
			
			dteWarningDate = "title=" & objRS("EmailSentDate") & ""

Dim strAddressStatus

			If objRS("PostalMessage") = "OK" Then
				
				strAddressStatus  = "<span class=""badge badge-pill badge-success"">Valid</span>"
			Else
				strAddressStatus  = "<span class=""badge badge-pill badge-danger"">Invalid</span>"
			End If
				
			strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Email&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-mail"" ></i> Email</button>"
	
	
			
			response.write "<TR><TD ><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("FirstName") & " " & objRS("Surname") & "</a></TD><TD><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
					"<TD >" & strCardNo & "</TD><TD >" & strStatus & "</TD>" & _
					"<TD >" & dteDateSubmitted & "</TD><TD " & strDaysColour & ">" & dteDateReviewed & "</TD><TD style=""font-size:12px;"" " & dteWarningDate & ">" & strEmailSent & "</TD><TD>" & strAddressStatus & "</TD>" & _
					"<TD>" & strAction & "</TD></TR>"
					
		
		'End If
		
		objRS.movenext
	Loop
	
	If y > 0 Then
		Response.Write "<TR><TH colspan=""10"">Total <input type=""HIDDEN"" id=""WhereClause"" name=""WhereClause"" value=""WHERE (Status='00' OR Status ='') " & strWhere & """ ></TH>" & _
				"<TH colspan=""3"" style=""text-align:center;"">" & x & "</TH></TR>"
	End If
	
				
objRS.Close

End Sub




Public Sub DisplayTableDetails_COPY()
Dim y
Dim strAction
Dim strStatus
Dim dteDateSubmitted
Dim dteDateReviewed
Dim strSearch
Dim strRecordMessage
Dim strCardNo
Dim strDaysColour
Dim strEmailSent
Dim dteWarningDate
Dim strPages
Dim strSort
Dim strOrderType
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
	
	If Session("ViewButton") = "Emailed" Then
		strWhere = " AND [EmailSent] = 'Email Expiring' "
	ElseIf Session("ViewButton") = "Removed" Then
		strWhere = " AND [EmailSent] = 'Removed Expiring' "
	ElseIf Session("ViewButton") = "AddedToCS" Then
		strWhere = " AND [EmailSent] = 'Added To CS' "
	Else
		'This catches ALL
		strWhere = ""
	End If
	
	'Determine the Daye to Expire being displayed
	strWhere = strWhere & " AND ([DaysToExpiry] > -30 AND [DaysToExpiry] < 120) AND (Status='00' OR Status ='') "
	
If strSearch = "" OR ISNull(strSearch) Then
	'If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 1000 * FROM qryCAPSCardsExpiry WITH(NOLOCK) WHERE [Status] = '00'" & strWhere & strSort
	'Else
	'	strSQL = "SELECT top 100 * FROM qryCAPSCardsExpiry WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND EmployeeID = '" & Session("EmployeeID") & "'"
	'End If
	
Else
	'If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 1000 * FROM qryCAPSCardsExpiry WITH(NOLOCK) WHERE [Status] = '00' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
	'Else
	'	strSQL = "SELECT top 100 * FROM qryCAPSCardsExpiry WITH(NOLOCK) WHERE [ActivationFlag] = 'N' AND EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
	'End If
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
	
	'Write a message in the list if there are no Expiring Cards
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Expiring Cards for " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	Else
		objRS.Movelast
		objRS.Movefirst
		lngTotalRecords = objRS.Recordcount
		
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Displaying 50 of " & lngTotalRecords & " expiring cards (" & lngStartingRecord & " to " & lngStartingRecord + 50 & ")</span></div></div>"
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=ApplicationID&SortType=" & strOrderType & """> Card ID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=EmployeeID&SortType=" & strOrderType & """> EID <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=Surname&SortType=" & strOrderType & """> Name <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=CardType&SortType=" & strOrderType & """> Card Type <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Card No.</th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=Status&SortType=" & strOrderType & """> Status  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=Expiry&SortType=" & strOrderType & """> Expiry Date <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"" title=""Days until Card Expires (up to today)""> Days </th>" & _
			"<th scope=""col""><a href=""ExpiringCards.asp?Sort=EmailSent&SortType=" & strOrderType & """> Process Status  <i class=""fa fa-sort""></i></a></th>" & _
			"<th scope=""col"">Action</th>" & _
			"</tr></thead><tbody class=""text-left"">"
					
	End If
	
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Expiring Cards " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	End If
    	
    Do until objRS.EOF 

		y = y + 1
		
		'Only write the first 50 records from the starting position
		If y <= lngStartingRecord + 50 AND y >= lngStartingRecord - 50 Then
		
			x = x + 1
			
			'Create the actions based on the Process Status of the card
			Select Case objRS("EmailSent")
			
			Case  "Removed Expiring"
				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=UnRemove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-list""></i> Re-List</button>"
			Case "Added to CS"

				strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportCS.asp?CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
				
			Case "Email Expiring"
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";> eeCancel</button>"
				'strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> ffCancel</button>"
			Case Else
				strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";> ggCancel</button>"
				'strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> hhCancel</button>"
				strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#EmailModal""><i class=""fa fa-minus-mail""></i> Email</button>"
				strAction = strAction & "<button type=""button"" class=""btn btn-outline-info btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Remove&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-cross""></i> bbRemove</button>"

				'strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='ExpiringCards.asp?Action=Email&CardID=" & objrs("CardID") & "&CardEID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-minus-mail""></i> Email</button>"
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
			
			If IsNull(objRS("Expiry")) Then
				dteDateSubmitted = ""
			Else
				dteDateSubmitted = FormatDateTime(objRS("Expiry"),vbShortDate)
			End If
			
			If IsNull(objRS("Expiry")) Then
				dteDateReviewed = ""
			Else
				dteDateReviewed = DateDiff("d",now(),objRS("Expiry"))
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
			
			If IsNull(objRS("EmailSent")) or objRS("EmailSent") = "" Then
				strEmailSent = ""
			Else
				If objRS("EmailSent") = "Email Expiring" Then
					strEmailSent = "Emailed"
				Elseif objRS("EmailSent") = "Removed Expiring" Then
					strEmailSent = "Removed"
				Elseif objRS("EmailSent") = "Added to CS" Then
					strEmailSent = "Added To CS"
				End If
			End If
			
			dteWarningDate = "title=" & objRS("Warning") & ""
			
			response.write "<TR><TD ><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:12px;""><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("FirstName") & " " & objRS("Surname") & "</a></TD><TD><a Target=""_self"" HREF=""CardDetail.asp?CardID=" & objRS(0) & """>" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
					"<TD >" & strCardNo & "</TD><TD >" & strStatus & "</TD>" & _
					"<TD >" & dteDateSubmitted & "</TD><TD " & strDaysColour & ">" & dteDateReviewed & "</TD><TD style=""font-size:12px;"" " & dteWarningDate & ">" & strEmailSent & "</TD>" & _
					"<TD>" & strAction & "</TD></TR>"
					
			
		End If
		
		objRS.movenext
	Loop
	
	If y > 0 Then
		Response.Write "<TR><TH colspan=""10"">Total <input type=""HIDDEN"" id=""WhereClause"" name=""WhereClause"" value=""" & strWhere & """ ></TH>" & _
				"<TH colspan=""3"" style=""text-align:center;"">" & x & "</TH></TR>"
	End If
	
	'Create the number of pages
	If IsNumeric(y) Then
		If y > 1 Then
			
			y = y / 50
			
			For x = 1 to y
				strPages = strPages & "<a href=""ExpiringCards.asp?StartingRecord=" & (x * 50) & """> " & x & " </a>"
			Next
			
		End If
	End If
	
	If y > 0 Then
		Response.Write "<TR><TH colspan=""9"" style=""text-align:center;""><a href=""ExpiringCards.asp?Previous&StartingRecord=" & lngStartingRecord -50 & """>Previous Page " & strPages & " <a href=""ExpiringCards.asp?Previous&StartingRecord=" & lngStartingRecord + 50 & """> Next Page</TH></TR>"
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
	
	strSQL = "UPDATE tblCAPSCard SET EmailSent = 'Added to CS', Warning = '" & Left(now(),20) & "' WHERE CardID = " & Session("CardID") & ""
	
	objCon.Execute strSQL
	
	'Call the procedure to save a message for the card
	Call SaveMessage("Card Cancelled from Expiring Cards List")
	
	'Call the function to save the Audit Log record
	Call SaveAuditLog(0,"Expiring Card","Card Cancelled",Request.QueryString("CardEID"),strCardType,strCardNo,Session("UserID"),"","","","Card Cancelled in Expiring Card screen",Request.QueryString("CardID"),0,0,0)
	
End Sub

Public Sub RemoveCard()
Dim strCardNo
Dim strCardType

	strSQL = "UPDATE tblCAPSCard SET EmailSent = 'Removed Expiring', Warning = '" & Left(now(),20) & "' WHERE CardID = " & Session("CardID") & ""
	
	objCon.Execute strSQL
	
	'Call the procedure to save a message for the card
	Call SaveMessage("Card Removed from Expiring Cards List")
	
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
	Call SaveAuditLog(0,"Expiring Card","Card Cancelled",Request.QueryString("CardEID"),strCardType,strCardNo,Session("UserID"),"","","","Card Removed from list in Expiring Card screen",Request.QueryString("CardID"),0,0,0)
	
End Sub


Public Sub EmailCard()
Dim strCardNo
Dim strCardType

	strSQL = "spCAPSResendCardExpiryEmails " & Request.QueryString("CardID") & "," & Session("UserID") & ",''"

	'Response.Write strSQL & "<BR><BR>"
	
	objCon.Execute strSQL
	
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


Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(4)

If Session("ViewButton") = "Emailed" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "Removed" Then
	arrButton(3) = "active"
ElseIf Session("ViewButton") = "AddedToCS" Then
	arrButton(4) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If



End Sub

Set objRS = Nothing
Set objCon = Nothing
%>
