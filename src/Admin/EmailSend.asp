
<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #Include file=../CC/CAPSFunctions.asp -->
<!-- #Include file=../ADOVBS.inc -->
<%

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

'Description:	Email Template Management Screen
'Author:		AB
'Date:			October 2020

	Response.Expires = -1500	

Dim objCon
Dim objRS
Dim objRS1
Dim objRS2
Dim objCmd

Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL

Dim strEmailSubject
Dim strCampaignSelected
Dim strCampaignSelected2

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objRS1 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")
	
    objCon.Open Session("DBConnection")	

    Session("CurrentPage") = "Admin/EmailError.asp"

	If IsNull(Session("EmailCampaignID")) OR Session("EmailCampaignID") = "" Then Session("EmailCampaignID")= 0

	If Not IsEmpty(Request.QueryString("StyleSheet")) Then
		Session("StyleSheet") = Request.QueryString("StyleSheet")
		
		Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""" & Request.QueryString("StyleSheet") & """>"
	Else
		If IsEmpty(Session("StyleSheet")) Then Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""../CAPSStyle.css"">"
	End If

	Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""../CAPSStyle.css"">"

	If Not IsEmpty(Request.QueryString("ViewButton")) Then
		Session("ViewButton") = Request.QueryString("ViewButton")
	End If
	
	If Not IsEmpty(Request.QueryString("EmailCampaignID")) Then
		Session("EmailCampaignID") = Request.QueryString("EmailCampaignID")
		Session("EmailSubject") = Request.QueryString("EmailSubject")
	End If
	
	If Not IsEmpty(Request.QueryString("EmailSubject")) Then
		Session("EmailSubjectSelected") = Request.QueryString("EmailSubject")
	End If
	
'Execute Action
If Request.QueryString("Action") = "SendCampaignEmails" Then
	Call SendEmails() 
End If

'Call the Delete function if the Delete button in the Delete Modal has been clicked
If Request.QueryString("Action") = "Delete" Then   
    Call DeleteData(Request.QueryString("EmailID"))
End If

	'Call the procedure to Load the Selected Email Campaign Details
	Call LoadSelectedEmailCampaign()
	
	
	'Call the procedure to Load the Logged in User's Email address for testing emails
	Call LoadUserEmail()
	'response.write Session("EmailSubject")
%>

<html>
<head>
<script LANGUAGE="javascript">
	function triggerModal(ModalID) {
				
		//var ModaRecipientFirstName = '#emailModal' + ModalID
		var ModaRecipientFirstName = '#emailModal'
		$(ModaRecipientFirstName).modal("show");

	}

	setTimeout( 'ShowTimeoutWarning();', 1080000 );

function ShowTimeoutWarning () {     
    window.alert( "********** Warning! **********' \n \n 'You will be automatically logged out in 2 minutes unless you change screens, Close or Save!" ); 
}

function DeleteModalClose(cb) {
   
	document.getElementById("ModalDelete").style.display = "none";
        
}  


function SaveEmailTemplate() {
	
	var id = document.getElementById('EmailErrorMsgID').value
	self.location = "EmailError.asp?Action=Save&EmailErrorMsgID=" + id;

}

function SearchUsers() {
	
	var id = document.getElementById('SearchInput').value
	self.location = "EmailSend.asp?SearchInput=" + id;

}

function SendEmails(varTest,varSubject) {
	
	self.location = "EmailSend.asp?Action=SendCampaignEmails&Mode="+varTest+"&Subject="+varSubject;

}

function ChangeSubject() {
	
	var e = document.getElementById("EmailSubjectSelect");
	var SubjectID = e.options[e.selectedIndex].value;
	
	self.location = "EmailSend.asp?EmailSubject=" + SubjectID;

}

function loadDocE(cb) {

  var id = cb.getAttribute('data-id');
  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("GetUserDetail").innerHTML = this.responseText;
    }
  };

  xhttp.open("GET", "../CC/AJAX/GetUserDetail.asp?UserID=" + id, true);
  xhttp.send();
}

function loadDelete(cb) {


var id = cb.getAttribute('data-id');
var name = cb.getAttribute('data-EmployeeName');

	document.getElementById("ModalDeleteMessage").innerHTML = 'Remove ' + name + ' from Email Campaign?';
	document.getElementById("ModalDelete").style.display = "block";
	document.getElementById("EmployeeDeleteID").value = id;
	
}

function deleteEmployee(cb) {

	var id = document.getElementById("EmployeeDeleteID").value
	//var id = cb.getAttribute('data-id');
	
	self.location = "EmailSend.asp?Action=Delete&EmailID=" + id;

}

	jQuery(document).ready(function($) {
    $(".clickable-row").click(function() {
        window.location = $(this).data("href");
    });
});
</script>

</head>
<body>

	<main class="main py-3">
		<div class="container">
		
<!-- Select Batch Number Modal -->
<div class="modal fade" id="ModalSendEmail" tabindex="-1" role="dialog" aria-labelledby="ModalSendEmail" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalSendEmailTitle" style="font-weight:bold;">Send Emails?</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
	  <div class="modal-header">
	  <h6 style="color:navy;">Email Campaign:</h6>
	  </div>
      <div class="modal-body" id="SendEmailMod">
			<div class="row"><div class="col-md-12">
				Send Emails for Email Campaign:<span style="font-weight:bold;""> <%=strCampaignSelected2%></span>?</br></br>
				<b>Clicking 'TEST'</b> -- All Emails will be sent to the Test recipient only: <span style="font-weight:bold;""><%=Session("UserEmail")%></span></br></br>
				<b>Clicking 'YES'</b> -- All Emails will be sent to the REAL recipients!</br></br>
				<span style="font-style:italic; font-size:12px;"">A maximum of 3,000 emails will be sent in any one Campaign each time "Yes" is clicked</span>
			</div></div>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
		<button type="button" class="btn btn-primary" onClick="SendEmails('Test','<%=strEmailSubject%>');" Title="Send ALL Emails to your account and not the intended recipients. For testing."><i class="fa fa-check"></i> TEST</button>
        <button type="button" class="btn btn-primary" onClick="SendEmails('Prod','<%=strEmailSubject%>');" ><i class="fa fa-check"></i> Yes</button>
		<input type="hidden" id="NewStatus" name="NewStatus" value=""/>
      </div>
    </div>
  </div>
</div>
<!-- End Select Batch Number Modal -->


	<!-- Start Delete Modal -->

	<div class="modal fade" id="ModalDelete" tabindex="-1" role="dialog" aria-labelledby="ModalDeleteCenterTitle" aria-hidden="true">
		<div class="modal-dialog modal-dialog-centered" role="document">
		  	<div class="modal-content">
				<div class="modal-header">
			  		<h5 class="modal-title" id="ModalDeleteLongTitle" style="font-weight:bold;">Delete Employee from Email Campaign</h5>
			  		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				</div>
				<div class="modal-body">			  
				  	<div class="row">
						<div class="col-md-12 mb-3">			
					 		<div class="alert alert-danger" role="alert" id="AlertDanger" style="display:block">
						  		<div id="ModalDeleteMessage"></div>
					  		</div>
						</div>
				  	</div>
				  	<div class="row">
					  	<div class="col-md-12 mb-3" style="text-align:right;">
						  	<input type="hidden" id="EmployeeDeleteID"></input>
						  	<button class="btn btn-primary btn-sm" onClick="deleteEmployee(this)"><i class="fa fa-check"></i> Yes</button>
							<button type="button" class="btn btn-secondary btn-sm" data-dismiss="modal" onClick="DeleteModalClose(this);"><i class="fa fa-times"></i> No</button>
					  	</div>
				  	</div>
			  	</div>
			</div>	
		</div>
		<div class="modal-footer"></div>
	</div>

	<!-- End Delete Modal -->

	<!-- Start Edit Modal -->
	<form action="EmailSend.asp?Action=Save" method="POST" id="frm" name="frm" class="needs-validation" novalidate>
		<div class="modal fade" id="emailModal" tabindex="-1" role="dialog" aria-labelledby="emailModalTitle" aria-hidden="true">
			<div class="modal-dialog modal-dialog-centered" role="document">
				<div class="modal-content">
					<div class="modal-header">
						<h5 class="modal-title" id="emailModalLabel">User Administration <i class="fa fa-chevron-right"></i></h5><button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					</div>
					<div class="modal-body">
						<div class="col-md-12">
							<table class="table table-bordered table-hover CAPS">						
								<div id="GetUserDetail"></div>						
							</table>
						</div>	
					</div>
					<div class="modal-footer"></div>
				</div>
			</div>
		</div>
	</form>
	<!-- End Edit Modal -->
	
<!-- End the first part of the Header Container -->

	
	
		  <div class="row mb-2">
			<div class="col-md-8">
			  <h2>Email Send Screen</h2>
			</div>
			<div class="col-md-4 text-right">
			  <!--<button class="btn btn-primary" data-toggle="modal" data-id="0" data-target="#emailModal" data-id="0" onClick="loadDocE(this);">
				<i class="fa fa-plus"></i> Create new email
			  </button>-->
			</div>
		  </div>
		  
		  <div class="row py-0">
            <div class="col-md-3">
				
              <%Call LoadActionButtons()%>
				</div>
				<div class="col-md-6">
				<%=strCampaignSelected%>
			  </div>
           
			<div class="col-md-3">
				<div class="form-group has-search">
					<span class="fa fa-search form-control-feedback"></span>
				 <input type="text" class="form-control" type="search" id="SearchInput" name="SearchInput" placeholder="Search Users by Keyword" onChange="SearchUsers();" value="<%=Request.QueryString("SearchInput")%>"/>
				 </div>
			</div>
          </div>
		  
		<div class="row py-0">
			<div class="col-md-3">
				  <% Call LoadCampaign() %>

			</div>
            <div class="col-md-9">
				  <% Call DisplayTable() %>

			</div>
        </div>

	
</main>

	



<!-- #Include file=../cc/CAPSFooter.asp -->

</body>
</html>

<%

Public Sub DisplayTable()

'Displays the main table with User details and results of searches
'Dim y
Dim strStatus
Dim strSelected
Dim intRecordCount

Dim strSearch
Dim strSort
Dim strWhere
Dim strRecordMessage
Dim lngStartingPage
Dim lngCurrentPage

Dim lngTotalPages
Dim lngTotalRecords
Dim arrPagecombo(6)
Dim strPageCombo
Dim strOrderType
Dim arrSort(8)
Dim strSortArrow
Dim arrNames
Dim strRecipientLastNameSearch
Dim strRecipientFirstNameSearch
Dim strSearchDate
Dim strSearchDay
Dim strSearchMonth
Dim strSearchYear
Dim strDateFrom
Dim strDateTo
Dim bolSkip
'Dim strStatus
Dim strPages2
Dim strRemove
Dim strStatusDisplay
Dim strDateCreated

strSearch = Replace(Request.QueryString("SearchInput"), "'", "''")
	
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
	If IsNull(strSort) Or strSort = "" Then strSort = " ORDER BY EmployeeID ASC"
	
	If Session("ViewButton") = "Admin" Then
		strWhere = " AND [UserTypeID] > 9 "
	ElseIf Session("ViewButton") = "New" Then
		strWhere = " AND [DateUpdated] > CONVERT(DATETIME, '" & DateAdd("d", -30, now()) & "', 103) "
		'CONVERT(DATETIME, '2020-12-20 00:00:00', 102)
	Else
		'This catches ALL
		strWhere = ""
	End If

	If Not IsNull(Session("EmailCampaignID")) Then
		strWhere = strWhere & " AND [EmailCampaignID] = " & Session("EmailCampaignID") & ""
	End If
	
If strSearch = "" OR ISNull(strSearch) Then
	'If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 100 * FROM tblCAPSEmail WITH(NOLOCK) WHERE [Active]='Y' " & strWhere & strSort
	'Else
	'	strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' " & strSort
	'End If
	
	
	'''UPDATED the SQL query to use the Subject to identify unique records
	strSQL = "SELECT top 100 * FROM qryCAPSEmail WITH(NOLOCK) WHERE [Status] = 'Created' AND [Active] = 'Y' AND [EmailCampaignID] = " & Session("EmailCampaignID") & " AND [EmailSubject] = '" & Session("EmailSubject") & "'"
	'response.write strsql
Else
	'If the user has entered the date lookup then process this, otherwise perform the EmployeeID and Name searches
	If Left(strSearch,2) = "d:" Then
	
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

			'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE (DateSubmitted > '" & strDateFrom & "' AND DateSubmitted < '" & strDateTo & "')" & strWhere & strSort
			strSQL = "SELECT TOP 100 * FROM tblCAPSEmail WITH(NOLOCK) WHERE (DateUpdated > '" & strDateFrom & "' AND DateUpdated < '" & strDateTo & "')" & strWhere & strSort
			
		Else
			'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE (DateSubmitted = '" & strSearchDate & "')" & strWhere & strSort
			strSQL = "SELECT TOP 100 * FROM tblCAPSEmail WITH(NOLOCK) WHERE (DateUpdated = '" & strSearchDate & "')" & strWhere & strSort
		End If
		
	Else
		'If Session("UserView") = "All" Then
			'If the user has entered a search term with a space then assume this is a first and last name so search on that only
			If Instr(1,strSearch," ")>0 Then
				arrNames = Split(strSearch," ")
				strRecipientLastNameSearch = arrNames(0)
				strRecipientFirstNameSearch = arrNames(1)
				
				'strWhere = " WHERE ([FirstName] Like '%" & strRecipientLastNameSearch & "%' AND [Surname] Like '%" & strRecipientFirstNameSearch & "%')"
				'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR ([FirstName] Like '%" & strRecipientLastNameSearch & "%' AND [Surname] Like '%" & strRecipientFirstNameSearch & "%'))" & strWhere & strSort
				strSQL = "SELECT TOP 100 * FROM tblCAPSEmail WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR (RecipientLastName Like '%" & strRecipientLastNameSearch & "%' AND RecipientFirstName Like '%" & strRecipientFirstNameSearch & "%')) " & strSort
			Else
				'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strWhere & strSort
				strSQL = "SELECT TOP 100 * FROM tblCAPSEmail WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR RecipientLastName Like '%" & strSearch & "%' OR RecipientFirstName Like '%" & strSearch & "%') " & strSort
				
			End If
		'Else
			'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strSort
		'	strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR RecipientLastName Like '%" & strSearch & "%' OR RecipientFirstName Like '%" & strSearch & "%') " & strSort
		'End If
	
	'End of 
	End If
	
End If

'Get the Sort Order for the sort order fontawesome image for the selected sort field
'For x = 1 to 8

	Select Case Request.QueryString("Sort")
	
		Case "EmployeeID"
			arrSort(1) = strSortArrow
		Case "RecipientLastName"
			arrSort(2) = strSortArrow
		Case "RecipientFirstName"
			arrSort(3) = strSortArrow
		Case "EmailToAddress"
			arrSort(4) = strSortArrow
		Case "UserTypeID"
			arrSort(5) = strSortArrow
		Case "Status"
			arrSort(6) = strSortArrow
		Case "SentByUser"
			arrSort(7) = strSortArrow
		Case "DateUpdated"
			arrSort(8) = strSortArrow
	End Select

'Next

'Build the message displayed at the bottom of the screen with the search details
'If Session("UserView") = "All" Then
	strRecordMessage = strSearch
'Else
'	strRecordMessage = Session("UserName") 
'End If

	'y = 0

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
	
	
	'Write a message in the list if there are no Users
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Email Campaign Selected " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	Else
		objRS.Movelast
		objRS.Movefirst
		lngTotalRecords = objRS.Recordcount
		
		objRS.AbsolutePage = lngStartingPage
		
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
                  "<span class=""panel-subheader""> Displaying " & Session("PageCombo") & " of " & lngTotalRecords & " users (" & (clng(lngStartingPage)*clng(Session("PageCombo")))-clng(Session("PageCombo"))+1 & " to " & clng(lngStartingPage)*clng(Session("PageCombo")) & ")</span><span class=""panel-subheader"" style=""float:right;"">Number of records per page: " & strPageCombo  & "</span></div></div>"
		
	
		'Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
        '          "<span class=""panel-subheader"">Displaying 50 of " & lngTotalRecords & " applications (" & lngStartingPage & " to " & lngStartingPage + 50 & ")</span></div></div>"
		
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""EmailSend.asp?Sort=EmployeeID&Link=AD&SortType=" & strOrderType & """> Employee ID <i class=""fa fa-sort" & arrSort(1) & """></i></a></th>" & _
			"<th scope=""col""><a href=""EmailSend.asp?Sort=RecipientLastName&Link=AD&SortType=" & strOrderType & """> Recipient <i class=""fa fa-sort" & arrSort(3) & """></i></a></th>" & _
			"<th scope=""col""><a href=""EmailSend.asp?Sort=EmailToAddress&Link=AD&SortType=" & strOrderType & """> Email To <i class=""fa fa-sort" & arrSort(4) & """></i></a></th>" & _
			"<th scope=""col""><a href=""EmailSend.asp?Sort=UserTypeID&Link=AD&SortType=" & strOrderType & """> Subject <i class=""fa fa-sort" & arrSort(5) & """></i></a></th>" & _
			"<th scope=""col""><a href=""EmailSend.asp?Sort=Status&Link=AD&SortType=" & strOrderType & """> Status  <i class=""fa fa-sort" & arrSort(6) & """></i></a> <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Card Process Status - the stage the application is at.""></i></th>" & _
			"<th scope=""col""><a href=""EmailSend.asp?Sort=SentByUser&Link=AD&SortType=" & strOrderType & """> Sent By <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
			"<th scope=""col""><a href=""EmailSend.asp?Sort=DateUpdated&Link=AD&SortType=" & strOrderType & """> Date Sent <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
			"<th scope=""col"">Action</th>" & _
			"</tr></thead><tbody class=""text-left"">"
				
	End If
    
	x = 0
	
	
'strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) Order By RecipientFirstName ASC"
intRecordCount = 0

If IsEmpty(Session("UserID")) Then Session("UserID") = 0 End If

'objRS.Open strSQL,objCon   
    	
	Do until objRS.EOF	
	
		If IsNull(objRS("Status")) OR objRS("Status") = "" Then
			strRemove = ""
		Else
			If objRS("Status") = "Created" Then
				strRemove = "<button class=""btn btn-primary btn-xs btn-danger"" data-toggle=""modal"" data-target=""#ModalDelete"" data-id=""" & objRS("EmailID") & """ data-EmployeeName=""" & objRS("RecipientLastName") & " " & objRS("RecipientFirstName") & """ onClick=""loadDelete(this);""><i class=""fa fa-times""></i> Remove</button>"
			Else
				strRemove = ""
			End If
		End If
		
		If IsNull(objRS("DateCreated")) OR objRS("DateCreated") = "" Then
			strDateCreated = ""
		Else
			strDateCreated = "Title=""Date Created: " & objRS("DateCreated") & """"
		End If
		
		'Determine the pill display based on the Status
		If IsNull(objRS("Status")) or objRS("Status") = "" Then
			strStatusDisplay = "<span class=""badge badge-info"">None</span>"
		Else
			If objRS("Status") = "Created" Then
				strStatusDisplay = "<span class=""badge badge-secondary"">" & objRS("Status") & "</span>"
			ElseIf objRS("Status") = "Deleted" Then
				strStatusDisplay = "<span class=""badge badge-danger"">" & objRS("Status") & "</span>"
			Else
				strStatusDisplay = "<span class=""badge badge-success"">" & objRS("Status") & "</span>"
			End If
		End If
			
	Response.Write "<tr><td style=""font-size:12px;"">" & objRS("EmployeeID") & "</td>" & _
		"<td style=""font-size:12px;"">" & objRS("RecipientFirstName") & " " & objRS("RecipientLastName") & "</td>" & _
		"<td style=""font-size:12px;"">" & objRS("EmailToAddress") & "</td>" & _
		"<td style=""font-size:12px;"">" & objRS("EmailSubject") & "</td>" & _
		"<td style=""font-size:12px;"" " & strDateCreated & " >" & strStatusDisplay & "</td>" & _ 
		"<td style=""font-size:12px;"">" & objRS("SentByUser") & "</td>" & _
		"<td style=""font-size:12px;"">" & objRS("DateSent") & "</td>" & _
		"<td style=""font-size:12px;"">" & strRemove & "</td></tr>"
	
	intRecordCount = intRecordCount + 1
	
	objRS.movenext
	
	Loop	
	
	
	If intRecordCount > 0 Then
		Response.Write "<TR><TH colspan=""7"">Total</TH>" & _
				"<TH colspan=""2"" style=""text-align:center;"">" & intRecordCount & "</TH></TR></tbody></table></div></div>"
				     
	End If

	bolSkip = False
	
	For x = 1 to lngTotalPages
		
		'Determine which page number is Status (displayed as Status)
		If clng(x) = clng(lngCurrentPage) Then
			strStatus = "Status"
		Else
			strStatus = ""
		End If
		
		If lngTotalPages > 20 Then
			If x > 1 Then
				'Add the Elipsis (...) to the start of the page numbers if there is more than 20 pages and the current place is beyond the first page
			'	If lngTotalPages > 20 Then
			'		strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""Applications.asp?StartingPage=" & lngTotalRecords - (clng(Session("PageCombo"))*20) & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
			'	End If
			End If
		
			If x > 20 Then
				If bolSkip = False Then
					strPages2 = strPages2 & "<li class=""page-item""><a class=""page-link"" href=""Applications.asp?StartingPage=" & lngTotalPages & """ aria-label=""More""><span aria-hidden=""true"">&hellip;</span><span class=""sr-only"">More</span></a></li>"
					bolSkip = True
				End If
			Else
				
			End If
		End If
		
		If bolSkip = True Then
		Else
			strPages2 = strPages2 & "<li class=""page-item " & strStatus & """><a class=""page-link"" href=""Applications.asp?StartingPage=" & x & """>" & x & "</a></li>"
		End If
		
	Next
	
	'Write the Pagination objects for all pages based on the total records and the number records displayed on screen
	If lngTotalPages > 0 Then
					
		Response.Write "<div class=""container""><div class=""row""><div class=""col-12 text-center"">" & _
			"<nav aria-label=""Page navigation""><ul class=""pagination""><li class=""page-item"">" & _      
			"<a class=""page-link"" href=""Applications.asp?StartingPage=1"" aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" & _
			strPages2 & _
			"<li class=""page-item"">" & _
			"<a class=""page-link"" href=""Applications.asp?StartingPage=" & lngTotalPages & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
			"</a></li></ul></nav></div></div></div>"

	End If
	
objRS.Close

End Sub

Sub SaveUser()

	Dim intModalID

	intModalID = Request.Form("ModalSaveID")

  		With objCmd

			.CommandType = 4
			.CommandText = "spUserSave"

			.Parameters.Append objCmd.CreateParameter("UserID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmailToAddress", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("RecipientLastName", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("RecipientFirstName", adVarChar, adParamInput, 50)              
			.Parameters.Append objCmd.CreateParameter("EmailAddress", adVarChar, adParamInput, 255)
			.Parameters.Append objCmd.CreateParameter("UserTypeID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Language", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("Comments", adLongVarChar, adParamInput, -1) 
			.Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput, 1)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarChar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("UserIDOutput", adInteger, adParamOutput)
			
			.Parameters("UserID") = Request.Form("UserID")
			.Parameters("EmailToAddress") = Request.Form("EmailToAddress")
			.Parameters("RecipientLastName") = Request.Form("FirstName")					
			.Parameters("RecipientFirstName") = Request.Form("LastName")
			.Parameters("EmailAddress") = Request.Form("EmailAddress")
			.Parameters("UserTypeID") = Request.Form("UserType")
			.Parameters("Language") = ""
			.Parameters("Comments") = Request.Form("Comments")
			.Parameters("Status") = Request.Form("Status")
			.Parameters("UpdatedBy") = Session("UserID")
			.Parameters("EmployeeID") = Request.Form("EmployeeID")

			.StatusConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'lngEmailErrorMsgID = objCmd.Parameters.Item("EmailErrorMsgIDOutput")
		'Session("EmailErrorMsgID") = lngEmailErrorMsgID		
	
End Sub

Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(3)

If Session("ViewButton") = "Admin" Then
	arrButton(2) = "Status"
ElseIf Session("ViewButton") = "New" Then
	arrButton(3) = "Status"
Else
	'This catches ALL
	arrButton(1) = "Status"
End If

	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='EmailSend.asp?Link=AD&ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='EmailSend.asp?Link=AD&ViewButton=Admin';""><i class=""fa fa-coffee""></i> Admin</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='EmailSend.asp?Link=AD&ViewButton=New';""><i class=""fa fa-clock""></i> New</button>" & _
				"</div>"

End Sub

Public Sub LoadActionButtons
'Load the the View Selector buttons depending on what has been clicked

	Response.Write "<div class=""btn-group aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary"" data-toggle=""modal"" data-target=""#ModalSendEmail"" data-EmailSubject=""" & strEmailSubject & """ HREF=""#"" title=""Send Emails to selected campaign group: " & strEmailSubject & " "">" & _
				"<i class=""fa fa-folder""></i> Send Emails</button>" & _
				"</div>"

	'"<button type=""button"" class=""btn btn-outline-primary"" data-toggle=""modal"" data-target=""#ModalSendEmail"" HREF=""#"" onClick=""loadCDMC(" & objRS("EmployeeID") & ")"" title=""Send Emails to selected campaign group"">" & _
	
End Sub


Public Sub LoadCampaign
'Load the Campaign details for emails already sent
Dim strStatusDisplay
Dim strSentBy
Dim strSelect
Dim strSelected
Dim strEmailSubject
Dim strObject
Dim strEmailSubjectShort
Dim strEmailSubjectTitle
Dim strDateCreated

	'Display all if the User is an Isidore Admin
	If Session("UserTypeID") = 99 Then
		'Get the Email Campaigns for the Business Group of the User logged in
		objRS1.Open "SELECT EmailSubject,BusinessArea From qryCAPSEmailSummary WITH(NOLOCK) WHERE [Active]='Y' GROUP BY EmailSubject,BusinessArea",objCon
		'objRS1.Open "SELECT EmailCampaignID,EmailSubject,BusinessArea From qryCAPSEmailSummary WITH(NOLOCK) WHERE [Active]='Y' GROUP BY EmailCampaignID,EmailSubject,BusinessArea",objCon
	Else
		'Get the Email Campaigns for the Business Group of the User logged in
		objRS1.Open "SELECT EmailSubject,BusinessArea From qryCAPSEmailSummary WITH(NOLOCK) WHERE BusinessArea = '" & Session("BusinessArea") & "' AND [Active]='Y' GROUP BY EmailSubject,BusinessArea",objCon
		'objRS1.Open "SELECT EmailCampaignID,EmailSubject,BusinessArea From qryCAPSEmailSummary WITH(NOLOCK) WHERE BusinessArea = '" & Session("BusinessArea") & "' AND [Active]='Y' GROUP BY EmailCampaignID,EmailSubject,BusinessArea",objCon
	End If
	
		If objRS1.EOF Then
			strSelect = "<option>No Subjects available</option>"
		Else
			strSelect = "<option>All Subjects</option>"
		End If
		
		'Loop through all Subjects and list them for filtering
		Do Until objRS1.EOF
			If Session("EmailSubjectSelected") = objRS1("EmailSubject") Then
				strSelected = " SELECTED "
			Else
				If Session("EmailSubjectSelected") = "...Created but not sent..." AND (IsNull(objRS1("EmailSubject")) OR objRS1("EmailSubject") = "") Then
					strSelected = " SELECTED "
				Else
					strSelected = ""
				End If
			End If
			
			'Get the short Subject for display
			
			
			'Catch any blank Subjects
			If IsNull(objRS1("EmailSubject")) Or objRS1("EmailSubject") = "" Then
				strSelect = strSelect & "<option name=""Blank"" id=""Blank"" " & strSelected & " style=""font-style:italic;"">...Created but not sent...</option>"
			Else
				If Len(objRS1("EmailSubject"))>25 Then
					strEmailSubjectShort = Left(objRS1("EmailSubject"),25)
				Else
					strEmailSubjectShort = objRS1("EmailSubject")
				End If
				
				'strSelect = strSelect & "<option name=""" & objRS1("EmailSubject") & """ id=""" & objRS1("EmailSubject") & """ " & strSelected & " >" & strEmailSubjectShort & "</option>"
				strSelect = strSelect & "<option name=""" & objRS1("EmailSubject") & """ id=""" & objRS1("EmailSubject") & """ " & strSelected & " >" & objRS1("EmailSubject") & "</option>"
			End If
			
		objRS1.Movenext
		Loop
		
	objRS1.Close

	strSelect = "<select class=""form-control"" id=""EmailSubjectSelect"" name=""EmailSubjectSelect"" onChange=""ChangeSubject();"" Style=""font-size:12px;"">" & strSelect & "</select>"
	
	If IsNull(Session("EmailSubjectSelected")) OR Session("EmailSubjectSelected") = "" OR Session("EmailSubjectSelected") = "All Subjects" Then
		strEmailSubject = ""
	Else
		strEmailSubject = " AND [EmailSubject] = '" & Session("EmailSubjectSelected") & "' "
	End If
	'Display all if the User is an Isidore Admin
	If Session("UserTypeID") = 99 Then
	
		'Description:	Loads CDMC Details onto the page called from
		objRS1.Open "SELECT TOP 30 * FROM qryCAPSEmailSummary WITH(NOLOCK) WHERE [Active]='Y' " & strEmailSubject & " AND [Status] <>'Deleted' ORDER By [Status],[EmailCampaignID] DESC",objCon
	Else
		'Description:	Loads CDMC Details onto the page called from
		objRS1.Open "SELECT TOP 30 * FROM qryCAPSEmailSummary WITH(NOLOCK) WHERE BusinessArea = '" & Session("BusinessArea") & "' AND [Active]='Y' " & strEmailSubject & " AND [Status] <>'Deleted' ORDER By [Status],[EmailCampaignID] DESC",objCon
	End If
	
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"" style=""font-weight:bold;"">Email Campaigns </span>" & strSelect & "</div></div>"
		
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact table-hover text-left""><thead>" & _
			"<tr><th style=""font-weight:bold; font-size:12px;"">Date Sent</th><th style=""font-weight:bold; font-size:12px;"">Subject</th>" & _
			"<th style=""font-weight:bold; font-size:12px;"">Status</th><th style=""font-weight:bold; font-size:12px;"">Emails</th></tr></thead><tbody>"
		
		
		Do Until objRS1.EOF
			
			'Determine the pill display based on the Status
			If IsNull(objRS1("Status")) or objRS1("Status") = "" Then
				strStatusDisplay = "<span class=""badge badge-info"">None</span>"
			Else
				If objRS1("Status") = "Created" Then
					strStatusDisplay = "<span class=""badge badge-secondary"">" & objRS1("Status") & "</span>"
				ElseIf objRS1("Status") = "Deleted" Then
					strStatusDisplay = "<span class=""badge badge-danger"">" & objRS1("Status") & "</span>"
				Else
					strStatusDisplay = "<span class=""badge badge-success"">" & objRS1("Status") & "</span>"
				End If
			End If
			
			'Get the Sent By for the Hover display
			If IsNull(objRS1("SentByUser")) or objRS1("SentByUser") = "" Then
				strSentBy = "title=""Not Yet Sent "
			Else
				strSentBy = "title=""Sent By: " & objRS1("SentByUser") & ""
			End If
			
			If IsNull(objRS1("Object")) or objRS1("Object") = "" Then
				strObject = ""
			Else
				strObject = " | Object: " & objRS1("Object") & ""
			End If
						
			If IsNull(objRS1("EmailSubject")) Then
				strEmailSubjectShort = ""
				strEmailSubjectTitle = ""
			Else
				If Len(objRS1("EmailSubject"))>20 Then
					strEmailSubjectShort = Left(objRS1("EmailSubject"),20)
				Else
					strEmailSubjectShort = objRS1("EmailSubject")
				End If
				strEmailSubjectTitle = " | Subject: " & objRS1("EmailSubject") & """"
			End If
			
			If IsNull(objRS1("DateCreated")) or objRS1("DateCreated") = "" Then
				strDateCreated = " ----"
			Else
				strDateCreated = " | DateCreated: " & objRS1("DateCreated") & ""
			End If
			
			strSentBy = strSentBy & strObject & strDateCreated & strEmailSubjectTitle
			
			Response.Write "<tr class='clickable-row' data-href='EmailSend.asp?EmailCampaignID=" & objRS1("EmailCampaignID") & "&EmailSubject=" & objRS1("EmailSubject") & "' data-target='_blank' " & strSentBy & "><td style=""font-size:12px;"">" & objRS1("DateSent") & "</td><td style=""font-size:12px;"">" & strEmailSubjectShort & "</td>" & _
				"<td style=""font-size:16px;"">" & strStatusDisplay & "</td><td style=""font-size:12px;"">" & objRS1("Emails") & "</td></TR>"
			
			'Response.Write "<tr class='clickable-row' data-href='EmailSend.asp?EmailCampaignID=" & objRS1("EmailCampaignID") & "&EmailSubject=" & objRS1("EmailSubject") & "' data-target='_blank' " & strSentBy & "><td style=""font-size:12px;"">" & objRS1("DateSent") & "</td><td style=""font-size:12px;"">" & objRS1("EmailSubject") & "</td>" & _
			'	"<td style=""font-size:16px;"">" & strStatusDisplay & "</td><td style=""font-size:12px;"">" & objRS1("Emails") & "</td></TR>"
				
			strEmailSubject = objRS1("EmailSubject")
			
		objRS1.Movenext
		Loop
		
		'Else
		'	Response.write "No campaigns"
	   'End If

	objRS1.Close
	
	Response.write "</tbody></table></div></div>"
	
End Sub

Public Sub SendEmails
'Procedure to send all emails in the email campaign
Dim intCountEmails
Dim strEmailBody
Dim strSubject
Dim strEmailSensitivity
Dim strEmailType
Dim strFromAddress
Dim strEmailAttachment
Dim strEmailAddressTo
Dim strFirstName
Dim strMode
Dim intWarningDays
Dim strError
Dim intErrorExists

intWarningDays = GetSystemAdmin("WarningDays")

If Session("EmailCampaignID") = 0 or IsNull(Session("EmailCampaignID")) Then

	Response.Write "<div class=""alert alert-danger"" role=""alert"">No Emails Sent. No Email Campaign Selected!</div>"
	Exit Sub
End If

'Check for the mode setting
If IsNull(Request.QueryString("Mode")) or Request.QueryString("Mode") = "" Then
	strMode = ""
	Response.Write "<div class=""alert alert-danger"" role=""alert"">No Emails Sent. No MODE Selected!</div>"
	Exit Sub
Else
	strMode = Request.QueryString("Mode")
End If

'response.write "SELECT top 1 * FROM qryCAPSEmail WITH(NOLOCK) WHERE [Status] = 'Created' AND [Active] = 'Y' AND [EmailCampaignID] = " & Session("EmailCampaignID") & " AND [EmailSubject] = '" & Session("EmailSubject") & "'"
'Exit sub

	'Description:	Loads CDMC Details onto the page called from
	objRS.Open "SELECT top 3000 * FROM qryCAPSEmail WITH(NOLOCK) WHERE [Status] = 'Created' AND [Active] = 'Y' AND [EmailCampaignID] = " & Session("EmailCampaignID") & " AND [EmailSubject] = '" & Session("EmailSubject") & "'" ,objCon
	'objRS.Open "SELECT top 1 * FROM qryCAPSEmail WITH(NOLOCK) WHERE [Status] = 'Created' AND [Active] = 'Y' AND [EmailCampaignID] = " & Session("EmailCampaignID") & "" ,objCon
	'response.write "SELECT top 1 * FROM qryCAPSEmail WITH(NOLOCK) WHERE [Status] = 'Created' AND [Active] = 'Y' AND [EmailCampaignID] = " & Session("EmailCampaignID") & " AND [EmailSubject] = '" & Session("EmailSubject") & "'"
		If objRS.EOF then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No Emails Sent. No Emails not yet sent for the Email Campaign: " & Session("EmailCampaignID") & "!</div>"
			Exit Sub
		End If
		
		'Set the Email Address From  --Overwritten below in the recordset ---**************
		strFromAddress = GetSystemAdmin("CAPSEmailAddressFrom")
		
		If strFromAddress = "" OR IsNull(strFromAddress) then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No Emails Sent. There is no System Parameter for the CAPSEmail Address FROM: CAPSEmailAddressFrom !</div>"
			Exit Sub
		End If
		
		intErrorExists = 0
		
		Do Until objRS.EOF
		
			strError = ""
			
			'Get the FROM email address for any different record address
			If NOT IsNull(objRS("FromAddress")) Then
				strFromAddress = objRS("FromAddress")
			End If
			
			'Get the email address of the Recipent
			If IsNull(objRS("EmailToAddress")) OR objRS("EmailToAddress") = "" Then
        
				strError = "Recipient does not have an email address"
				'GoTo EndLoop 'Send to the operator so that the email still gets processed and the card removed/cancelled ---- TO DO ----
			Else
				'If the user has selected test then send to their own account instead of the users
				If strMode = "Test" Then
					strEmailAddressTo = Session("UserEmail")
				Else
					strEmailAddressTo = objRS("EmailToAddress") '"michael.giacomin@defence.gov.au" '
				End If
			End If
	'response.write "strFromAddress=" &  strFromAddress
	'response.write " strEmailAddressTo=" &  strEmailAddressTo
			'Titus security tag (Official, Sensitive etc..)
			If IsNull(objRS("EmailSensitivity")) Or objRS("EmailSensitivity") = "" Then
				strEmailSensitivity = "[SEC=OFFICIAL]"
			Else
				strEmailSensitivity = objRS("EmailSensitivity")
			End If
			
			'Get the email subject for the email
			If IsNull(objRS("EmailSubject")) Then
				strSubject = ""
			Else
				strSubject = objRS("EmailSubject") & " " & strEmailSensitivity
			End If
	
			'HTML or Plain
			If IsNull(objRS("EmailType")) Or objRS("EmailType") = "" Then
				strEmailType = "HTML"
			Else
				strEmailType = objRS("EmailType")
			End If
	
			'Any attachments as files and locations should be in UNC format
			If IsNull(objRS("EmailAttachment")) Then
				strEmailAttachment = ""
			Else
				strEmailAttachment = objRS("EmailAttachment")
			End If
	
			'Get the Email Detail for the Body of the email
			If IsNull(objRS("EmailBody")) Then
				strEmailBody = ""
			Else
				strEmailBody = objRS("EmailBody")
				strEmailBody = "<HTML><SPAN Style=""font-family:arial; font-size:14px;"">" & strEmailBody & "</SPAN></HTML>"
			End If
			
			'Get the recipient first name and replace this is the Email Body if it appears
			If IsNull(objRS("RecipientFirstName")) Then
				strFirstName = ""
			Else
				strFirstName = objRS("RecipientFirstName")
				'Replace the First name field in square brackets with the first name of the recipient
				strEmailBody = Replace(strEmailBody, "[FirstName]", strFirstName)
				strEmailBody = Replace(strEmailBody, "[Date]", Date() + intWarningDays)
			End If
	
			If strError = "" Then
			
				'NEW Jan 2022 -- If there is an error then stop the process and display an error
				On Error Resume Next
				
				'Call the procedure (in the Functions include page) to send the email
				Send_Email strFromAddress,strEmailAddressTo,strSubject,strEmailBody,strEmailAttachment,strEmailType
				'Send_Email("capsapp@defence.gov.au","andrew.bull3@defence.gov.au","<I>Test","Body of email</I>")
	
				If Err.Number <> 0 Then
					Response.Write "<div class=""alert alert-danger"" role=""alert"">EmailID: " & objRS("EmailID") & " EmpID:" & objRS("EmployeeID") & " ~ Error: " & Err.Number & " " & Err.Description & ". Not all emails sent. Try Sending Batch again.</div>"
					intErrorExists = 1
				End If
				
				On Error GoTo 0
				
			Else
				
			End If

			'If there has been an error in the send function above then skip the database update of the record and end this procedure (intErrorExists=1 is an error and the else is at the bottom)
			If intErrorExists = 0 Then
			
			
				'If the user has selected test then send to their own account instead of the users
				If strMode = "Test" Then
					'Do not mark emails as sent
				Else
				
					'If there is an error then mark the record with the error
					If strError <> "" Then
					
						Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: " & strError & "  for EmailID: " & Session("EmailID") & "</div>"
						
						'Update the record as an error
						objCon.Execute("UPDATE tblCAPSEmail SET [Status] = 'Error',SendError='" & strError & "',DateUpdated=GetDate() WHERE EmailID=" & objRS("EmailID") & "")
						
					Else
					
						'Update the record as sent
						objCon.Execute("UPDATE tblCAPSEmail SET [Status] = 'EmailSent',SentByUser = '" & Session("UserName") & "',SentByApplication='CAPSSMTP',FromAddress='" & strFromAddress & "',DateSent=GetDate() WHERE EmailID=" & objRS("EmailID") & "")
						
						If objRS("Object") = "CAPSSMTPError" Then
							'Response.Write "UPDATE tblCAPSApplication SET [EmailSent] = GetDate(), [ErrorEmailSent] = GetDate() ,[WarningDate] = GetDate() WHERE ApplicationID=" & objRS("ObjectID") & ""
							objCon.Execute("UPDATE tblCAPSApplication SET [EmailSent] = GetDate(), [ErrorEmailSent] = GetDate() ,[WarningDate] = GetDate() WHERE ApplicationID=" & objRS("ObjectID") & "")
						
						End If
					End If
					
				'END Test MODE
				End If
				
				intCountEmails = intCountEmails + 1
				
				objRS.Movenext
				
			'Else for If intErrorExists = 0 (so intErrorExists = 1) Then
			Else
				objRS.Movelast
			End If
			
		'objRS.Movenext
		Loop
		
	objRS.Close
	
	If intCountEmails = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Emails Sent for Email Campaign ID: " & Session("EmailCampaignID") & "</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intCountEmails & " Emails Sent for Email Campaign ID: " & Session("EmailCampaignID") & " - " & strSubject & "</div>"
	End If
		
End Sub

Public Sub LoadSelectedEmailCampaign()
'Load the Campaign details for emails already sent
Dim strStatusDisplay

	'Description:	Loads CDMC Details onto the page called from
	objRS1.Open "SELECT * FROM qryCAPSEmailSummary WITH(NOLOCK) WHERE [EmailCampaignID]=" & Session("EmailCampaignID") & " AND EmailSubject = '" & Session("EmailSubject") & "'",objCon
	'objRS1.Open "SELECT * FROM qryCAPSEmailSummary WITH(NOLOCK) WHERE [EmailCampaignID]=" & Session("EmailCampaignID"),objCon

		strCampaignSelected = "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"" style=""font-weight:bold;"">"
				  
		'Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
        '          "<span class=""panel-subheader"" style=""font-weight:bold;"">"
		
		If objRS1.EOF Then
			'Response.Write "No Email Campaign Selected"
		Else
			strEmailSubject = objRS1("EmailSubject")
			strCampaignSelected2 = "" & objRS1("EmailSubject") & " : " & objRS1("Emails") & " Emails"
			strCampaignSelected = strCampaignSelected & "ID: <span style=""font-weight:normal;"">" & objRS1("EmailCampaignID") & " </span>DateSent: <span style=""font-weight:normal;"">" & objRS1("DateSent") & "</span> Subject: <span style=""font-weight:normal;"">" & _
				objRS1("EmailSubject") & "</span> Type: <span style=""font-weight:bold; color:green;"">" & objRS1("Object") & "</span> Emails: <span style=""font-weight:normal;"">" & objRS1("Emails") & " </span>"
			
			'strCampaignSelected = strCampaignSelected & "Email Campaign Selected</br>ID: <span style=""font-weight:normal;"">" & objRS1("EmailCampaignID") & " </span>DateSent: <span style=""font-weight:normal;"">" & objRS1("DateSent") & "</span> Subject: <span style=""font-weight:normal;"">" & _
			'	objRS1("EmailSubject") & "</span> Status: <span style=""font-weight:normal;"">" & objRS1("EmailSubject") & "</span> Emails: <span style=""font-weight:normal;"">" & objRS1("Emails") & " </span>"
				
			'Response.Write "Email Campaign Selected</br>ID: <span style=""font-weight:normal;"">" & objRS1("EmailCampaignID") & " </span>DateSent: <span style=""font-weight:normal;"">" & objRS1("DateSent") & "</span> Subject: <span style=""font-weight:normal;"">" & _
			'	objRS1("EmailSubject") & "</span> Status: <span style=""font-weight:normal;"">" & objRS1("EmailSubject") & "</span> Emails: <span style=""font-weight:normal;"">" & objRS1("Emails") & " </span>"
		End If
		

	objRS1.Close
	
	strCampaignSelected = strCampaignSelected & "</span></div></div>"
	'Response.write "</span></div></div>"
	
End Sub


Public Sub LoadUserEmail()
'Load the logged in User's Email address for sending test emails to

	'Description:	Loads CDMC Details onto the page called from
	objRS1.Open "SELECT TOP 1 * FROM tblUsers WITH(NOLOCK) WHERE [UserID]=" & Session("UserID"),objCon
		
		If objRS1.EOF Then
			Session("UserEmail") = "No Email"
		Else
			Session("UserEmail") = objRS1("EmailAddress")
		End If
		

	objRS1.Close
	
End Sub


Public Sub DeleteData(lngEmailID)


	'Update the record as sent
	objCon.Execute("UPDATE tblCAPSEmail SET [Status] = 'Deleted',DateUpdated=GetDate(),UpdatedBy = '" & Session("UserI") & "',Active='N' WHERE EmailID=" & lngEmailID & "")
	
	
	If IsNull(lngEmailID) OR lngEmailID = "" OR lngEmailID = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error. No Employee selected to remove from Email Campaign.</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & lngEmailID & " Removed From Email Campaign : " & strEmailSubject & "</div>"
	End If

End Sub

Set objRS = Nothing
Set objRS1 = Nothing
Set objCon = Nothing
%>
