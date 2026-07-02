
<!-- #Include file=../cc/CAPSHeader.asp -->
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

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objRS1 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")
	
    objCon.Open Session("DBConnection")	

    Session("CurrentPage") = "Admin/EmailError.asp"

	If IsNull(Session("ApplicationID")) OR Session("ApplicationID") = "" Then Session("ApplicationID")= 0

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
	
'Execute Action

If Request.QueryString("Action") = "Save" Then

	Call SaveUser() 

End If

If Request.QueryString("Action") = "Delete" Then   
	
    Call DeleteData(Request.QueryString("EmployeeID"))
   
End If

%>

<html>
<head>
<script LANGUAGE="javascript">
	function triggerModal(ModalID) {
				
		//var ModalName = '#emailModal' + ModalID
		var ModalName = '#emailModal'
		$(ModalName).modal("show");

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
	self.location = "Users.asp?SearchInput=" + id;

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

	document.getElementById("ModalDeleteMessage").innerHTML = 'Deactivate User - ' + name + '?';
	document.getElementById("ModalDelete").style.display = "block";
	document.getElementById("EmployeeDeleteID").value = id;
	
}

function deleteEmployee(cb) {

	var id = document.getElementById("EmployeeDeleteID").value

	self.location = "Users.asp?Action=Delete&EmployeeID=" + id;

}

	(function() {
		
	'use strict';
	window.addEventListener('load', function() {
	// Fetch all the forms we want to apply custom Bootstrap validation styles to
	var forms = document.getElementsByClassName('needs-validation');
	// Loop over them and prevent submission
	var validation = Array.prototype.filter.call(forms, function(form) {
	form.addEventListener('submit', function(event) {
	if (form.checkValidity() === false) {
	event.preventDefault();
	event.stopPropagation();
	document.getElementById("AlertDanger").style.display = "block";
	}
	form.classList.add('was-validated');
	//Success
	}, false);

	form.addEventListener('change', function(event) {
	if (form.checkValidity() === false) {
	event.preventDefault();
	event.stopPropagation();
	}
	form.classList.add('was-validated');
	}, false);
	
	});
	}, false);
	})();

</script>

</head>
<body>

	<main class="main py-3">
		<div class="container">
		

	<!-- Start Delete Modal -->

	<div class="modal fade" id="ModalDelete" tabindex="-1" role="dialog" aria-labelledby="ModalDeleteCenterTitle" aria-hidden="true">
		<div class="modal-dialog modal-dialog-centered" role="document">
		  	<div class="modal-content">
				<div class="modal-header">
			  		<h5 class="modal-title" id="ModalDeleteLongTitle" style="font-weight:bold;">Delete User</h5>
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
	<form action="Users.asp?Action=Save" method="POST" id="frm" name="frm" class="needs-validation" novalidate>
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
			  <h2>User Administration Screen</h2>
			</div>
			<div class="col-md-4 text-right">
			  <button class="btn btn-primary" data-toggle="modal" data-id="0" data-target="#emailModal" data-id="0" onClick="loadDocE(this);">
				<i class="fa fa-plus"></i> Create new user
			  </button>
			</div>
		  </div>
		  
		  <div class="row py-2">
            <div class="col-md-9">
              <%Call LoadViewButtons()%>
            </div>
			<div class="col-md-3">
				<div class="form-group has-search">
					<span class="fa fa-search form-control-feedback"></span>
				 <input type="text" class="form-control" type="search" id="SearchInput" name="SearchInput" placeholder="Search Users by Keyword" onChange="SearchUsers();" value="<%=Request.QueryString("SearchInput")%>"/>
				 </div>
			</div>
          </div>
		  

				  <% Call DisplayTable() %>



	
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
Dim strFNameSearch
Dim strLNameSearch
Dim strSearchDate
Dim strSearchDay
Dim strSearchMonth
Dim strSearchYear
Dim strDateFrom
Dim strDateTo
Dim bolSkip
Dim strActive
Dim strPages2

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
	If IsNull(strSort) Or strSort = "" Then strSort = " ORDER BY LName ASC"
	
	If Session("ViewButton") = "Admin" Then
		strWhere = " AND [UserTypeID] > 9 "
	ElseIf Session("ViewButton") = "New" Then
		strWhere = " AND [DateUpdated] > CONVERT(DATETIME, '" & DateAdd("d", -30, now()) & "', 103) "
		'CONVERT(DATETIME, '2020-12-20 00:00:00', 102)
	Else
		'This catches ALL
		strWhere = ""
	End If

	
If strSearch = "" OR ISNull(strSearch) Then
	'If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) WHERE [UserID] > 0 " & strWhere & strSort
	'Else
	'	strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' " & strSort
	'End If
	
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
			strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) WHERE (DateUpdated > '" & strDateFrom & "' AND DateUpdated < '" & strDateTo & "')" & strWhere & strSort
			
		Else
			'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE (DateSubmitted = '" & strSearchDate & "')" & strWhere & strSort
			strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) WHERE (DateUpdated = '" & strSearchDate & "')" & strWhere & strSort
		End If
		
	Else
		'If Session("UserView") = "All" Then
			'If the user has entered a search term with a space then assume this is a first and last name so search on that only
			If Instr(1,strSearch," ")>0 Then
				arrNames = Split(strSearch," ")
				strFNameSearch = arrNames(0)
				strLNameSearch = arrNames(1)
				
				'strWhere = " WHERE ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%')"
				'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%'))" & strWhere & strSort
				strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR (FName Like '%" & strFNameSearch & "%' AND LName Like '%" & strLNameSearch & "%')) " & strSort
			Else
				'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strWhere & strSort
				strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FName Like '%" & strSearch & "%' OR LName Like '%" & strSearch & "%') " & strSort
			End If
		'Else
			'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strSort
		'	strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FName Like '%" & strSearch & "%' OR LName Like '%" & strSearch & "%') " & strSort
		'End If
	
	'End of 
	End If
	
End If

'Get the Sort Order for the sort order fontawesome image for the selected sort field
'For x = 1 to 8

	Select Case Request.QueryString("Sort")
	
		Case "EmployeeID"
			arrSort(1) = strSortArrow
		Case "FName"
			arrSort(2) = strSortArrow
		Case "LName"
			arrSort(3) = strSortArrow
		Case "UserLogon"
			arrSort(4) = strSortArrow
		Case "UserTypeID"
			arrSort(5) = strSortArrow
		Case "Active"
			arrSort(6) = strSortArrow
		Case "UpdatedByName"
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
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No User record for " & strRecordMessage & "</TH>" & _
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
                  "<span class=""panel-subheader"">Displaying " & Session("PageCombo") & " of " & lngTotalRecords & " users (" & (clng(lngStartingPage)*clng(Session("PageCombo")))-clng(Session("PageCombo"))+1 & " to " & clng(lngStartingPage)*clng(Session("PageCombo")) & ")</span><span class=""panel-subheader"" style=""float:right;"">Number of records per page: " & strPageCombo  & "</span></div></div>"
		
	
		'Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
        '          "<span class=""panel-subheader"">Displaying 50 of " & lngTotalRecords & " applications (" & lngStartingPage & " to " & lngStartingPage + 50 & ")</span></div></div>"
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col""><a href=""Users.asp?Sort=EmployeeID&Link=AD&SortType=" & strOrderType & """> Employee ID <i class=""fa fa-sort" & arrSort(1) & """></i></a></th>" & _
			"<th scope=""col""><a href=""Users.asp?Sort=LName&Link=AD&SortType=" & strOrderType & """> Last Name <i class=""fa fa-sort" & arrSort(2) & """></i></a></th>" & _
			"<th scope=""col""><a href=""Users.asp?Sort=FName&Link=AD&SortType=" & strOrderType & """> First Name <i class=""fa fa-sort" & arrSort(3) & """></i></a></th>" & _
			"<th scope=""col""><a href=""Users.asp?Sort=UserLogon&Link=AD&SortType=" & strOrderType & """> User Logon <i class=""fa fa-sort" & arrSort(4) & """></i></a></th>" & _
			"<th scope=""col""><a href=""Users.asp?Sort=UserTypeID&Link=AD&SortType=" & strOrderType & """> User Type <i class=""fa fa-sort" & arrSort(5) & """></i></a></th>" & _
			"<th scope=""col""><a href=""Users.asp?Sort=Active&Link=AD&SortType=" & strOrderType & """> Active  <i class=""fa fa-sort" & arrSort(6) & """></i></a> <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Only Users who are Active (Active=Y) can access the system.""></i></th>" & _
			"<th scope=""col""><a href=""Users.asp?Sort=UpdatedByName&Link=AD&SortType=" & strOrderType & """> Updated By <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
			"<th scope=""col""><a href=""Users.asp?Sort=DateUpdated&Link=AD&SortType=" & strOrderType & """> Date Updated <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
			"<th scope=""col"">Action</th>" & _
			"</tr></thead><tbody class=""text-left"">"
				
	End If
    
	x = 0
	
	
'strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) Order By LName ASC"
intRecordCount = 0

If IsEmpty(Session("UserID")) Then Session("UserID") = 0 End If

'objRS.Open strSQL,objCon   
    	
	Do until objRS.EOF	
	
	Response.Write "<tr><td>" & objRS("EmployeeID") & "</td>" & _
		"<td>" & objRS("LName") & "</td>" & _
		"<td>" & objRS("FName") & "</td>" & _
		"<td>" & objRS("UserLogon") & "</td>" & _
		"<td>" & objRS("UserTypeName") & "</td>" & _
		"<td>" & objRS("Active") & "</td>" & _ 
		"<td>" & objRS("UpdatedByName") & "</td>" & _
		"<td>" & objRS("DateUpdated") & "</td>" & _
		"<td><button	data-toggle=""modal"" data-target=""#emailModal"" data-id=""" & objRS("UserID") & """ data-EmployeeName=""" & objRS("LName") & """ onClick=""loadDocE(this);"" class=""btn btn-primary btn-xs"">" & _
		"<i class=""fa fa-pen""></i> Edit</button>" & _
		"<button class=""btn btn-primary btn-xs btn-danger"" data-toggle=""modal"" data-target=""#ModalDelete"" data-id=""" & objRS("UserID") & """ data-EmployeeName=""" & objRS("FName") & " " & objRS("LName") & """ onClick=""loadDelete(this);""><i class=""fa fa-times""></i> Remove</button></td></tr>"
	
	intRecordCount = intRecordCount + 1
	
	objRS.movenext
	
	Loop	
	
	
	If intRecordCount > 0 Then
		Response.Write "<TR><TH colspan=""7"">Total</TH>" & _
				"<TH colspan=""2"" style=""text-align:center;"">" & intRecordCount & "</TH></TR></tbody></table></div></div>"
				     
	End If

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
			strPages2 = strPages2 & "<li class=""page-item " & strActive & """><a class=""page-link"" href=""Applications.asp?StartingPage=" & x & """>" & x & "</a></li>"
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
			.Parameters.Append objCmd.CreateParameter("UserLogon", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("FName", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("LName", adVarChar, adParamInput, 50)              
			.Parameters.Append objCmd.CreateParameter("EmailAddress", adVarChar, adParamInput, 255)
			.Parameters.Append objCmd.CreateParameter("UserTypeID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Language", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("Comments", adLongVarChar, adParamInput, -1) 
			.Parameters.Append objCmd.CreateParameter("Active", adVarChar, adParamInput, 1)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarChar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("UserIDOutput", adInteger, adParamOutput)
			
			.Parameters("UserID") = Request.Form("UserID")
			.Parameters("UserLogon") = Request.Form("UserLogon")
			.Parameters("FName") = Request.Form("FirstName")					
			.Parameters("LName") = Request.Form("LastName")
			.Parameters("EmailAddress") = Request.Form("EmailAddress")
			.Parameters("UserTypeID") = Request.Form("UserType")
			.Parameters("Language") = ""
			.Parameters("Comments") = Request.Form("Comments")
			.Parameters("Active") = Request.Form("Active")
			.Parameters("UpdatedBy") = Session("UserID")
			.Parameters("EmployeeID") = Request.Form("EmployeeID")

			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'lngEmailErrorMsgID = objCmd.Parameters.Item("EmailErrorMsgIDOutput")
		'Session("EmailErrorMsgID") = lngEmailErrorMsgID		
	
End Sub

Public Function Validate_Access(UserTypeID,Screen)

    If Session("UserTypeID") = 99 Then
        
        Validate_Access = "Y"
        
    Else
        
        objRS.Open "SELECT ScreenID FROM qryScreenAccess WHERE UserTypeID = " & UserTypeID & " AND PageName = '" & Screen & "?TransactionType=" & Session("TransactionType") & "'",objCon

            If objRS.EOF Then
                Validate_Access = "N" 
            Else
                Validate_Access = "Y"
            End If
    
        objRS.Close
    
    End If

End Function

Public Sub DeleteData(ID)

		objCon.Execute "DELETE tblUsers WHERE UserID = '" & ID & "'"
		
End Sub


Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(3)

If Session("ViewButton") = "Admin" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "New" Then
	arrButton(3) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='Users.asp?Link=AD&ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='Users.asp?Link=AD&ViewButton=Admin';""><i class=""fa fa-coffee""></i> Admin</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='Users.asp?Link=AD&ViewButton=New';""><i class=""fa fa-clock""></i> New</button>" & _
				"</div>"

End Sub


Set objRS = Nothing
Set objCon = Nothing
%>
