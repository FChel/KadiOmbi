
<!-- #Include file=../cc/CAPSHeader.asp -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #include file="../CC/CAPSFunctions.asp" -->
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

    Session("CurrentPage") = "Admin/CAPSFileLoad.asp"

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

	Call SaveFileLoad() 

End If

If Request.QueryString("Action") = "Delete" Then   
	
    Call DeleteData(Request.QueryString("FileLoadID"))
   
End If

%>

<html>
<head>
<script LANGUAGE="javascript">
	function triggerModal(ModalID) {
				
		//var ModaFileName = '#emailModal' + ModalID
		var ModaFileName = '#emailModal'
		$(ModaFileName).modal("show");

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
	self.location = "CAPSFileLoad.asp?SearchInput=" + id;

}

function loadDocE(cb) {

  var id = cb.getAttribute('data-id');
  var xhttp = new XMLHttpRequest(); 
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("GetDetail").innerHTML = this.responseText;
    }
  };

  xhttp.open("GET", "../CC/AJAX/GetCAPSFileLoad.asp?FileLoadID=" + id, true);
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

	self.location = "CAPSFileLoad.asp?Action=Delete&FileLoadID=" + id;

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
	<form action="CAPSFileLoad.asp?Action=Save" method="POST" id="frm" name="frm" class="needs-validation" novalidate>
		<div class="modal fade" id="emailModal" tabindex="-1" role="dialog" aria-labelledby="emailModalTitle" aria-hidden="true">
			<div class="modal-dialog modal-dialog-centered" role="document">
				<div class="modal-content">
					<div class="modal-header">
						<h5 class="modal-title" id="emailModalLabel">CAPS File Load Administration zz<i class="fa fa-chevron-right"></i></h5><button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					</div>
					<div class="modal-body">
						<div class="col-md-12">
							<table class="table table-bordered table-hover CAPS">						
								<div id="GetDetail"></div>						
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
			  <h2>CAPS File Load Administration Screen</h2>
			</div>
		
		  </div>
		  
		  <div class="row py-2">
            <div class="col-md-9">
              <%Call LoadViewButtons()%>
            </div>
			<div class="col-md-3">
				<div class="form-group has-search">
					<span class="fa fa-search form-control-feedback"></span>
				 <input type="text" class="form-control" type="search" id="SearchInput" name="SearchInput" placeholder="Search by Date Loaded" onChange="SearchUsers();" value="<%=Request.QueryString("SearchInput")%>"/>
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
Dim arrSort(9)
Dim strSortArrow
Dim arrNames
Dim strFileTypeSearch
Dim strFileNameSearch
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
	If IsNull(strSort) Or strSort = "" Then strSort = " ORDER BY FileName ASC"
	
	If Session("ViewButton") = "Admin" Then
		strWhere = " AND [UserTypeID] > 9 "
	ElseIf Session("ViewButton") = "New" Then
		strWhere = " AND [DateLoaded] > CONVERT(DATETIME, '" & DateAdd("d", -30, now()) & "', 103) "
		'CONVERT(DATETIME, '2020-12-20 00:00:00', 102)
	Else
		'This catches ALL
		strWhere = ""
	End If
		
If strSearch = "" OR ISNull(strSearch) Then

		If Session("ViewButton") <> "All" Then
			strSQL = "SELECT TOP 100 * FROM qryCAPSFileLoad WITH(NOLOCK) WHERE FileType = '" & Session("ViewButton") & "' AND [FileLoadID] > 0 " & strWhere & strSort
		Else
			strSQL = "SELECT TOP 100 * FROM qryCAPSFileLoad WITH(NOLOCK) WHERE [FileLoadID] > 0 " & strWhere & strSort
		End If
		
Else
	'If the user has entered the date lookup then process this, otherwise perform the FileLoadID and Name searches
	If Left(strSearch,2) = "d:" Then	
		
		strSearchDate = Right(strSearch,Len(strSearch)-2)
		
		If IsDate(strSearchDate) Then 'DateAdd("d", -1, strSearchDate)		
			
			strSearchDate = MediumDate2(strSearchDate) 
			'strDateFrom = DateAdd("d", -1, strSearchDate)
			strDateFrom = strSearchDate
			strDateTo = DateAdd("d", +1, strSearchDate)
			strDateTo = MediumDate2(strDateTo)

			'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE (DateSubmitted > '" & strDateFrom & "' AND DateSubmitted < '" & strDateTo & "')" & strWhere & strSort
			If Session("ViewButton") = "All" Then
				strSQL = "SELECT TOP 100 * FROM qryCAPSFileLoad WITH(NOLOCK) WHERE (DateLoaded > '" & strDateFrom & "' AND DateLoaded < '" & strDateTo & "')" & strWhere & strSort
			Else
				strSQL = "SELECT TOP 100 * FROM qryCAPSFileLoad WITH(NOLOCK) WHERE FileType = '" & Session("ViewButton") & "' AND (DateLoaded > '" & strDateFrom & "' AND DateLoaded < '" & strDateTo & "')" & strWhere & strSort
			End If			
			
		Else
			'strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE (DateSubmitted = '" & strSearchDate & "')" & strWhere & strSort
			strSQL = "SELECT TOP 100 * FROM qryCAPSFileLoad WITH(NOLOCK) WHERE FileType = '" & Session("ViewButton") & "'" & strWhere & strSort
			
		End If
		
	Else
	
		strSQL = "SELECT TOP 100 * FROM qryCAPSFileLoad WITH(NOLOCK) WHERE FileType = '" & Session("ViewButton") & "' AND (DateLoaded = '" & strSearchDate & "')" & strWhere & strSort
					
	End If
	
End If

'Get the Sort Order for the sort order fontawesome image for the selected sort field
'For x = 1 to 8

	Select Case Request.QueryString("Sort")
	
		Case "FileLoadID"
			arrSort(1) = strSortArrow
		Case "FileType"
			arrSort(2) = strSortArrow
		Case "FileName"
			arrSort(3) = strSortArrow	
		Case "FileSeqNum"
			arrSort(4) = strSortArrow
		Case "Status"
			arrSort(5) = strSortArrow
		Case "Deleted"
			arrSort(6) = strSortArrow
		Case "DateLoaded"
			arrSort(7) = strSortArrow
		Case "LoadedBy"
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
			"<th scope=""col""><a href=""CAPSFileLoad.asp?Sort=FileLoadID&Link=AD&SortType=" & strOrderType & """>ID <i class=""fa fa-sort" & arrSort(1) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSFileLoad.asp?Sort=FileType&Link=AD&SortType=" & strOrderType & """> File Type <i class=""fa fa-sort" & arrSort(2) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSFileLoad.asp?Sort=FileName&Link=AD&SortType=" & strOrderType & """> File Name <i class=""fa fa-sort" & arrSort(3) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSFileLoad.asp?Sort=FileSeqNum&Link=AD&SortType=" & strOrderType & """>Seq Num  <i class=""fa fa-sort" & arrSort(4) & """></i></a> <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Card Process Status - the stage the application is at.""></i></th>" & _
			"<th scope=""col""><a href=""CAPSFileLoad.asp?Sort=Status&Link=AD&SortType=" & strOrderType & """> Status <i class=""fa fa-sort" & arrSort(5) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSFileLoad.asp?Sort=Status&Link=AD&SortType=" & strOrderType & """> Deleted <i class=""fa fa-sort" & arrSort(6) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSFileLoad.asp?Sort=DateLoaded&Link=AD&SortType=" & strOrderType & """> Date Loaded <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
			"<th scope=""col""><a href=""CAPSFileLoad.asp?Sort=LoadedBy&Link=AD&SortType=" & strOrderType & """> Loaded By <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
			"<th scope=""col"">Action</th>" & _
			"</tr></thead><tbody class=""text-left"">"
				
	End If
    

	
	
'strSQL = "SELECT TOP 100 * FROM qryUsers WITH(NOLOCK) Order By FileName ASC"
intRecordCount = 0

If IsEmpty(Session("UserID")) Then Session("UserID") = 0 End If

'objRS.Open strSQL,objCon   
    	
	Do until objRS.EOF	
	
	Response.Write "<tr><td>" & objRS("FileLoadID") & "</td>" & _
		"<td>" & objRS("FileType") & "</td>" & _
		"<td>" & objRS("FileName") & "</td>" & _		
		"<td>" & objRS("FileSeqNum") & "</td>" & _ 
		"<td>" & objRS("Status") & "</td>" & _
		"<td>" & objRS("Deleted") & "</td>" & _
		"<td>" & objRS("DateLoaded") & "</td>" & _
		"<td>" & objRS("LName") & "</td>" & _
		"<td><button	data-toggle=""modal"" data-target=""#emailModal"" data-id=""" & objRS("FileLoadID") & """ onClick=""loadDocE(this);"" class=""btn btn-primary btn-xs"">" & _
		"<i class=""fa fa-pen""></i> Edit</button>" & _
		"</tr>"
	
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
			"<a class=""page-link"" href=""CAPSFileLoad.asp?StartingPage=" & lngTotalPages & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span>" & _
			"</a></li></ul></nav></div></div></div>"

	End If
	
objRS.Close

End Sub

Sub SaveFileLoad()

	Dim intModalID

	intModalID = Request.Form("FileLoadID")

	objCon.Execute "UPDATE tblCAPSFileLoad SET Status = '" & Request.Form("Status") & "', Deleted = '" & Request.Form("Deleted") & "' WHERE FileLoadID = " & intModalID & ""
	
	
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

		objCon.Execute "DELETE qryCAPSFileLoad WHERE FileLoadID = '" & ID & "'"
		
End Sub


Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(7)

SELECT CASE Session("ViewButton") 

	CASE "All"
		arrButton(1) = "active"
	CASE "ANZCardlist" 
		arrButton(2) = "active"
	CASE "CSFromDiners"
		arrButton(3) = "active"
	CASE "NAfile"
		arrButton(4) = "active"
	CASE "ProMasterUser"
		arrButton(5) = "active"
	CASE "ROMANCostCentres"
		arrButton(6) = "active"
	CASE "Training"
		arrButton(7) = "active"
		
END SELECT


	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='CAPSFileLoad.asp?Link=AD&ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='CAPSFileLoad.asp?Link=AD&ViewButton=ANZCardlist';""><i class=""fa fa-coffee""></i> ANZ Card List</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='CAPSFileLoad.asp?Link=AD&ViewButton=CSFromDiners';""><i class=""fa fa-clock""></i> CS From Diners</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='CAPSFileLoad.asp?Link=AD&ViewButton=NAfile';""><i class=""fa fa-clock""></i> NA File</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(5) & """ onClick=""self.location.href='CAPSFileLoad.asp?Link=AD&ViewButton=ProMasterUser';""><i class=""fa fa-clock""></i> Promaster User</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(6) & """ onClick=""self.location.href='CAPSFileLoad.asp?Link=AD&ViewButton=ROMANCostCentres';""><i class=""fa fa-clock""></i> ROMAN Cost Centres</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(7) & """ onClick=""self.location.href='CAPSFileLoad.asp?Link=AD&ViewButton=Training';""><i class=""fa fa-clock""></i> Training</button>" & _
				"</div>"

End Sub


Set objRS = Nothing
Set objCon = Nothing
%>
