
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=ADOVBS.inc -->
<!-- #Include file=CAPSFunctions.asp -->
<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp?State=Expired")

'Description:	View and process credit limit applications
'Author:		Tiff P (plagiarising some MG code)
'Date:			January 2025

	Response.Expires = -1500	
	
Dim objCon
Dim objCmd
Dim objRS

Dim x 
Dim strsearchKeyword
Dim strsearchCardType
Dim strsearchChangeType
Dim strsearchRequestType
Dim strsearchStatus
Dim strSortOrder
Dim strsearchChangeOption

'stuff for pagination
Dim y
Dim strPages
Dim strPages2
Dim strActive
Dim lngPage
Dim lngStartingRecord
Dim lngEndingRecord
Dim lngTotalRecords
Dim lngPageRecordsMax
Dim lngTotalPages
Dim lngNextRecord
Dim lngCurrentList
Dim lngCurrentRecord
Dim strTableStructure
Dim lngCurrentPage
Dim strActiveItem
Dim strPagination
Dim strDisplaying
Dim i
Dim strShowGroup
'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection
objCon.Open Session("DBConnection")

strServerPath = Session("ServerPath")

If IsEmpty(Request.QueryString("CurrentPage")) OR IsNull(Request.QueryString("CurrentPage")) OR (Request.QueryString("CurrentPage")) = " " Then
	lngCurrentRecord = 1
	Else
	lngCurrentRecord = clng(Request.QueryString("CurrentPage"))
End If

'use session variable for managing current page and records
If IsEmpty(Request.QueryString("SearchKeyword")) OR IsNull(Request.QueryString("SearchKeyword")) OR (Request.QueryString("SearchKeyword")) = " " Then
	Else
	Session("SearchKeyword") = Request.QueryString("SearchKeyword")	
End If

If IsEmpty(Request.QueryString("SearchCardType")) OR IsNull(Request.QueryString("SearchCardType")) OR (Request.QueryString("SearchCardType")) = " " Then
	ElseIf Request.QueryString("SearchCardType") = "All" Then
		Session("SearchCardType") = ""
	Else
	Session("SearchCardType") = Request.QueryString("SearchCardType")	
End If

If IsEmpty(Request.QueryString("SearchChangeType")) OR IsNull(Request.QueryString("SearchChangeType")) OR (Request.QueryString("SearchChangeType")) = " " Then
	Session("SearchChangeType") = " ApplicationTypeName LIKE '%Limit%' "
	strsearchChangeType = Session("SearchChangeType")
	strsearchChangeOption = ""
	Else
	Session("SearchChangeType") = " ApplicationTypeName = '" & Request.QueryString("SearchChangeType") & "' "
	strsearchChangeType = Session("SearchChangeType")
	strsearchChangeOption = Request.QueryString("SearchChangeType")
End If

If IsEmpty(Request.QueryString("SearchRequestType")) OR IsNull(Request.QueryString("SearchRequestType")) OR (Request.QueryString("SearchRequestType")) = " " Then
	ElseIf Request.QueryString("SearchRequestType") = "All" Then
		Session("SearchRequestType") = ""
	Else
	Session("SearchRequestType") = Request.QueryString("SearchRequestType")	
End If

If IsEmpty(Request.QueryString("SearchStatus")) OR IsNull(Request.QueryString("SearchStatus")) OR (Request.QueryString("SearchStatus")) = " " Then
	ElseIf Request.QueryString("SearchStatus") = "All" Then
		Session("SearchStatus") = ""
	Else
	Session("SearchStatus") = Request.QueryString("SearchStatus")	
End If

If IsEmpty(Request.QueryString("SortOrder")) OR IsNull(Request.QueryString("SortOrder")) OR (Request.QueryString("SortOrder")) = " " Then
	Else
	Session("SortOrder") = Request.QueryString("SortOrder")	
End If



%>
<script>
function OpenCLModal(cb) {
	
	var id = cb.getAttribute('data-CLAppID');
	//document.getElementById('CLAppID').value=id;

  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ModalReviewLimitBody").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetLimitDetails.asp?AppID=" + id + "", true);
  xhttp.send();

}
</script>
<body>
	<main class="main py-3">
    <div class="container-fluid justify-content-center align-items-center flex flex-wrap ">
        <div class="row p-3 m-3 overflow-auto">
			<div class="row p-auto m-auto overflow-auto w-100">
				<div class="col d-flex flex-column w-100">
					<div class="row row-cols-auto pt-1 pb-1 m-auto overflow-auto w-100" id="creditlimitlist">
						<div class="col d-flex flex-column w-100" id="searchbar">
							<h2 class="float-start">Credit Limits</h2>
						</div>
						<form>
							<div class="row row-cols-auto pt-3 pb-3 m-auto overflow-auto w-100 align-items-end text-center" id="search">
								<div class="col form-group">
									<label for="SearchKeyword">Keyword</label>
									<input class="form-control" type="text" id="SearchKeyword" placeholder="Search for a keyword" name="SearchKeyword" value="<%=Session("SearchKeyword")%>"></input>
								</div>
								<div class="col form-group">
									<% Call CardTypeOptions()%>
								</div>
								<div class="col form-group">
									<% Call ChangeTypeOptions()%>
								</div>								
								<div class="col form-group">
									<% Call RequestTypeOptions()%>
								</div>
								<div class="col form-group">
									<% Call StatusTypeOptions()%>
								</div>
								<div class="col form-group">
									<% Call SortOptions()%>
								</div>
								<button type="submit" class="btn btn-primary form-group" onClick="Search()">Search</button>
								</div>
							</div>
						</form>
					</div>
					<% If IsEmpty(strTableStructure) OR IsNull(strTableStructure) Then
						Call DisplayCreditLimits()
						Call Pagination(lngCurrentPage, lngTotalPages)
					End If %>
						<section id="content" class="content flex-fill w-100 text-center">
						<div class="col d-flex flex-column w-100" id="paginationBar">
						<h6><%=strDisplaying%></h6><%=strPagination%>
						</div>
						<table class="table flex-fill w-100 text-center table-striped">
							<thead>
								<tr>
									<th scope="col">Application ID</th>
									<th scope="col">Card Type</th>
									<th scope="col">Change Request</th>
									<th scope="col">Request Type</th>
									<th scope="col">Name on Card</th>
									<th scope="col">Submitted</th>
									<th scope="col">Status</th>
									<th scope="col">Action</th>
								</tr>
							</thead>
							<tbody>				

								<%=strTableStructure%>
							</tbody>
						</table>
						<%=strPagination%>
						</section>
				</div>
			</div>
		</div>		  
    </div>
	</main>
	<!-- Review Credit Limit Application Modal -->
	<div class="modal fade" id="ModalReviewLimit" tabindex="-1" role="dialog" aria-labelledby="ModalReviewLimit" aria-hidden="true">
	  <div class="modal-dialog modal-dialog-centered" role="document">
		<div class="modal-content">
		  <div class="modal-header">
			<h5 class="modal-title" id="ModalReviewAppLimitTitle" style="font-weight:bold;">View Credit Limit</h5>
			<button type="button" class="close" data-dismiss="modal" aria-label="Close">
			  <span aria-hidden="true">&times;</span>
			</button>
		  </div>
		  <div class="modal-body" id="ModalReviewLimitBody">

		  </div>
		  <div class="modal-footer">
			<button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
		  </div>
		</div>
	  </div>
	</div>
	<!-- End Credit Limit Application Modal -->
	<div class="container-fluid">
<!-- #Include file=CAPSFooter.asp -->
	</div>
  </body>
</html>

<%
'filter/search for card type
Public Sub CardTypeOptions()
Dim strSearchCard 

strSearchCard = "<label for=""SearchCardType"">Card Type</label><select class=""form-control"" id=""SearchCardType"" name=""SearchCardType""><option value=""All"">All</option>"

	Select Case Session("SearchCardType")
		Case "NAB DTC"
			strSearchCard = strSearchCard & "<option selected>NAB DTC</option><option>NAB Lodge</option><option>NAB DPC</option>"
		Case "NAB Lodge"
			strSearchCard = strSearchCard & "<option>NAB DTC</option><option selected>NAB Lodge</option><option>NAB DPC</option>"
		Case "NAB DPC"
			strSearchCard = strSearchCard & "<option>NAB DTC</option><option>NAB Lodge</option><option selected>NAB DPC</option>"
		Case Else
			strSearchCard = strSearchCard & "<option>NAB DTC</option><option>NAB Lodge</option><option>NAB DPC</option>"
	End Select
	
	Response.Write strSearchCard & "</select>"
	
End Sub 

'filter/search for change type
Public Sub ChangeTypeOptions()
Dim strsearchChangeSelect

strsearchChangeSelect = "<label for=""SearchChangeType"">Change Request</label><select class=""form-control"" id=""SearchChangeType"" name=""SearchChangeType""><option value="" "">All</option>"
	
	Select Case strsearchChangeOption
		Case "DTC Limit Change"
			strsearchChangeSelect = strsearchChangeSelect & "<option selected>DTC Limit Change</option><option>Lodge Limit Change</option><option>DPC Limit Change</option>"
		Case "Lodge Limit Change"
			strsearchChangeSelect = strsearchChangeSelect & "<option>DTC Limit Change</option><option selected>Lodge Limit Change</option><option>DPC Limit Change</option>"
		Case "DPC Limit Change"
			strsearchChangeSelect = strsearchChangeSelect & "<option>DTC Limit Change</option><option >Lodge Limit Change</option><option selected>DPC Limit Change</option>"
		Case Else
			strsearchChangeSelect = strsearchChangeSelect & "<option>DTC Limit Change</option><option >Lodge Limit Change</option><option>DPC Limit Change</option>"
	End Select
	
	Response.Write strsearchChangeSelect & "</select>"
End Sub 

'filter/search for request type
Public Sub RequestTypeOptions()
Dim strSearchRequest

strSearchRequest = "<label for=""SearchRequestType"">Request Type</label><select class=""form-control"" id=""SearchRequestType"" name=""SearchRequestType""><option value=""All"">All</option>"
	
	Select Case Session("SearchRequestType")
		Case "No"
			strSearchRequest = strSearchRequest & "<option selected value=""No"">Temporary</option><option value=""Yes"">Permanent</option>"
		Case "Yes"
			strSearchRequest = strSearchRequest & "<option value=""No"">Temporary</option><option selected value=""Yes"">Permanent</option>"
		Case Else
			strSearchRequest = strSearchRequest & "<option value=""No"">Temporary</option><option value=""Yes"">Permanent</option>"
	End Select
	
	Response.Write strSearchRequest & "</select>"

End Sub

'filter/search for Status 
Public Sub StatusTypeOptions()
Dim strseachStatusType

strseachStatusType = "<label for=""SearchStatus"">Status</label><select class=""form-control"" id=""SearchStatus"" name=""SearchStatus""><option value=""All"">All</option>"
	
	Select Case Session("SearchStatus")
		Case "On Hold"
			strseachStatusType = strseachStatusType & "<option selected>On Hold</option><option>Awaiting Review</option><option >Awaiting Export</option><option >Exported</option><option >Done</option><option >Rejected</option><option >Deleted</option>"
		Case "Awaiting Review"
			strseachStatusType = strseachStatusType & "<option >On Hold</option><option selected>Awaiting Review</option><option >Awaiting Export</option><option >Exported</option><option >Done</option><option >Rejected</option><option >Deleted</option>"
		Case "Awaiting Export"
			strseachStatusType = strseachStatusType & "<option >On Hold</option><option>Awaiting Review</option><option selected>Awaiting Export</option><option >Exported</option><option >Done</option><option >Rejected</option><option >Deleted</option>"
		Case "Exported"
			strseachStatusType = strseachStatusType & "<option >On Hold</option><option>Awaiting Review</option><option>Awaiting Export</option><option selected>Exported</option><option >Done</option><option >Rejected</option><option >Deleted</option>"
		Case "Done"
			strseachStatusType = strseachStatusType & "<option >On Hold</option><option>Awaiting Review</option><option>Awaiting Export</option><option>Exported</option><option selected>Done</option><option >Rejected</option><option >Deleted</option>"
		Case "Rejected"
			strseachStatusType = strseachStatusType & "<option >On Hold</option><option>Awaiting Review</option><option>Awaiting Export</option><option>Exported</option><option>Done</option><option selected>Rejected</option><option >Deleted</option>"
		Case "Deleted"
			strseachStatusType = strseachStatusType & "<option >On Hold</option><option>Awaiting Review</option><option>Awaiting Export</option><option>Exported</option><option>Done</option><option>Rejected</option><option selected>Deleted</option>"
		Case Else
			strseachStatusType = strseachStatusType & "<option >On Hold</option><option>Awaiting Review</option><option>Awaiting Export</option><option>Exported</option><option>Done</option><option>Rejected</option><option>Deleted</option>"
	End Select
	
	Response.Write strseachStatusType & "</select>"
	
End Sub 

'filter for Sort order
Public Sub SortOptions()
Dim strSortOrderType

strSortOrderType = "<label for=""SortOrder"">Sort</label><select class=""form-control"" id=""SortOrder"" name=""SortOrder""><option value=""DateSubmitted DESC"">Submitted Desc</option>"
	
	Select Case Session("SortOrder")
		Case "Submitted Asc"
			strSortOrderType = strSortOrderType & "<option value=""DateSubmitted ASC"" selected>Submitted Asc</option><option value=""ApplicationID DESC"">ApplicationID Desc</option><option value=""ApplicationID DESC"">ApplicationID Asc</option>"
		Case "ApplicationID Desc"
			strSortOrderType = strSortOrderType & "<option value=""DateSubmitted ASC"">Submitted Asc</option><option value=""ApplicationID DESC"" selected>ApplicationID Desc</option><option value=""ApplicationID DESC"">ApplicationID Asc</option>"
		Case "ApplicationID Asc"
			strSortOrderType = strSortOrderType & "<option value=""DateSubmitted ASC"">Submitted Asc</option><option value=""ApplicationID DESC"">ApplicationID Desc</option><option value=""ApplicationID DESC"" selected>ApplicationID Asc</option>"
		Case Else
			strSortOrderType = strSortOrderType & "<option value=""DateSubmitted ASC"">Submitted Asc</option><option value=""ApplicationID DESC"">ApplicationID Desc</option><option value=""ApplicationID DESC"">ApplicationID Asc</option>"
	End Select
	
	Response.Write strSortOrderType & "</select>"
End Sub 

'display credit limits
Public Sub DisplayCreditLimits()

Dim strSQL
Dim strApplicationID
Dim strCardType
Dim strChangeRequest 
Dim strRequestType
Dim strNameOnCard
Dim strDateSubmitted
Dim strStatus
Dim strAction

'max records shown per page
lngPageRecordsMax = 50


If IsEmpty(Session("SearchKeyword")) OR IsNull(Session("SearchKeyword")) OR Session("SearchKeyword") = ""  Then
		strsearchKeyword = " "
	Else
		strsearchKeyword = " AND (EmployeeID LIKE '%" & Session("SearchKeyword") &"%' OR FirstName LIKE '%" &Session("SearchKeyword")& "%' OR Surname LIKE '%" &Session("SearchKeyword")& "%' OR NameOnCard LIKE '%" & Session("SearchKeyword")& "%' )"
End If

If IsEmpty(Session("SearchCardType")) OR IsNull(Session("SearchCardType")) OR Session("SearchCardType") = "" Then
		strsearchCardType = " "
	Else
		strsearchCardType = " AND (CardTypeSub = '" & Session("SearchCardType") & "') "
End If

If IsNull(Session("SearchRequestType")) OR Session("SearchRequestType") = "" Then
		strsearchRequestType = " "
	Else
		strsearchRequestType = " AND (ChangesPermanent = '" & Session("SearchRequestType") & "') "
End If

If IsNull(Session("SearchStatus")) OR Session("SearchStatus") = "" Then
		strsearchStatus = " "
	Else
		strsearchStatus = " AND ([Status] = '" & Session("SearchStatus") & "') "
End If

If IsNull(Session("SortOrder")) OR Session("SortOrder") = "" Then
		strSortOrder = " DateSubmitted DESC"
	Else
		strSortOrder = Session("SortOrder")
End If

	strSQL = "SELECT * FROM qryCAPSApplicationsList WITH(NOLOCK) WHERE " & strsearchChangeType & strsearchKeyword & strsearchCardType & strsearchRequestType & strsearchStatus & " AND (Left(CardTypeSub,3)='NAB') ORDER BY " & strSortOrder & ""
'response.write strSQL
objRS.Open strSQL,objCon, 3, 1
'to count records
y = 0

	If objRS.EOF Then
'		'something something
	Else
'		'count the total records in the dataset
		objRS.Movelast
		objRS.Movefirst
		lngTotalRecords = objRS.Recordcount
	End If
	
	'count total pages of data
	IF IsNumeric(lngTotalRecords) Then
		lngTotalPages = int(lngTotalRecords/lngPageRecordsMax)+1
	Else
		lngTotalPages = 1
	End If
	
	'if there is less than the max number of records, handle it
	If lngTotalPages < 1 Then
		lngPage = 1
		lngCurrentList = ((lngPage - 1) * lngPageRecordsMax)
		lngStartingRecord = 0
		lngEndingRecord = lngTotalRecords
			If lngTotalRecords < 1 Then
				strDisplaying = "No results found"
			Else
				strDisplaying = "Displaying "& lngTotalRecords & " of " &lngTotalRecords& " results"
			End If
	Else
		'loop through the total pages to create page numbers, record groups and buttons
		do until x = lngTotalPages
			x=x+1
			lngPage = x
			lngCurrentList = int(((lngPage - 1) * lngPageRecordsMax))
				'set the current page variables based on the page number
				If lngCurrentRecord = lngPage  Then
					lngStartingRecord = lngCurrentList
					lngEndingRecord = lngCurrentList + lngPageRecordsMax
						If lngEndingRecord > lngTotalRecords Then
							lngEndingRecord = lngTotalRecords
						End If
					lngCurrentPage = lngPage
					strDisplaying = "Displaying "& lngEndingRecord & " of " &lngTotalRecords& " results"
				End If
		Loop
		
	End If
		'loop through the records and write the table
		Do until objRS.EOF
			'count the records
			y = y + 1
			
			'if y is equal to or less than the total of the current list to display
			If y <= lngEndingRecord and y => lngStartingRecord Then
			
				If IsNull(objRS("ApplicationID")) Then
					strApplicationID = "Unknown"
				Else	
					strApplicationID = objRS("ApplicationID")
				End If
				
				If IsNull(objRS("CardTypeSub")) Then
					strCardType = "Unknown"
				Else	
					Select Case objRS("CardTypeSub")
						Case "NAB DTC"
							strCardType = "<span class=""badge badge-pill badge-primary"">" & objRS("CardTypeSub") & "</span>"
						Case "NAB Lodge"
							strCardType = "<span class=""badge badge-pill badge-dark"">" & objRS("CardTypeSub") & "</span>"
						Case "NAB DPC"
							strCardType = "<span class=""badge badge-pill text-light"" style=""background-color: #cb15a6;"">" & objRS("CardTypeSub") & "</span>"
						Case Else
							strCardType = objRS("CardTypeSub")
					End Select
				End If
				
				If IsNull(objRS("ApplicationTypeName")) Then
					strChangeRequest = "Unknown"
				Else	
					strChangeRequest = objRS("ApplicationTypeName")
				End If
				
				If IsNull(objRS("ChangesPermanent")) Then
					strRequestType = "Unknown"
				Else
					Select Case objRS("ChangesPermanent")
						Case "No"
							strRequestType = "Temporary"
						Case "Yes"
							strRequestType = "Permanent"
						Case Else
							strRequestType = objRS("ChangesPermanent")
					End Select
				End If
				
				If IsNull(objRS("Status")) Then
					strStatus = "Unknown"
				Else
					'change this to call GetApplicationDetails for the modal
					Select Case objRS("Status")
						Case "On Hold"	
							strStatus = "<span class=""badge badge-pill badge-dark"">" & objRS("Status") & "</span>"
							'strAction = "<button type=""button"" class=""btn btn-warning btn-md text-light"" data-toggle=""modal"" data-target=""#ModalReviewLimitApp"" data-CLAppID=" & strApplicationID & " title=""View"" onClick=OpenCLModal(this);><i class=""fa fa-eye""></i></button>"
							strAction = "<a type=""button"" title=""Review Application"" class=""btn btn-outline-secondary btn-md"" href=""ApplicationDetail.asp?ApplicationID=" & strApplicationID  & """><i class=""fa fa-file""></i></a>"
							'strAction = strAction & "<button type=""button"" class=""btn btn-danger btn-md"" href=""#"" title=""Reject""><i class=""fa fa-times""></i></button>"
						Case "Awaiting Review"	
							strStatus = "<span class=""badge badge-pill badge-warning"">" & objRS("Status") & "</span>"
							'strAction = "<button type=""button"" class=""btn btn-warning btn-md text-light"" data-toggle=""modal"" data-target=""#ModalReviewLimitApp"" data-CLAppID=" & strApplicationID & """title=""View"" onClick=OpenCLModal(this);><i class=""fa fa-eye""></i></button>"
							strAction = "<a type=""button"" title=""Review Application"" class=""btn btn-outline-secondary btn-md"" href=""ApplicationDetail.asp?ApplicationID=" & strApplicationID  & """><i class=""fa fa-file""></i></a>"
							'strAction = strAction & "<button type=""button"" class=""btn btn-danger btn-md"" href=""#"" title=""Reject""><i class=""fa fa-times""></i></button>"
						Case "Done"	
							strStatus = "<span class=""badge badge-pill badge-success"">" & objRS("Status") & "</span>"
							strAction = "<a type=""button"" title=""Review Application"" class=""btn btn-warning text-light btn-md"" href=""ApplicationDetail.asp?ApplicationID=" & strApplicationID  & """><i class=""fa fa-eye""></i></a>"
							strAction = strAction & "<button type=""button"" class=""btn btn-primary btn-md"" onClick=OpenCLModal(this); data-toggle=""modal"" data-target=""#ModalReviewLimit"" data-CLAppID=" & strApplicationID & " title=""View Limit Details""><i class=""fa fa-money-bill""></i></button>"
												
						Case "Rejected"	
							strStatus = "<span class=""badge badge-pill badge-danger"">" & objRS("Status") & "</span>"
							strAction = "<a type=""button"" title=""Review Application"" class=""btn btn-warning text-light btn-md"" href=""ApplicationDetail.asp?ApplicationID=" & strApplicationID  & """><i class=""fa fa-eye""></i></button>"
						Case Else
							strStatus = "<span class=""badge badge-pill badge-light"">" & objRS("Status") & "</span>"
							strAction = ""
					End Select
				End If
				
				If IsNull(objRS("DateSubmitted")) Then
					strDateSubmitted = "Unknown"
				Else	
					strDateSubmitted = Left(objRS("DateSubmitted"),10)
				End If
				
				If IsNull(objRS("NameOnCard")) Then
					strNameOnCard = "Unknown"
				Else	
					strNameOnCard = objRS("NameOnCard")
				End If
			
				strTableStructure = strTableStructure & "<tr><th scope=""row"">" & strApplicationID & "</th><td>" & strCardType & "</td><td>" & strChangeRequest & "</td><td>" & strRequestType & "</td><td>" & strNameOnCard & "</td><td>" & strDateSubmitted & "</td><td>" & strStatus & "</td><td>" & strAction & "</td></tr>"				
				
			End If
			objRS.movenext
		Loop
objRS.Close
End Sub

'write the pagination
Public Sub Pagination(lngCurrentPage, lngTotalPages)
Dim maxPages
Dim startPage
Dim endPage

'max number of pages on each screen
maxPages = 20	
'start writing the pagination container		
strPagination = strPagination & "<div class=""container""><div class=""row""><div class=""col-12 text-center""><nav aria-label=""Page navigation""><ul class=""pagination show"">"

'get the starting page for the screen
startPage = lngCurrentPage
'get the ending page for the screen
endPage = startPage + maxPages - 1 
	'if there are more pages counted than the total number of pages
	If endPage > lngTotalPages Then
		endPage = lngTotalPages
		startPage = endPage - maxPages + 1
		'if startPage ends up less than 1, don't print that
		If startPage < 1 Then
			startPage = 1 
		End If
	End If
	
	'if the current page is not the starting page, write the next button
	If lngCurrentPage > 1 Then
		strPagination = strPagination & "<li class=""page-item""><a class=""page-link"" href=""CreditLimits.asp?CurrentPage=" & lngCurrentPage - 1 & """ aria-label=""Previous""><span aria-hidden=""true"">&laquo;</span><span class=""sr-only"">Previous</span></a></li>" 
	End If

	'write the buttons from start to end for the page
	For i = startPage to endPage 
		'if its the active page, highlight it
		If i = lngCurrentPage Then
			strPagination = strPagination & "<li class=""page-item active""><a class=""page-link "" href=""CreditLimits.asp?CurrentPage=" & i & """ aria-label=""Next""><span aria-hidden=""true"">" & i & "</a></li>"		
		Else 
			strPagination = strPagination & "<li class=""page-item""><a class=""page-link"" href=""CreditLimits.asp?CurrentPage=" & i & """ aria-label=""Next""><span aria-hidden=""true"">" & i & "</a></li>"		
		End If
	Next
	'f the current page is not the last page, make the back button
	If lngCurrentPage < lngTotalPages Then
		strPagination = strPagination & "<li class=""page-item""><a class=""page-link"" href=""CreditLimits.asp?CurrentPage=" & lngCurrentPage + 1 & """ aria-label=""Next""><span aria-hidden=""true"">&raquo;</span><span class=""sr-only"">Next</span></a></li>"
	End If
	'close the pagination container
	strPagination = strPagination & "</ul></nav></div></div></div>"

End Sub 


Set objRS = Nothing
Set objCon = Nothing

%>