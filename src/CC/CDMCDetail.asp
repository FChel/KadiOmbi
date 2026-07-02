
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=ADOVBS.inc -->
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
Dim objCmd
Dim objRS
Dim objRS1
Dim strSelected
Dim x 
Dim strMessage
Dim strColour
Dim strViewButton

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection
objCon.Open Session("DBConnection")

'Collect the action results to save details
If Not IsEmpty(Request.QueryString("Task")) Then
	'Response.Write Request.QueryString("TaskID") & "," & Request.QueryString("Task") & "," & Request.QueryString("TaskLogID")
	Call SaveTaskLog(Request.QueryString("TaskID"),Request.QueryString("Task"),Request.QueryString("TaskLogID"))
End If

If Not IsEmpty(Request.QueryString("ViewButton")) Then
	strViewButton = Request.QueryString("ViewButton")
End If

If Not IsEmpty(Request.QueryString("CDMCID")) Then
	Session("CDMCID") = Request.QueryString("CDMCID")
End If

If Not IsEmpty(Request.QueryString("EmployeeSearchID")) Then
	Session("EmployeeSearchID") = Request.QueryString("EmployeeSearchID")
	
	'If the search has come from the Card screen then clear the Session CDMCID and get the top one in the load procedure
	Session("CDMCID") = 0
	
	'Get the Employees Top CDMCID from the local function
	Session("CDMCID") = GETCMDCID
End If

If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "Save" Then
		Call SaveMessage
	End If	
End If
	
%>
<script>
function OpenSs(cb) {

	//alert("asas");
	//var e = document.getElementById(this.cb);
	//var result = e.options[e.selectedIndex].value;
	
	//document.getElementById('ContinueMod').value=result;
	document.getElementById('ModApp').showModal();
}




</script>
<!-- Modal -->
<div class="modal fade" id="ModApp" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">CAPS Admin Tasks</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>Continue With Process</label><br>
            <input type="text" name="ContinueMod" id="ContinueMod" class="form-control input-md" value="">
			<input type="text" name="ContinueModTask" id="ContinueModTask" class="form-control input-md" value="">
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
      </div>
    </div>
  </div>
</div>


    <main class="main py-3">
      <div class="container">
		<form action="ApplicationDetail.asp?Action=Save" method="POST" id="frm" name="frm">
        <div class="row">
          <div class="col-md-6">
  
				<div class="panel-header py-2">
					<h4>CDMC Details</h4>
				 </div>
				 
						<%
						Call DisplayTableDetails()
						
						%>
				<button type="button" class="btn btn-primary btn-sm" onclick="self.location='CDMCList.asp'";> Close</button>
            </div>
			
			<div class="col-md-6 sidebar">
		  
      
              <div class="panel panel-shadow panel-validation mb-3">
				  <div class="panel-header">
				  <div class="row">
					<div class="col-md-10">
					<h4>CDMC History Updates</h4>
				
					</div>
					<div class="col-md-2">
						<button type="button" class="btn btn-primary btn-sm" onclick="self.location='CDMCList.asp'";> Close</button>
					</div>
			
				</div>
					<span class="panel-subheader">History of Changes to CDMC</span>
				  </div>
				  <div class="panel-content mb-8">
					<div class="table-responsive">
						<table class="table table-compact" id="dataTable" width="100%" cellspacing="0">
						  <thead>
							<tr>
							  <th>CMDCID</th>
							  <th>EID</th>
							  <th>First Name</th>
							  <th>Last Name</th>
							  <th>First Updated</th>
							  <th>Last Updated</th>
							  <th>Deleted</th>
							  
							</tr>
						  </thead>
						  <tbody>
						   
							<%
					
								  DisplayTableSummary()
									
							%>	

						  </tbody>
						  
						</table>

					  </div>
					  
				</div>
			  </div>
     
			
            </div>
          </div>
		  </form>
        </div>
    </main>

	
    <script>
     $(".ActionButton1").on("show.bs.dropdown", function(event){
	  var x = $(event.relatedTarget).text(); // Get the text of the element
	  alert(x);
	});
    </script>
	
	
<!-- #Include file=CAPSFooter.asp -->
  </body>
</html>

<%


Public Sub DisplayTableDetails()

Dim strAction
Dim strStatus
Dim strAddress
Dim dteDateUpdated
Dim strSQL
Dim strHeader
Dim strTitle
Dim strTitle2
Dim strTitle3
Dim strTitle4
Dim strValidPostal
Dim strAddressBGColour
Dim strPostalMessage

'If Session("EmployeeID") = "" OR ISNull(Session("EmployeeID")) Then
'	strSQL = "SELECT * FROM qryCAPSApplications WHERE ApplicationID = '" & Session("ApplicationID") & "'"
'Else
	strSQL = "SELECT * FROM qryCAPSCDMCHistory WITH(NOLOCK) WHERE CDMCID = '" & Session("CDMCID") & "'"
	'strSQL = "SELECT * FROM qryCAPSCDMC WHERE CDMCID = '" & Session("CDMCID") & "'"
'End If

objRS.Open strSQL,objCon

	
    If Not objRS.EOF Then

		Select Case objRS("hasChanged")
		
			Case  "Y"
				strAction = "<button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Release&CDMCID=" & objrs("CDMCID") & "'"";><i class=""fa fa-check""></i> Release</button>"
				strAction = strAction & " <button type=""button"" class=""btn btn-outline-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Reject&CDMCID=" & objrs("CDMCID") & "'"";><i class=""fa fa-times""></i> Reject</button>"
			Case "N"
				strAction = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='CSToDiners.asp?CDMCID=" & objrs("CDMCID") & "'"";><i class=""fa fa-file-o""></i> View CS</button>"
			
			Case Else
				strAction = "<button type=""button"" class=""btn btn-outline-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Cancel&CDMCID=" & objrs("CDMCID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				'strAction = "Rejected"
				strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationDetail.asp?CDMCID=" & objrs("CDMCID") & "'"";>Submitted</button>"
		End Select

		'strAddress = Trim(objRS("Address1")) & " " & Trim(objRS("Address2")) & " " & Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode"))
		
		'If len(strAddress) > 15 Then strAddress = left(strAddress,15) & "..."
		
		If IsNull(objRS("DateUpdated")) Then
			dteDateUpdated = ""
		Else
			dteDateUpdated = FormatDateTime(objRS("DateUpdated"),vbShortDate)
		End If
		
		'Set the style/pill for the Is Valid Postal fields
		If IsNull(objRS("IsValidPostal")) Then
			strValidPostal = ""
		Else
			strValidPostal = objRS("IsValidPostal")
		End If
		
		'Set the Postal message style based on the Valid Postal below
		If IsNull(objRS("PostalMessage")) Then
			strPostalMessage = ""
		Else
			strPostalMessage = objRS("PostalMessage")
		End If
		
		'Set the style and colour for the Valid Postal display field
		If strValidPostal = "Y" Then
			strValidPostal = "<span class=""badge badge-pill badge-success""> Yes </span>"
			
			''Set the Background colour for the Address fields
			strAddressBGColour = "#ccffcc"
			
			'Set the Postal Message Style
			strPostalMessage = "<span class=""badge badge-pill badge-success"">" & strPostalMessage & "</span>"
		Else
			strValidPostal = "<span class=""badge badge-pill badge-danger""> No </span>"
			
			''Set the Background colour for the Address fields
			strAddressBGColour = "#FFCDD2"
			
			'Set the Postal Message Style
			strPostalMessage = "<span class=""badge badge-pill badge-danger"">" & strPostalMessage & "</span>"
		End If
		
			
		
		
		'Set the Style and Title for the fields which CAPS uses
		strTitle = "title=""Used by CAPS (from the Corporate Directory) to re-format as the Address sent to card providers (3 lines x 30 characters for NAB)"" style=""background-color:#ccffcc;"""
		strTitle2 = "title=""Used by CAPS (from the Corporate Directory) to re-format as the Phone numbers sent to card providers (10 digits or +61 followed by 9 digits)"" style=""background-color:#ccffcc;"""
		
		strTitle3 = "title=""Re-formatted by CAPS from Postal Address as the Address sent to card providers (3 lines x 30 characters for NAB)"" style=""background-color:" & strAddressBGColour & ";"""
		strTitle4 = "title=""Re-formatted by CAPS from Telephone and Mobile) as the Phone numbers sent to card providers (10 digits or +61 followed by 9 digits)"" style=""background-color:" & strAddressBGColour & ";"""
				
		
		'Set the header to the Department for all
		strHeader = "<img src=""../images/logo_coa.png"" Title=""Department of Defence""> " & objRS("Firstname") & " " & objRS("Surname")
		
		Response.write "<div class=""panel-content row""><div class=""mb-3 col-md-5""><h4>" & strHeader & "</h4></div><div class=""mb-3 col-md-7"">" &  _
		"<div class=""btn-group btn-selector table-tabs-selector"" role=""group"" aria-label=""Basic example"">" &  _
		"<button type=""button"" data-target=""table-tabs"" data-type=""as-tabs"" class=""btn btn-outline-primary active"">" &  _
		"<i class=""fa fa-list""></i> View as Tabs</button>" &  _
		"<button type=""button"" data-target=""table-tabs"" data-type=""as-table"" class=""btn btn-outline-primary"">" &  _
		"<i class=""fa fa-table""></i> View as Table</button></div></div></div>" &  _
		"<div id=""table-tabs"" class=""as-tabs""><ul class=""nav nav-tabs"" id=""myFiTab"" role=""tablist""><li class=""nav-item"" role=""presentation"">" &  _
		"<a class=""nav-link active"" id=""overview-tab"" data-toggle=""tab"" href=""#overview"" role=""tab"" aria-controls=""overview"" aria-selected=""true"">Organisation Details</a>" &  _
		"</li><li class=""nav-item"" role=""presentation"">" &  _
		"<a class=""nav-link"" id=""address-details-tab"" data-toggle=""tab"" href=""#address-details"" role=""tab"" aria-controls=""card-details"" aria-selected=""false"">Address Details</a>" &  _
		"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-Phone-tab"" data-toggle=""tab"" href=""#my-Phone"" role=""tab"" aria-controls=""my-Phone"" aria-selected=""false"">Phone</a>" &  _
		"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-caps-tab"" data-toggle=""tab"" href=""#my-caps"" role=""tab"" aria-controls=""my-caps"" aria-selected=""false"">CAPS Details</a>" &  _
		"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-update-tab"" data-toggle=""tab"" href=""#my-update"" role=""tab"" aria-controls=""my-update"" aria-selected=""false"">Updates</a>" &  _
		"</li></ul><div class=""tab-content panel panel-light p-3"" id=""myFiTabContent"">" &  _
		"<div class=""tab-pane fade show active"" id=""overview"" role=""tabpanel"" aria-labelledby=""overview-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>CDMCID</th><td style=""font-weight:bold;"">" & objRS("CDMCID") & "</td></tr>" &  _
		"<tr><th>GroupName</th><td>" & objRS("GroupName") & "</td></tr>" &  _
		"<tr><th>DivisionName</th><td>" & objRS("DivisionName") & "</td></tr>" &  _
		"<tr><th>BranchName</th><td>" & objRS("BranchName") & "</td></tr>" &  _
		"<tr><th>DepartmentName</th><td>" & objRS("DepartmentName") & "</td></tr>" &  _
		"<tr><th>DepartmentNumber</th><td>" & objRS("DepartmentNumber") & "</td></tr>" &  _
		"<tr><th>CostCentre</th><td>" & objRS("CostCentre") & "</td></tr>" &  _
		"<tr><th>EmployeeID</th><td>" & objRS("EmployeeID") & "</td></tr>" &  _
		"<tr><th>EmployeeType</th><td>" & objRS("EmployeeType") & "</td></tr>" &  _
		"<tr><th>Firstname</th><td>" & objRS("Firstname") & "</td></tr>" &  _
		"<tr><th>Surname</th><td>" & objRS("Surname") & "</td></tr>" &  _
		"<tr><th>Title</th><td>" & objRS("Title") & "</td></tr>" &  _
		"</table></div>" &  _
		"<div class=""tab-pane fade"" id=""my-Phone"" role=""tabpanel"" aria-labelledby=""my-Phone-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>Email_Address</th><td>" & objRS("Email_Address") & "</td></tr>" &  _
		"<tr><th " & strTitle2 & ">TelephoneNumber</th><td " & strTitle2 & ">" & objRS("TelephoneNumber") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("TelephoneNumber"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle2 & ">MobileNumber</th><td " & strTitle2 & ">" & objRS("MobileNumber") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("MobileNumber"))) & " chars</span></td></tr>" &  _
		"<tr><th>DateofBirth</th><td>" & objRS("DateofBirth") & "</td></tr>" &  _
		"<tr><th>Gender</th><td>" & objRS("Gender") & "</td></tr>" &  _
		"<tr><th>ActualRankLvl</th><td>" & objRS("ActualRankLvl") & "</td></tr>" &  _
		"<tr><th>Site</th><td>" & objRS("Site") & "</td></tr>" &  _
		"<tr><th>Unit</th><td>" & objRS("Unit") & "</td></tr>" &  _
		"<tr><th>ReportsTo</th><td>" & objRS("ReportsTo") & "</td></tr>" &  _
		"</table></div>" &  _
		"<div class=""tab-pane fade"" id=""address-details"" role=""tabpanel"" aria-labelledby=""address-details-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th " & strTitle & ">PostalAddress_Unit</th><td " & strTitle & ">" & objRS("PostalAddress_Unit") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("PostalAddress_Unit"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle & ">PostalAddress_ClientLocation</th><td " & strTitle & ">" & objRS("PostalAddress_ClientLocation") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("PostalAddress_ClientLocation"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle & ">PostalAddress_DeliveryLocation</th><td " & strTitle & ">" & objRS("PostalAddress_DeliveryLocation") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("PostalAddress_DeliveryLocation"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle & ">PostalAddress_City</th><td " & strTitle & ">" & objRS("PostalAddress_City") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("PostalAddress_City"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle & ">PostalAddress_State</th><td " & strTitle & ">" & objRS("PostalAddress_State") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("PostalAddress_State"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle & ">PostalAddress_PostCode</th><td " & strTitle & ">" & objRS("PostalAddress_PostCode") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("PostalAddress_PostCode"))) & " chars</span></td></tr>" &  _
		"<tr><th>PostalAddress_Country</th><td>" & objRS("PostalAddress_Country") & "</td></tr>" &  _
		"<tr><th>DCD_PostalAddress</th><td>" & objRS("DCD_PostalAddress") & "</td></tr>" &  _
		"<tr><th>addressline1</th><td>" & objRS("addressline1") & "</td></tr>" &  _
		"<tr><th>addressline2</th><td>" & objRS("addressline2") & "</td></tr>" &  _
		"<tr><th>addressline3</th><td>" & objRS("addressline3") & "</td></tr>" &  _
		"<tr><th>addressline4</th><td>" & objRS("addressline4") & "</td></tr>" &  _
		"<tr><th>addressline5</th><td>" & objRS("addressline5") & "</td></tr>" &  _
		"<tr><th>addressline6</th><td>" & objRS("addressline6") & "</td></tr>" &  _
		"<tr><th>ClientLocation</th><td>" & objRS("ClientLocation") & "</td></tr>" &  _
		"<tr><th>StreetAddress</th><td>" & objRS("StreetAddress") & "</td></tr>" &  _
		"<tr><th>City</th><td>" & objRS("City") & "</td></tr>" &  _
		"<tr><th>State</th><td>" & objRS("State") & "</td></tr>" &  _
		"<tr><th>PostCode</th><td>" & objRS("PostCode") & "</td></tr>" &  _
		"<tr><th>DCDProtectedIdentity</th><td>" & objRS("DCDProtectedIdentity") & "</td></tr>" &  _
		"</table></div>" &  _
		"<div class=""tab-pane fade"" id=""my-caps"" role=""tabpanel"" aria-labelledby=""my-caps-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>IsValidPostal</th><td>" & strValidPostal & "</td></tr>" &  _
		"<tr><th " & strTitle3 & ">OutAddr1</th><td " & strTitle3 & ">" & objRS("OutAddr1") & "<span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("OutAddr1"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle3 & ">OutAddr2</th><td " & strTitle3 & ">" & objRS("OutAddr2") & "<span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("OutAddr2"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle3 & ">OutAddr3</th><td " & strTitle3 & ">" & objRS("OutAddr3") & "<span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("OutAddr3"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle3 & ">OutSuburb</th><td " & strTitle3 & ">" & objRS("OutSuburb") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("OutSuburb"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle3 & ">OutState</th><td " & strTitle3 & ">" & objRS("OutState") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("OutState"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle3 & ">OutPostCode</th><td " & strTitle3 & ">" & objRS("OutPostCode") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("OutPostCode"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle4 & ">OutDinersWorkPhone</th><td " & strTitle4 & ">" & objRS("OutDinersWorkPhone") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("OutDinersWorkPhone"))) & " chars</span></td></tr>" &  _
		"<tr><th " & strTitle4 & ">OutDinersMobilePhone</th><td " & strTitle4 & ">" & objRS("OutDinersMobilePhone") & " <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(Trim(objRS("OutDinersMobilePhone"))) & " chars</span></td></tr>" &  _
		"<tr><th>PostalMessage</th><td>" & strPostalMessage & "</td></tr>" &  _
		"<tr><th>hasChanged</th><td>" & objRS("hasChanged") & "</td></tr>" &  _
		"<tr><th>FormalFirstName</th><td>" & objRS("FormalFirstName") & "</td></tr>" &  _
		"<tr><th>FormalLastName</th><td>" & objRS("FormalLastName") & "</td></tr>" &  _
		"<tr><th>FormalMiddleName</th><td>" & objRS("FormalMiddleName") & "</td></tr>" &  _
		"<tr><th>OutDinersAddress1</th><td>" & objRS("OutDinersAddress1") & "</td></tr>" &  _
		"<tr><th>OutDinersAddress2</th><td>" & objRS("OutDinersAddress2") & "</td></tr>" &  _
		"<tr><th>OutTitle</th><td>" & objRS("OutTitle") & "</td></tr>" &  _
		"<tr><th>OutANZPhone</th><td>" & objRS("OutANZPhone") & "</td></tr>" &  _
		"<tr><th>RemoveCountdown</th><td>" & objRS("RemoveCountdown") & "</td></tr>" &  _
		"</table></div>" &  _
		"<div class=""tab-pane fade"" id=""my-update"" role=""tabpanel"" aria-labelledby=""my-update-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>FirstUpdated</th><td>" & objRS("FirstUpdated") & "</td></tr>" &  _
		"<tr><th>LastUpdated</th><td>" & objRS("LastUpdated") & "</td></tr>" &  _
		"<tr><th>ActiveEmployee</th><td>" & objRS("ActiveEmployee") & "</td></tr>" &  _
		"<tr><th>UpdatedBy</th><td>" & objRS("UpdatedBy") & "</td></tr>" &  _
		"<tr><th>DateUpdated</th><td>" & objRS("DateUpdated") & "</td></tr>" &  _
		"<tr><th>FileID</th><td>" & objRS("FileID") & "</td></tr>" &  _
		"<tr><th>Loaded</th><td>" & objRS("Loaded") & "</td></tr>" &  _
		"<tr><th>Deleted</th><td>" & objRS("Deleted") & "</td></tr>" &  _
		"</table>" &  _
		"</div></div></div>"
		
		
		
	End If
	
	'Response.Write strAction & " " & strStatus
	
objRS.Close

End Sub


Public Sub DisplayTableSummary()
Dim y
Dim strAction, strStatus
Dim strAddr1, strAddr2, strAddr3
Dim arrNames
Dim strFNameSearch
Dim strLNameSearch
Dim strWhere
Dim dteFirstUpdated
Dim dteLastUpdated
Dim strSQL
Dim strStyle

'If Session("EmployeeID") = "" Then
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
'	strSQL = "SELECT Top 100 * FROM qryCAPSCDMC WITH(NOLOCK) " & strWhere
'Else
	strSQL = "SELECT Top 100 * FROM qryCAPSCDMCHistory WITH(NOLOCK) WHERE [EmployeeID] = '" & Session("EmployeeSearchID") & "' ORDER By CDMCID DESC"
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
'End If

objRS.Open strSQL,objCon
    y = 0
    	
    Do until objRS.EOF 
		If isNull(objRS("FirstUpdated")) Then
			dteFirstUpdated = ""
		Else
			dteFirstUpdated = FormatDateTime(objRS("FirstUpdated"),vbShortDate)
		End If
		
		If isNull(objRS("LastUpdated")) Then
			dteLastUpdated = ""
		Else
			dteLastUpdated = FormatDateTime(objRS("LastUpdated"),vbShortDate)
		End If

		'Get the Address details based on the address type selected
		'strAddr1 = Left(objRS("Addressline1") & " " & objRS("Addressline2") & " " & objRS("Addressline3") & " " & objRS("Addressline4") & " " & objRS("Addressline5") & " " & objRS("Addressline6"),30) & "..."
		'strAddr2 = Left(objRS("PostalAddress_Unit") & " " & objRS("PostalAddress_ClientLocation") & " " & objRS("PostalAddress_DeliveryLocation") & " " & objRS("Postaladdress_City") & " " & objRS("Postaladdress_State") & " " & objRS("Postaladdress_PostCode"),30) & "..."
		'strAddr3 = Left(objRS("OutAddr1") & " " & objRS("OutAddr2") & " " & objRS("OutAddr3") & " " & objRS("OutSuburb") & " " & objRS("OutState") & " " & objRS("OutPostCode"),30) & "..."
		
		'Highlight the row selected
		If clng(Session("CDMCID")) = clng(objRS("CDMCID")) Then
			strStyle = "title=""<-- Currently Selected"" style=""background-color:#e6e6e6; text-align:center;"""
		Else
			strStyle = " style=""text-align:center;"" "
		End If
		
		Response.Write "<TR><TD " & strStyle & "><a Target=""_self"" HREF=""CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & """>" & objRS("CDMCID") & "</a></TD>" & _
				"<TD " & strStyle & "><a Target=""_self"" HREF=""CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & """>" & objRS("EmployeeID") & "</a></TD>" & _
				"<TD " & strStyle & ">" & objRS("Firstname") & "</TD><TD " & strStyle & ">" & objRS("Surname") & "</TD>" & _
				"<TD " & strStyle & ">" & dteFirstUpdated & "</TD>"  & _
				"<TD " & strStyle & ">" & dteLastUpdated & "</TD><TD " & strStyle & ">" & objRS("Deleted") & "</TD></TR>" 
				
		'response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & """>" & objRS("EmployeeID") & "</a></TD>" & _
		'		"<TD><a Target=""_self"" HREF=""CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""CDMCList.asp?CDMCID=" & objRS("CDMCID") & """>" & objRS("Title") & "</a></TD>" & _
		'		"<TD style=""text-align:center;"">" & objRS("Firstname") & "</TD><TD style=""text-align:center;"">" & objRS("Surname") & "</TD>" & _
		'		"<TD style=""text-align:center;"">" & strAddr1 & "</TD><TD style=""text-align:center;"">" & strAddr2 & "</TD>" & _
		'		"<TD style=""text-align:center;"">" & strAddr3 & "</TD><TD style=""text-align:center;"">" & dteFirstUpdated & "</TD>"  & _
		'		"<TD style=""text-align:center;"">" & dteLastUpdated & "</TD><TD style=""text-align:center;"">" & objRS("Active") & "</TD></TR>" 
		'		'"<TD style=""text-align:center;"">" & strStatus & "</TD><TD style=""text-align:center;"">" & objRS(14) & "</TD><TD style=""text-align:center;"">" & objRS(15) & "</TD></TR>"
			
			y = y + 1
			
		objRS.movenext
	Loop
	
	
	response.write "<TR><TH colspan=""5"">Total</TH>" & _
				"<TH colspan=""2"" style=""text-align:center;"">" & y & "</TH></TR>"
				
objRS.Close

End Sub

Public Sub LoadMessages()
'Procedure to load any messages relating to the application
Dim strSQL
Dim strPerson

	strSQL = "SELECT * FROM qryCAPSMessage WITH(NOLOCK) WHERE [Object] = 'Application' AND [ObjectID] = '" & Session("ApplicationID") & "'"

	objRS.Open strSQL,objCon
	
    Do Until objRS.EOF
		
		If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
			strPerson = ""
		Else
			If objRS("MessageFrom") = Session("UserID") Then
				strPerson = "(You)"
			Else
				If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
					If objRS("MessageFrom") = 0 Then
						strPerson = "(Admin)"
					End If
				End If
			End If
		End If
		
		Response.write "<div class=""panel panel-light col-9""><div class=""panel-header"">" & _
			"<h6>" & objRS("UserFrom") & " " & strPerson & "</h6><span class=""panel-subheader"">" & objRS("MessageDetail") & "</span></div></div>"

		objRS.Movenext
	Loop
				
objRS.Close

End Sub

Public Sub SaveMessage()

Dim lngMessageID
Dim lngAdminID
Dim strMessage
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

If Request.Form("MessageS") = "" or IsNull(Request.Form("MessageS")) Then
	strMessage = ""
Else
	strMessage = Request.Form("MessageS")
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
			.Parameters("Object") = "Application"
			.Parameters("ObjectID") = Session("ApplicationID")
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


Public Function GETCMDCID()

Dim strSQL

	strSQL = "SELECT Top 1 [CDMCID] FROM qryCAPSCDMCHistory WITH(NOLOCK) WHERE [EmployeeID] = '" & Session("EmployeeSearchID") & "' ORDER By CDMCID DESC"
	
objRS.Open strSQL,objCon

    If NOT objRS.EOF Then
		GETCMDCID = objRS("CDMCID")
	Else
		GETCMDCID=0
	End If

objRS.Close
	
End Function

Set objRS = Nothing
Set objCon = Nothing

%>