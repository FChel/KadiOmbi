<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #Include file=../CC/CAPSFunctions.asp -->
<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp?State=Expired")


'Description:	Create and view System Parameters
'Author:		MG
'Date:			January 2020

	Response.Expires = -1500	

Dim objCon
Dim objCmd
Dim objRS
Dim objRS1
Dim strSelected
Dim x,y
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

If Not IsEmpty(Request.QueryString("Action")) Then
	If isNumeric(Request.Form("TotalRecords")) Then
		y=0

		For x = 1 to Request.Form("TotalRecords")
			If IsNumeric(Request.Form("ID"&x)) Then
				y = y + 1

				Call SaveParameter(Request.Form("ID"&x),x,y)
			End If
		Next
		
		If y = 0 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Nothing changed to SAVE!</div>"
		End If
	Else
		Response.Write "Nothing to save"
		'Call SaveMessage
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

function UpdateInput(varID) {

	document.getElementById('ID'+varID).value = varID
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
		<form action="SystemParameters.asp?Action=Save" method="POST" id="frm" name="frm">

           <!--<div class="row mb-2">
			  <div class="col-md-10">
				<h4 class="text-left">System Parameters</h4>
			  </div>
			  <div class="col-md-2 text-right"></div>
			</div>
			
			 <div class="row mb-2">
			  <div class="col-12">
				<ul class="nav nav-tabs" id="myFiTab" role="tablist">
				  <li class="nav-item" role="presentation">
					<a class="nav-link active" href="#System">System Parameters</a>
				  </li>
				  <li class="nav-item" role="presentation">
					<a class="nav-link" href="#Admin">Admin Parameters</a>
				  </li>
				  <li class="nav-item" role="presentation">
					<a class="nav-link" href="#Application">Application Parameters</a>
				  </li>
				</ul>
			  </div>
			</div>-->
			
			<div class="row mb-4">
				<div class="col-12">
				 
						<%
						Call DisplayTableDetails()
						
						%>

            <!--</div>-->
			
			
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
	
	
<!-- #Include file=../CC/CAPSFooter.asp -->
  </body>
</html>

<%


Public Sub DisplayTableDetails2()

Dim x
Dim strActive
Dim strActiveSelect
Dim strDateUpdated
Dim strSQL
Dim strParameterType

	strSQL = "SELECT * FROM qryCAPSSystemParameters WITH(NOLOCK) ORDER BY [ParameterType],[ParameterName]"

objRS.Open strSQL,objCon

	If NOT objRS.EOF Then
	
		'Response.Write "<div class=""panel-content row"" >" & _
		'	"<div class=""form-row col-md-2""><label style=""font-weight:bold;"">Parameter Name</label></div><div class=""form-row col-md-3""><label style=""font-weight:bold;"">Parameter Value</label></div>" & _
		'	"<div class=""form-row col-md-4""><label style=""font-weight:bold;"">Parameter Description</label></div><div class=""form-row col-md-1""><label style=""font-weight:bold;"">Active</label></div>" & _
		'	"<div class=""form-row col-md-1""><label style=""font-weight:bold;"">Updated By</label></div><div class=""form-row col-md-1""><label style=""font-weight:bold;"">Date Updated</label></div></div>"
			
		Response.Write "<div class=""row information-row""><div class=""col-sm-4""><h6>Parameter Name</h6></div><div class=""col-sm-5""><h6>Parameter value</h6></div>" & _
			"<div class=""col-md-1""><h6>Active</h6></div><div class=""col-md-1 my-auto""><h6>Updated By</h6></div>" & _
			"<div class=""col-md-1 my-auto""><h6>Date Updated</h6></div></div>"

	End If
	
	strParameterType = ""
	
	Do Until objRS.EOF
	
		x = x + 1
		
		If IsNull(objRS("Active")) OR objRS("Active") = "" Then
			strActive = ""
		Else
			strActive = objRS("Active")
		End If
		
		strActiveSelect = ""
		
		If strActive = "Y" Then strSelected = " SELECTED "
		
		strActiveSelect = strActiveSelect & "<option " & strSelected & " value=""Y"">Y</option>"
		
		If strActive = "N" Or strActive = "" Then 
			strSelected = " SELECTED "
		Else
			strSelected = ""
		End If
		
		strActiveSelect = strActiveSelect & "<option " & strSelected & " value=""N"">N</option>"
		
		If IsNull(objRS("DateUpdated")) OR objRS("DateUpdated") = "" Then
			strDateUpdated = ""
		Else
			If IsDate(objRS("DateUpdated")) Then
				strDateUpdated = FormatDateTime(objRS("DateUpdated"),vbShortDate)
			End If
		End If
		'If the Parameter Type has changed then write the tab details for displaying details by Parameter Type (for clicking on the tab links)
		If strParameterType <> objRS("ParameterType") Then
			'End the previous div if this is not the first chnage in system parameters
			If strParameterType <> "" Then Response.Write "</DIV>"
			
			Response.Write "<div class=""tab-pane fade show"" id="""  &objRS("ParameterType") & """ role=""tabpanel"" aria-labelledby="""  &objRS("ParameterType") & "-tab"">"
		End If
		
		'Response.Write "<div class=""panel-content row"" >" & _
		'		"<div class=""form-row col-md-2""><label>" & objRS("ParameterName") & "</label></div>" & _
		'		"<div class=""form-row col-md-3""><input style=""font-weight:bold;"" type=""text"" id=""" & objRS("ParameterName") & """ name=""" & objRS("ParameterName") & """ class=""form-control input-md"" value=""" & objRS("ParameterValue") & """></div>" & _
		'		"<div class=""form-row col-md-4""><input type=""text"" id=""" & objRS("ParameterDescription") & """ name=""" & objRS("ParameterDescription") & """ class=""form-control input-md"" value=""" & objRS("ParameterDescription") & """></div>" & _
		'		"<div class=""form-row col-md-1""><SELECT class=""form-control"" name=""Active" & objRS("SystemParameterId") & """ id=""Active" & objRS("SystemParameterId") & """ >" & strActiveSelect & "</Select></div>" & _
		'		"<div class=""form-row col-md-1""><input style=""border: 0px;"" type=""text"" id=""" & objRS("UpdatedBy") & """ name=""" & objRS("UpdatedBy") & """ class=""form-control input-md"" value=""" & objRS("UpdatedBy") & """></div>" & _
		'		"<div class=""form-row col-md-1""><input style=""border: 0px;"" type=""text"" id=""" & objRS("DateUpdated") & """ name=""" & objRS("DateUpdated") & """ class=""form-control input-md"" value=""" & objRS("DateUpdated") & """></div></div>"
		
		Response.Write "<div class=""row information-row"" Style=""padding:10px;""><div class=""col-sm-4""><h6>" & objRS("SystemParameterID") & " - " & objRS("ParameterName") & "</h6><span class=""description"" style=""font-size:12px;"">" & objRS("ParameterDescription") & "</span></div>" & _
				"<div class=""col-sm-5 my-auto""><input onChange=""UpdateInput(" & objRS("SystemParameterID") & ");"" type=""text"" class=""form-control"" id=""ParameterValue" & objRS("SystemParameterID") & """ name=""ParameterValue" & objRS("SystemParameterID") & """ placeholder=""Parameter value"" value=""" & objRS("ParameterValue") & """></div>" & _
				"<div class=""col-md-1 my-auto""><SELECT onChange=""UpdateInput(" & objRS("SystemParameterID") & ");"" class=""form-control"" name=""Active" & objRS("SystemParameterID") & """ id=""Active" & objRS("SystemParameterID") & """ >" & strActiveSelect & "</Select></div>" & _
				"<div class=""col-md-1 my-auto"">" & objRS("UpdatedByName") & "</div><div class=""col-md-1 my-auto"" title=""" & objRS("DateUpdated") & """>" & strDateUpdated & "<input type=""hidden"" class=""form-control"" id=""ID" & objRS("SystemParameterID") & """ name=""ID" & objRS("SystemParameterID") & """ >" & _
				"<input type=""hidden"" class=""form-control"" id=""ParameterName" & objRS("SystemParameterID") & """ name=""ParameterName" & objRS("SystemParameterID") & """ value=""" & objRS("ParameterName") & """></div></div>"
		
	objRS.Movenext
	
	Loop
	
	
				
objRS.Close

'Response.Write "<div class=""panel-content row"" ><div class=""form-row col-md-12""><button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='SystemParameters.asp'"";> Save</button></div></div>"

Response.Write "</div><input type=""text"" id=""TotalRecords"" name=""TotalRecords"" value=""" & x & """><div class=""row mt-3""><div class=""col-12 text-right""><button class=""btn btn-primary"">Save changes</button></div></div>"

End Sub



Public Sub DisplayTableDetails()

Dim x, y
Dim strActive
Dim strActive2
Dim strActiveSelect
Dim strDateUpdated
Dim strSQL
Dim strParameterType
Dim strHeader
Dim arrHeader(10)

	Response.write "<div class=""panel-content row""><div class=""mb-3 col-md-4""><h4>System Parameters</h4></div><div class=""mb-3 col-md-5"">" &  _
		"<div class=""btn-group btn-selector table-tabs-selector"" role=""group"" aria-label=""Basic example"">" &  _
		"<button type=""button"" data-target=""table-tabs"" data-type=""as-tabs"" class=""btn btn-outline-primary active"">" &  _
		"<i class=""fa fa-list""></i> View as Tabs</button>" &  _
		"<button type=""button"" data-target=""table-tabs"" data-type=""as-table"" class=""btn btn-outline-primary"">" &  _
		"<i class=""fa fa-table""></i> View as Table</button></div></div></div>" &  _
		"<div class=""panel-content row""><div class=""mb-3 col-md-4""><h6></h6></div></div>"
	
	Response.write  "<div id=""table-tabs"" class=""as-tabs""><ul class=""nav nav-tabs"" id=""myFiTab"" role=""tablist"">"'<li class=""nav-item"" role=""presentation"">" &  _
		'"<a class=""nav-link active"" id=""overview-tab"" data-toggle=""tab"" href=""#overview"" role=""tab"" aria-controls=""overview"" aria-selected=""true"">Application Details</a>" &  _
		'"</li><li class=""nav-item"" role=""presentation"">" &  _
		'"<a class=""nav-link"" id=""card-details-tab"" data-toggle=""tab"" href=""#card-details"" role=""tab"" aria-controls=""card-details"" aria-selected=""false"">Contact Details</a>" &  _
		'"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-limits-tab"" data-toggle=""tab"" href=""#my-limits"" role=""tab"" aria-controls=""my-limits"" aria-selected=""false"">Limits</a>" &  _
		'"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-system-tab"" data-toggle=""tab"" href=""#my-system"" role=""tab"" aria-controls=""my-system"" aria-selected=""false"">System Admin</a>" &  _
		'"</li>"
	
		
	strSQL = "SELECT [ParameterType] FROM qryCAPSSystemParameters WITH(NOLOCK) GROUP BY [ParameterType] ORDER BY [ParameterType]"

	objRS.Open strSQL,objCon

		Do Until objRS.EOF
			
			x = x + 1
			
			If x = 1 Then
				strActive = "active"
				strActive2 = "active show"
			Else
				strActive = ""
				strActive2 = ""
			End If
			
			Response.Write "<li class=""nav-item"" role=""presentation""><a class=""nav-link " & strActive & """ id=""my-"  &objRS("ParameterType") & "-tab"" data-toggle=""tab"" href=""#my-"  &objRS("ParameterType") & """ role=""tab"" aria-controls=""my-"  &objRS("ParameterType") & """ aria-selected=""false"">"  &objRS("ParameterType") & " Admin</a></li>"
			
			arrHeader(x) = "<div class=""tab-pane fade " & strActive2 & """ id=""my-"  &objRS("ParameterType") & """ role=""tabpanel"" aria-labelledby=""my-"  &objRS("ParameterType") & "-tab"">"
			
			objRS.Movenext
		Loop
	
		Response.Write "</ul><div class=""tab-content panel panel-light p-3"" id=""myFiTabContent"">"
		
		Response.Write "<div class=""row information-row""><div class=""col-sm-4""><h6>Parameter Name</h6></div><div class=""col-sm-5""><h6>Parameter value</h6></div>" & _
			"<div class=""col-md-1""><h6>Active</h6></div><div class=""col-md-1 my-auto""><h6>Updated By</h6></div>" & _
			"<div class=""col-md-1 my-auto""><h6>Date Updated</h6></div></div>"
			
	objRS.Close
	
	x = 0
	strActive = ""
	
	'Response.write "</ul><div class=""tab-content panel panel-light p-3"" id=""myFiTabContent"">" &  _
	'	"<div class=""tab-pane fade show active"" id=""overview"" role=""tabpanel"" aria-labelledby=""overview-tab"">"
		
	'Response.write strHeader
	
	strSQL = "SELECT * FROM qryCAPSSystemParameters WITH(NOLOCK) ORDER BY [ParameterType],[ParameterName]"

objRS.Open strSQL,objCon

	If NOT objRS.EOF Then
	
		'Response.Write "<div class=""panel-content row"" >" & _
		'	"<div class=""form-row col-md-2""><label style=""font-weight:bold;"">Parameter Name</label></div><div class=""form-row col-md-3""><label style=""font-weight:bold;"">Parameter Value</label></div>" & _
		'	"<div class=""form-row col-md-4""><label style=""font-weight:bold;"">Parameter Description</label></div><div class=""form-row col-md-1""><label style=""font-weight:bold;"">Active</label></div>" & _
		'	"<div class=""form-row col-md-1""><label style=""font-weight:bold;"">Updated By</label></div><div class=""form-row col-md-1""><label style=""font-weight:bold;"">Date Updated</label></div></div>"
			
'		Response.Write "<div class=""row information-row""><div class=""col-sm-4""><h6>Parameter Name</h6></div><div class=""col-sm-5""><h6>Parameter value</h6></div>" & _
'			"<div class=""col-md-1""><h6>Active</h6></div><div class=""col-md-1 my-auto""><h6>Updated By</h6></div>" & _
'			"<div class=""col-md-1 my-auto""><h6>Date Updated</h6></div></div>"

	End If
	
	
	strParameterType = ""
	
	Do until objRS.EOF
	
		x = x + 1
		
		If IsNull(objRS("Active")) OR objRS("Active") = "" Then
			strActive = ""
		Else
			strActive = objRS("Active")
		End If
		
		strActiveSelect = ""
		
		If strActive = "Y" Then strSelected = " SELECTED "
		
		strActiveSelect = strActiveSelect & "<option " & strSelected & " value=""Y"">Y</option>"
		
		If strActive = "N" Or strActive = "" Then 
			strSelected = " SELECTED "
		Else
			strSelected = ""
		End If
		
		strActiveSelect = strActiveSelect & "<option " & strSelected & " value=""N"">N</option>"
		
		If IsNull(objRS("DateUpdated")) OR objRS("DateUpdated") = "" Then
			strDateUpdated = ""
		Else
			If IsDate(objRS("DateUpdated")) Then
				strDateUpdated = FormatDateTime(objRS("DateUpdated"),vbShortDate)
			End If
		End If
		
		
		''''New June 2025 to check the Service Account DateUpdated as it needs to be reset every 12 months
		If objRS("ParameterName")="CAPSServiceAccountPassword" Then
			
			''Display a warning to the user if the Service Account is nearing expiry (12 months - the Service account password must be changed)
			If DateDiff("d",now(),strDateUpdated) < 30 Then
				Response.write "<div class=""content-body"" style=""position:absolute; top:10; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-warning alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
				"<span aria-hidden=""true"">&times;</span></button>" & _
				"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
				"<span>The CAPS Service Account is due to expire on: " & DateAdd("d",365,strDateUpdated) &  " (in " & DateDiff("d",now(),strDateUpdated) & " days)<br><a target=""_new"" href=""https://dsms.dpe.protected.mil.au/sm/ess/offeringPage/Service%20Account%20ResetUnlock?query=Service%20Account&TENANTID=863607281"">A Job must be logged to change it </a>" & _
				"</span></div></div></div>"
			End If
		
			''Display a danger to the user if the Service Account is past expiry (12 months - the Service account password must be changed)
			If DateDiff("d",now(),strDateUpdated) < 0 Then
				Response.write "<div class=""content-body"" style=""position:absolute; top:0; left:200; z-index:3000;  width:75%;""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
				"<span aria-hidden=""true"">&times;</span></button>" & _
				"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
				"<span>The CAPS Service Account expired on: " & DateAdd("d",365,strDateUpdated) &  " (in " & DateDiff("d",now(),strDateUpdated) & " days ago)<br><a target=""_new"" href=""https://dsms.dpe.protected.mil.au/sm/ess/offeringPage/Service%20Account%20ResetUnlock?query=Service%20Account&TENANTID=863607281"">A Job must be logged to change it </a>" & _
				"</span></div></div></div>"
			End If
		
		End If
		
		
		'If the Parameter Type has changed then write the tab details for displaying details by Parameter Type (for clicking on the tab links)
		If strParameterType <> objRS("ParameterType") Then
			'End the previous div if this is not the first chnage in system parameters
			'If strParameterType <> "" Then Response.Write "</DIV>"
			y = y + 1
			
			If y = 1 Then 
				'Response.write "<div class=""tab-pane fade show active"" id=""overview"" role=""tabpanel"" aria-labelledby=""overview-tab"">"
				'Response.Write "</div></div>"
			Else
			
				If y = 1 Then
				'Response.Write "</div></div>"
				Else
				Response.Write "</div>"
				End If
			End If
			
			Response.write arrHeader(y)
			'Response.Write "<div class=""tab-pane fade show"" id="""  &objRS("ParameterType") & """ role=""tabpanel"" aria-labelledby="""  &objRS("ParameterType") & "-tab"">"
			
		End If
		
		
		Response.write "<div class=""row information-row"" Style=""padding:10px;""><div class=""col-sm-4""><h6>" & objRS("SystemParameterID") & " - " & objRS("ParameterName") & "</h6><span class=""description"" style=""font-size:12px;"">" & objRS("ParameterDescription") & "</span></div>" & _
				"<div class=""col-sm-5 my-auto""><input onChange=""UpdateInput(" & objRS("SystemParameterID") & ");"" type=""text"" class=""form-control"" id=""ParameterValue" & objRS("SystemParameterID") & """ name=""ParameterValue" & objRS("SystemParameterID") & """ placeholder=""Parameter value"" value=""" & objRS("ParameterValue") & """></div>" & _
				"<div class=""col-md-1 my-auto""><SELECT onChange=""UpdateInput(" & objRS("SystemParameterID") & ");"" class=""form-control"" name=""Active" & objRS("SystemParameterID") & """ id=""Active" & objRS("SystemParameterID") & """ >" & strActiveSelect & "</Select></div>" & _
				"<div class=""col-md-1 my-auto"">" & objRS("UpdatedByName") & "</div><div class=""col-md-1 my-auto"" title=""" & objRS("DateUpdated") & """>" & strDateUpdated & "<input type=""hidden"" class=""form-control"" id=""ID" & objRS("SystemParameterID") & """ name=""ID" & objRS("SystemParameterID") & """ >" & _
				"<input type=""hidden"" class=""form-control"" id=""ParameterName" & objRS("SystemParameterID") & """ name=""ParameterName" & objRS("SystemParameterID") & """ value=""" & objRS("ParameterName") & """></div></div>" 

		

		'Response.Write "<div class=""row information-row"" Style=""padding:10px;""><div class=""col-sm-4""><h6>" & objRS("SystemParameterID") & " - " & objRS("ParameterName") & "</h6><span class=""description"" style=""font-size:12px;"">" & objRS("ParameterDescription") & "</span></div>" & _
		'		"<div class=""col-sm-5 my-auto""><input onChange=""UpdateInput(" & objRS("SystemParameterID") & ");"" type=""text"" class=""form-control"" id=""ParameterValue" & objRS("SystemParameterID") & """ name=""ParameterValue" & objRS("SystemParameterID") & """ placeholder=""Parameter value"" value=""" & objRS("ParameterValue") & """></div>" & _
		'		"<div class=""col-md-1 my-auto""><SELECT onChange=""UpdateInput(" & objRS("SystemParameterID") & ");"" class=""form-control"" name=""Active" & objRS("SystemParameterID") & """ id=""Active" & objRS("SystemParameterID") & """ >" & strActiveSelect & "</Select></div>" & _
		'		"<div class=""col-md-1 my-auto"">" & objRS("UpdatedByName") & "</div><div class=""col-md-1 my-auto"" title=""" & objRS("DateUpdated") & """>" & strDateUpdated & "<input type=""hidden"" class=""form-control"" id=""ID" & objRS("SystemParameterID") & """ name=""ID" & objRS("SystemParameterID") & """ >" & _
		'		"<input type=""hidden"" class=""form-control"" id=""ParameterName" & objRS("SystemParameterID") & """ name=""ParameterName" & objRS("SystemParameterID") & """ value=""" & objRS("ParameterName") & """></div></div>"
		
		strParameterType = objRS("ParameterType")
		
	objRS.Movenext
	
	Loop
	
objRS.Close

	'Response.Write "</div></div>"
	
	Response.Write "</div><input type=""hidden"" id=""TotalRecords"" name=""TotalRecords"" value=""" & x & """><div class=""row mt-3""><div class=""col-12 text-right""><button class=""btn btn-primary"">Save changes</button></div></div>"

End Sub



Public Sub DisplayMessages()

Dim strAction
Dim strStatus
Dim strAddress
Dim dteDateSubmitted
Dim dteDateReviewed
Dim strSQL

	strSQL = "SELECT * FROM qryCAPSCards WHERE CardID = '" & Session("CardID") & "'"

	objRS.Open strSQL,objCon
	
    Do Until objRS.EOF
		If isNull(objRS(9)) Then
			dblEmpCont = 0
		Else
			dblEmpCont = objRS(9)
		End If

		Select Case objRS("Status")
		
		Case  "Received"
			strAction = "<button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
		Case "Added To CS"

			strAction = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='CSToDiners.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
		
		Case "Submitted"
			strAction = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"

			strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Submitted to GCFO</button>"
		Case "Cancelled"
			strAction = "Cancelled - " & FormatDateTime(objRS("DateUpdated"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"

			strStatus  = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
		
		Case "GCFO Approved"
			strAction = "<button type=""button"" title=""Approved by GCFO"" class=""btn btn-secondary btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-check""></i>GCFO Approved</button>"
		
			strStatus  = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Approved by GCFO</button>"
		
		Case Else
			strAction = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
			'strAction = "Rejected"
			strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Submitted</button>"
		End Select

		strAddress = Trim(objRS("Address1")) & " " & Trim(objRS("Address2")) & " " & Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode"))
		
		If len(strAddress) > 15 Then strAddress = left(strAddress,15) & "..."
		
		If IsNull(objRS("DateLoaded")) Then
			dteDateSubmitted = ""
		Else
			dteDateSubmitted = FormatDateTime(objRS("DateLoaded"),vbShortDate)
		End If
		
		If IsNull(objRS("PMLoadDate")) Then
			dteDateReviewed = ""
		Else
			dteDateReviewed = FormatDateTime(objRS("PMLoadDate"),vbShortDate)
		End If
		
		Response.Write "<div class=""panel-content row"" >" & _
				"<div class=""form-row col-md-4""><label>Application ID</label></div><div class=""form-row col-md-8""><input type=""text"" id=""CardID"" name=""CardID"" class=""form-control input-md"" value=""" & objRS("CardID") & """></div></div>" & _
				"<div class=""panel-content row"" >" & _
				"<div class=""form-row col-md-4""><label>Status</label></div><div class=""form-row col-md-8""><input type=""text"" id=""Status"" name=""Status"" class=""form-control input-md"" value=""" & objRS("Status") & """></div></div>" & _
				"<div class=""panel-content row"" >" & _
				"<div class=""form-row col-md-4""><label>Credit Limit</label></div><div class=""form-row col-md-8""><input type=""text"" id=""CreditLimit"" name=""CreditLimit"" class=""form-control input-md"" value=""" & objRS("CreditLimit") & """></div></div>" & _
				"<div class=""panel-content row"" >" & _
				"<div class=""form-row col-md-4""><label>Date Submitted</label></div><div class=""form-row col-md-8""><input style=""border: 0px; font-weight:bold;"" type=""text"" id=""DateSubmitted"" name=""DateSubmitted"" class=""form-control input-md"" value=""" & objRS("DateSubmitted") & """></div></div>" & _
				"<div class=""panel-content row"" >" & _
				"<div class=""form-row col-md-4""><label>Group</label></div><div class=""form-row col-md-8""><input type=""text"" id=""ReportGroup"" name=""ReportGroup"" class=""form-control input-md"" value=""" & objRS("ReportGroup") & """></div></div>" & _
				"<div class=""panel-content row"" >" & _
				"<div class=""form-row col-md-4""><label>Name On Card</label></div><div class=""form-row col-md-8""><input type=""text"" id=""NameOnCard"" name=""NameOnCard"" class=""form-control input-md"" value=""" & objRS("NameOnCard") & """></div></div>" & _
				"<div class=""panel-content row"" >" & _
				"<div class=""form-row col-md-4""><label>Address 1</label></div><div class=""form-row col-md-8""><input type=""text"" id=""Address1"" name=""Address1"" class=""form-control input-md"" value=""" & objRS("Address1") & """></div></div>" & _
				"<div class=""panel-content row"" >" & _
				"<div class=""form-row col-md-4""><label>Address 2</label></div><div class=""form-row col-md-8""><input type=""text"" id=""Address2"" name=""Address2"" class=""form-control input-md"" value=""" & objRS("Address2") & """></div></div>"
		
		objRS.Movenext
	Loop
	
				
objRS.Close

End Sub

Public Sub LoadMessages()
'Procedure to load any messages relating to the application
Dim strSQL
Dim strPerson

	strSQL = "SELECT * FROM qryCAPSMessage WITH(NOLOCK) WHERE [Object] = 'Card' AND [ObjectID] = '" & Session("CardID") & "'"

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

Public Sub SaveParameter(lngSystemParameterID,x,y)

'Dim lngSystemParameterID
Dim intRecord
Dim strParameterName

'If Session("MessageID") = "" or IsNull(Session("MessageID")) Then
'	lngSystemParameterID = 0
'Else
'	lngSystemParameterID = Session("MessageID")
'End If

'response.write Request.Form("ParameterValue"&lngSystemParameterID) & " - "
'response.write lngSystemParameterID & " - " & Request.Form("Active"&lngSystemParameterID) & " s=" & y


	With objCmd

		.CommandType = 4
		.CommandText = "spCAPSSystemParametersUpdate"

		If y = 1 Then
		.Parameters.Append objCmd.CreateParameter("SystemParameterID", adInteger)
		'.Parameters.Append objCmd.CreateParameter("ParameterName", adVarChar, adParamInput,50)
		'.Parameters.Append objCmd.CreateParameter("ParameterDescription", adVarChar, adParamInput, 300)
		.Parameters.Append objCmd.CreateParameter("ParameterValue", adVarChar, adParamInput, 200)
		'.Parameters.Append objCmd.CreateParameter("ParameterType", adVarChar, adParamInput, 20)
		.Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput, 1)
		.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
		.Parameters.Append objCmd.CreateParameter("SystemParameterIDOutput", adInteger, adParamOutput)
		End If
		
		.Parameters("SystemParameterID") = lngSystemParameterID
		'.Parameters("ParameterName") = Request.Form("ParameterName"&x)
		'.Parameters("ParameterDescription") = strParameterDescription
		.Parameters("ParameterValue") = Request.Form("ParameterValue"&lngSystemParameterID)
		'.Parameters("ParameterType") = strParameterType
		.Parameters("Active") = Request.Form("Active"&lngSystemParameterID)
		.Parameters("UpdatedBy") = Session("UserID")
		.ActiveConnection = objCon
		 
	End With
   
	objCmd.Execute        
  
	'Return the result of the Save Function.
	intRecord = objCmd.Parameters.Item("SystemParameterIDOutput") 
 
	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">System Parameter " & Request.Form("ParameterName"&lngSystemParameterID) & " NOT Saved! Error with Parameter as it doesn't exist. See Admin.</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">System Parameter " & Request.Form("ParameterName"&lngSystemParameterID) & " Saved!</div>"
	End If
	
End Sub



Set objRS = Nothing
Set objCon = Nothing

%>