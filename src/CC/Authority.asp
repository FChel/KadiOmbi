
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=../ADOVBS.inc -->
<%

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"


'Description:	Authority Management screen for users to assign authority
'Author:		MG
'Date:			Janaury 2020

	Response.Expires = -1500	

Dim objCon
Dim objRS
Dim objRS1
Dim objRS2
Dim objCmd

Dim strForeColour
Dim intMode
Dim dblTotal
Dim intFinYearPart1
Dim intFinYearPart2
Dim arrHeadings(5)
Dim strCostCentreName
Dim strVersionName
Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL

Dim intEmployeeID
Dim dteEndDate

Dim lngApplicationID
Dim strEmployeeID

Dim strPostCode
Dim dteDateReceived
Dim strStatus
Dim strReviewedBy
Dim dteDateReviewed
Dim lngCreditLimit
Dim strFirstName
Dim strLastName
Dim strAddress1
Dim strSearchTerm
Dim strName
Dim strUserAuthority

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objRS1 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")
	
    objCon.Open Session("DBConnection")	

    Session("CurrentPage") = "CC/Authority.asp"

	If IsNull(Session("ApplicationID")) OR Session("ApplicationID") = "" Then Session("ApplicationID")= 0

	If Not IsEmpty(Request.QueryString("StyleSheet")) Then
		Session("StyleSheet") = Request.QueryString("StyleSheet")
		
		Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""" & Request.QueryString("StyleSheet") & """>"
	Else
		If IsEmpty(Session("StyleSheet")) Then Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""../CAPSStyle.css"">"
	End If

		Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""../CAPSStyle.css"">"
		
If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	Session("ApplicationID") = Request.QueryString("ApplicationID")
End If

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	Session("EmployeeID") = Request.QueryString("EmployeeID")
	Session("CarParkingID") = 0
End If

If Not IsEmpty(Request.QueryString("Filter")) Then
	Session("Filter") = Request.QueryString("Filter")
End If

If Not IsEmpty(Request.QueryString("TransactionType")) Then
	Session("TransactionType") = Request.QueryString("TransactionType")
End If
 Session("InputSheetID") = 1
 

If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "Reject" Then
		Call RejectApplication()
	End If
	
	If Request.QueryString("Action") = "Release" Then
		Call ReleaseApplication()
	End If
	
	If Request.QueryString("Action") = "Email" Then
		'Call the procedure to email
		Response.write EMailSMTP("CAPS@vbmrsn05.drn.mil.au", "michael.giacomin@defence.gov.au", "CAPS", "Hi Michael, this is your email", "", "", "vbmrsn05.drn.mil.au", "25")
		'EMailSMTP( myFrom, myTo, mySubject, myTextBody, myHTMLBody, myAttachment, mySMTPServer, mySMTPPort )
	End If
End If

'Execute Action
If Request.QueryString("Action") = "Save" Then

   If Session("StatusID") = 1 Then
        'Do not allow Read Only users to make changes
        If Session("UserTypeID") = 4 Then
            strMessage = "NOT SAVED. You are a READ ONLY User and cannot make any changes."
        Else
            Call SaveCarParking()
        End If
   Else
        strMessage = "Budget is closed, no changes can be made!"
   End If
End If

If Request.QueryString("Action") = "Delete" Then
    If Session("StatusID") = 1 Then
        Call DeleteData(Request.QueryString("GeneralExpenseID"))
    Else
          Response.Write "&nbsp;&nbsp;<img src=""../images/warning.gif"" /><B><FONT Color=""Red"">&nbsp;&nbsp;WARNING - BUDGET IS NOT OPEN, CHANGES CANNOT BE MADE.</FONT></B>" 
          strMessage = "<FONT Color=""Red""><B>BUDGET IS NOT OPEN, CHANGES CANNOT BE MADE.</B></FONT>"
          strMessageIcon = "<img src=""../images/warning.gif"" />"
    End If
End If

If Not IsEmpty(Request.QueryString("SearchTerm")) Then
	strSearchTerm = Request.Form("UserSearch")
	Session("Filter") = "UserSearch"
	'Request.QueryString("Action")
	'Response.write "SearchTerm=" & Request.form("UserSearch") &  " " & Request.QueryString("Action")
End If

If Not IsEmpty(Request.QueryString("SendRequest")) Then
	
	Call SaveAuthority(Request.QueryString("SendRequest"))
End If

If Not IsEmpty(Request.QueryString("ApproveRequest")) Then
	Call ApproveAuthority(Request.QueryString("ApproveRequest"))
End If

If Not IsEmpty(Request.QueryString("RemoveRequest")) Then
	Call RemoveAuthority(Request.QueryString("RemoveRequest"))
End If

If Not IsEmpty(Request.QueryString("ChangeLogin")) Then
	Call ChangeLogin(Request.QueryString("ChangeLogin"))	
End If

If Session("UserIDAuthority") <> Session("UserID") Then 
	strUserAuthority = "<button type=""button"" class=""btn btn-secondary btn-xs"" data-Loginid=""" & Session("UserIDAuthority") & """ data-UserID=""" & Session("UserIDAuthority") & """  onClick=""LoginAs(this);"" Title=""Click to change back to your Login""><i class=""fa fa-globe""></i> Return To Your Login</button>"
End If

  Call LoadDetails()
  
  'Response.Write "AuthUser=" & Session("UserIDAuthority")
  'Response.Write "</br> User=" & Session("EmployeeID")
%>

<html>
<head>

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
    { self.location = 'HomeCC.asp'; }
}

setTimeout( 'ShowTimeoutWarning();', 1080000 );

function ShowTimeoutWarning () {     
    window.alert( "********** Warning! **********' \n \n 'You will be automatically logged out in 2 minutes unless you change screens, Close or Save!" ); 
}

function DeleteData(GEXPID) {
   
        if (window.confirm('Would you like to DELETE the selected record?') == true) {

            self.location = "Loans.asp?Action=Delete&GeneralExpenseID=" + GEXPID;
        }
        
}

   
function UserSearch(cb) {

	alert(asas);
	self.location = "Authority.asp?Action=UserSearch&SearchID=" + cb.value;

}

function SendRequest() {

	self.location = "Authority.asp?SendRequest=" + document.getElementById('UserIDTo').value;

}

function ApproveRequest() {

	self.location = "Authority.asp?ApproveRequest=" + document.getElementById('AuthorityIDApprove').value;

}

function RemoveRequest() {

	self.location = "Authority.asp?RemoveRequest=" + document.getElementById('AuthorityIDRemove').value;

}

function OpenSs(cb) {

	varTextLen=cb.name.length*8
	
	document.getElementById('NameTo').value=cb.name;
	document.getElementById('NameTo1').value=cb.name;
	document.getElementById('NameTo2').value=cb.name;
	document.getElementById('NameTo1').style.width=+varTextLen+"px";
	document.getElementById('NameTo2').style.width=+varTextLen+"px";
	
	var id = cb.getAttribute('data-id');
	document.getElementById('IDTo').value=id;
	
	var Userid = cb.getAttribute('data-userid');
	document.getElementById('UserIDTo').value=Userid;
}

function ApproveAuth(cb) {

	varTextLen=cb.name.length*8
	
	document.getElementById('NameApprove').value=cb.name;
	document.getElementById('NameApprove').style.width=+varTextLen+"px";
	
	var id = cb.getAttribute('data-Authid');
	document.getElementById('AuthorityIDApprove').value=id;
	
	var Userid = cb.getAttribute('data-userid');
	document.getElementById('ApproveUserIDTo').value=Userid;
}

function RemoveAuth(cb) {

	var nameid = cb.getAttribute('data-nameid');
	varTextLen=nameid.length*8
	
	//document.getElementById('NameRemove').value=cb.name;
	document.getElementById('NameRemove').style.width=+varTextLen+"px";
	
	var id = cb.getAttribute('data-Authid');
	document.getElementById('AuthorityIDRemove').value=id;
	
	//var nameid = cb.getAttribute('data-nameid');
	document.getElementById('NameRemove').value=nameid;
}

function LoginAs(cb) {

	var nameid = cb.getAttribute('data-Loginid');
	var Userid = cb.getAttribute('data-UserID');
	self.location = "Authority.asp?ChangeLogin=" + Userid;
}


</script>

</head>
<body >

<main class="main py-3">
    <div class="container">

<!-- Request Modal -->
<div class="modal fade" id="exampleModalCenter" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle" style="font-weight:bold;">CAPS - Assign Authority in CAPS Form</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <span style="font-weight:bold;">NOTE: This will send a request to <input type="text" name="NameTo1" id="NameTo1" style="border:0;" READONLY> requesting access to login as them in CAPS only.</span><br><br>
		<div class="col-md-12">
			<table class="table table-bordered table-hover CAPS">
						
					<tr>
						<td>Request To:</td><td><input type="text" name="NameTo" id="NameTo" class="form-control input-md" style="border:0;" READONLY></td>
						<td><input type="text" name="IDTo" id="IDTo" class="form-control input-md" style="border:0;" READONLY></td>
						<td><input type="text" name="UserIDTo" id="UserIDTo" HIDDEN></td>
					</tr>
					<tr>
						<td>Request From:</td><td><input type="text" name="NameFrom" id="NameFrom" class="form-control input-md" style="border:0;" READONLY></td>
						<td><input type="text" name="IDFrom" id="IDFrom" class="form-control input-md" style="border:0;" READONLY></td>
						<td><input type="text" name="UserIDFrom" id="UserIDFrom" HIDDEN></td>
					</tr>
						
			</table>
			
		</div>
		
		<span>Dear <input type="text" name="NameTo2" id="NameTo2" style="border:0;" READONLY>, I would like to request access to login and perform actions on your behalf <span style="font-weight:bold; color:red;">in CAPS only.</span></span><br>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-close"></i> Close</button>
        <button type="button" class="btn btn-primary" onClick="SendRequest();"><i class="fa fa-envelope"></i> Send</button>
      </div>
    </div>
  </div>
</div>
<!-- End Request Modal -->

<!-- Approve Modal -->
<div class="modal fade" id="ModalApprove" tabindex="-1" role="dialog" aria-labelledby="ModalApprove" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalApproveTitle" style="font-weight:bold;">CAPS - Assign Authority in CAPS Form</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <span>Approve <input type="text" name="NameApprove" id="NameApprove" style="border:0; font-weight:bold;" READONLY> to login as YOU and perform actions on your behalf in CAPS only.</span><br><br>
		<input type="text" name="ApproveUserIDTo" id="ApproveUserIDTo" HIDDEN><input type="text" name="ApproveUserIDFrom" id="ApproveUserIDFrom" HIDDEN><input type="text" name="AuthorityIDApprove" id="AuthorityIDApprove" HIDDEN>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-close"></i> No</button>
        <button type="button" class="btn btn-primary" onClick="ApproveRequest();"><i class="fa fa-check"></i> Yes</button>
      </div>
    </div>
  </div>
</div>
<!-- End Approve Modal -->

<!-- Remove Modal -->
<div class="modal fade" id="ModalRemove" tabindex="-1" role="dialog" aria-labelledby="ModalRemove" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalApproveTitle" style="font-weight:bold;">CAPS - Assign Authority in CAPS Form</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <span>Remove authority for <input type="text" name="NameRemove" id="NameRemove" style="border:0; font-weight:bold;" READONLY> to login as YOU and perform actions on your behalf in CAPS only.</span><br><br>
		<input type="text" name="AuthorityIDRemove" id="AuthorityIDRemove" HIDDEN>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-close"></i> No</button>
        <button type="button" class="btn btn-primary" onClick="RemoveRequest();"><i class="fa fa-check"></i> Yes</button>
      </div>
    </div>
  </div>
</div>
<!-- End Remove Modal -->

<form action="Authority.asp?Action=UserSearch&SearchTerm=" method="POST" id="frm" name="frm" class="form-inline">
<div class="content-wrapper">
    <div class="container-fluid">
	
		<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-10">
				<h2 class="text-left">Authority for <%=Session("UserName")%> <i class="fa help-tooltip fa-question-circle" data-toggle="tooltip" title="Current Authorities are listed on the left-hand side. Authority can be added from the right-hand side."></i></h2>
			</div>
		</div>
      </section>
	  
		<div class="row">
			<div class="col-md-8">
				<div class="card">
					<div class="card-header card-header-icon">
					<i class="fa fa-address-card"></i> Manage User Authority in CAPS &nbsp;&nbsp; <%=strUserAuthority %></div>
					<div class="card-body">
					<table class="table table-bordered table-hover table-compact">
						<thead class="CAPS">
							<tr>
								<th>Employee ID</th><th>Name</th><th>Start Date</th><th>End Date</th><th>Requested</th><th>Active</th><th>Action</th><th>Log In</th>
							</tr>
						</thead>
						<tbody class="CAPS2">
							
						<%Call DisplayTableDetails("Owner")%>
						</tbody>
					</table>
					
					</div>
		  
				</div>
			</div>
			
			<div class="col-md-4">
				<div class="card">
					<div class="card-header card-header-icon">&nbsp;&nbsp;&nbsp;
					<i class="fa fa-user"></i> Search for Users to Add &nbsp;<input type="text" name="UserSearch" id="UserSearch" style="width:200px;" placeholder="..Enter name or Employee ID" value="<%=strSearchTerm%>"><button type="submit" style="height:26px;"><i class="fa fa-search fa-1.5x"></i></button><% If Not IsNull(strSearchTerm) Then Response.Write "<span style=""color:red; font-weight:bold;""> &nbsp;" & strSearchTerm & "</span>" %></div>

					<div class="card-body">
					<table class="table table-bordered table-hover table-compact">
						<thead class="CAPS2">
							<tr>
								<th>Employee ID</th><th>Name</th><th></th>
							</tr>
						</thead>
						<tbody>
							
							
							<%Call DisplayTableDetails("All")%>
						</tbody>
					</table>
					
					</div>
				</div>
			</div>
		</div>
		</form>
	</div>
</div>



<%=strMessageIcon %>
 

</form>
</div>
</main>

<!-- #Include file=CAPSFooter.asp -->
	
</body>
</html>
<%

Public Sub DisplayTableDetails(strType)
Dim y
Dim lngEmployeeID
Dim strName
Dim dteStartDate
Dim dteEndDate
Dim strWhere
Dim arrNames
Dim strFNameSearch
Dim strLNameSearch
Dim lngUserID
Dim dteAuthorised
Dim strAction
Dim strActive
Dim strRequested
Dim dteRequested
Dim strAuthorised
Dim strStatus 
Dim strLogin
Dim strUserLogin

If Session("Filter") = "UserSearch" Then
	If IsNull(strSearchTerm) or IsEmpty(strSearchTerm) Then
	Else
		'If the user has entered a search term with a space the assume this is a first and last name so search on that only
		If Instr(1,strSearchTerm," ")>0 Then
			arrNames = Split(strSearchTerm," ")
			strFNameSearch = arrNames(0)
			strLNameSearch = arrNames(1)
			
			strWhere = " AND ([FName] Like '%" & strFNameSearch & "%' AND [LName] Like '%" & strLNameSearch & "%')"
		Else
			strWhere = " AND ([FName] Like '%" & strSearchTerm & "%' OR [LName] Like '%" & strSearchTerm & "%' OR [UserLogon] Like '%" & strSearchTerm & "%')"
		End If
	End If
End If

If strType = "Owner" Then
	strSQL = "SELECT Top 100 * FROM qryCAPSAuthorityUsers WHERE [AuthorityID] Is Not Null AND [AuthUserID] = " & Session("UserIDAuthority") & " " '& strWhere
	'strSQL = "SELECT Top 100 * FROM qryCAPSAuthorityUsers WHERE [AuthorityID] Is Not Null AND [AuthEmployeeID] = '" & Session("UserIDAuthority") & "' " '& strWhere
Else
	strSQL = "SELECT Top 100 * FROM qryCAPSAuthorityUsers WHERE [AuthorityID] Is Null " & strWhere
End If

objRS.Open strSQL,objCon
    y = 0
    	
    Do until objRS.EOF 
		If isNull(objRS("EmployeeID")) Then
			lngEmployeeID = 0
		Else
			lngEmployeeID = objRS("EmployeeID")
		End If
		
		If isNull(objRS("UserID")) Then
			lngUserID = 0
		Else
			lngUserID = objRS("UserID")
		End If

		If isNull(objRS("FName")) OR objRS("LName") = "" Then
			strName = ""
		Else
			strName = objRS("FName")  & " " & objRS("LName")
			'dteDateReviewed = FormatDateTime(objRS(16),vbShortDate)
		End If
		
		If isNull(objRS("StartDate")) OR objRS("StartDate") = "" Then
			dteStartDate = ""
		Else
			dteStartDate = FormatDateTime(objRS("StartDate"),vbShortDate)
		End If
		
		If isNull(objRS("EndDate")) OR objRS("EndDate") = "" Then
			dteEndDate = ""
		Else
			dteEndDate = FormatDateTime(objRS("EndDate"),vbShortDate)
		End If
		
		If isNull(objRS("Authorised")) OR objRS("Authorised") = "" Then
			dteAuthorised = ""
		Else
			dteAuthorised = FormatDateTime(objRS("Authorised"),vbShortDate)
		End If
		
		If isNull(objRS("Status")) OR objRS("Status") = "" Then
			strStatus = ""
		Else
			strStatus = objRS("Status")
		End If
		
		If isNull(objRS("UserLogon")) OR objRS("UserLogon") = "" Then
			strUserLogin = ""
		Else
			strUserLogin = objRS("UserLogon")
		End If
		
		
		If isNull(objRS("Requested")) OR objRS("Requested") = "" Then
			dteRequested = ""
		Else
			'If strStatus = "Requested" Then
			If objRS("Requested") = "" or objRS("Requested") = "01/01/1900" Then
				dteRequested = ""
				strRequested = ""
			Else
				dteRequested = FormatDateTime(objRS("Requested"),vbShortDate)
				'strRequested = "Title=""Requested"""
				strRequested = "Title=""Requested " & DateDiff("d",objRS("Requested"),now()) & " day's ago"""
			End If
			
		End If

		If strStatus = "Approved" Then
			strLogin = "<button type=""button"" class=""btn btn-secondary btn-xs"" data-Loginid=""" & strUserLogin & """ data-UserID=""" & lngUserID & """ data-nameid=""" & strName & """ onClick=""LoginAs(this);"" Title=""Click to Login as " & strName & """><i class=""fa fa-globe""></i> Login As</button>"
		Else
			strLogin = ""
		End If
		
		If strType = "Owner" Then
			'If strStatus = "Approved" Then
			If dteAuthorised = "" or left(dteAuthorised,9) = "1/01/1900" Then
								
				strAction = "<button type=""button"" class=""btn btn-warning btn-xs"" data-toggle=""modal"" data-target=""#ModalApprove"" data-Authid=""" & objRS("AuthorityID") & """  data-name=""" & strName & """ onClick=""ApproveAuth(this);"" Title=""Not Active - Awaiting Your Approval""><i class=""fa fa-plus""></i> Approve</button>"
				strActive = "<span class=""badge badge-pill badge-danger"" Title=""Not Active - Awaiting Your Approval""><i class=""fa fa-times""></i></span>"
			Else
				If strStatus = "Removed" Then	
					strAction = "<button type=""button"" class=""btn btn-warning btn-xs"" data-toggle=""modal"" data-target=""#ModalApprove"" data-Authid=""" & objRS("AuthorityID") & """  data-name=""" & strName & """ onClick=""ApproveAuth(this);"" Title=""Not Active - Awaiting Your Approval""><i class=""fa fa-plus""></i> Approve</button>"
					strActive = "<span class=""badge badge-pill badge-danger"" Title=""Not Active (Removed)""><i class=""fa fa-times""></i></span>"
				Else
					
					strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" data-toggle=""modal"" data-target=""#ModalRemove"" data-Authid=""" & objRS("AuthorityID") & """  data-nameid=""" & strName & """ onClick=""RemoveAuth(this);"" Title=""Active - Click to Remove Authority""><i class=""fa fa-minus""></i> Remove</button>"
					strActive = "<span class=""badge badge-pill badge-success"" Title=""Currently Active (Authorised - " & dteAuthorised & ")""><i class=""fa fa-check""></i></span>"
				End If
				
			End If
			
			Response.Write "<tr><td style=""font-size:13px;"">" & lngEmployeeID & "</td><td style=""font-size:13px;"">" & strName & "</td><td style=""font-size:13px;"">" & dteStartDate & "</td><td style=""font-size:13px;"">" & dteEndDate & "</td>" & _
						"<td style=""font-size:13px;"" " & strRequested & ">" & dteRequested & "</td><td style=""text-align:center; vertical-align:middle;"">" & strActive & "</td>" & _
						"<td style=""text-align:center;"">" & strAction & "</td><td style=""text-align:center;"">" & strLogin & "</td></tr>"
		Else
		Response.Write "<tr><td style=""font-size:13px;"">" & lngEmployeeID & "</td><td style=""font-size:13px;"">" & strName & "</td>" & _
						"<td style=""text-align:center;""><button type=""button"" class=""btn btn-success btn-xs openDialog"" data-toggle=""modal"" data-target=""#exampleModalCenter"" data-id=""" & lngEmployeeID & """ data-userid=""" & lngUserID & """ name=""" & strName & """ onClick=""OpenSs(this);""><i class=""fa fa-plus""></i> Add</button></td></tr>"
			
		End If
			y = y + 1
			
		objRS.movenext
	Loop

	If strType = "Owner" Then
		response.write "<TR><TD colspan=""5"" Style=""font-weight:bold;"">Total</TD>" & _
					"<TD colspan=""3"" style=""text-align:center; font-weight:bold;"">" & y & "</TD></TR>"
	Else
		response.write "<TR><TH colspan=""2"">Total</TH>" & _
					"<TH colspan=""1"" style=""text-align:center;"">" & y & "</TH></TR>"
	End If
	
objRS.Close

End Sub


Sub LoadDetails()

       'Description:	Loads Position details into page if applicable.
		objRS.Open "SELECT * FROM tblApplication WHERE ApplicationID = " & Session("ApplicationID") & "",objCon

			If Not objRS.EOF Then
               
				lngApplicationID = objRS("ApplicationID")
				strEmployeeID = objRS("EmployeeID")
				
    		Else
				Session("ApplicationID") = 0
			  	lngApplicationID = 0'objRS("ApplicationID")
				strEmployeeID = ""
				
           End If

		objRS.Close
	
End Sub



Sub SaveAuthority(strEIDOwner)

Dim lngAuthorityID

	If IsNull(Request.Form("AuthorityID")) or Request.Form("AuthorityID") = "" Then
		lngAuthorityID = 0
	Else
		lngAuthorityID =  Request.Form("AuthorityID")
	End If
	
  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSAuthoritySave"

			.Parameters.Append objCmd.CreateParameter("AuthorityID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("UserIDOwner", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("UserIDAuth", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Requested", adDate, adParamInput)                
			.Parameters.Append objCmd.CreateParameter("Authorised", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("StartDate", adDate, adParamInput) 
			.Parameters.Append objCmd.CreateParameter("EndDate",adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Ongoing", adInteger, adParamInput) 
			.Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput, 20) 
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("AuthorityIDOutput", adInteger, adParamOutput)
			
			.Parameters("AuthorityID") = lngAuthorityID
			.Parameters("UserIDOwner") = strEIDOwner
			.Parameters("UserIDAuth") = Session("UserIDAuthority")					
			.Parameters("Requested") = Now()
			.Parameters("Authorised") = "01-Jan-1900"'Request.Form("Authorised")
			.Parameters("StartDate") = "01-Jan-1900"'dteStartDate
			.Parameters("EndDate") = "01-Jan-1900"'dteEndDate             
			.Parameters("Ongoing") = 0'Request.Form("Ongoing") 
			.Parameters("Status") = "Requested"
			.Parameters("UpdatedBy") = Session("UserIDAuthority")
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		lngAuthorityID = objCmd.Parameters.Item("AuthorityIDOutput")
		
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
		strMessageColour = "Black"
          
	
End Sub

Sub ApproveAuthority(lngAuthorityID)

	If IsNull(lngAuthorityID) or lngAuthorityID = "" Then
		lngAuthorityID = 0
	End If
	
  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSAuthorityApprove"

			.Parameters.Append objCmd.CreateParameter("AuthorityID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("AuthorityIDOutput", adInteger, adParamOutput)
			
			.Parameters("AuthorityID") = lngAuthorityID
			.Parameters("UpdatedBy") = Session("UserIDAuthority")
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		lngAuthorityID = objCmd.Parameters.Item("AuthorityIDOutput")
		
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
		strMessageColour = "Black"
          

End Sub

Sub RemoveAuthority(lngAuthorityID)

	If IsNull(lngAuthorityID) or lngAuthorityID = "" Then
		lngAuthorityID = 0
	End If
	
  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSAuthorityRemove"

			.Parameters.Append objCmd.CreateParameter("AuthorityID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("AuthorityIDOutput", adInteger, adParamOutput)
			
			.Parameters("AuthorityID") = lngAuthorityID
			.Parameters("UpdatedBy") = Session("UserIDAuthority")
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		lngAuthorityID = objCmd.Parameters.Item("AuthorityIDOutput")
		
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
		strMessageColour = "Black"
          

End Sub

Public Sub RejectApplication()

	strSQL = "UPDATE tblApplication SET Status = 'Rejected' WHERE ApplicationID = " & Session("ApplicationID") & ""
	
	objCon.Execute strSQL
	
End Sub

Public Sub ReleaseApplication()

Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spApplicationToCS"

			.Parameters.Append objCmd.CreateParameter("ApplicationID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CSToDinernIDOutput", adInteger, adParamOutput)
			
			.Parameters("ApplicationID") = Session("ApplicationID")
			.Parameters("UpdatedBy") = Session("UserIDAuthority")
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CSToDinernIDOutput") 
	 
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" /> " & intRecord & " Added to CS"
		strMessageColour = "Black"
		
End Sub

Public Sub AddTask(strUserFrom, lngUserFrom, lngUserTo)

	strSQL = "INSERT tblTasks (TaskTypeID,TaskSubject,TaskInstructions,AssignedTo,AssignedBy,TaskDateRaised,TaskPriorityID,UpdatedBy,DateUpdated) " & _
			"VALUES (1,""Authority Request from " & strUserFrom & """,""A request to assign your authority within CAPS has been made by " & strUserFrom & ". Please approve if you authorise this""," & _
				"lngUserFrom,lngUserTo,Now(),1,""" & Session("UserIDAuthority") & """,now())"
	
	objCon.Execute strSQL
	
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


Public Sub ChangeLogin(strNewLogin)
'Changes the current login to the one selected
'response.write "SELECT * FROM qryUserLogon WHERE UserLogon = '" & strNewLogin & "'"
Dim strSQL

'If the login request is for the Main user logged in then use the UserID
'If strType = "Auth" Then
	'strSQL = "SELECT * FROM qryUserLogon WHERE EmployeeID = '" & strNewLogin & "'"
	strSQL = "SELECT * FROM qryUserLogon WHERE UserID = " & strNewLogin & ""
'Else
'	strSQL = "SELECT * FROM qryUserLogon WHERE UserLogon = '" & strNewLogin & "'"
'End If

'Reset all the assigned variables (global variables)
Session("CMSUserApplication") = ""
Session("LimitChangeCardID") = 0
Session("ApplicationEmployeeID") = 0
Session("UserIDAuthority") = 0		 
			
			
objRS.Open strSQL,objCon,3,1

	If Not objRS.EOF Then
		'Session("UserIDAuthority") = Session("EmployeeID")
		
		Session("Logon") = objRS("UserLogon")
		Session("EmployeeID") = objRS("EmployeeID")
		Session("UserID") = objRS("UserID")
		Session("UserName") = objRS("FName") & " " & objRS("LName")
		'Session("UserIDAuthority") = objRS("EmployeeID")
	Else
		
   End If

objRS.Close
		
End Sub

Public Sub SendEmail(strEmailType)


 Dim ObjSendMail
 Set ObjSendMail = CreateObject("CDO.Message") 
      
 'This section provides the configuration information for the remote SMTP server.
      
 ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using the network (SMTP over the network).
 ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.iinet.net.au"'"mail.grapevine.net.au"'"mail.bizbudg.com"'"mail.bizbudg.com.au"
 ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
 ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
 ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
      
 ' If your server requires outgoing authentication uncomment the lines bleow and use a valid email address and password.
 ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'basic (clear-text) authentication
 ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "michael.giacomin@defence.gov.au"
 ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "1234!"
      
 ObjSendMail.Configuration.Fields.Update
      
 'End remote SMTP server configuration section==

If strEmailType = "User" Then
 ObjSendMail.To = Request.Form("email")
Else
 ObjSendMail.To = "michael.giacomin@isidore.com.au"
End If

 ObjSendMail.Subject = "Isidore Contact Email"
 ObjSendMail.From = "noreply@isidore.com"
      
 ' we are sending a text email.. simply switch the comments around to send an html email instead
 'ObjSendMail.HTMLBody = "this is the body"
 
If strEmailType = "User" Then
    ObjSendMail.TextBody = "Hi " & Request.Form("name") & chr(13) & chr(13) & "Thanks for your enquiry with Isidore." & chr(13) & chr(13) & _
                        "We will be in touch with you as soon as possible" & chr(13) & chr(13) & _
                        "Thanks," & chr(13) & chr(13) & _
                        "Isidore Support"
Else                                        
    ObjSendMail.TextBody = "Hi Isidore Admin," & chr(13) & chr(13) & "Person name: " & Request.Form("name") & " " & Request.Form("email") & " " & chr(13) & chr(13) & _
                        "Phone: " & Request.Form("phone") & " " & chr(13) & chr(13) & _
                        "Message: " & Request.Form("message") & " " & chr(13) & chr(13) & _
                        "<-- END MESSAGE" & chr(13) & chr(13) & _
                        "Has just submitted a contact form on the website " & chr(13) & chr(13) & _
                        "Thanks," & chr(13) & chr(13) & _
                        "Isidore Support"
End If
      
 ObjSendMail.Send
      
 Set ObjSendMail = Nothing 

End Sub


Function EMailSMTP( myFrom, myTo, mySubject, myTextBody, myHTMLBody, myAttachment, mySMTPServer, mySMTPPort )
' This function sends an e-mail message using CDOSYS
'
' Arguments:
' myFrom       = Sender's e-mail address ("John Doe <jdoe@mydomain.org>" or "jdoe@mydomain.org")
' myTo         = Receiver's e-mail address ("John Doe <jdoe@mydomain.org>" or "jdoe@mydomain.org")
' mySubject    = Message subject (optional)
' myTextBody   = Actual message (text only, optional)
' myHTMLBody   = Actual message (HTML, optional)
' myAttachment = Attachment as fully qualified file name, either string or array of strings (optional)
' mySMTPServer = SMTP server (IP address or host name)
' mySMTPPort   = SMTP server port (optional, default 25)
'
' Returns:
' status message
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com

    ' Standard housekeeping
    Dim i, objEmail

    ' Use custom error handling
    On Error Resume Next

    ' Create an e-mail message object
    Set objEmail = CreateObject( "CDO.Message" )

    ' Fill in the field values
    With objEmail
        .From     = myFrom
        .To       = myTo
        ' Other options you might want to add:
        ' .Cc     = ...
        ' .Bcc    = ...
        .Subject  = mySubject
        .TextBody = myTextBody
        .HTMLBody = myHTMLBody
        If IsArray( myAttachment ) Then
            For i = 0 To UBound( myAttachment )
                .AddAttachment Replace( myAttachment( i ), "\", "\\" ),"",""
            Next
        ElseIf myAttachment <> "" Then
            .AddAttachment Replace( myAttachment, "\", "\\" ),"",""
        End If
        If mySMTPPort = "" Then
            mySMTPPort = 25
        End If
        With .Configuration.Fields
            .Item( "http://schemas.microsoft.com/cdo/configuration/sendusing"      ) = 2
            .Item( "http://schemas.microsoft.com/cdo/configuration/smtpserver"     ) = mySMTPServer
            .Item( "http://schemas.microsoft.com/cdo/configuration/smtpserverport" ) = mySMTPPort
            .Update
        End With
        ' Send the message
        .Send
    End With
    ' Return status message
    If Err Then
        EMailSMTP = "ERROR " & Err.Number & ": " & Err.Description
        Err.Clear
    Else
        EMailSMTP = "Message sent ok"
    End If

    ' Release the e-mail message object
    Set objEmail = Nothing
    ' Restore default error handling
    On Error Goto 0
End Function


Set objRS = Nothing
Set objCon = Nothing
%>
