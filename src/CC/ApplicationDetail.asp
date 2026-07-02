
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=ADOVBS.inc -->
<!-- #Include file=CAPSFunctions.asp -->
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
Dim objCmd2
Dim objCmd3
Dim objCmd5
Dim objCmd6
Dim objCmd7
Dim objCmd8
Dim objRS
Dim objRS1
Dim strSelected
Dim x 
Dim strMessage
Dim strColour
Dim strViewButton

Dim strSubmittedDate
Dim strReviewdDate
Dim strExportDate
Dim strASFINDate
Dim strBankResponseDate
Dim strStatusSum
Dim strEmailDate
Dim strErrorDate
Dim strEmailErrorID

Dim strCMSUserName
Dim strCMSAccountHolder
Dim strCMSAHEID
Dim strCMSLocation
Dim strCMSAdminCentre
Dim strCMSSupervisor
Dim strSQL
Dim intRecord

Dim strErrorSum
Dim intErrorTotal

Dim strApplicationType
Dim strEID

Dim lngApplicationVersion

Dim lngTransactionAmountXML
Dim lngCreditAmountXML
Dim strAppTypeGlobal

Dim strCreditLimitFrom
Dim strCreditLimitTo
Dim strCreditLimitPermanent
Dim dblTransactionLimitOriginal
Dim dblTransactionLimitNew
Dim strGetLimitDates

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objCmd2 = Server.CreateObject("ADODB.Command")
Set objCmd3 = Server.CreateObject("ADODB.Command")
Set objCmd5 = Server.CreateObject("ADODB.Command")
Set objCmd6 = Server.CreateObject("ADODB.Command")
Set objCmd7 = Server.CreateObject("ADODB.Command")
Set objCmd8 = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

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

If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	Session("ApplicationID") = Request.QueryString("ApplicationID")
	
	'Get the Applicants Name also for display ---- Function GetApplicantName is in CAPSFunctions.asp file
	Session("ApplicationName") = GetApplicantName(Session("ApplicationID"),"EID")
End If



If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "Save" Then
		Call SaveMessage
	End If	
	'If the User has clicked the Release button
	If Request.QueryString("Action") = "Release" Then
		Call SubmitApplication()
		
		'Response.Write "Save=" & Request.QueryString("CreditLimit") & ", " & Request.QueryString("CMSUserName")
	End If
	'If the User has clicked the Release button
	If Request.QueryString("Action") = "Reject" Then
		Call SubmitApplication()
	End If
	
	'If the User has clicked the Name Change button
	If Request.QueryString("Action") = "NameChange" Then
		'ChangeNameOnCard Request.QueryString("NewName"), Request.QueryString("AppStatus")	
		Call ChangeNameOnCard(Request.QueryString("AppStatus"),Request.QueryString("NewTitle"),Request.QueryString("NewFirstName"),Request.QueryString("NewSurname"),Request.QueryString("NewNOC"),Request.QueryString("NCAppID"))
	End If
	
	'If the User has clicked the Limit Change button
	If Request.QueryString("Action") = "LimitChange" Then
		ChangeLimit Request.QueryString("NewLimit"), Request.QueryString("AppStatus")	
	End If
	
	'If the User has clicked the Update Missing CMC Details button
	If Request.QueryString("Action") = "UpdateCMC" Then
		Call UpdateApplicationMissingDetails(Request.QueryString("EID"))
	End If
	
	'Resend Error Email
	If Request.QueryString("Action") = "ResendEmail" Then
		Call ReSendErrorEmail(Session("ApplicationID"))
	End If
	
	
End If

If Request.QueryString("Action")="ProcessApplicationContacts" Then
	'Call the procedures to update the application with CDMC details, then contact details, then Errors
	Call UpdateApplicationCDMC(Request.QueryString("EID"))
	Call UpdateApplicationContacts(Request.QueryString("EID"))
	Call UpdateApplicationErrors(Request.QueryString("EID"))
	Call UpdateEmailErrorTemplate()
	
	'If the Application is a CMC only then update the Gender and Date of Birth as they are missing on old applications (will only update if they are missing)
	'Call UpdateApplicationMissingDetails(Request.QueryString("EID"))
	''''The above is run above via a button next to DOB field
	
End If

	Call LoadCMSDetails(1)
	
%>
<script>
function OpenSs(cb) {

	//alert("asas");
	//var e = document.getElementById(this.cb);
	//var result = e.options[e.selectedIndex].value;
	
	//document.getElementById('ContinueMod').value=result;
	document.getElementById('ModApp').showModal();
}

function loadCDMC(varID) {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CDMCDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetCDMCDetails.asp?EmployeeID=" + varID + "", true);
  xhttp.send();
}


function loadXML(varID) {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("XMLDetail").innerHTML = this.responseText;
    }
  };
  
  xhttp.open("GET", "../CC/AJAX/GetXMLDetails.asp?ApplicationID=" + varID + "", true);
  xhttp.send();
}

function loadLimitChange(varID, varAppID) {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("LChangeDetail").innerHTML = this.responseText;
    }
  };
 
  xhttp.open("GET", "../CC/AJAX/GetLimitChangeDetails.asp?EmployeeID=" + varID + "&ApplicationID=" + varAppID  +"", true);
  xhttp.send();
}


function TestMe() {

	alert(frm.CreditLimit.value);
	alert(document.getElementById("CMSUserName").value);
	
}

function OpenNCOLD(cb) {

	var id = cb.getAttribute('data-NCAppID');
	document.getElementById('AppChangeID').value=id;

	var Userid = cb.getAttribute('data-NCAppName');
	document.getElementById('AppName').value=Userid;

	var AppStat = cb.getAttribute('data-NCAppNameStatus');
	document.getElementById('AppNameStatus').value=AppStat;
	
	document.getElementById('NameOnCardChange').value=Userid;
	
	document.getElementById('demoModalLabel').value='Status Change for ' + Userid + ':';
	
}

function OpenCreditChange(cb) {

	var id = cb.getAttribute('data-CLAppID');
	document.getElementById('CreditLimitChangeID').value=id;

	var Userid = cb.getAttribute('data-CLCreditLimit');
	document.getElementById('CreditLimitChange').value=Userid;

	var AppName = cb.getAttribute('data-CLAppName');
	document.getElementById('AppNameCL').value=AppName;
	
	var AppStat = cb.getAttribute('data-CLAppNameStatus');
	document.getElementById('CreditLimitChangeStatus').value=AppStat;
	
	//document.getElementById('CreditLimitChange').value=Userid;
	
	//document.getElementById('demoModalLabel').value='Status Change for ' + Userid + ':';
	
}

function SaveNameChangeOLD() {

	self.location = "ApplicationDetail.asp?Action=NameChange&NewName="+document.getElementById("NameOnCardChange").value+"&AppStatus="+document.getElementById("AppNameStatus").value;
}

function SaveLimitChange() {

	self.location = "ApplicationDetail.asp?Action=LimitChange&NewLimit="+document.getElementById("CreditLimitChange").value+"&AppStatus="+document.getElementById("CreditLimitChangeStatus").value;
}

function OpenNC(cb) {

	//alert('Please see System Admin as this is not yeat available');
	
	var id = cb.getAttribute('data-NCAppID');
	document.getElementById('NCAppID').value=id;

	var Userid = cb.getAttribute('data-NCAppName');
	document.getElementById('AppName').value=Userid;

	var AppStat = cb.getAttribute('data-NCAppNameStatus');
	document.getElementById('NCAppNameStatus').value=AppStat;
	
	var CardType = cb.getAttribute('data-NCCardType');
	document.getElementById('NCCardType').value=CardType;
	
	//var Userid = cb.getAttribute('data-NCAppName');
	document.getElementById('NewNOC').value=Userid;
	
	var newtitle = cb.getAttribute('data-NCTitle');
	document.getElementById('NewTitle').value=newtitle;
	
	var firstname= cb.getAttribute('data-NCFirstName');
	document.getElementById('NewFirstName').value=firstname;
	
	var surname = cb.getAttribute('data-NCSurname');
	document.getElementById('NewSurname').value=surname;

	document.getElementById('NameOnCardChange').value=Userid;
	
	document.getElementById('demoModalLabel').value='Status Change for ' + Userid + ':';
	
}


function SaveNameChange() {

	self.location = "ApplicationDetail.asp?Action=NameChange&NewNOC="+document.getElementById("NewNOC").value+"&AppStatus="+document.getElementById("NCAppNameStatus").value+"&NewFirstName="+document.getElementById("NewFirstName").value+"&NewTitle="+document.getElementById("NewTitle").value+"&NewSurname="+document.getElementById("NewSurname").value+"&NCCardType="+document.getElementById("NCCardType").value+"&NCAppID="+document.getElementById("NCAppID").value;
}


function LoadRelease(varStatus) {

	document.getElementById("ErrorDetailsMod").innerHTML = document.getElementById("my-errors").innerHTML;
	document.getElementById("ModalReleaseTitle").innerHTML = 'Update Application to: ' + varStatus + '?';
	document.getElementById("NewStatus").value = varStatus;
}

function ReleaseApp() {
varCheck = true;

	//if (frm.CreditLimit.value)='' {
	//	alert("Please Enter a Credit Limit");
	//	varCheck = false;
	//}
	//if (frm.CMSUserName.value)='' {
	//	alert("Please Enter a CMS USer Name");
	//	varCheck = false;
	//}
	
	if (varCheck = true) {
		self.location = "ApplicationDetail.asp?Action=Release&Status="+document.getElementById("NewStatus").value;
		//self.location = "ApplicationDetail.asp?Action=Release&CreditLimit=" + frm.CreditLimit.value+"&CMSUserName="+frm.CMSUserName.value+"&Status="+document.getElementById("NewStatus").value;
		//self.location="ApplicationDetail.asp?Action=Release&CreditLimit="+ frm.CreditLimit.value + "&CMSUserName=" + frm.CMSUerName.value;
	}
	
}

function RejectApp() {
varCheck = true;

	if (varCheck = true) {
		self.location = "ApplicationDetail.asp?Action=Reject&Status=Rejected"
	}
	
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

<!-- Select Batch Number Modal -->
<div class="modal fade" id="ModalRelease" tabindex="-1" role="dialog" aria-labelledby="ModalRelease" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalReleaseTitle" style="font-weight:bold;">Release Application?</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
	  <div class="modal-header">
	  <h6 style="color:navy;">Applicant: <%=Session("ApplicationName")%></h6>
	  </div>
      <div class="modal-body" id="ErrorDetailsMod">
        
      </div>
      <div class="modal-footer">
		<button type="button" class="btn btn-danger" onClick="RejectApp();" ><i class="fa fa-check"></i> Reject</button>
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
        <button type="button" class="btn btn-primary" onClick="ReleaseApp();" ><i class="fa fa-check"></i> Yes</button>
		<!--<button type="button" class="btn btn-primary" onClick='window.location="ApplicationDetail.asp?Action=Release&CreditLimit='+frm.CreditLimit.value+" ><i class="fa fa-check"></i> Yes</button>-->
		<input type="hidden" id="NewStatus" name="NewStatus" value=""/>
      </div>
    </div>
  </div>
</div>
<!-- End Select Batch Number Modal -->

<!-- Modal -->
<div class="modal fade" id="LoadMod" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
   <div class="loader">
        <div class="wrap">
            <div class="spinner"></div>
            <span class="loading-message">Loading...</h6>
        </div>
    </div>
</div>


	
<!-- Modal -->
<div class="modal fade" id="CDMCModal" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-dialog-right modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  CDMC Detail
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="CDMCDetail">
               
				  
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>

 
 <!-- XML APplicaiton Modal -->
<div class="modal fade" id="XMLModal" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
    
	 <div class="modal-dialog modal-dialog-right modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  XML Application Detail
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="XMLDetail">
               
				  
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
 <!-- END XML Applicaiton Modal -->
 
 <!-- LChange Application Modal -->
<div class="modal fade" id="LChangeModal" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
    <div class="modal-dialog modal-large modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="LChangeModalLabel">
                  Application Limit Change Detail
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="LChangeDetail">
               
				  
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
 <!-- END LChange Applicaiton Modal -->
 
 <!-- Change Name on Card Modal -->
    <div class="modal fade" id="NameOnCardModalOLD" tabindex="-1" role="dialog" aria-labelledby="demoModalLabel" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="demoModalLabel">Name On Card</h5>
            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
              <span aria-hidden="true">&times;</span>
            </button>
          </div>
          <div class="modal-body">
		  <div class="col-md-12">
			<table class="table table-hover">
			<tr><td style="font-weight:bold;">Changing Name on Card for:</td><td><input id="AppName" name = "AppName" value="" style="border: 0;" class="form-control"/></td></tr>
           <%
		   
			Response.Write "<tr><td><label for""NameOnCardChange""  style=""font-weight:bold;"">Change to:</label></td>" & _
					"<td><INPUT id=""NameOnCardChange"" name=""NameOnCardChange"" class=""form-control"" value=""""/><td></tr>"
			
		   %>
		   <tr><td><input id="AppChangeID" name = "AppChangeID" value="" HIDDEN /></td><input id="AppNameStatus" name = "AppNameStatus" value="" HIDDEN /></td></tr>
			</table>
		   </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
            <button type="button" class="btn btn-primary" onClick="SaveNameChange();"><i class="fa fa-check"></i> Save changes</button>
          </div>
        </div>
      </div>
    </div>
	<!-- End Change Name on Card Modal -->
	
	 <!-- Change Name on Card Modal -->
    <div class="modal fade" id="NameOnCardModal" tabindex="-1" role="dialog" aria-labelledby="demoModalLabel" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="demoModalLabel">Name On Card</h5>
            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
              <span aria-hidden="true">&times;</span>
            </button>
          </div>
          <div class="modal-body">
		  <div class="col-md-12">
			<table class="table table-hover">
			<tr><td style="font-weight:bold;">Changing Name on Card for:</td><td><input id="AppName" name = "AppName" value="" style="border: 0;" class="form-control"/></td></tr>
           <%
		   
			Response.Write "<tr><td><label for""NewNOC"" style=""font-weight:bold;"">New Name On Card:</label></td>" & _
					"<td><INPUT id=""NewNOC"" name=""NewNOC"" class=""form-control"" value=""""/><td></tr>" & _	
					"<tr><td><label for""NewTitle"" style=""font-weight:bold;"">New Title:</label></td>" & _
					"<td><INPUT id=""NewTitle"" name=""NewTitle"" class=""form-control"" value=""""/><td></tr>" & _
					"<tr><td><label for""NewFirstName""  style=""font-weight:bold;"">New First Name:</label></td>" & _
					"<td><INPUT id=""NewFirstName"" name=""NewFirstName"" class=""form-control"" value=""""/><td></tr>" & _
					"<tr><td><label for""NewSurname""  style=""font-weight:bold;"">New Surname:</label></td>" & _
					"<td><INPUT id=""NewSurname"" name=""NewSurname"" class=""form-control"" value=""""/><td></tr>" & _
					"<tr><td><input id=""NCAppID"" name=""NCAppID"" value="""" HIDDEN /></td><td><input id=""NCAppNameStatus"" name=""NCAppNameStatus"" value="""" HIDDEN /></td><td><input id=""NCCardType"" name=""NCCardType"" value="""" HIDDEN /></td></tr>"
		   
		   %>
			</table>
		   </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
            <button type="button" class="btn btn-primary" onClick="SaveNameChange();"><i class="fa fa-check"></i> Save changes</button>
          </div>
        </div>
      </div>
    </div>
	<!-- End Change Name on Card Modal -->
	
	<!-- Change Credit Limit Modal -->
    <div class="modal fade" id="CreditLimitModal" tabindex="-1" role="dialog" aria-labelledby="demoModalLabel" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="demoModalLabel">Credit Limit</h5>
            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
              <span aria-hidden="true">&times;</span>
            </button>
          </div>
          <div class="modal-body">
		  <div class="col-md-12">
			<table class="table table-hover">
			<tr><td style="font-weight:bold;">Changing Credit Limit for:</td><td><input id="AppNameCL" name = "AppNameCL" value="" style="border: 0;" class="form-control"/></td></tr>
           <%
		   
			Response.Write "<tr><td><label for""CreditLimitChange""  style=""font-weight:bold;"">Change to:</label></td>" & _
					"<td><INPUT id=""CreditLimitChange"" name=""CreditLimitChange"" class=""form-control"" value=""""/><td></tr>"
			
		   %>
		   <tr><td><input id="CreditLimitChangeID" name = "CreditLimitChangeID" value="" HIDDEN /></td><input id="CreditLimitChangeStatus" name = "CreditLimitChangeStatus" value="" HIDDEN /></td></tr>
			</table>
		   </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
            <button type="button" class="btn btn-primary" onClick="SaveLimitChange();"><i class="fa fa-check"></i> Save changes</button>
          </div>
        </div>
      </div>
    </div>
	<!-- End Change Name on Card Modal -->
	
	

    <main class="main py-5">
      <div class="container">
		<form action="ApplicationDetail.asp?Action=Save" method="POST" id="frm" name="frm">
        <div class="row">
          <div class="col-md-8">
			<% Call DisplayTableDetails() %>	
				
            </div>
			
			 <div class="col-md-4 sidebar">
		  
              <div class="panel panel-shadow panel-validation mb-3">
				  <div class="panel-header">
					<h4><i class="fa fa-file-signature"></i> Application Notes</h4>
					<span class="panel-subheader">Recent Notes</span>
				  </div>
				  <div class="panel-content mb-8">
					<% Call LoadMessages() %>
					
					<textarea rows="4" id="MessageS" name="MessageS" class="form-control input-md" value="" placeholder="Type a note"></textarea>
					<div class="col-md-12 text-right my-auto">
					<button type="button" class="btn btn-primary" onClick="frm.submit();">Save Note</button>
					</div>
				</div>
			  </div>
			
			
			  <div class="panel panel-shadow mb-3">
              <div class="panel-header">
                <h4><i class="fa fa-clipboard"></i> Card Progress</h4>
                <span class="panel-subheader">Card Progress</span>
              </div>
              <div class="panel-content">
					<% Call LoadApplicationSummary %>
              </div>
            </div>
			
			
			<div class="panel panel-shadow mb-3">
              <div class="panel-header">
                <h4><i class="fa fa-book-open"></i> Application Change Summary</h4>
                
              </div>
			  <% Call LoadAudit() %>
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
Dim dteDateSubmitted
Dim dteDateReviewed
Dim strCreditLimit
Dim strHeader
Dim strCardType
Dim strErrors
Dim strErrorsHeader
Dim strAppType
Dim strUpdate
Dim arrErrors(4,15)
Dim intErrors
Dim strErrorDisplay
Dim strRelease
Dim strResolved

Dim strLimitChangeDetails
Dim strLimitChangeLast4
Dim strCardDetails
Dim strCurrentLimit 
Dim strXMLApplication
Dim strCMSUserType
Dim strUpdateCMC
Dim strTransactionLimit
Dim strBranded
Dim strLimitChangeButton
Dim strAppDOBAge
Dim strDTCSignedByApplicant
Dim strDPCSignedByApplicant
Dim strCMCSignedByApplicant
Dim strDTCDuelSignedByApplicant
Dim strDPCSignedBySuper
Dim strDTCLimitChangedByApplicant
Dim strDPCLimitChangedByApplicant
Dim strResponseWrite
Dim strTxnLimit
'First get any error for the application and pass into an array
strSQL = "SELECT * FROM tblCAPSApplicationError WITH(NOLOCK) WHERE ApplicationID = '" & Session("ApplicationID") & "'"

objRS.Open strSQL,objCon

intErrorTotal = 0

    Do Until objRS.EOF
	
		intErrors = intErrors + 1
		
		intErrorTotal = intErrorTotal + 1
		
		arrErrors(0,intErrors) = objRS("ErrorID")
		arrErrors(1,intErrors) = objRS("ErrorName")
		arrErrors(2,intErrors) = objRS("ErrorDescription")
		
		strResolved = ""
		
		'Determine what to display depending on the Resolved value
		If IsNull(objRS("Resolved")) or objRS("Resolved") = "" Then
			arrErrors(3,intErrors) = ""
		Else
			If objRS("Resolved") = "Y" Then
				arrErrors(3,intErrors) = "style=""text-decoration: line-through; color:grey;"""
				
				strResolved = "<span style=""color:red; font-weight:bold;"">Resolved</span>"
				
				intErrorTotal = intErrorTotal - 1
			End If
		End If
		
		'Determine what to display depending on the Resolved Date
		If IsNull(objRS("Resolved")) or objRS("Resolved") = "" Then
			arrErrors(4,intErrors) = ""
		Else
			arrErrors(4,intErrors) = "title=""Date resolved: " & objRS("DateResolved") & """"
		End If
		
		
		'Build the errors section for the Application details tab section
		strErrorDisplay = strErrorDisplay & "<tr><th " & arrErrors(3,intErrors) & ">" & arrErrors(0,intErrors) & ": " & arrErrors(1,intErrors) & "</th><td " & arrErrors(4,intErrors) & ">" & arrErrors(2,intErrors) & " " & strResolved & "</td></tr>"
		
		objRS.Movenext
		
	Loop

objRS.Close

'If Session("EmployeeID") = "" OR ISNull(Session("EmployeeID")) Then
'	strSQL = "SELECT * FROM qryCAPSApplications WHERE ApplicationID = '" & Session("ApplicationID") & "'"
'Else
	strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE ApplicationID = '" & Session("ApplicationID") & "'"
'End If

'Response.Write strSQL

objRS.Open strSQL,objCon

	
    If Not objRS.EOF Then
		'If isNull(objRS(9)) Then
		'	dblEmpCont = 0
		'Else
		'	dblEmpCont = objRS(9)
		'End If

		If isNull(objRS("CardTypeSub")) Then
			strHeader = ""
			strCardType = ""
		Else
			strHeader = objRS("CardTypeSub")
			strCardType  = objRS("CardType") & " " & objRS("CardTypeSub")
		End If
		
		''New for the DPC Change
		strBranded = objRS("BankCardType")
		
		'Get the Application Type Name
		If IsNull(objRS("ApplicationTypeName")) Then
			strApplicationType = ""
		Else
			strApplicationType = objRS("ApplicationTypeName")
			'---If the application is a DPC Mastercard then the Branding is different
			If trim(strApplicationType) = "DPC Mastercard" Then 
				strApplicationType = strApplicationType & " (Diners)"
				
				'Display a Badge based on the branding type
				If strBranded = "K5" Then
					strBranded = "<span class=""badge badge-pill badge-info"">Branded</span>"
				ElseIf strBranded = "PW" Then
					strBranded = "<span class=""badge badge-pill badge-dark"">UnBranded</span>"
				Else
					'Do not add a badge as he value should be on of the 2 above
				End If
				
			End If
			'Branding for the other Diners cards -DTCs
			If trim(Left(strApplicationType,3)) = "DTC" OR trim(Left(strApplicationType,3)) = "CMC" Then 
				
				'Display a Badge based on the branding type
				If Right(Trim(strBranded),1) = "1" Then
					strBranded = "<span class=""badge badge-pill badge-info"">Branded</span>"
				ElseIf Right(Trim(strBranded),1) = "2" Then
					strBranded = "<span class=""badge badge-pill badge-dark"">UnBranded</span>"
				Else
					'Do not add a badge as he value should be on of the 2 above
				End If
				
			End If
			'Get the branding for DPC cards - which already has the branding in the field name
			If trim(Left(strApplicationType,3)) = "DPC" AND trim(strApplicationType) = "DPC Mastercard" Then 
				strBranded = "<span class=""badge badge-pill badge-info"">" & strBranded & "</span>"
			End If
			
	
			
			If trim(Mid(strApplicationType,5,3)) = "NAB" Then
			
				If strBranded = "LU" or strBranded = "KH" or strBranded = "LW" Then
					strBranded = "<span class=""badge badge-pill badge-info"">Branded</span>"
				ElseIf strBranded = "EK" or strBranded = "EV" Then
					strBranded = "<span class=""badge badge-pill badge-dark"">UnBranded</span>"
				Else
					'Do not add a badge as he value should be on of the 2 above
				End If
			
			
			End If
			
			'Do not display 0 as it will be displayed by the recordset value (below)
			If strBranded = "0" Then strBranded = ""
		End If

		'Get the CMS User Type for loading CMS details
		If IsNull(objRS("CMSUserType")) Then
			strCMSUserType = 1
		Else
			strCMSUserType = objRS("CMSUserType")
		End If
		
		'Call the procedure to load the CMS User Details
		Call LoadCMSDetails(strCMSUserType)
		
		'Determine the image and title based on the card type
		If Trim(strHeader) = "Diners" Then
			strHeader = "<img src=""../images/icon_diners.png"" title=""" & strCardType & """> " & strCardType & " " & strApplicationType
		ElseIf strHeader = "ANZ" Then
			strHeader = "<img src=""../images/logo_ANZ.png"" Title=""" & strCardType & """> " & strCardType & " " & strApplicationType
		ElseIf strHeader = "Mastercard" Then
			strHeader = "<img src=""../images/logo_mc.png"" Title=""" & strCardType & """> " & strCardType & " " & strApplicationType
		Else
			strHeader = "<img src=""../images/logo_coa.png"" Title=""" & strCardType & """> " & strCardType & " " & strApplicationType
		End If
		
		'Write the header section based on the Card Type and related image
		'Response.Write "<h4>" & strHeader & "</h4></div>"
		
		
		Select Case objRS("Status")
		
			Case  "Awaiting Review"
				strAction = "<button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-check""></i> Release</button>"
				strAction = strAction & " <button type=""button"" class=""btn btn-outline-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-times""></i> Reject</button>"
				
				strRelease = "<div class=""mb-3 col-md-3""><button type=""button"" data-toggle=""modal"" data-target=""#ModalRelease"" onClick=""LoadRelease('Awaiting Export');"" class=""btn btn-outline-secondary"" title=""Click to Release and add to the NA file for export to Diners""><i class=""fa fa-check-circle""></i> Release</button></div>"
				
				strStatus = "<span class=""badge badge-pill badge-warning"">" & objRS("Status") & "</span>"
			Case "Added To CS"
				strAction = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-file-o""></i> View CS</button>"
			
			Case "Submitted"
				strAction = "<button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-check""></i> Release</button>"
				strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-times""></i> Reject</button>"
			
				strStatus = "<span class=""badge badge-pill badge-success"">Submitted to GCFO</span>"
			Case "On Hold"
				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-file-o""></i> View CS</button>"
			
				strRelease = "<div class=""mb-3 col-md-3""><button type=""button"" data-toggle=""modal"" data-target=""#ModalRelease"" onClick=""LoadRelease('Awaiting Review');"" class=""btn btn-outline-warning"" title=""Release from ON HOLD to Awaiting Review so it can be released""><i class=""fa fa-check-circle""></i> Progress</button></div>"
				
				strStatus = "<span class=""badge badge-pill badge-secondary"" data-toggle=""modal"" data-target=""#StatusModal"" data-AppStat=""" & objRS("Status") & """ data-AppID=""" & objRS("ApplicationID") & """  data-AppName=""" & objRS("FirstName") & " " & objRS("Surname") & " - " & objRS("CardType") & " " & objRS("CardTypeSub") & " Application"" onClick=""OpenSs(this);"">" & objRS("Status") & "</span>"
			Case "Cancelled"
				strAction = "Cancelled - " & FormatDateTime(objRS("DateUpdated"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationDetail.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				'strStatus  = "<button type=""button"" class=""btn btn-outline-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Cancelled</button>"
			
			Case "Awaiting issue"
				strAction = "<button type=""button"" title=""Approved by GCFO"" class=""btn btn-outline-success btn-sm"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-check""></i> Approved</button>"
				'strStatus  = "<button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Awaiting Issue</button>"
			
				strStatus = "<span class=""badge badge-pill badge-info"">" & objRS("Status") & "</span>"
				
			Case "Deleted"
				
				strStatus = "<span class=""badge badge-pill badge-danger"">Deleted</span>"
			Case "Rejected"
				
				strStatus = "<span class=""badge badge-pill badge-danger"">Rejected</span>"
			Case "Approved by ASFIN"
				
				strStatus = "<span class=""badge badge-pill badge-success"">Approved by GCFO</span>"
			
			Case "Added To NA"
				strAction = "<button type=""button"" class=""btn btn-outline-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				'strAction = "Rejected"
				'strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted</button>"
				strStatus = "<span class=""badge badge-pill badge-info"">" & objRS("Status") & "</span>"
			
			Case "Temp Hold"
				'strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-file-o""></i> View CS</button>"
			
				'strRelease = "<div class=""mb-3 col-md-3""><button type=""button"" data-toggle=""modal"" data-target=""#ModalRelease"" onClick=""LoadRelease('Awaiting Review');"" class=""btn btn-outline-warning"" title=""Release from ON HOLD to Awaiting Review so it can be released""><i class=""fa fa-check-circle""></i> Progress</button></div>"
				
				strStatus = "<span class=""badge badge-pill badge-danger"" data-toggle=""modal"" data-target=""#StatusModal"" data-AppStat=""" & objRS("Status") & """ data-AppID=""" & objRS("ApplicationID") & """  data-AppName=""" & objRS("FirstName") & " " & objRS("Surname") & " - " & objRS("CardType") & " " & objRS("CardTypeSub") & " Application"" onClick=""OpenSs(this);"">" & objRS("Status") & "</span>"
			
			Case "Old App Version"
				strAction = "<button type=""button"" class=""btn btn-outline-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				strStatus = "<span class=""badge badge-pill badge-danger"" title=""Application has been rejected due to version being too old"">" & objRS("Status") & "</span>"
			Case Else
				strAction = "<button type=""button"" class=""btn btn-outline-danger btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				'strAction = "Rejected"
				'strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted</button>"
				strStatus = "<span class=""badge badge-pill badge-success"">" & objRS("Status") & "</span>"
		End Select

		strAddress = Trim(objRS("Address1")) & " " & Trim(objRS("Address2")) & " " & Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode"))
		
		If len(strAddress) > 15 Then strAddress = left(strAddress,15) & "..."
		
		If IsNull(objRS("DateSubmitted")) Then
			dteDateSubmitted = ""
		Else
			dteDateSubmitted = FormatDateTime(objRS("DateSubmitted"),vbShortDate)
		End If
		
		If IsNull(objRS("DateReviewed")) Then
			dteDateReviewed = ""
		Else
			dteDateReviewed = FormatDateTime(objRS("DateReviewed"),vbShortDate)
		End If

		If IsNull(objRS("CreditLimit")) Then
			strCreditLimit = ""
		Else
			If IsNumeric(objRS("CreditLimit")) Then
				If objRS("CreditLimit")  > 0 Then
					'Determine the Format by the Card Type
					If Right(strApplicationType,12) = "Limit Change" Then
						strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
					Else
						If Left(strCardType,3) = "DTC" Then
							'Mastercard needs to be divided by 100
							If strCardType = "DTC Mastercard" Then
								If objRS("CreditLimit") = "30000" Then
									strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
								Else
									strCreditLimit = FormatCurrency(objRS("CreditLimit")/100,0)
								End If
							Else
							'strCreditLimit = FormatCurrency(objRS("CreditLimit")/100,0)
							strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
							End If
						Else
							strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
							'strCreditLimit = FormatCurrency(objRS("CreditLimit")/100,0)
						End If
					End If
				Else
					strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
				End If
			Else
				strCreditLimit = objRS("CreditLimit")
			End If
			
		End If

		If IsNull(objRS("CurrentLimit")) Then
			strCurrentLimit = ""
		Else
			If IsNumeric(objRS("CurrentLimit")) Then
				If Left(strCardType,3) = "DTC" Then
				'If strCardType = "DTC Diners" Then
					'strCurrentLimit = FormatCurrency(objRS("CurrentLimit")/100,0)
					strCurrentLimit = FormatCurrency(objRS("CurrentLimit"),0)
				Else
					strCurrentLimit = FormatCurrency(objRS("CurrentLimit"),0)
				End If
			End If
		End If
		
		'Get the transaction limit and format
		If IsNull(objRS("TransactionLimit")) Then
			strTransactionLimit = ""
		Else
		
			If IsNumeric(objRS("TransactionLimit")) Then
				If Left(strCardType,3) = "DTC" Then
				'If strCardType = "DTC Diners" Then
					'strCurrentLimit = FormatCurrency(objRS("TransactionLimit")/100,0)
					strTransactionLimit = FormatCurrency(objRS("TransactionLimit"),0)
				Else
					strTransactionLimit = FormatCurrency(objRS("TransactionLimit"),0)
				End If
			End If
		End If
		
		'Set the dates to global variables for use in the Summary Section
		strSubmittedDate = objRS("DateSubmitted")
		strReviewdDate = objRS("DateReviewed")
		strExportDate = objRS("DateExported")
		strASFINDate = objRS("ASFINSignedDate")
		strBankResponseDate = objRS("BankResponseDate")
		strStatusSum = objRS("Status") 
		strEmailDate = objRS("EmailSent")
		strErrorDate = objRS("ErrorsChecked")
		strEmailErrorID = objRS("EmailErrorID")
		
		strEID = objRS("EmployeeID")
		 
	'*****NEW Table Style with TABS START----
	
	'Check to see what the Application Type is to display errors if an AE602 form
	If IsNull(objRS("ApplicationType")) Then
		strAppType = ""
		strAppTypeGlobal = ""
	Else
		strAppType = objRS("ApplicationType")
		strAppTypeGlobal = objRS("ApplicationType")

		'Select the logo/image/font-awesome image to go with the application type to match the Applications List screen
		If strAppType = "AE602 XML" Then
			strAppType = strAppType & " <i title=""Old style AE602 PDF/XML form completed and emailed to Credit Cards"" class=""fa fa-file-pdf""></i>"
		ElseIf strAppType = "Portal" Then
			strAppType = strAppType & " <i title=""Defence Credit Card Portal Application submitted online"" class=""fa fa-globe""></i>"
		Else
			strAppType = strAppType & " <i title=""Service Connect Application"" class=""fa fa-link""></i>"
		End If
	
	End If
	
	'Date of Birth
	If Not IsNull(objRS("DateOfBirth")) Then
		'Date of birth id formatted different by differenet applications
		strAppDOBAge = Right(Trim(objRS("DateOfBirth")),2) & "/" & Mid(Trim(objRS("DateOfBirth")),5,2) & "/" & Left(Trim(objRS("DateOfBirth")),4)
		
		If IsDate(strAppDOBAge) Then
			''Only display of Birth is more than OR less than 18
			If DateDiff("yyyy",strAppDOBAge,now()) <18 Then
				strAppDOBAge = strAppDOBAge & "&nbsp;&nbsp;&nbsp; <span title=""" & objRS("DateOfBirth") & " from application"" class=""badge badge-pill badge-danger"">Under 18 Years of Age today</span>"
				'strAppDOBAge = strAppDOBAge & "&nbsp;&nbsp;&nbsp; <span class=""badge badge-pill badge-info"">" & DateDiff("yyyy",strAppDOBAge,now()) & " Years of Age today</span>"
			Else
				strAppDOBAge = strAppDOBAge & "&nbsp;&nbsp;&nbsp; <span title=""" & objRS("DateOfBirth") & " from application"" class=""badge badge-pill badge-success"">Over 18 Years of Age today</span>"
			End If
		Else
			strAppDOBAge = objRS("DateOfBirth")
		End If
	Else
		strAppDOBAge = ""
	End If
			
	'If strAppType = "AE602 XML" Then
	
		'Create the detail to be displayed in the tab
		strErrors = "<div class=""tab-pane fade"" id=""my-errors"" role=""tabpanel"" aria-labelledby=""my-errors-tab"">" &  _
			"<table class=""table"">" &  _
			"<tr><th>Email Error ID</th><td>" & objRS("EmailErrorID") & "</td></tr>" &  _
			"<tr><th>Warning Date</th><td>" & objRS("WarningDate") & "</td></tr>" &  _
			"<tr><th>Email Error Sent</th><td>" & objRS("ErrorEmailSent") & "</td></tr>" &  _
			"<tr><th>Errors</th><td>" & objRS("Notes") & "</td></tr>" & strErrorDisplay & "" & _
			"</table>" &  _
			"</div>" '<button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=ResendEmail&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-envelope-open-text""></i> Re-Email Errors</button>"

			'Temporarily removed from above (just before the end of table
			'"<tr><th></th><td><button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=ReEmail&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-envelope""></i> Re-Email</button></td></tr>" & _
			
		'Create the Tab Header
		strErrorsHeader = "<li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-errors-tab"" data-toggle=""tab"" href=""#my-errors"" role=""tab"" aria-controls=""my-errors"" aria-selected=""false"">Errors</a></li>"
		
		'Set the global variable to Yes for use in later functions/procedures in this page (Application Summary)
		strErrorSum = "Y"
	
	'If the Application Type Name is Limit Change then add the Limit Change details to the view
	If Not IsNull(strApplicationType) Then
		If Len(strApplicationType) > 12 Then
			If Right(strApplicationType,12) = "Limit Change" Then
	
				strLimitChangeDetails = "<tr><th>ASFIN</th><td>" & objRS("ASFinEmployeeID") & " - " & objRS("ASFinFirstName") & " " & objRS("ASFinSurname") & "</td></tr>" & _
					"<tr><th>ASFIN Signed</th><td>" & objRS("ASFINSigned") & " " & objRS("ASFINSignedDate") & "</td></tr>"
				
				strCardDetails = GetCardDetails(objRS("EmployeeID"),Left(strCardType,3))
				
				strLimitChangeLast4 = "<tr><th>Last 4 Digits</th><td>" & objRS("LastFourDigits") & "</td></tr>" & strCardDetails
				
				strLimitChangeButton = "<button type=""button"" class=""btn btn-outline-secondary btn-sm"" data-toggle=""modal"" data-target=""#LChangeModal"" HREF=""#"" onClick=""loadLimitChange(" & objRS("EmployeeID") & "," & objRS("ApplicationID") & ")""><i class=""fa fa-eye""></i> View Limit Change Detail</button>"
				
			End If
		End If
	End If
	
	'If the Application Type Name is CMC Only then allow the user to update the Gender and DOB fields
	If Not IsNull(strApplicationType) Then
		If strApplicationType = "CMC Only" Then
			
			strUpdateCMC = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=UpdateCMC&ApplicationID=" & objrs("ApplicationID") & "&EID=" & objRS("EmployeeID") & "'"";><i class=""fa fa-cogs""></i> Update Details</button>"
			
		End If
	End If
	
	'Create the re-process application contact details button to update the application contact details from CDMC
	strUpdate = "<tr><th></th><td><button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='ApplicationDetail.asp?Action=ProcessApplicationContacts&EID=" & objrs("EmployeeID") & "'"";  data-toggle=""modal"" data-target=""#LoadMod""><i class=""fa fa-retweet""></i> Update Details from CDMC</button>" & _
		"&nbsp;&nbsp;<button type=""button"" class=""btn btn-outline-secondary btn-sm"" data-toggle=""modal"" data-target=""#CDMCModal"" HREF=""#"" onClick=""loadCDMC(" & objRS("EmployeeID") & ")""><i class=""fa fa-eye""></i> View CDMC</button></td></tr>"
	
	'End If
	
	strXMLApplication = "<tr><th></th><td>" & strLimitChangeButton & " <button type=""button"" class=""btn btn-outline-secondary btn-sm"" data-toggle=""modal"" data-target=""#XMLModal"" HREF=""#"" onClick=""loadXML(" & objRS("ApplicationID") & ")""><i class=""fa fa-eye""></i> View XML Application</button></td></tr>"
	
	'''New for DCCP moved this function here so that transaction limit amounts can also be retrieved (rather than writing it further down)
	strGetLimitDates = GetLimitDates

	Response.write "<div class=""panel-content row""><div class=""mb-3 col-md-4""><h4>" & strHeader & "</h4></div><div class=""mb-3 col-md-5"">" &  _
		"<div class=""btn-group btn-selector table-tabs-selector"" role=""group"" aria-label=""Basic example"">" &  _
		"<button type=""button"" data-target=""table-tabs"" data-type=""as-tabs"" class=""btn btn-outline-primary active"">" &  _
		"<i class=""fa fa-list""></i> View as Tabs</button>" &  _
		"<button type=""button"" data-target=""table-tabs"" data-type=""as-table"" class=""btn btn-outline-primary"">" &  _
		"<i class=""fa fa-table""></i> View as Table</button></div></div>" & strRelease & "</div>" &  _
		"<div class=""panel-content row""><div class=""mb-3 col-md-8""><h6>" & Session("ApplicationName") & "</h6></div><div class=""mb-3 col-md-4"" style=""align:right; text-align:right;""><button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='Applications.asp'"";><i class=""fa fa-times""></i> Close </button>&nbsp;<button type=""button"" class=""btn btn-outline-primary btn-sm"" title=""Close and return to Application List with " & objRS("FirstName") & " " & objRS("Surname") & " selected"" onclick=""self.location='Applications.asp?EmployeeID=" & objRS("EmployeeID") & "&Link=AP'"";><i class=""fa fa-times""></i> <i class=""fa fa-user""></i> Close Employee</button></div></div>" & _
		"<div id=""table-tabs"" class=""as-tabs""><ul class=""nav nav-tabs"" id=""myFiTab"" role=""tablist""><li class=""nav-item"" role=""presentation"">" &  _
		"<a class=""nav-link active"" id=""overview-tab"" data-toggle=""tab"" href=""#overview"" role=""tab"" aria-controls=""overview"" aria-selected=""true"">Application Details</a>" &  _
		"</li><li class=""nav-item"" role=""presentation"">" &  _
		"<a class=""nav-link"" id=""card-details-tab"" data-toggle=""tab"" href=""#card-details"" role=""tab"" aria-controls=""card-details"" aria-selected=""false"">Contact Details</a>" &  _
		"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-limits-tab"" data-toggle=""tab"" href=""#my-limits"" role=""tab"" aria-controls=""my-limits"" aria-selected=""false"">Limits</a>" &  _
		"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-cms-tab"" data-toggle=""tab"" href=""#my-cms"" role=""tab"" aria-controls=""my-cms"" aria-selected=""false"">CMS Details</a>" &  _
		"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-workflow-tab"" data-toggle=""tab"" href=""#my-workflow"" role=""tab"" aria-controls=""my-workflow"" aria-selected=""false"">Workflow</a>" &  _
		"</li>" & strErrorsHeader & "</ul><div class=""tab-content panel panel-light p-3"" id=""myFiTabContent"">" &  _
		"<div class=""tab-pane fade show active"" id=""overview"" role=""tabpanel"" aria-labelledby=""overview-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>Application ID</th><td>" & objRS("ApplicationID") & "</td></tr>" &  _
		"<tr><th>Application Type</th><td>" & strApplicationType & "</td></tr>" &  _
		"<tr><th>Status</th><td>" & strStatus & "</td></tr>" &  _
		"<tr><th>Date submitted</th><td>" & objRS("DateSubmitted") & "</td></tr>" &  _
		"<tr><th>Employee ID</th><td style=""font-weight:bold;"">" & objRS("EmployeeID") & "</td></tr>" &  _
		"<tr><th>Name On Card</th><td style=""font-weight:bold;""><a data-toggle=""modal"" data-target=""#NameOnCardModal"" data-NCCardType=""" & objRS("CardTypeSub") & """ data-NCAppNameStatus=""" & objRS("Status") & """ data-NCAppID=""" & objRS("ApplicationID") & """ data-NCAppName=""" & objRS("NameOnCard") & """ data-NCTitle=""" & objRS("Title") & """ data-NCFirstName=""" & objRS("FirstName") & " " & objRS("MiddleName") & """ data-NCSurname=""" & objRS("Surname") & """ onClick=""OpenNC(this);"">" & objRS("NameOnCard") & " <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(Trim(objRS("NameOnCard"))) & " chars</a></span> &nbsp;&nbsp;<span style=""font-size:11px; color:gray;"">Click to change</span></td></tr>" &  _
		"<tr><th>Title</th><td>" & objRS("Title") & "</td></tr>" &  _
		"<tr><th>First Name(s)</th><td>" & objRS("FirstName") & " " & objRS("MiddleName") & "</td></tr>" &  _
		"<tr><th>Surname</th><td>" & objRS("Surname") & "</td></tr>" & _
		"<tr><th>Gender</th><td>" & objRS("Gender") & " " & strUpdateCMC & "</td></tr>" & _
		"<tr><th>DOB</th><td>" & strAppDOBAge & " </td></tr>" & _
		"<tr><th>Branded</th><td>" & strBranded & " " & objRS("BankCardType") & "</td></tr>" & strLimitChangeLast4 & "" &  _
		"<tr><th>Application Form</th><td>" & strAppType & " </td></tr></table></div>" & _
		"<div class=""tab-pane fade"" id=""my-limits"" role=""tabpanel"" aria-labelledby=""my-limits-tab"">" &  _
		"<table class=""table"">"
		
		
		'NEW update for the DPC change ---May 2023
		'Only dispaly limit details from the application if the application is for a new card, not for Limit Changes
		If Right(Trim(strApplicationType),12) = "Limit Change" Then
		
			'''New update for DCCP portal to get the limit detail from the application form not XML (as there is no XML)
			If Left(strAppType,6)="Portal" Then
				If IsNull(objRS("TransactionLimit")) Then
					lngTransactionAmountXML = "none"
				Else
					If IsNumeric(objRS("TransactionLimit")) Then
						lngTransactionAmountXML = FormatCurrency(objRS("TransactionLimit"),0)
					Else
						lngTransactionAmountXML = objRS("TransactionLimit")
					End If
				End If
				

				'strTransactionLimit = dblTransactionLimitOriginal
				'lngTransactionAmountXML = dblTransactionLimitNew

				'''Swap the Credit limit amountss ---might be a temporary fix for DCCP values------
				lngCreditAmountXML = strCurrentLimit

				If IsNull(objRS("CurrentLimit")) Then
					'lngCreditAmountXML = "none"
					strCurrentLimit = "none"
				Else
					If IsNumeric(objRS("CurrentLimit")) Then
						'lngCreditAmountXML = FormatCurrency(objRS("CurrentLimit"),0)
						strCurrentLimit = FormatCurrency(objRS("CurrentLimit"),0)
					Else
						'lngCreditAmountXML = objRS("CurrentLimit")
						strCurrentLimit =  objRS("CurrentLimit")
					End If
				End If
				
				''''Card Limit on card is the same as Current Limit for Portal applications
				strCreditLimit = strCurrentLimit 

				If IsNull(objRS("CreditLimit")) Then
					lngCreditAmountXML = "none"
					'strCurrentLimit = "none"
				Else
					If IsNumeric(objRS("CreditLimit")) Then
						lngCreditAmountXML = FormatCurrency(objRS("CreditLimit"),0)
						'strCurrentLimit = FormatCurrency(objRS("CreditLimit"),0)
					Else
						lngCreditAmountXML = objRS("CreditLimit")
						'strCurrentLimit =  objRS("CreditLimit")
					End If
				End If				
			
				''Get the CreditLimit Dates
				If IsNull(objRS("LimitDateFrom")) Then
					strCreditLimitFrom = ""
				Else
					strCreditLimitFrom = objRS("LimitDateFrom") 
				End If

				
				If IsNull(objRS("LimitDateTo")) Then
					strCreditLimitTo = ""
				Else
					strCreditLimitTo = objRS("LimitDateTo")
				End If

				''Get the Credit Limit Permanent field
				If IsNull(objRS("ChangesPermanent")) Then
					strCreditLimitPermanent = ""
				Else
					strCreditLimitPermanent = objRS("ChangesPermanent")
				End If

				'lngTransactionAmountXML = objRS("TransactionLimit")
				'lngCreditAmountXML = objRS("CreditLimit")
			Else
				'First call the procedure to get detial from the XML application table for limit changes. This will populate global variables
				Call GetLimitXML
			End If
		End If	
			Response.Write "<tr><th>Current Credit Limit (on Card)</th><td><a data-toggle=""modal"" data-target=""#CreditLimitModal"" data-CLAppNameStatus=""" & objRS("Status") & """ data-CLAppID=""" & objRS("ApplicationID") & """ data-CLAppName=""" & objRS("NameOnCard") & """ data-CLCreditLimit=""" & strCreditLimit & """ onClick=""OpenCreditChange(this);"">" & strCreditLimit & " </a></span> &nbsp;&nbsp;<span style=""font-size:11px; color:gray;"">Click to change</span></td></tr>" & _
				"<tr><th>Current Credit Limit (from Application)</th><td>" & strCurrentLimit & "</td></tr>" &  _
				"<tr><th class=""updated"">New Credit Limit</th><td class=""updated"">" & lngCreditAmountXML & " </td></tr>" &  _
				"<tr><th>Current Transaction Limit</th><td>" & strTransactionLimit & "</td></tr>" &  _
				"<tr><th class=""updated"">New Transaction Limit</th><td class=""updated"">" & lngTransactionAmountXML & "</td></tr>" &  _
				"<tr><th>Cash Daily</th><td>" & objRS("CashDaily") & "</td></tr>" &  _
				"<tr><th>Cash OTC</th><td>" & objRS("CashOTC") & "</td></tr>"
		'End If
		
			
		Response.Write "<tr><th>Justification</th><td>" & objRS("Justification") & "</td></tr>" '&  _

		Response.Write strGetLimitDates	
		
		Response.Write "<tr><th>Limit Date Reduced</th><td>" & objRS("LimitDateReduced") & "</td></tr>" & _
		Get_Signatures &  strXMLApplication & _
		"</table></div>" &  _
		"<div class=""tab-pane fade"" id=""card-details"" role=""tabpanel"" aria-labelledby=""card-details-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>Address 1</th><td>" & objRS("Address1") & "</td></tr>" &  _
		"<tr><th>Address 2</th><td>" & objRS("Address2") & "</td></tr>" &  _
		"<tr><th>Address 3</th><td>" & objRS("Address3") & "</td></tr>" &  _
		"<tr><th>Suburb</th><td>" & objRS("Suburb") & "</td></tr>" &  _
		"<tr><th>State</th><td>" & objRS("State") & "</td></tr>" &  _
		"<tr><th>PostCode</th><td>" & objRS("PostCode") & "</td></tr>" &  _
		"<tr><th>Home Phone</th><td>" & objRS("HomePhone") & "</td></tr>" &  _
		"<tr><th>Work Phone</th><td>" & objRS("WorkPhone") & "</td></tr>" &  _
		"<tr><th>Mobile Phone</th><td>" & objRS("MobilePhone") & "</td></tr>" &  _
		"<tr><th>Email</th><td>" & objRS("Email") & "</td></tr>" & strUpdate &  _
		"</table></div>" &  _
		"<div class=""tab-pane fade"" id=""my-cms"" role=""tabpanel"" aria-labelledby=""cms-details-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>CMS User</th><td>" & objRS("CMSUser") & " &nbsp;(from Application)&nbsp;&nbsp;&nbsp;&nbsp;<span style=""font-size:12px;font-type:italic; color:gray;"">(" & strCMSUserName & ") from CMS for EmployeeID = " & strCMSAHEID & "</span></td></tr>" &  _ 
		"<tr><th>Account Holder</th><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style=""font-size:12px;font-type:italic; color:gray;"">(" & strCMSAccountHolder & ") from CMS for EmployeeID = " & strCMSAHEID & "</span></td></tr>" &  _
		"<tr><th>Account Holder EID</th><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span style=""font-size:12px;font-type:italic; color:gray;"">(" & strCMSAHEID & ") from CMS for EmployeeID = " & strCMSAHEID & "</span></td></tr>" &  _
		"<tr><th>Report Group</th><td>" & objRS("ReportGroup") & "</td></tr>" &  _
		"<tr><th>Default Company</th><td>" & objRS("DefaultCompany") & "</td></tr>" &  _
		"<tr><th>Default Cost Centre</th><td>" & objRS("DefaultCC") & "</td></tr>" &  _
		"<tr><th>Default WBS</th><td>" & objRS("DefaultWBS") & "</td></tr>" &  _
		"<tr><th>CMS User Type</th><td>" & objRS("CMSUserType") & "</td></tr>" &  _
		"</table></div>" &  _
		"<div class=""tab-pane fade"" id=""my-workflow"" role=""tabpanel"" aria-labelledby=""my-workflow-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>Application Version</th><td>" & lngApplicationVersion & "</td></tr>" &  _
		"<tr><th>Date Submitted</th><td>" & objRS("DateSubmitted") & "</td></tr>" &  _
		"<tr><th>Submitted By</th><td>" & objRS("SubmittedByName") & "</td></tr>" &  _
		"<tr><th>Errors Checked</th><td>" & objRS("ErrorsChecked") & "</td></tr>" &  _
		"<tr><th>Email Sent</th><td>" & objRS("EmailSent") & "</td></tr>" &  _
		"<tr><th>Reviewed By</th><td>" & objRS("ReviewedByName") & "</td></tr>" &  _
		"<tr><th>Date Reviewed</th><td>" & objRS("DateReviewed") & "</td></tr>" &  _
		"<tr><th>UpdatedBy</th><td>" & objRS("UpdatedByName") & "</td></tr>" &  _
		"<tr><th>DateUpdated</th><td>" & objRS("DateUpdated") & "</td></tr>" & strLimitChangeDetails & _
		"</table>" &  _
		"</div>" & strErrors & "</div></div>"
'Removed default fund and io from screen
'"<tr><th>Default IO</th><td>" & objRS("DefaultIO") & "</td></tr>" &  _
'"<tr><th>Default Fund</th><td>" & objRS("DefaultFund") & "</td></tr>" &  _

 '<input value=""" & strCMSUserName & """ id=""CMSUserName"" name=""CMSUserName"" style=""border:0px;""></td></tr>
 '"<tr><th>Name On Card</th><td style=""font-weight:bold;""><a data-toggle=""modal"" data-target=""#NameOnCardModal"" data-NCCardType=""" & objRS("CardTypeSub") & """ data-NCAppNameStatus=""" & objRS("Status") & """ data-NCAppID=""" & objRS("ApplicationID") & """ data-NCAppName=""" & objRS("NameOnCard") & """ data-NCTitle=""" & objRS("Title") & """ data-NCFirstName=""" & objRS("FirstName") & " " & objRS("MiddleName") & """ data-NCSurname=""" & objRS("Surname") & """ onClick=""OpenNC(this);"">" & objRS("NameOnCard") & " <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(Trim(objRS("NameOnCard"))) & " chars</a></span> &nbsp;&nbsp;<span style=""font-size:11px; color:gray;"">Click to change</span></td></tr>" &  _
 ''OLD Name on ard Change from above------Updated the above line from CardDetails.asp
 '"<tr><th>Name On Card</th><td><a data-toggle=""modal"" data-target=""#NameOnCardModal"" data-NCAppNameStatus=""" & objRS("Status") & """ data-NCAppID=""" & objRS("ApplicationID") & """ data-NCAppName=""" & objRS("NameOnCard") & """ onClick=""OpenNC(this);"">" & objRS("NameOnCard") & " <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(Trim(objRS("NameOnCard"))) & " chars</a></span> &nbsp;&nbsp;<span style=""font-size:11px; color:gray;"">Click to change</span></td></tr>" &  _
 
 
	'*****NEW Table Style with TABS END----
	
	End If
	
	'Response.Write strAction & " " & strStatus
	Response.Write "<button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='Applications.asp'"";><i class=""fa fa-times""></i> Close </button>"
	
objRS.Close

End Sub

Public Sub LoadMessages()
'Procedure to load any messages relating to the application
Dim strSQL
Dim strPerson
Dim dteMessageDate
Dim strMessage

	strSQL = "SELECT * FROM qryCAPSMessage WITH(NOLOCK) WHERE [Object] = 'Application' AND [ObjectID] = '" & Session("ApplicationID") & "'"

	objRS.Open strSQL,objCon
	
    Do Until objRS.EOF
		
		'Get the message creator
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
		
		'Get the message date and format it
		If IsNull(objRS("DateUpdated")) or objRS("DateUpdated")  = "" Then
			dteMessageDate = ""
		Else
			dteMessageDate = "<span style=""font-size:12px; color:grey;"" title=""" & FormatDateTime(objRS("DateUpdated"),1) & " " & FormatDateTime(objRS("DateUpdated"),3) & """>" & FormatDateTime(objRS("DateUpdated"),1) & "</span>"
		End If
		
		'Replace any Carriage Returns in the text with a line break in HTML
		If IsNull(objRS("MessageDetail")) or objRS("MessageDetail")= "" Then
			strMessage=""
		Else
			strMessage = objRS("MessageDetail")
			strMessage = Replace(strMessage,chr(13),"</BR>")
		End If
		
		
		Response.write "<div class=""panel panel-light col-12""><div class=""panel-header"">" & _
			"<h6>" & objRS("UserFrom") & " " & strPerson & " - " & dteMessageDate & "</h6><span class=""panel-subheader"">" & strMessage & "</span></div></div>"

		objRS.Movenext
	Loop
				
objRS.Close

End Sub


Public Sub LoadApplicationSummary()
'Procedure to load the progress of the application
Dim strDODColour
Dim arrComplete(5)
Dim strASFinLine
Dim strErrLine
Dim strBGCol
Dim strBGColErr
Dim strTitleErr

	If IsNull(strSubmittedDate) Or strSubmittedDate = "" Then
		strSubmittedDate = "Not Submitted"
	Else
		strSubmittedDate = strSubmittedDate
		arrComplete(1) = "complete"
	End If
	
	'Get the Error EmailID for the Submitted date area to show if the app was old (red background)
	If IsNull(strEmailErrorID) Or strEmailErrorID = "" Then
		strBGColErr = ""
		strTitleErr = ""
	Else
		If strEmailErrorID = 99 Then
			strBGColErr = "style=""background-color:red"" "
			strTitleErr = "title=""When Application was submitted the application version was old"" "
		Else
			strBGColErr = ""
			strTitleErr = ""
		End If
	End If
	
	If IsNull(strReviewdDate) Or strReviewdDate = "" Then
		If strSubmittedDate = "Not Submitted" Then
			strReviewdDate = "Awaiting Review"
		Else
			strReviewdDate = "Awaiting Review"
		End If
	Else
		If strSubmittedDate = "Not Submitted" Then
			strReviewdDate = strReviewdDate
		Else
			strReviewdDate = strReviewdDate'DateDiff("d",strReviewdDate,strSubmittedDate)
		End If
		
		arrComplete(2) = "complete"
	End If
	
	strDODColour = "green"
	
	'If the card is a Limit Change then it needs to show the ASFin approval
	If 1 = 2 Then
		If IsNull(strASFINDate) Or strASFINDate = "" Then
			strASFINDate = "Awaiting Approval"
		Else
			strASFINDate = strASFINDate
			arrComplete(0) = "complete"
			
			strASFinLine = "<div class=""timeline-item " & arrComplete(2) & """><div class=""dot-container""><span class=""dot""></span></div>" & _
							"<div class=""content""><h6>ASFin Approval</h6><span class=""date"">" & strASFINDate & "</span></div></div>"
		End If
	End If
	
	'Build the Error Summary section if errors exist and the application is an AE602
	If strErrorSum = "Y" Then
	
		'If there is an error date then the application has been checked for errors (AE602 applications)
		If IsNull(strErrorDate) Or strErrorDate = "" Then
		Else
			strEmailDate = "Errors Checked " & strErrorDate
			arrComplete(5) = "complete"
			
			'If there are errors then change the background colour to red
			If intErrorTotal > 0 Then
				strBGCol = "style=""background-color:red"" "
				'strBGCol = "style=""background-color:#03b5fc;"" "
			End If
			
		End If
		
		If IsNull(strEmailDate) Or strEmailDate = "" Then
			strEmailDate = "Awaiting Check"
		Else
			strEmailDate = strEmailDate
			'arrComplete(5) = "complete"
		End If
		
		strErrLine = "<div class=""timeline-item " & arrComplete(5) & """><div class=""dot-container""><span class=""dot"" " & strBGCol & "></span></div>" & _
							"<div class=""content""><h6>Errors</h6><span class=""date"">" & strEmailDate & "</span></div></div>"
							
	End If
	
	If IsNull(strExportDate) Or strExportDate = "" Then
		strExportDate = "Awaiting Export"
	Else
		strExportDate = strExportDate
		arrComplete(3) = "complete"
	End If
	
	If IsNull(strBankResponseDate) Or strBankResponseDate = "" Then
		strBankResponseDate = "Not Sent"
	Else
		strBankResponseDate = strBankResponseDate
		arrComplete(4) = "complete"
	End If
	
	
	Response.Write "<div class=""timeline""><div class=""timeline-item " & arrComplete(1) & """ " & strTitleErr & "><div class=""dot-container""><span class=""dot"" " & strBGColErr & "></span></div>" & _
		"<div class=""content"" " & strTitleErr & "><h6>Submitted</h6><span class=""date"">" & strSubmittedDate & "</span></div></div>" & strErrLine & _
		"<div class=""timeline-item " & arrComplete(2) & """><div class=""dot-container""><span class=""dot""></span></div>" & _
		"<div class=""content""><h6>Reviewed</h6><span class=""date"">" & strReviewdDate & "</span></div></div>" & strASFinLine & _
		"<div class=""timeline-item " & arrComplete(3) & """><div class=""dot-container""><span class=""dot""></span></div>" & _
		"<div class=""content""><h6>Exported</h6><span class=""date"">" & strExportDate & "</span></div></div>" & _
		"<div class=""timeline-item " & arrComplete(4) & """><div class=""dot-container""><span class=""dot""></span></div>" & _
		"<div class=""content""><h6>Mailed By Bank</h6><span class=""date"">" & strBankResponseDate & "</span></div></div></div>"

End Sub


Public Sub LoadCMSDetails(strType)
'Procedure to load CMS details depending on the Type Passed in (as the details based on the Applicant can use the EmployeeID
'where if it is another account holder the application details have to be loaded first and the query to CMS is different)
on error resume next

Dim objCon2
Dim strCMSConnect

'''Get the System Parameter to see if the connection to CMS should be used...this is mainly forr UAT/Test environment.  System Parameter ConnectToCMS must exist
strCMSConnect = GetSystemAdmin("CMSConnect")

If IsNull(strCMSConnect) or strCMSConnect  = "CMSConnect" Then

Else
	If strCMSConnect  = "Yes" Then

	'ProMaster Connection details
	Set objCon2 = Server.CreateObject("ADODB.Connection")
	Session("DBConnection2") = "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	objCon2.ConnectionTimeout=1
	objCon2.Open Session("DBConnection2")
	End If

End If

'If there is no connection to ProMaster (Card Management System) then do not try to use the connection
	If objCon2.State = 1 Then
		'If the CMS Account Holder is different to the APplicant then search by CMS UserID
		If strType = "2" Then
		
			'Open a recordset in the ProMaster (CMS) database to check the Employee has a CMS Account
			objRS1.Open "SELECT [user_name],[employee_id] FROM procharge_user WITH(NoLock) WHERE employee_id = '" & Session("ApplicationEmployeeID") & "' AND [active_indicator ] = 'Y'",objCon2
			
				If objRS1.EOF Then
					strMessage = Session("ApplicationEmployeeID") & " has no active CMS Account"
					strCMSUserName = ""
					
				Else
					strMessage = "CMS Account for " & Session("ApplicationEmployeeID") & ": " & objRS1("user_name")
					strCMSUserName = objRS1("user_name")
				End If
				
			objRS1.Close
		End If
		
		'If the applicant is the CMS User then search by EID
		If strType = "1" Then
			'Open the Account Table to get account holder details
			strSQL = "SELECT TOP 1 [PU].[user_name] AS UserName, [PU].[employee_id] AS EmplID, [ca].[account_ref_no] AS AccountHolder FROM card_account ca (nolock) LEFT JOIN payment_cards pc (nolock) on pc.card_type = ca.card_type and pc.card_account_number = ca.card_account_number left join procharge_user pu (nolock) on ca.user_name = pu.user_name " & _
					"WHERE PU.employee_id = '" & Session("ApplicationEmployeeID") & "'"
		
			objRS1.Open strSQL,objCon2
			
				If objRS1.EOF Then
					'strMessage = Session("ApplicationEmployeeID") & " has no active CMS Account"
					strCMSUserName = "None"
					strCMSAccountHolder = "None"
					strCMSAHEID = Session("ApplicationEmployeeID")
					strCMSLocation = ""
					strCMSAdminCentre = ""
					strCMSSupervisor = ""
				Else
					'strMessage = "CMS Account for " & Session("ApplicationEmployeeID") & ": " & objRS("user_name")
					strCMSUserName = objRS1("UserName")
					strCMSAccountHolder = objRS1("AccountHolder")
					strCMSAHEID = objRS1("EmplID")
					strCMSLocation = objRS1("EmplID")
					strCMSAdminCentre = objRS1("EmplID")
					strCMSSupervisor = objRS1("EmplID")
				End If
				
			objRS1.Close
		End If
	Else
		strMessage = "CMS database currently unavailable, please try again in 1 hour"
		strCMSUserName = "CMS database unavailable"
	End If	
	'response.write "err=" & err.number
	If err.Number="3704" Then Response.Write "<div class=""alert alert-danger"" role=""alert""><i class=""fa fa-exclamation""></i> CMS is currently not available, therefore CMS details in CAPS may not be accurate until CMS is back up.</div>"'"CMS is not available. CMS Details not accurate"
	'if err.number <>0 Then reposnse.write "dfdfdfddf" & err.number
	on error goto 0
	
End Sub


Public Sub LoadAudit()
'Procedure to load any audit records relating to the card
Dim strSQL
Dim strPerson

	strSQL = "SELECT * FROM tblCAPSAuditLog WITH(NOLOCK) WHERE [ApplicationID] = '" & Session("ApplicationID") & "' ORDER By [ChangeDate] DESC"

	objRS.Open strSQL,objCon

	If objRS.EOF THEN
		Response.write "<div class=""panel col-12""><span class=""panel-subheader"" style=""color:grey; font-style:italic;"">No Application history details..</span></div></div>"
	End If
		
    Do Until objRS.EOF
		
		'If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
		'	strPerson = ""
		'Else
		'	If objRS("MessageFrom") = Session("UserID") Then
		'		strPerson = "(You)"
		'	Else
		'		If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
		'			If objRS("MessageFrom") = 0 Then
		'				strPerson = "(Admin)"
		'			End If
		'		End If
		'	End If
		'End If
		
		'Response.write "<div class=""panel col-12""><a href=""AuditLog.asp?ApplicationID=" & objRS("ApplicationID") & """><span class=""panel-subheader"">" & objRS("ChangeDate") & " - " & objRS("ChangeDetails") & "</span></a></div></div>"
		Response.write "<div class=""panel col-12""><a href=""AuditLog.asp?CardID=0&EmployeeID=" & objRS("EID") & "&ApplicationID=" & objRS("ApplicationID") & """><span class=""panel-subheader"">" & objRS("ChangeDate") & " - " & objRS("ChangeDetails") & "</span></a></div>"

		objRS.Movenext
	Loop
		
			Response.Write "</div>"

			
objRS.Close

End Sub



Public Sub SaveMessage()

Dim lngMessageID
Dim lngAdminID
Dim strMessage
Dim intRecord
Dim strMessageTitle

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


Public Sub UpdateApplicationContacts(strEIDInput)
'Procedure to run a stored procedure which updates all applications (based on their status [default="on hold"]) and adds CDMC details to them as they are not collected as part of the AE602 form.
Dim strStatusProcess
Dim intRecord

	'The status passed into the procedure determines which status of applications is processed (have CDMC details updated into their application)
	strStatusProcess = "On Hold"
	
  	With objCmd
  	
		.CommandType = 4
		.CommandText = "spCAPSApplicationXMLUpdateContact"
		
		.Parameters.Append objCmd.CreateParameter("EIDInput", adVarchar, adParamInput,12)
		.Parameters.Append objCmd.CreateParameter("StatusInput", adVarchar, adParamInput,20)
		.Parameters.Append objCmd.CreateParameter("UpdateCountOutput", adInteger, adParamOutput)
		
		.Parameters("EIDInput") = strEIDInput 'The Employee currently viewed.  -- Can be empty to processes all applications (AE602) rather than just the EID passed in
		.Parameters("StatusInput") = strStatusProcess 'This can be left empty to process "on hold" applications only or made another specific Status to process ("Awaiting review", 
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd.Execute        
	
	'Return the result of the Save Function.
     intRecord = objCmd.Parameters.Item("UpdateCountOutput")  

	If intRecord = -1 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Application contact details NOT updated for " & strEIDInput & " they are not in the CDMC OR have no valid CDMC address details!</div>"
	ElseIf intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Application contact details NOT updated for " & strEIDInput & " with status " & strStatusProcess & "!</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">Applications for " & strEIDInput & " with Status " & strStatusProcess & " updated with CDMC contact details!</div>"
	End If
	
End Sub

Public Sub UpdateApplicationErrors(strEID)
'Procedure to run a stored procedure which checks all AE602 XML applications (on hold or awaiting review) for errors and adds errors or resolves them.
	
  	With objCmd2
  	
		.CommandType = 4
		.CommandText = "spCAPSApplicationErrors"
		
		.Parameters.Append objCmd2.CreateParameter("UserID", adInteger, adParamInput)
		.Parameters.Append objCmd2.CreateParameter("EmployeeID", adVarchar, adParamInput,20)
		.Parameters.Append objCmd2.CreateParameter("ApplicationErrorsOutput", adInteger, adParamOutput)
		
		.Parameters("UserID") = Session("UserID") 'The User who had performed the error checks
		.Parameters("EmployeeID") = strEID'"0" 'Set EmployeeID to 0 so that procedure checks all applications
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd2.Execute        
	
	'Return the result of the Save Function.
     intRecord = objCmd2.Parameters.Item("ApplicationErrorsOutput")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Applications checked for errors with status On Hold or Awaiting Review!</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Applications with Status On Hold or Awaiting Review checked for Errors!</div>"
	End If
	
End Sub

Public Sub UpdateApplicationCDMC(strEID)
'Procedure to run a stored procedure which checks all AE602 XML applications (on hold or awaiting review) for errors and adds errors or resolves them.
	
  	With objCmd3
  	
		.CommandType = 4
		.CommandText = "spCAPSCDMCProcessContactDetails"
		
		.Parameters.Append objCmd3.CreateParameter("UserID", adInteger, adParamInput)
		.Parameters.Append objCmd3.CreateParameter("EmployeeID", adVarchar, adParamInput,20)
		.Parameters.Append objCmd3.CreateParameter("CDMCProcessOutput", adInteger, adParamOutput)
		
		.Parameters("UserID") = Session("UserID") 'The User who had performed the error checks
		.Parameters("EmployeeID") = strEID'"0" 'Set EmployeeID to 0 so that procedure checks all applications
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd3.Execute        
	
	'Return the result of the Save Function.
     intRecord = objCmd3.Parameters.Item("CDMCProcessOutput")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Applications CDMC details updated!</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Application for " & strEID & " updated from CDMC details!</div>"
	End If
	
End Sub


Public Sub UpdateApplicationMissingDetails(strEID)
'Procedure to run a stored procedure which checks Applications in tblApplications when they are CMC only and adds missing Gender of Date of Birth (which old applications don't have).
	
  	With objCmd6
  	
		.CommandType = 4
		.CommandText = "spCAPSApplicationUpdateMisingDetails"
		
		.Parameters.Append objCmd6.CreateParameter("ApplicationID", adInteger, adParamInput)
		.Parameters.Append objCmd6.CreateParameter("ApplicationUpdateOutput", adInteger, adParamOutput)
		
		.Parameters("ApplicationID") = Session("ApplicationID") 'The User who had performed the error checks
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd6.Execute        
	
	'Return the result of the Save Function.
     intRecord = objCmd6.Parameters.Item("ApplicationUpdateOutput")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Applications with Missing details updated!</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Application for " & strEID & " updated with Missing Gender and Date of Birth from CDMC details!</div>"
	End If
	
End Sub


Public Sub ChangeNameOnCardOLD(strNewName, strAppProcess)
'Procedure to run a stored procedure which updates the Cardholder Name on Card
Dim intRecord

	'Makes sure that the application is only on hold or awaiting review
	If strAppProcess = "On Hold" OR strAppProcess = "Awaiting Review"  Then
	
		With objCmd
		
			.CommandType = 4
			.CommandText = "spCAPSNameOnCardSave"
			
			.Parameters.Append objCmd.CreateParameter("UniqueID", adInteger,adParamInput)
			.Parameters.Append objCmd.CreateParameter("NameOnCard", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("TableName", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("NameChangeIDOutput", adInteger, adParamOutput)
			
			.Parameters("UniqueID") = Session("ApplicationID") 'The Application currently viewed.
			.Parameters("NameOnCard") = strNewName 'The new name on card
			.Parameters("TableName") = "tblCAPSApplication"
			.Parameters("UpdatedBy") = Session("UserID")
			
			.ActiveConnection = objCon
					
		End With
					
		objCmd.Execute        
	
		'Return the result of the Save Function.
		 intRecord = objCmd.Parameters.Item("NameChangeIDOutput")  

		If intRecord = 0 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Application Name On Card NOT updated for " & strNewName & " for application " & Session("ApplicationID") & "!</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Applications Name on Card updated to " & strNewName & " !</div>"
		End If
	Else
	'If the application is not on hold or awaiting review then display error message
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Application Name On Card NOT updated for " & strNewName & ".  The application can only be 'On Hold' or 'Awaiting Review'!</div>"
	End If
	
End Sub

Public Sub ChangeNameOnCard(strAppProcess,strTitle,strFirstName,strSurname,strNewName,intID)
'Procedure to run a stored procedure which updates the Cardholder Name on Card
Dim intRecord

	'Makes sure that the application is only on hold or awaiting review
	If strAppProcess = "On Hold" OR strAppProcess = "Awaiting Review" Then
	
	
		With objCmd
		
			.CommandType = 4
			.CommandText = "spCAPSNameOnCardSave"
			
			.Parameters.Append objCmd.CreateParameter("UniqueID", adInteger,adParamInput)
			.Parameters.Append objCmd.CreateParameter("Title", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("FirstName", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("Surname", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("NameOnCard", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("TableName", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("NameChangeIDOutput", adInteger, adParamOutput)
			
			.Parameters("UniqueID") = intID'Session("CardID") 'The Application currently viewed.
			.Parameters("Title") = strTitle 'The new name on card
			.Parameters("FirstName") = strFirstName 'The new name on card
			.Parameters("Surname") = strSurname 'The new name on card
			.Parameters("NameOnCard") = strNewName 'The new name on card --NO LONGER USED
			.Parameters("TableName") = "tblCAPSApplication"
			.Parameters("UpdatedBy") = Session("UserID")
			
			.ActiveConnection = objCon
					
		End With
					
		objCmd.Execute        
	
		'Return the result of the Save Function.
		 intRecord = objCmd.Parameters.Item("NameChangeIDOutput")  

		If intRecord = -1 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Application Name On Card NOT updated for " & strTitle & " " & strFirstName & " " & strSurname & " !</div>"
		ElseIf intRecord = -2 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Application Name On Card NOT updated for " & strTitle & " " & strFirstName & " " & strSurname & " !</div>"
		ElseIf intRecord = 0 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Application Name On Card NOT updated for " & strTitle & " " & strFirstName & " " & strSurname & " !</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Card Name on Card updated to " & strTitle & " " & strFirstName & " " & strSurname & "!</div>"
		End If
	Else
	'If the application is not on hold or awaiting review then display error message
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Only On Hold or Awaiting Review Applications can have a name change. Name On Card NOT updated for " & strTitle & " " & strFirstName & " " & strSurname & "!</div>"
	End If
	
End Sub


Public Sub ChangeLimit(strNewLimit, strAppProcess)
'Procedure to run a stored procedure which updates the Cardholder Name on Card
Dim intRecord
Dim dblLimitNumber
Dim x

	'Makes sure that the application is only on hold or awaiting review
	If strAppProcess = "On Hold" OR strAppProcess = "Awaiting Review"  Then
	
		'Format the Value passed in to get the float/double value only
		If Len(strNewLimit) <1 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Application Limit NOT updated.  The Credit Limit entered is empty!</div>"
			Exit Sub
		End If
		
		'Loop through the string a get only numbers
		For x = 1 to Len(strNewLimit)
		
			If IsNumeric(Mid(strNewLimit,x,1)) Then
				dblLimitNumber = dblLimitNumber & Mid(strNewLimit,x,1)
			End If
		
		Next
		
		'Make sure there is a value for the resulting number
		If Len(dblLimitNumber) <1 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Application Limit NOT updated.  The Credit Limit entered is NOT a number!</div>"
			Exit Sub
		End If
		
		'The Diners Credit Limit is in cents so multiply the number entered
		dblLimitNumber = dblLimitNumber * 100
		
		
		With objCmd
		
			.CommandType = 4
			.CommandText = "spCAPSLimitSave"
			
			.Parameters.Append objCmd.CreateParameter("UniqueID", adInteger,adParamInput)
			.Parameters.Append objCmd.CreateParameter("Limit", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("TableName", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("LimitChangeIDOutput", adInteger, adParamOutput)
			
			.Parameters("UniqueID") = Session("ApplicationID") 'The Application currently viewed.
			.Parameters("Limit") = dblLimitNumber 'The new Credit Limit
			.Parameters("TableName") = "tblCAPSApplication"
			.Parameters("UpdatedBy") = Session("UserID")
			
			.ActiveConnection = objCon
					
		End With
					
		objCmd.Execute        
	
		'Return the result of the Save Function.
		 intRecord = objCmd.Parameters.Item("LimitChangeIDOutput")  

		If intRecord = 0 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Application Credit Limit NOT updated for " & strNewLimit & " for application " & Session("ApplicationID") & "!</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Applications Limit updated to " & strNewLimit & " !</div>"
		End If
	Else
	'If the application is not on hold or awaiting review then display error message
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Application Limit NOT updated.  The application can only be 'On Hold' or 'Awaiting Review'!</div>"
	End If
	
End Sub


Public Sub SubmitApplication()
'Procedure to Release applications from Awaiting Review to Awaiting Export
Dim intRecord
Dim dblCreditLimit
Dim strSysParamCSOneEID
Dim strEIDOnCS
Dim strAppStatus


	'Divide the Credit Limit by 100 ------- NOT USED HERE AS THIS SHOULD BE DONE ON DISPLAY NOT SAVE
	'If IsNumeric(Request.QueryString("CreditLimit")) Then
	'	dblCreditLimit = Request.QueryString("CreditLimit")/100
	'Else
		'dblCreditLimit = GetSystemAdmin("DTCDefaultCreditLimit")
	'End If
	
	'-----New functionality added October 2021 to check whether there is already a record on the CS File for the Employee before releasing Limit Changes
	'Get the One EID System Parameter, which sets whether there can be only one record per EID on the CS File each file/day
	'A value of 'N' will mean that if the Employee already has a record on the CS file they cannot be added (or their Limit Change released)
	strApplicationType = GetApplicationTypeName(Session("ApplicationID"))
	
	strAppStatus = GetApplicationStatus(Session("ApplicationID"))
	
	
	
	If Right(strApplicationType,12) = "Limit Change" AND strAppStatus = "Awaiting Review" Then	
	
	With objCmd8
  	
		.CommandType = 4
		.CommandText = "spCAPSLimitChangeApplicationsInsert"
		
		.Parameters.Append objCmd8.CreateParameter("ApplicationID", adInteger, adParamInput)
		.Parameters.Append objCmd8.CreateParameter("UpdatedBy", adInteger, adParamInput)
		.Parameters.Append objCmd8.CreateParameter("LimitChangeAppOutputID", adInteger, adParamOutput)
		
		.Parameters("ApplicationID") = Session("ApplicationID")
		.Parameters("UpdatedBy") = Session("UserID") 'The User who had performed the error checks
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd8.Execute     
	'Return the result of the Save Function.
     intRecord = objCmd8.Parameters.Item("LimitChangeAppOutputID")  
	
	End If
	'''END New check for Credit Limit Changes for EID already being on the current CS file to be exported
	
  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSApplicationProcess"

			.Parameters.Append objCmd.CreateParameter("ApplicationID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("ReviewedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CreditLimit", adDouble, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CMSUserID", adVarChar, adParamInput, 50) 
			.Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("CDMCApplicationProcessIDOutput", adInteger, adParamOutput)
			
			.Parameters("ApplicationID") = Session("ApplicationID")
			.Parameters("ReviewedBy") = Session("UserID")
			.Parameters("CreditLimit") = 0'Request.QueryString("CreditLimit")
			.Parameters("CMSUserID") = NULL 'Request.QueryString("CMSUserName")
			.Parameters("Status") = Request.QueryString("Status")
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
	  'Response.Write "EXEC spCAPSApplicationProcess " & Session("ApplicationID") & "," & Session("UserID") & ",0,NULL," & Request.QueryString("Status")
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCApplicationProcessIDOutput") 
		
		'Response.Write "XXX" & Request.QueryString("Status")
	 
		If intRecord = 0 Then
			'strMessageIcon = "&nbsp;&nbsp;<img src=""../images/cross.png"" /> Application " & intRecord & " NOT approved! ERROR..."
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Application " & Session("ApplicationName") & " NOT updated! An Error has occurred. See System Admin with Application ID: " & Session("ApplicationID") & " </div>"
			
			'strMessageColour = "Red"
		Else
			'strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" /> Application " & intRecord & " approved!"
			Response.Write "<div class=""alert alert-success"" role=""alert"">Application " & Session("ApplicationName") & " Updated to: " & Request.QueryString("Status") & "!</div>"
			
			'strMessageColour = "Black"
		End If
		
End Sub

Public Function GetLimitChange()
'New procedure to load Credit Limit data from tblCAPSLimitChange
Dim strSQL
Dim strLimitDateFromLC
Dim strLimitDateToLC
Dim strTransactionLimitDateFromLC
Dim strTransactionLimitDateToLC
Dim objRSLimC

Set objRSLimC = Server.CreateObject("ADODB.Recordset")

	strSQL = "SELECT * FROM tblCAPSLimitChange WITH(NOLOCK) WHERE [ApplicationID] = " & Session("ApplicationID") & ""

	objRSLimC.Open strSQL,objCon

	If objRSLimC.EOF THEN
		GetLimitChange = "<tr><th>Credit Limit Date From</th><td></td></tr>" &  _
				"<tr><th>Credit Limit Date To</th><td></td></tr>" &  _
				"<tr><th>Transaction Limit Date From</th><td></td></tr>" &  _
				"<tr><th>Transaction Limit Date To</th><td></td></tr>"
	Else
		strLimitDateFromLC = "None"
		strLimitDateToLC = "None"
		
		'Do not display the Limit Date if it is the default of 01/01/1900
		If IsNull(objRSLimC("LimitDateFrom")) Then
		Else
			
			If DateDiff("d",objRSLimC("LimitDateFrom"),"1-JAN-1900") =0 Then
			Else
				strLimitDateFromLC = FormatDateTime(objRSLimC("LimitDateFrom"),2) & " <span style=""font-size:12px;font-type:italic; color:gray;"">("  & objRSLimC("LimitDateFrom") & ")</span>"
			End If
		End If
		
		If IsNull(objRSLimC("LimitDateTo")) Then
		Else
			If DateDiff("d",objRSLimC("LimitDateTo"),"1-JAN-1900") =0 Then
			Else
				strLimitDateToLC = FormatDateTime(objRSLimC("LimitDateTo"),2) & " <span style=""font-size:12px;font-type:italic; color:gray;"">("  & objRSLimC("LimitDateTo") & ")</span>"
			End If
		End If
		
		'Do not display the Transaction Limit Date if it is the default of 01/01/1900
		If IsNull(objRSLimC("TransactionLimitDateFrom")) Then
		Else
			
			If DateDiff("d",objRSLimC("TransactionLimitDateFrom"),"1-JAN-1900") =0 Then
			Else
				strTransactionLimitDateFromLC = FormatDateTime(objRSLimC("TransactionLimitDateFrom"),2) & " <span style=""font-size:12px;font-type:italic; color:gray;"">("  & objRSLimC("TransactionLimitDateFrom") & ")</span>"
			End If
		End If
		
		If IsNull(objRSLimC("TransactionLimitDateTo")) Then
		Else
			If DateDiff("d",objRSLimC("TransactionLimitDateTo"),"1-JAN-1900") =0 Then
			Else
				strTransactionLimitDateToLC = FormatDateTime(objRSLimC("TransactionLimitDateTo"),2) & " <span style=""font-size:12px;font-type:italic; color:gray;"">("  & objRSLimC("TransactionLimitDateTo") & ")</span>"
			End If
		End If
		
		GetLimitChange = "<tr><th>Credit Limit Date From</th><td>" & strLimitDateFromLC & "</td></tr>" &  _
				"<tr><th>Credit Limit Date To</th><td>" & strLimitDateToLC & "</td></tr>" & _
				"<tr><th>Transaction Limit Date From</th><td>" & strTransactionLimitDateFromLC & "</td></tr>" &  _
				"<tr><th>Tranasction Limit Date To</th><td>" & strTransactionLimitDateToLC & "</td></tr>"
	End If
				
objRSLimC.Close

End Function

Public Function GetLimitDates()
'Tiffany - new procedure to load Credit Limit data from xml because I can't find it by the names in the limit table
Dim strSQL
Dim strLimitDateFromLD
Dim strLimitDateToLD
Dim strTransactionLimitDateFromLD
Dim strTransactionLimitDateToLD
Dim objRSLimD
Dim strTxnLimitFlag
Dim strMonthlyPerm
Dim dblTransactionLimitOrig
Dim strDCCPSignatures

Set objRSLimD = Server.CreateObject("ADODB.Recordset")

''New for the DCCP ----May 2026
If Left(strAppTypeGlobal,6) = "Portal" Then

	strSQL = "SELECT * FROM tblCAPSLimitDetailsPortal WITH(NOLOCK) WHERE [ApplicationID] = " & Session("ApplicationID") & ""

	objRSLimD.Open strSQL,objCon

	If objRSLimD.EOF THEN

		GetLimitDates = "<tr><th>Changes Permanent (Monthly Limit)?</th><td></td></tr>" & _
		"<tr><th>Credit Limit Date From</th><td></td></tr>" &  _
		"<tr><th>Credit Limit Date To</th><td></td></tr>" &  _
		"<tr><th>Transaction Limit Change?</th><td></td></tr>" & _
		"<tr><th>Transaction Limit Date From</th><td></td></tr>" &  _
		"<tr><th>Transaction Limit Date To</th><td></td></tr>"
	Else
		If IsNull(objRSLimD("TransactionLimitDateFrom")) Then
			strTxnLimitFlag = ""
		Else
			strTxnLimitFlag = "Yes"
		End If
		
		'''Set the global variables fo rTransaction Amounts, which are written in a different palce to the other fields in this function
		If IsNull(objRSLimD("TransactionLimitOriginal")) Then
			dblTransactionLimitOriginal = ""
		Else
			If IsNumeric(objRSLimD("TransactionLimitOriginal")) Then
				dblTransactionLimitOriginal = FormatCurrency(objRSLimD("TransactionLimitOriginal"),2)
			Else
				dblTransactionLimitOriginal = objRSLimD("TransactionLimitOriginal")
			End If
		End If

		If IsNull(objRSLimD("TransactionLimitNew")) Then
			dblTransactionLimitNew = ""
		Else
			If IsNumeric(objRSLimD("TransactionLimitNew")) Then
				dblTransactionLimitNew = FormatCurrency(objRSLimD("TransactionLimitNew"),2)
			Else
				dblTransactionLimitNew = objRSLimD("TransactionLimitNew")
			End If
		End If

		strCreditLimitFrom = objRSLimD("CreditLimitDateFrom")
		strCreditLimitTo  = objRSLimD("CreditLimitDateTo")

	
		GetLimitDates = "<tr><th>Changes Permanent (Monthly Limit)?</th><td>" & strCreditLimitPermanent  & "</td></tr>" & _
		"<tr><th>Credit Limit Date From</th><td>" & strCreditLimitFrom & "</td></tr>" &  _
		"<tr><th>Credit Limit Date To</th><td>" & strCreditLimitTo & "</td></tr>" &  _
		"<tr><th>Transaction Limit Change?</th><td>" & strTxnLimitFlag & "</td></tr>" & _
		"<tr><th>Transaction Limit Date From</th><td>" & objRSLimD("TransactionLimitDateFrom") & "</td></tr>" &  _
		"<tr><th>Transaction Limit Date To</th><td>" & objRSLimD("TransactionLimitDateTo") & "</td></tr>"

		'''Write out the new fields in the table that relate to the DCCP
		If NOT IsNull(objRSLimD("TransactionLimitOriginal")) Then
			If IsNumeric(objRSLimD("TransactionLimitOriginal")) Then
				dblTransactionLimitOrig = FormatCurrency(objRSLimD("TransactionLimitOriginal"),2)
			Else
				dblTransactionLimitOrig = objRSLimD("TransactionLimitOriginal")
			End If

			GetLimitDates = GetLimitDates & "<tr><th>Transaction Limit Original</th><td>" & dblTransactionLimitOrig  & "</td></tr>"
		End If

		
				
		
		If NOT IsNull(objRSLimD("EMAIL")) Then
			'GetLimitDates = GetLimitDates & "<tr><th>EMAIL</th><td>" & objRSLimD("EMAIL")  & "</td></tr>"
			strDCCPSignatures = strDCCPSignatures & "<tr><th>EMAIL</th><td>" & objRSLimD("EMAIL")  & "</td></tr>"
		End If

		If NOT IsNull(objRSLimD("DIRECTOR")) Then
			'GetLimitDates = GetLimitDates & "<tr><th>DIRECTOR</th><td>" & objRSLimD("DIRECTOR")  & "</td></tr>"
			strDCCPSignatures = strDCCPSignatures &  "<tr><th>DIRECTOR</th><td>" & objRSLimD("DIRECTOR")  & "</td></tr>"
		End If

		If NOT IsNull(objRSLimD("ASFIN")) Then
			'GetLimitDates = GetLimitDates & "<tr><th>ASFIN</th><td>" & objRSLimD("ASFIN")  & "</td></tr>"
			strDCCPSignatures = strDCCPSignatures & "<tr><th>ASFIN</th><td>" & objRSLimD("ASFIN")  & "</td></tr>"
		End If

		If NOT IsNull(objRSLimD("CFO")) Then
			'GetLimitDates = GetLimitDates & "<tr><th>CFO</th><td>" & objRSLimD("CFO")  & "</td></tr>"
			strDCCPSignatures = strDCCPSignatures & "<tr><th>CFO</th><td>" & objRSLimD("CFO")  & "</td></tr>"
		End If

		If strDCCPSignatures = "" Then
		Else
			GetLimitDates = GetLimitDates & "<TR><TH class=""updated"">Signatures Required - DCCP</TH><TH class=""updated""></TH></TR>" & strDCCPSignatures
		End If

	End If

	objRSLimD.Close

Else


	strSQL = "SELECT * FROM tblCAPSXMLApplication WITH(NOLOCK) WHERE [ApplicationID] = " & Session("ApplicationID") & ""

	objRSLimD.Open strSQL,objCon

	If objRSLimD.EOF THEN
		GetLimitDates = "<tr><th>Changes Permanent (Monthly Limit)?</th><td></td></tr>" & _
				"<tr><th>Credit Limit Date From</th><td></td></tr>" &  _
				"<tr><th>Credit Limit Date To</th><td></td></tr>" &  _
				"<tr><th>Transaction Limit Change?</th><td></td></tr>" & _
				"<tr><th>Transaction Limit Date From</th><td></td></tr>" &  _
				"<tr><th>Transaction Limit Date To</th><td></td></tr>"
	Else
		strLimitDateFromLD = "None"
		strLimitDateToLD = "None"
		strTransactionLimitDateFromLD = ""
		strTransactionLimitDateToLD = ""
		strTxnLimitFlag = ""
		strMonthlyPerm = ""
					
		If IsNull(objRSLimD("sublimitChange_chkIncreaseDecreaseLimit")) OR objRSLimD("sublimitChange_chkIncreaseDecreaseLimit") = "" Then
		Else
			strTxnLimitFlag = objRSLimD("sublimitChange_chkIncreaseDecreaseLimit")			
		End If
		
		If IsNull(objRSLimD("subLimitChange_grpChangesPermanent")) OR objRSLimD("subLimitChange_grpChangesPermanent") = "" Then
		Else
			strMonthlyPerm = objRSLimD("subLimitChange_grpChangesPermanent")			
		End If
		
		'Do not display the Limit Date if it is the default of 01/01/1900
		If IsNull(objRSLimD("subLimitChange_subPeriodofChangeFromDateToDate_dteStart")) Then
		Else
				strLimitDateFromLD = objRSLimD("subLimitChange_subPeriodofChangeFromDateToDate_dteStart")
				'strLimitDateFromLD = "<tr><th>Credit Limit Date From</th><td>" & strLimitDateFromLD & "</td></tr>"
		End If
		
		If IsNull(objRSLimD("subLimitChange_subPeriodofChangeFromDateToDate_dteEnd")) Then
		Else
				strLimitDateToLD = objRSLimD("subLimitChange_subPeriodofChangeFromDateToDate_dteEnd")
				'strLimitDateToLD = "<tr><th>Credit Limit Date To</th><td>" & strLimitDateToLD & "</td></tr>"
		End If
		
		'Do not display the Transaction Limit Date if it is the default of 01/01/1900
		If IsNull(objRSLimD("dteRequestedDateFrom")) Then
		Else
				strTransactionLimitDateFromLD = Left(objRSLimD("dteRequestedDateFrom"),11)
				'strTransactionLimitDateFromLD = "<tr><th>Transaction Limit Date From</th><td>" & strTransactionLimitDateFromLD & "</td></tr>" 
		End If
		
		If IsNull(objRSLimD("dteRequestedDateTo")) Then
		Else
				strTransactionLimitDateToLD = Right(objRSLimD("dteRequestedDateTo"),11)
				'strTransactionLimitDateToLD = "<tr><th>Transaction Limit Date To</th><td>" & strTransactionLimitDateToLD & "</td></tr>"
		End If
				
		Select CASE strMonthlyPerm
			CASE "Yes"
				strMonthlyPerm = "<tr><th>Changes Permanent (Monthly Limit)?</th><td>" & strMonthlyPerm & "</td></tr>"
				strLimitDateFromLD = ""
				strLimitDateToLD = ""
			CASE "No"
				strMonthlyPerm = "<tr><th>Changes Permanent (Monthly Limit)?</th><td>" & strMonthlyPerm & "</td></tr>"
				strLimitDateFromLD = "<tr><th>Credit Limit Dates</th><td> <b>Start:</b>&nbsp;&nbsp;" & strLimitDateFromLD & " &nbsp;<b>End:</b>&nbsp; " & strLimitDateToLD & "</td></tr>"
			CASE Else
				strMonthlyPerm = ""
				strMonthlyPerm = "<tr><th>Changes Permanent (Monthly Limit)?</th><td>" & strMonthlyPerm & "</td></tr>"
		End Select
		
		Select CASE strTxnLimitFlag
			CASE "1"
				strTxnLimitFlag = "Yes"
				strTxnLimitFlag = "<tr><th>Transaction Limit Change?</th><td>" & strTxnLimitFlag & "</td></tr>"
				strTransactionLimitDateFromLD = "<tr><th>Transaction Limit Dates</th><td> <b>Start:</b>&nbsp;&nbsp;" & strTransactionLimitDateFromLD & " &nbsp;<b>End:</b>&nbsp; " & strTransactionLimitDateToLD & " </td></tr>" 
			CASE "0"
				strTxnLimitFlag = "No"	
				strTxnLimitFlag = "<tr><th>Transaction Limit Change?</th><td>" & strTxnLimitFlag & "</td></tr>"
				strTransactionLimitDateToLD = ""
				strTransactionLimitDateFromLD = ""
			CASE Else
				strTxnLimitFlag = ""
				strTxnLimitFlag = "<tr><th>Transaction Limit Change?</th><td>" & strTxnLimitFlag & "</td></tr>"
		End Select
		
		GetLimitDates =  strMonthlyPerm & strLimitDateFromLD & strTxnLimitFlag & strTransactionLimitDateFromLD
				
				
	End If

	objRSLimD.Close

SET objRSLimD = Nothing

End If				


End Function

Public Function Get_Signatures()

'New procedure to load Signatures from tblCAPSXMLApplication
Dim strSQL
Dim objRSRank
Dim objRSSig
Dim strResponse
Dim strColour
Dim strClass
Dim strClass1
Dim strClass2
Dim strClass3
Dim strClass4
Dim dblNewCreditLimit
Dim strSigCheck
Dim strSESApprover
Dim strASFINSigName
Dim strCFOSigName
Dim strSESSigName
Dim strSESRank
Dim objRSLimD

Set objRSSig = Server.CreateObject("ADODB.Recordset")
Set objRSLimD = Server.CreateObject("ADODB.Recordset")
	
	strSQL = "SELECT * FROM tblCAPSXMLApplication WHERE ApplicationID = " & Session("ApplicationID") & ""


	objRSSig.Open strSQL,objCon
	
		If Not objRSSig.EOF Then
		
			dblNewCreditLimit = objRSSig("subLimitChange_numNewLimit")
			strSigCheck = ""'objRSSig("ffApprover")
			strSESApprover = objRSSig("subSESApprover")
			
			
			''''New Jan 2025 --- for the display of application version on the screen get the Version Number from the XML application
			If IsNull(objRSSig("AppVersion")) Then 
				lngApplicationVersion = 0
			Else
				lngApplicationVersion = objRSSig("AppVersion")
			End If
						
			Select Case objRSSig("grpIAmApplyingFor")
			
				Case 1
				
					If objRSSig("subApply_subDigitalSignature1_fldSignFlag") = "true" Then						
						strClass = "alert alert-success"						
					Else					
						strClass = "alert alert-danger"							
					End If
					
					strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subApply_subDigitalSignature1_fldSignFlag") & "</div></TR>"
				
				Case 2
				
					If objRSSig("subApply_subDigitalSignature1_fldSignFlag") = "true" Then						
						strClass = "alert alert-success"						
					Else					
						strClass = "alert alert-danger"							
					End If
				
					strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subApply_subDigitalSignature1_fldSignFlag") & "</div></TR>"
				
				Case 7
				
					If objRSSig("subApply_subDigitalSignature1_fldSignFlag") = "true" Then						
						strClass = "alert alert-success"						
					Else					
						strClass = "alert alert-danger"							
					End If
				
					strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subApply_subDigitalSignature1_fldSignFlag") & "</div></TR>"
				
				
				Case 3
				
					If objRSSig("subApply_subDigitalSignature1_fldSignFlag") = "true" Then						
						strClass = "alert alert-success"						
					Else					
						strClass = "alert alert-danger"							
					End If
					
					strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subApply_subDigitalSignature1_fldSignFlag") & "</div></TR>"
				
				Case 4
				
					If objRSSig("subApply_subDigitalSignature1_fldSignFlag") = "true" Then						
						strClass = "alert alert-success"						
					Else					
						strClass = "alert alert-danger"							
					End If
					
					If objRSSig("subApply_subDigitalSignature2_fldSignFlag") = "true" Then						
						strClass1 = "alert alert-success"						
					Else					
						strClass1 = "alert alert-danger"							
					End If
					
					strResponse = "<TR><TH>Signatures Required</TH></TR><TR><TH>" & strSigCheck & "</TH><TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subApply_subDigitalSignature1_fldSignFlag") & "</div></TR>" & _
					"<TR><TH>Supervisor Signature</TH><TH><div class=""" & strClass1 & """>" & objRSSig("subApply_subDigitalSignature2_fldSignFlag") & "</div></TR>"
				Case 5				
					
						If dblNewCreditLimit < 100001 Then
						
							If objRSSig("fldIsSigned") = "true" Then						
								strClass = "alert alert-success"						
							Else					
								strClass = "alert alert-danger"							
							End If
							
							If objRSSig("subDigitalSignature5_fldSignFlag") = "true" Then						
								strClass2 = "alert alert-success"						
							Else					
								strClass2 = "alert alert-danger"							
							End If
						
							strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature  </TH><TH><div class=""" & strClass & """>" & objRSSig("fldIsSigned") & "</div></TR>" & _
							"<TR><TH>SES Signature</TH><TH><div class=""" & strClass2 & """>" & objRSSig("subDigitalSignature5_fldSignFlag") & " :  " & strSESApprover & "</div></TR>"
						
						
						End If
						
						
					
						If dblNewCreditLimit > 100000 AND dblNewCreditLimit < 500000 Then 
						
							If objRSSig("fldIsSigned") = "true" Then						
								strClass = "alert alert-success"						
							Else					
								strClass = "alert alert-danger"							
							End If
							
							
							If objRSSig("subDigitalSignature5_fldSignFlag") = "true" Then						
								strClass3 = "alert alert-success"						
							Else					
								strClass3 = "alert alert-danger"							
							End If
						
							strASFINSigName = "(" & objRSSig("subDigitalSignature5_fldSAEmployeeID") & " : " & objRSSig("subDigitalSignature5_fldCMAGivenNames") & " " & objRSSig("subDigitalSignature5_fldCMAFamilyName") & ")"
																					
							strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("fldIsSigned") & "</div></TR>" & _
							"<TR><TH>AS FIN Signature</TH><TH><div class=""" & strClass3 & """>" & objRSSig("subDigitalSignature5_fldSignFlag") & " " & strASFinSigName & "</div></TR>"
						
						
						End If
						
						If dblNewCreditLimit >= 500000 Then
						
							If objRSSig("fldIsSigned") = "true" Then						
								strClass = "alert alert-success"						
							Else					
								strClass = "alert alert-danger"							
							End If
							
							If objRSSig("subDigitalSignature5_fldSignFlag") = "true" Then						
								strClass3 = "alert alert-success"						
							Else					
								strClass3 = "alert alert-danger"							
							End If
							
							If objRSSig("subDigitalSignature7_fldSignFlag") = "true" Then						
								strClass4 = "alert alert-success"						
							Else					
								strClass4 = "alert alert-danger"							
							End If
							
							strASFINSigName = "(" & objRSSig("subDigitalSignature5_fldSAEmployeeID") & " : " & objRSSig("subDigitalSignature5_fldCMAGivenNames") & " " & objRSSig("subDigitalSignature5_fldCMAFamilyName") & ")"
							strCFOSigName = "(" & objRSSig("subDigitalSignature6_fldSAEmployeeID") & " : " & objRSSig("subDigitalSignature6_fldCMAGivenNames") & " " & objRSSig("subDigitalSignature6_fldCMAFamilyName") & ")"					
														
							strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("fldIsSigned") & "</div></TR>" & _
							"<TR><TH>AS FIN Signature</TH><TH><div class=""" & strClass3 & """>" & objRSSig("subDigitalSignature5_fldSignFlag") & " " & strASFinSigName & "</div></TR>" & _
							"<TR><TH>CFO Signature</TH><TH><div class=""" & strClass4 & """>" & objRSSig("subDigitalSignature6_fldSignFlag") & " " & strCFOSigName & "</div></TR>"	
						
						End If
				
				Case 6
				
						If dblNewCreditLimit < 100001 Then
						
							If objRSSig("subLimitChange_subDigitalSignature4_fldSignFlag") = "true" Then						
								strClass = "alert alert-success"						
							Else					
								strClass = "alert alert-danger"							
							End If
							
							If objRSSig("subDigitalSignature5_fldSignFlag") = "true" Then						
								strClass2 = "alert alert-success"						
							Else					
								strClass2 = "alert alert-danger"							
							End If
						
							strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subLimitChange_subDigitalSignature4_fldSignFlag") & "</div></TR>" & _
							"<TR><TH>SES Signature</TH><TH><div class=""" & strClass2 & """>" & objRSSig("subDigitalSignature5_fldSignFlag") & " :  " & strSESApprover & "</div></TR>"
						
						
						End If
						
						
					
						If dblNewCreditLimit > 100000 AND dblNewCreditLimit < 500000 Then 
						
							If objRSSig("subLimitChange_subDigitalSignature4_fldSignFlag") = "true" Then						
								strClass = "alert alert-success"						
							Else					
								strClass = "alert alert-danger"							
							End If
							
							
							If objRSSig("subDigitalSignature5_fldSignFlag") = "true" Then						
								strClass3 = "alert alert-success"						
							Else					
								strClass3 = "alert alert-danger"							
							End If
						
							strASFINSigName = "(" & objRSSig("subDigitalSignature5_fldSAEmployeeID") & " : " & objRSSig("subDigitalSignature5_fldCMAGivenNames") & " " & objRSSig("subDigitalSignature5_fldCMAFamilyName") & ")"
											
														
							strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subLimitChange_subDigitalSignature4_fldSignFlag") & "</div></TR>" & _
							"<TR><TH>AS FIN Signature</TH><TH><div class=""" & strClass3 & """>" & objRSSig("subDigitalSignature5_fldSignFlag") & " " & strASFinSigName & "</div></TR>"
							
						
						
						End If
						
						If dblNewCreditLimit >= 500000 Then
						
							If objRSSig("subLimitChange_subDigitalSignature4_fldSignFlag") = "true" Then						
								strClass = "alert alert-success"						
							Else					
								strClass = "alert alert-danger"							
							End If
							
							If objRSSig("subDigitalSignature5_fldSignFlag") = "true" Then						
								strClass3 = "alert alert-success"						
							Else					
								strClass3 = "alert alert-danger"							
							End If
							
							If objRSSig("subDigitalSignature6_fldSignFlag") = "true" Then						
								strClass4 = "alert alert-success"						
							Else					
								strClass4 = "alert alert-danger"							
							End If
							
							strASFINSigName = "(" & objRSSig("subDigitalSignature5_fldSAEmployeeID") & " : " & objRSSig("subDigitalSignature5_fldCMAGivenNames") & " " & objRSSig("subDigitalSignature5_fldCMAFamilyName") & ")"
							strCFOSigName = "(" & objRSSig("subDigitalSignature6_fldSAEmployeeID") & " : " & objRSSig("subDigitalSignature6_fldCMAGivenNames") & " " & objRSSig("subDigitalSignature6_fldCMAFamilyName") & ")"					
														
							strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subLimitChange_subDigitalSignature4_fldSignFlag") & "</div></TR>" & _
							"<TR><TH>AS FIN Signature</TH><TH><div class=""" & strClass3 & """>" & objRSSig("subDigitalSignature5_fldSignFlag") & " " & strASFinSigName & "</div></TR>" & _
							"<TR><TH>CFO Signature</TH><TH><div class=""" & strClass4 & """>" & objRSSig("subDigitalSignature6_fldSignFlag") & " " & strCFOSigName & "</div></TR>"	
						
						End If
						
				Case 8
				
						If dblNewCreditLimit < 100001 Then
						
							If objRSSig("subLimitChange_subDigitalSignature3_fldSignFlag") = "true" Then						
								strClass = "alert alert-success"						
							Else					
								strClass = "alert alert-danger"							
							End If
							
							If objRSSig("subDigitalSignature5_fldSignFlag") = "true" Then						
								strClass2 = "alert alert-success"						
							Else					
								strClass2 = "alert alert-danger"							
							End If
						
							strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subLimitChange_subDigitalSignature3_fldSignFlag") & "</div></TR>" & _
							"<TR><TH>SES Signature</TH><TH><div class=""" & strClass2 & """>" & objRSSig("subDigitalSignature5_fldSignFlag") & " :  " & strSESApprover & "</div></TR>"
						
						
						End If
						
						
					
						If dblNewCreditLimit > 100000 AND dblNewCreditLimit < 500000 Then 
						
							If objRSSig("subLimitChange_subDigitalSignature3_fldSignFlag") = "true" Then						
								strClass = "alert alert-success"						
							Else					
								strClass = "alert alert-danger"							
							End If
							
							
							If objRSSig("subDigitalSignature5_fldSignFlag") = "true" Then						
								strClass3 = "alert alert-success"						
							Else					
								strClass3 = "alert alert-danger"							
							End If
						
							strASFINSigName = "(" & objRSSig("subDigitalSignature5_fldSAEmployeeID") & " : " & objRSSig("subDigitalSignature5_fldCMAGivenNames") & " " & objRSSig("subDigitalSignature5_fldCMAFamilyName") & ")"
											
														
							strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subLimitChange_subDigitalSignature3_fldSignFlag") & "</div></TR>" & _
							"<TR><TH>AS FIN Signature</TH><TH><div class=""" & strClass3 & """>" & objRSSig("subDigitalSignature5_fldSignFlag") & " " & strASFinSigName & "</div></TR>"
							
						
						
						End If
						
						If dblNewCreditLimit >= 500000 Then
						
							If objRSSig("subLimitChange_subDigitalSignature3_fldSignFlag") = "true" Then						
								strClass = "alert alert-success"						
							Else					
								strClass = "alert alert-danger"							
							End If
							
							If objRSSig("subDigitalSignature5_fldSignFlag") = "true" Then						
								strClass3 = "alert alert-success"						
							Else					
								strClass3 = "alert alert-danger"							
							End If
							
							If objRSSig("subDigitalSignature6_fldSignFlag") = "true" Then						
								strClass4 = "alert alert-success"						
							Else					
								strClass4 = "alert alert-danger"							
							End If
							
							strASFINSigName = "(" & objRSSig("subDigitalSignature5_fldSAEmployeeID") & " : " & objRSSig("subDigitalSignature5_fldCMAGivenNames") & " " & objRSSig("subDigitalSignature5_fldCMAFamilyName") & ")"
							strCFOSigName = "(" & objRSSig("subDigitalSignature6_fldSAEmployeeID") & " : " & objRSSig("subDigitalSignature6_fldCMAGivenNames") & " " & objRSSig("subDigitalSignature6_fldCMAFamilyName") & ")"					
														
							strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """>" & objRSSig("subLimitChange_subDigitalSignature3_fldSignFlag") & "</div></TR>" & _
							"<TR><TH>AS FIN Signature</TH><TH><div class=""" & strClass3 & """>" & objRSSig("subDigitalSignature5_fldSignFlag") & " " & strASFinSigName & "</div></TR>" & _
							"<TR><TH>CFO Signature</TH><TH><div class=""" & strClass4 & """>" & objRSSig("subDigitalSignature6_fldSignFlag") & " " & strCFOSigName & "</div></TR>"	
						
						End If
				
			End Select

			

			Get_Signatures = strResponse
		
		End If
	
	objRSSig.Close

'Get Portal 

strSQL = "SELECT * FROM tblCAPSLimitDetailsPortal WITH(NOLOCK) WHERE [ApplicationID] = " & Session("ApplicationID") & ""

				objRSLimD.Open strSQL,objCon

				If objRSLimD.eof then

				
				Else

					strClass = "alert alert-success"
					strResponse = "<TR><TH class=""updated"">Signatures Required</TH><TH class=""updated"">" & strSigCheck & "</TH></TR><TR><TH>Applicant Signature</TH><TH><div class=""" & strClass & """></div></TR>" & _
							"<TR><TH>AS FIN Signature</TH><TH><div class=""" & strClass3 & """>true " & strASFinSigName & "</div></TR>" & _
							"<TR><TH>CFO Signature</TH><TH><div class=""" & strClass4 & """>true " & strCFOSigName & "</div></TR>"	
						

				End If

			objRSLimD.Close

	Get_Signatures = strResponse
	

End Function

Public Sub GetLimitXML()
'New procudure to load Credit Limit data from tblCAPSXMLApplication
Dim strSQL
Dim objRSLimXML

Set objRSLimXML = Server.CreateObject("ADODB.Recordset")


	strSQL = "SELECT [subLimitChange_subRequestedTransactionAmountDates_ddlRequestedLimit],[subLimitChange_numNewLimit] FROM tblCAPSXMLApplication WITH(NOLOCK) WHERE [ApplicationID] = " & Session("ApplicationID") & ""

	objRSLimXML.Open strSQL,objCon

	If objRSLimXML.EOF THEN
		lngTransactionAmountXML = "none"
		lngCreditAmountXML = "none"
	Else
		'Get the New Transaction Limit applied for from the XML application table
		If IsNull(objRSLimXML("subLimitChange_subRequestedTransactionAmountDates_ddlRequestedLimit")) Then
			lngTransactionAmountXML = "none"
		Else
			'Get the Transaction Limit applied for
			If objRSLimXML("subLimitChange_subRequestedTransactionAmountDates_ddlRequestedLimit") = "" Then
				lngTransactionAmountXML = "$0"
			Else
				lngTransactionAmountXML = FormatCurrency(objRSLimXML("subLimitChange_subRequestedTransactionAmountDates_ddlRequestedLimit"),0)
			End If
			
		End If
		
		If IsNull(objRSLimXML("subLimitChange_numNewLimit")) Then
			lngCreditAmountXML = "none"
		Else			
			'Get the Credit Limit Amount applied for
			If objRSLimXML("subLimitChange_numNewLimit") = "" Then
				lngCreditAmountXML = "$0"
			Else
				lngCreditAmountXML = FormatCurrency(objRSLimXML("subLimitChange_numNewLimit"),0)
			End If
			'If IsNumeric(lngCreditAmountXML) Then
			'	lngCreditAmountXML = FormatCurrency(lngCreditAmountXML,2)
			'End If
		End If
		
	End If
	
		
objRSLimXML.Close


End Sub


Public Function GetCardDetails(strEmployeeID,strCardType)
'Procedure to load any audit records relating to the card
Dim strSQL
Dim strPerson
Dim objRSCard
Dim x
Dim strStatus
Dim strActive
Dim strCreditLimit

If IsNull(strEmployeeID) OR strEmployeeID = "" Then Exit Function

Set objRSCard = Server.CreateObject("ADODB.Recordset")

	'Determine the recordset based on the Card Type
	If strCardType = "DTC" Then
		strSQL = "SELECT [CardType],[CardTypeSub],[Status],[ActivationFlag],[Expiry],[CardNumberShort],[CardNumber],[CreditLimit] FROM tblCAPSCard WITH(NOLOCK) WHERE CardType = 'DTC' AND [EmployeeID] = '" & strEmployeeID & "' AND (Status = '00' or Status = '') "'AND Active = 'Y'"
	Else
		strSQL = "SELECT [CardType],[CardTypeSub],[Status],[ActivationFlag],[Expiry],[CardNumber],[CreditLimit] FROM tblCAPSCard WITH(NOLOCK) WHERE CardType = 'DPC' AND [EmployeeID] = '" & strEmployeeID & "' AND Status In ('','00') "'" AND Active = 'Y'"
	End If
	
	objRSCard.Open strSQL,objCon

	If objRSCard.EOF THEN
		GetCardDetails = "<tr><th>Existing Card(s)</th><td style=""color:grey; font-style:italic;"">No Card details for " & Session("ApplicationEmployeeID") & "</td></tr>"
	End If
		
    Do Until objRSCard.EOF
		x = x + 1
		
		strStatus = objRSCard("Status")
		
		'Make the status Active if the card is a DPC and has no value (empty string) and do not display Active Flag
		If strCardType = "DPC" Then
		
			strActive = ""
			
			If strStatus = "" Then
				strStatus = "<span style=""color:green; font-weight:bold;"">Active</span>"' "Active"
			Else	
				strStatus = objRSCard("Status")
			End If
			
			'Set the Status Colour
			If strStatus = "00" Then
				strStatus = "<span style=""color:green; font-weight:bold;"">" & strStatus & "</span>"
			ElseIf strStatus = "01" OR strStatus = "02" OR strStatus = "03" Then
				strStatus = "<span style=""color:red; font-weight:bold;"">" & strStatus & "</span>"
			Else
			
			End If
			
		Else
			strActive = "<b>ActiveFlag:</b>" & objRSCard("ActivationFlag")
		End If
		
		'''Update for new DPC change April 2023 --added the card limit to the display
		If IsNull(objRSCard("CreditLimit")) Then
			strCreditLimit = ""
		Else
			If IsNumeric(objRSCard("CreditLimit")) Then
				If objRSCard("CreditLimit") > 0 Then
					If strCardType = "DPC ANZ" Then
						'strCreditLimit = FormatCurrency(objRSCard("CreditLimit"),0)
						strCreditLimit = FormatCurrency(objRSCard("CreditLimit")/100,0)

					Else
						strCreditLimit = FormatCurrency(objRSCard("CreditLimit"),0)
					End If
					
				Else
					strCreditLimit = FormatCurrency(objRSCard("CreditLimit"),0)
				End If
			Else
				strCreditLimit = objRSCard("CreditLimit")
			End If
			
		End If
		
		GetCardDetails = GetCardDetails & "<tr><th style=""color:grey; font-style:italic;"">Existing Card(s)</th><td Title=""Limit is Current Credit Limit for Card"">" & x & ". " & objRSCard("CardType") & " " & objRSCard("CardTypeSub") & " <b>Status:</b>" & strStatus & " " & strActive & " <b>Expiry:</b>" & objRSCard("Expiry") & " <b>No:</b>" & MaskCard(objRSCard("CardNumber")) & " </BR>&nbsp;&nbsp;<b>Currrent Credit Limit:</b>" & strCreditLimit & "</td></tr>"
		'GetCardDetails = GetCardDetails & "<tr><th style=""color:grey; font-style:italic;"">Existing Card(s)</th><td>" & x & ". " & objRSCard("CardType") & " " & objRSCard("CardTypeSub") & " <b>Status:</b>" & strStatus & " " & strActive & " <b>Expiry:</b>" & objRSCard("Expiry") & " <b>No:</b>" & MaskCard(objRSCard("CardNumber")) & "</td></tr>"
	
		objRSCard.Movenext
	Loop
	
	
objRSCard.Close
Set objRSCard = Nothing

End Function


Public Sub UpdateEmailErrorTemplate()
'Procedure to run a stored procedure which checks all AE602 XML applications (on hold or awaiting review) for errors and adds errors or resolves them.
	
  	With objCmd5
  	
		.CommandType = 4
		.CommandText = "spCAPSProcessEmailErrorTemplate"
		
		.Parameters.Append objCmd5.CreateParameter("ApplicationID", adInteger, adParamInput)
				
		.Parameters("ApplicationID") = Session("ApplicationID") 'Set Application to 0 so that procedure checks all applications
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd5.Execute        
	
End Sub

Public Sub ReSendErrorEmail()

'Resend the Error Email

Dim strResponse
Dim strEmailSentTo
	
  	With objCmd7
  	
		.CommandType = 4
		.CommandText = "spCAPSReSendErrorEmail"
		
		.Parameters.Append objCmd7.CreateParameter("ApplicationID", adInteger, adParamInput)
		.Parameters.Append objCmd7.CreateParameter("UpdatedBy", adInteger, adParamInput)
		.Parameters.Append objCmd7.CreateParameter("Response", adVarChar, adParamOutput, 50)
		.Parameters.Append objCmd7.CreateParameter("EmailSentTo", adVarChar, adParamOutput, 500)	
				
		.Parameters("ApplicationID") = Session("ApplicationID") 'Set Application to 0 so that procedure checks all applications
		.Parameters("UpdatedBy") = Session("UserID") 'Set Application to 0 so that procedure checks all applications
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd7.Execute

		'Return the result of the Save Function.
		strResponse = objCmd7.Parameters.Item("Response")
		strEmailSentTo = objCmd7.Parameters.Item("EmailSentTo")  		

		If strResponse = "No Email to Send." Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No Email to send for application " & Session("ApplicationID") & ".</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Email sent to " & strEmailSentTo & " for application " & Session("ApplicationID") & ".</div>"
		End If
	
End Sub



Set objRS = Nothing
Set objCon = Nothing



%>