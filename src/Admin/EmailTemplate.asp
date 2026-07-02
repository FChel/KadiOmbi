
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
	'Response.Write "EMAILLLLLLL"
	''Send_Email "michael.giacomin@defence.gov.au","andrew.bull3@defence.gov.au","Subject","Body of email</I>","","HTML"
	'Send_Email "andrew.bull@isidore.com","andrew.bull3@defence.gov.au","Subject","Body of email</I>","","HTML"
	

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

    Session("CurrentPage") = "Admin/EmailTemplate.asp"

	If IsNull(Session("ApplicationID")) OR Session("ApplicationID") = "" Then Session("ApplicationID")= 0

	If Not IsEmpty(Request.QueryString("StyleSheet")) Then
		Session("StyleSheet") = Request.QueryString("StyleSheet")
		
		Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""" & Request.QueryString("StyleSheet") & """>"
	Else
		If IsEmpty(Session("StyleSheet")) Then Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""../CAPSStyle.css"">"
	End If

		Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""../CAPSStyle.css"">"


'Execute Action

If Request.QueryString("Action") = "Save" Then

	Call SaveEmailTemplate() 

End If

If Request.QueryString("Action") = "Delete" Then   
	
    Call DeleteData(Request.QueryString("EmailDetailID"))
   
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
	
	var id = document.getElementById('EmailDetailID').value
	self.location = "EmailTemplate.asp?Action=Save&EmailDetailID=" + id;

}

function loadDocE(cb) {

  var id = cb.getAttribute('data-id');
  var xhttp = new XMLHttpRequest();

  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("GetEmailDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetEmailDetail.asp?EmailDetailID=" + id, true);
  xhttp.send();
}

function loadDelete(cb) {

var id = cb.getAttribute('data-id');
var name = cb.getAttribute('data-EmailTemplateName');

	document.getElementById("ModalDeleteMessage").innerHTML = 'Delete the selected Email Template - ' + name + '?';
	document.getElementById("ModalDelete").style.display = "block";
	document.getElementById("EmailTemplateDeleteID").value = id;
	
}

function deleteEmailTemplate(cb) {

	var id = document.getElementById("EmailTemplateDeleteID").value
	self.location = "EmailTemplate.asp?Action=Delete&EmailDetailID=" + id;

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
<main class="main py-4">
	<div class="container">			
		<div class="row mb-3">
			<div class="col-md-10">
					<h3 class="text-left">Email Template Configuration </h3>
			</div>

			<div class="col-md-2 text-right"></div>
		
				<div class="row">
					<div class="col-md-3">
						<h6 class="mb-3">Select Email Template</h6>
						<div
							class="nav flex-column nav-pills"
							id="v-pills-tab"
							role="tablist"
							aria-orientation="vertical">
								<%Call DisplayTableLeft()%>
						</div>
					</div>
					<div class="col-md-9">
						<div class="tab-content" id="v-pills-tabContent">
							<%Call DisplayTableRight()%>
						</div>
					</div>
				</div>			
			</div>
		</div>	
	</div>

	<!-- Start Delete Modal -->

	<div class="modal fade" id="ModalDelete" tabindex="-1" role="dialog" aria-labelledby="ModalDeleteCenterTitle" aria-hidden="true">
		<div class="modal-dialog modal-dialog-centered" role="document">
		  	<div class="modal-content">
				<div class="modal-header">
			  		<h5 class="modal-title" id="ModalDeleteLongTitle" style="font-weight:bold;">Delete Email Template</h5>
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
					  	<div class="col-md-12 mb-3">
						  	<input type="hidden" id="EmailTemplateDeleteID"></input>
						  	<button class="btn btn-primary btn-sm" onClick="deleteEmailTemplate(this)">Yes</button>
						 	 <button type="button" class="btn btn-secondary" data-dismiss="modal" onClick="DeleteModalClose(this);"><i class="fa fa-close"></i>No</button>
					  	</div>
				  	</div>
			  	</div>
			</div>	
		</div>
		<div class="modal-footer"></div>
	</div>

	<!-- End Delete Modal -->

	<!-- Start Edit Modal -->
	<form action="EmailTemplate.asp?Action=Save" method="POST" id="frm" name="frm" class="needs-validation" novalidate>
		<div class="modal fade" id="emailModal" tabindex="-1" role="dialog" aria-labelledby="emailModalTitle" aria-hidden="true">
		<div class="modal-dialog modal-dialog-centered" role="document">
			<div class="modal-content">
				<div class="modal-header">
					<h5 class="modal-title" id="emailModalLabel">
					  Edit Template <i class="fa fa-chevron-right"></i></h5>
					<button  type="button" class="close" data-dismiss="modal" aria-label="Close">
					  <span aria-hidden="true">&times;</span>
					</button>
				  </div>
			<div class="modal-body">
				<div class="col-md-12">
					<table class="table table-bordered table-hover CAPS">						
						<div id="GetEmailDetail">			  
						</div>						
					</table>
				</div>	
			</div>
					<div class="modal-footer">
					</div>
				</div>
			</div>
		</div>
		</form>
	<!-- End Edit Modal -->


</main>

	

<div id="mfoot"></div>
<%=strMessageIcon %> 

<!-- #Include file=../cc/CAPSFooter.asp -->

</body>
</html>

<%

Public Sub DisplayTableLeft()

'Writes out the left side table

Dim y
Dim lngEmailDetailID
Dim strEmailTemplateName
Dim strEmailSubject
Dim strEmailHeader
Dim strEmailBody
Dim strEmailFooter
Dim strEmailAttachment
Dim strEmailImportance
Dim strEmailSensitivity
Dim strFromAddress
Dim strStatus
Dim strSelected
Dim intRecordCount


strSQL = "SELECT * FROM tblCAPSEmailDetail WHERE BusinessArea = '" & Session("BusinessArea") & "'"
intRecordCount = 1

objRS.Open strSQL,objCon   
    	
	Do until objRS.EOF 

	strSelected = "false"

		If cint(objRS("EmailDetailID")) = cint(Session("EmailDetailID")) Then
			strSelected = "true"		
		End If

		If IsEmpty(Session("EmailDetailID")) and intRecordCount =  1 Then
				strSelected = "true"
		End If		
	
			lngEmailDetailID = objRS("EmailDetailID")
			strEmailTemplateName = objRS("EmailTemplateName")
			strEmailSubject = objRS("EmailSubject")
			strEmailHeader = objRS("EmailHeader")
			strEmailBody = objRS("EmailBody")
			strEmailFooter = objRS("EmailFooter")
			strEmailAttachment = objRS("EmailAttachment")
			strEmailImportance = objRS("EmailImportance")
			strEmailSensitivity = objRS("EmailSensitivity")
			strFromAddress = objRS("FromAddress")

		If strSelected = "true" Then			
			Response.Write "<a class=""nav-link active"" data-id=""" & intRecordCount & """ id=""v-pills-template-" & intRecordCount & "-tab"" data-toggle=""pill"" href=""#v-pills-template-" & intRecordCount & """ role=""tab"" aria-controls=""v-pills-template-1"" aria-selected=""" & strSelected & """ onClick=""loadDocE(this);"">" & strEmailTemplateName & "</a>"
		Else
			Response.Write "<a class=""nav-link"" id=""v-pills-template-" & intRecordCount & "-tab"" data-toggle=""pill"" href=""#v-pills-template-" & intRecordCount & """ role=""tab"" aria-controls=""v-pills-template-1"" aria-selected=""" & strSelected & """ >" & strEmailTemplateName & "</a>"
		End If

			intRecordCount = intRecordCount + 1

		objRS.movenext
	
	Loop	
	
objRS.Close

End Sub

Public Sub DisplayTableRight()

'Writes out the right side table

Dim y
Dim lngEmailDetailID
Dim strEmailTemplateName
Dim strEmailSubject
Dim strEmailHeader
Dim strEmailBody
Dim strEmailFooter
Dim strEmailAttachment
Dim strEmailImportance
Dim strEmailSensitivity
Dim strFromAddress
Dim strStatus
Dim strSelected
Dim intRecordCount
Dim strEmailType
Dim strSensePill

strSQL = "SELECT * FROM tblCAPSEmailDetail"
intRecordCount = 1



If IsEmpty(Session("EmailDetailID")) Then Session("EmailDetailID") = 0 End If

objRS.Open strSQL,objCon   
    	
	Do until objRS.EOF	
	
		If cint(objRS("EmailDetailID")) = cint(Session("EmailDetailID")) Then
			strSelected = "true"			
		Else
			If intRecordCount = 1 and Session("EmailDetailID") = 0 Then
				strSelected = "true"
			Else
				strSelected = "false"	
			End If		
		End If		 

		If Not IsEmpty(Request.Form("ModalSaveID")) Then
		
			If cint(Request.Form("ModalSaveID")) = cint(intRecordCount) Then
				strSelected = "true"
			End If

		End If
	
			lngEmailDetailID = objRS("EmailDetailID")
			strEmailTemplateName = objRS("EmailTemplateName")
			strEmailSubject = objRS("EmailSubject")
			strEmailHeader = objRS("EmailHeader")
			strEmailBody = objRS("EmailBody")
			strEmailFooter = objRS("EmailFooter")
			strEmailAttachment = objRS("EmailAttachment")
			strEmailImportance = objRS("EmailImportance")

			If objRS("EmailImportance") = "Low" Then
				strEmailImportance = "<span class=""badge badge-pill badge-success"">" & objRS("EmailImportance") & "</span>"
			End If

			If objRS("EmailImportance") = "Normal" Then
				strEmailImportance = "<span class=""badge badge-pill badge-danger"">" & objRS("EmailImportance") & "</span>"
			End If

			If objRS("EmailImportance") = "High" Then
				strEmailImportance = "<span class=""badge badge-pill badge-warning"">" & objRS("EmailImportance") & "</span>"
			End If		
			
			'Get the Email Sensitivity
			If IsNull(objRS("EmailSensitivity")) Then
				strEmailSensitivity = ""
			Else
				strEmailSensitivity = Left(objRS("EmailSensitivity"),14)
			End If
			
			'Get the Pill/butotn colour based on the Email Sensitivity
			Select CASE strEmailSensitivity
				CASE "[SEC=UNOFFICIA"
					strSensePill = "primary"
				CASE "[SEC=OFFICIAL]"
					strSensePill = "secondary"
				CASE "[SEC=OFFICIAL:"
					strSensePill = "warning"
				CASE "[SEC=PROTECTED"
					strSensePill = "danger"
				CASE ELSE
			
					strSensePill = "primary"
			END Select
			
			'If objRS("EmailSensitivity") = "Normal" Then
			'	strEmailSensitivity = "<span class=""badge badge-pill badge-primary"">" & objRS("EmailSensitivity") & "</span>"
			'End If

			'If objRS("EmailSensitivity") = "Personal" Then
			'	strEmailSensitivity = "<span class=""badge badge-pill badge-secondary"">" & objRS("EmailSensitivity") & "</span>"
			'End If

			'If objRS("EmailSensitivity") = "Private" Then
			'	strEmailSensitivity = "<span class=""badge badge-pill badge-info"">" & objRS("EmailSensitivity") & "</span>"
			'End If

			'If objRS("EmailSensitivity") = "Confidential" Then
			'	strEmailSensitivity = "<span class=""badge badge-pill badge-dark"">" & objRS("EmailSensitivity") & "</span>"
			'End If
			
			strEmailSensitivity = "<span class=""badge badge-pill badge-" & strSensePill & """>" & objRS("EmailSensitivity") & "</span>"
			
			strFromAddress = objRS("FromAddress")

	If strSelected = "true" Then
		Response.Write "<div class=""tab-pane fade show active"" id=""v-pills-template-" & intRecordCount & """ role=""tabpanel"" aria-labelledby=""v-pills-template-" & intRecordCount & "-tab"">"
	Else
		Response.Write "<div class=""tab-pane fade"" id=""v-pills-template-" & intRecordCount & """ role=""tabpanel"" aria-labelledby=""v-pills-template-" & intRecordCount & "-tab"">"
	End If
			
			Response.Write "<div class=""content-spacer"">"
				Response.Write "<div class=""row tab-content-header"">"
					  	Response.Write "<div class=""col-md-6 my-auto""><h4>" & strEmailTemplateName & "</h4></div>"
						Response.Write "<div class=""col-md-6 text-md-right""><button class=""btn btn-sm btn-outline-primary"" data-toggle=""modal"" data-target=""#emailModal"" data-id=""" & lngEmailDetailID & """ data-EmailType=""" & strEmailType & """ onClick=""loadDocE(this);""><i class=""fa fa-pen""></i> Edit template </button>"
						
						Response.Write "<button class=""btn btn-sm btn-outline-danger"" data-toggle=""modal"" data-target=""#ModalDelete"" data-id=""" & lngEmailDetailID & """ data-EmailTemplateName=""" & strEmailTemplatename & """ onClick=""loadDelete(this);""><i class=""fa fa-trash""></i> Delete template</button></div>"
					
						Response.Write "<div class=""col-12""><hr /></div>"
						
						  	Response.Write "<div class=""col-12 information-rows"">"
								
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-3""><b>Template&nbsp;Name</b></div><div class=""col-lg-10 col-md-9""><span >" & strEmailTemplateName & "</span></div></div>"
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-3""><b>Subject</b></div><div class=""col-lg-10 col-md-9""><span >" & strEmailSubject & "</span></div></div>"
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-3""><b>Importance</b></div><div class=""col-lg-10 col-md-9""><span >" & strEmailImportance & "</span></div></div>"
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-3""><b>Sensitivity</b></div><div class=""col-lg-10 col-md-9""><span >" & strEmailSensitivity & "</span></div></div>"			
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-3""><b>Header</b></div><div class=""col-lg-10 col-md-9""><span >" & strEmailHeader & "</span></div></div>"
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-3""><b>Body</b></div><div class=""col-lg-10 col-md-9""><span >" & strEmailBody & "</span></div></div>"
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-3""><b>Footer</b></div><div class=""col-lg-10 col-md-9""><span >" & strEmailFooter & "</span></div></div>"
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-3""><b>From&nbsp;Address</b></div><div class=""col-lg-10 col-md-9"">" & strFromAddress & "</div></div>"
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-3""><b>Attachment</b></div><div class=""col-lg-10 col-md-9"">" & strEmailAttachment & "</div></div>"

							Response.Write "</div>"					
				Response.Write "</div>"
			Response.Write "</div>"
		Response.Write "</div>"
			
		intRecordCount = intRecordCount + 1

		objRS.movenext
	
	Loop	
	
objRS.Close

End Sub

Public Sub WriteEditModals()

'Writes out the Edit Template Modals

Dim y
Dim lngEmailDetailID
Dim strEmailTemplateName
Dim strEmailSubject
Dim strEmailHeader
Dim strEmailBody
Dim strEmailFooter
Dim strEmailAttachment
Dim strEmailImportance
Dim strEmailSensitivity
Dim strFromAddress
Dim strStatus
Dim strSelected
Dim intRecordCount


strSQL = "SELECT * FROM tblCAPSEmailDetail"
intRecordCount = 1

objRS.Open strSQL,objCon   
    	
	Do until objRS.EOF	
	
			lngEmailDetailID = objRS("EmailDetailID")
			strEmailTemplateName = objRS("EmailTemplateName")
			strEmailSubject = "<span>" & objRS("EmailSubject") & "</span>"
			strEmailHeader = "<span>" & objRS("EmailHeader") & "</span>"
			strEmailBody = "<span>" & objRS("EmailBody") & "</span>"
			strEmailFooter = "<span>" & objRS("EmailFooter") & "</span>"
			strEmailAttachment = objRS("EmailAttachment")
			strEmailImportance = objRS("EmailImportance")
			strEmailSensitivity = objRS("EmailSensitivity")
			strFromAddress = objRS("FromAddress")
			
			Response.Write "<div class=""modal fade"" id=""emailModal" & intRecordCount & """ name=""emailModal" & intRecordCount & """ tabindex=""-1"" role=""dialog"" aria-labelledby=""emailModal"" aria-hidden=""true"">"
				Response.Write "<div class=""modal-dialog modal-dialog-centered modal-dialog-scrollable"">"
					Response.Write "<div class=""modal-content"">"
						Response.Write "<div class=""modal-header"">"
							Response.Write "<h5 class=""modal-title"" id=""emailModalLabel""> Edit Template <i class=""fa fa-chevron-right""></i>" & strEmailTemplateName & "</h5>"
							Response.Write "<button type=""button"" class=""close"" data-dismiss=""modal"" aria-label=""Close"">"
								Response.Write "<span aria-hidden=""true"">&times;</span>"
								Response.Write "</button>"
						Response.Write "</div>"
								Response.Write "<div class=""modal-body"">"
									Response.Write "<div class=""col-12 information-rows"">"																
																		
										Response.Write "<div class=""row"">"
											Response.Write "<div class=""col-lg-2 col-md-3""><b>Subject</b></div>"
											Response.Write "<div class=""col-lg-10 col-md-9"">"
												Response.Write "<input type=""text"" class=""form-control"" id=""EmailSubject" & intRecordCount & """ placeholder=""Subject"" aria-describedby=""EmailSubject""  value=""" & strEmailSubject & """ required>"
												Response.Write "<div class=""valid-feedback"">Valid</div>"
												Response.Write "<div class=""invalid-feedback"">Subject must be entered.</div>"
											Response.Write "</div>"
										Response.Write "</div>"
									Response.Write "</div>"									
																				
								Response.Write "</div>"

								Response.Write "<div class=""modal-footer"">"
									Response.Write "<button type=""button"" class=""btn btn-link"" data-dismiss=""modal"" >Cancel</button>"
									'Response.Write "<button type=""button"" class=""btn btn-primary"" onclick=""SaveEmailTemplate(" & intRecordCount & ");"">Save changes</button>"
									Response.Write "<button type=""submit"" class=""btn btn-primary""  onclick=""SaveEmailTemplate(" & intRecordCount & ");"">Save changes</button>"
								Response.Write "</div>"
											
					Response.Write "</div>"
				Response.Write "</div>"
			Response.Write "</div>"
		
		intRecordCount = intRecordCount + 1

		objRS.movenext
	
	Loop	
	
objRS.Close

End Sub

Sub SaveEmailTemplate()

	Dim intModalID
	Dim lngEmailDetailID
	Dim lngEmailErrrorID
	
	intModalID = Request.Form("ModalSaveID")

	'Get the Error Email ID or make it zero if empty 
	If IsNull(Request.Form("EmailErrorID" & intModalID & "")) Or Request.Form("EmailErrorID" & intModalID & "") = "" Then
		lngEmailErrrorID = 0
	Else
		lngEmailErrrorID = Request.Form("EmailErrorID" & intModalID & "")
	End If
	
  		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSEmailDetailSave"

			.Parameters.Append objCmd.CreateParameter("EmailDetailID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmailTemplateName", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("EmailSubject", adVarChar, adParamInput, 200)
			.Parameters.Append objCmd.CreateParameter("EmailHeader", adLongVarChar, adParamInput, -1)              
			.Parameters.Append objCmd.CreateParameter("EmailBody", adLongVarChar, adParamInput, -1)
			.Parameters.Append objCmd.CreateParameter("EmailFooter", adLongVarChar, adParamInput, -1)
			.Parameters.Append objCmd.CreateParameter("EmailAttachment", adVarChar, adParamInput, 500)
			.Parameters.Append objCmd.CreateParameter("EmailImportance", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("EmailSensitivity", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("FromAddress", adVarChar, adParamInput, 500) 
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("BusinessArea", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("EmailErrorID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmailDetailIDOutput", adInteger, adParamOutput)
			
			.Parameters("EmailDetailID") = Request.Form("EmailDetailID" & intModalID & "")
			.Parameters("EmailTemplateName") = Request.Form("EmailTemplateName" & intModalID & "")
			.Parameters("EmailSubject") = Request.Form("EmailSubject" & intModalID & "")					
			.Parameters("EmailHeader") = Request.Form("EmailHeader" & intModalID & "")
			.Parameters("EmailBody") = Request.Form("EmailBody" & intModalID & "")
			.Parameters("EmailFooter") = Request.Form("EmailFooter" & intModalID & "")
			.Parameters("EmailAttachment") = Request.Form("EmailAttachments" & intModalID & "")            
			.Parameters("EmailImportance") = Request.Form("EmailImportance" & intModalID & "") 
			.Parameters("EmailSensitivity") = Request.Form("EmailSensitivity" & intModalID & "")
			.Parameters("FromAddress") = Request.Form("FromAddress" & intModalID & "")
			.Parameters("UpdatedBy") = Session("UserID")
			.Parameters("BusinessArea") = Session("BusinessArea")
			.Parameters("EmailErrorID") = lngEmailErrrorID
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		lngEmailDetailID = objCmd.Parameters.Item("EmailDetailIDOutput")
		Session("EmailDetailID") = lngEmailDetailID		
	
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

		objCon.Execute "DELETE tblCAPSEmailDetail WHERE EmailDetailID= " & ID & ""

End Sub

Set objRS = Nothing
Set objCon = Nothing
%>
