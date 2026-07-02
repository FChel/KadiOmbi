
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


'Execute Action

If Request.QueryString("Action") = "Save" Then

	Call SaveEmailTemplate() 

End If

If Request.QueryString("Action") = "Delete" Then   
	
    Call DeleteData(Request.QueryString("EmailErrorMsgID"))
   
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

function loadDocE(cb) {

  var id = cb.getAttribute('data-id');
  var xhttp = new XMLHttpRequest();

  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("GetEmailDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetErrorMsg.asp?EmailErrorMsgID=" + id, true);
  xhttp.send();
}

function loadDelete(cb) {

var id = cb.getAttribute('data-id');
var name = cb.getAttribute('data-EmailTemplateName');

	document.getElementById("ModalDeleteMessage").innerHTML = 'Do you wish to delete the selected Email Template - ' + name + '?';
	document.getElementById("ModalDelete").style.display = "block";
	document.getElementById("EmailTemplateDeleteID").value = id;
	
}

function deleteEmailTemplate(cb) {

	var id = document.getElementById("EmailTemplateDeleteID").value
	self.location = "EmailError.asp?Action=Delete&EmailErrorMsgID=" + id;

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
			<div class="col-md-10"><h3 class="text-left">Email Error Messages Configuration </h3></div>
			<div class="col-md-2 text-right"></div>
				<div class="row">
					<div class="col-md-3"><h6 class="mb-3">Select Error Message</h6><div class="nav flex-column nav-pills" id="v-pills-tab" role="tablist" aria-orientation="vertical"><%Call DisplayTableLeft()%></div>
				</div>
				<div class="col-md-9">
						<div class="tab-content" id="v-pills-tabContent"><%Call DisplayTableRight()%></div>
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
	<form action="EmailError.asp?Action=Save" method="POST" id="frm" name="frm" class="needs-validation" novalidate>
		<div class="modal fade" id="emailModal" tabindex="-1" role="dialog" aria-labelledby="emailModalTitle" aria-hidden="true">
			<div class="modal-dialog modal-dialog-centered" role="document">
				<div class="modal-content">
					<div class="modal-header">
						<h5 class="modal-title" id="emailModalLabel">Edit Error Message <i class="fa fa-chevron-right"></i></h5><button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					</div>
					<div class="modal-body">
						<div class="col-md-12">
							<table class="table table-bordered table-hover CAPS">						
								<div id="GetEmailDetail"></div>						
							</table>
						</div>	
					</div>
					<div class="modal-footer"></div>
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
Dim lngEmailErrorMsgID
Dim intEmailErrorMsgNo
Dim strEmailErrorMsg
Dim strEmailErrorMsgFriendly
Dim strCardType
Dim strCardTypeSub
Dim strStatus
Dim strSelected
Dim intRecordCount

strSQL = "SELECT * FROM tblCAPSEmailErrorMsg"
intRecordCount = 1

objRS.Open strSQL,objCon   
    	
	Do until objRS.EOF 

	strSelected = "false"

		If cint(objRS("EmailErrorMsgID")) = cint(Session("EmailErrorMsgID")) Then
			strSelected = "true"		
		End If

		If IsEmpty(Session("EmailErrorMsgID")) and intRecordCount =  1 Then
				strSelected = "true"
		End If		
	
			lngEmailErrorMsgID = objRS("EmailErrorMsgID")
			intEmailErrorMsgNo = objRS("EmailErrorMsgNo")
			strEmailErrorMsg = objRS("EmailErrorMsg")
			strEmailErrorMsgFriendly = objRS("EmailErrorMsgFriendly")
			strCardType = objRS("CardType")
			strCardTypeSub = objRS("CardTypeSub")		

		If strSelected = "true" Then			
			Response.Write "<a class=""nav-link active"" data-id=""" & intRecordCount & """ id=""v-pills-template-" & intRecordCount & "-tab"" data-toggle=""pill"" href=""#v-pills-template-" & intRecordCount & """ role=""tab"" aria-controls=""v-pills-template-1"" aria-selected=""" & strSelected & """ onClick=""loadDocE(this);"">" & intEmailErrorMsgNo & " " & strCardType & " " & strCardTypeSub & "</a>"
		Else
			Response.Write "<a class=""nav-link"" id=""v-pills-template-" & intRecordCount & "-tab"" data-toggle=""pill"" href=""#v-pills-template-" & intRecordCount & """ role=""tab"" aria-controls=""v-pills-template-1"" aria-selected=""" & strSelected & """ >" & intEmailErrorMsgNo & " " & strCardType & " " & strCardTypeSub & "</a>"
		End If

			intRecordCount = intRecordCount + 1

		objRS.movenext
	
	Loop	
	
objRS.Close

End Sub

Public Sub DisplayTableRight()

'Writes out the right side table

Dim y
Dim lngEmailErrorMsgID
Dim intEmailErrorMsgNo
Dim strEmailErrorMsg
Dim strEmailErrorMsgFriendly
Dim strCardType
Dim strCardTypeSub
Dim strStatus
Dim strSelected
Dim intRecordCount
Dim strEmailType
Dim strEmailTemplatename

strSQL = "SELECT * FROM tblCAPSEmailErrorMsg"

intRecordCount = 1

If IsEmpty(Session("EmailErrorMsgID")) Then Session("EmailErrorMsgID") = 0 End If

objRS.Open strSQL,objCon   
    	
	Do until objRS.EOF	
	
		If cint(objRS("EmailErrorMsgID")) = cint(Session("EmailErrorMsgID")) Then
			strSelected = "true"			
		Else
			If intRecordCount = 1 and Session("EmailErrorMsgID") = 0 Then
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
	
				lngEmailErrorMsgID = objRS("EmailErrorMsgID")
				intEmailErrorMsgNo = objRS("EmailErrorMsgNo")
				strEmailErrorMsg = objRS("EmailErrorMsg")
				strEmailErrorMsgFriendly = objRS("EmailErrorMsgFriendly")
				strCardType = objRS("CardType")
				strCardTypeSub = objRS("CardTypeSub")

			If objRS("CardType") = "DTC" Then
				strCardType = "<span class=""badge badge-pill badge-success"">" & objRS("CardType") & "</span>"
			End If

			If objRS("CardType") = "DPC" Then
				strCardType = "<span class=""badge badge-pill badge-danger"">" & objRS("CardType") & "</span>"
			End If

			If objRS("CardTypeSub") = "Diners" Then
				strCardTypeSub = "<span class=""badge badge-pill badge-primary"">" & objRS("CardTypeSub") & "</span>"
			End If

			If objRS("CardTypeSub") = "CTS" Then
				strCardTypeSub = "<span class=""badge badge-pill badge-secondary"">" & objRS("CardTypeSub") & "</span>"
			End If

			If objRS("CardTypeSub") = "Mastercard" Then
				strCardTypeSub = "<span class=""badge badge-pill badge-info"">" & objRS("CardTypeSub") & "</span>"
			End If

			If objRS("CardTypeSub") = "ANZ" Then
				strCardTypeSub = "<span class=""badge badge-pill badge-dark"">" & objRS("CardTypeSub") & "</span>"
			End If			

If strSelected = "true" Then
	Response.Write "<div class=""tab-pane fade show active"" id=""v-pills-template-" & intRecordCount & """ role=""tabpanel"" aria-labelledby=""v-pills-template-" & intRecordCount & "-tab"">"
Else
	Response.Write "<div class=""tab-pane fade"" id=""v-pills-template-" & intRecordCount & """ role=""tabpanel"" aria-labelledby=""v-pills-template-" & intRecordCount & "-tab"">"
End If
			
		Response.Write "<div class=""content-spacer"">"
			Response.Write "<div class=""row tab-content-header"">"
				Response.Write "<div class=""col-md-6 my-auto""><h4>" & intEmailErrorMsgNo & "</h4></div>"
					Response.Write "<div class=""col-md-6 text-md-right""><button class=""btn btn-sm btn-outline-primary"" data-toggle=""modal"" data-target=""#emailModal"" data-id=""" & lngEmailErrorMsgID & """ data-EmailType=""" & strEmailType & """ onClick=""loadDocE(this);""><i class=""fa fa-pen""></i> Edit template </button><button class=""btn btn-sm btn-outline-danger"" data-toggle=""modal"" data-target=""#ModalDelete"" data-id=""" & lngEmailErrorMsgID & """ data-EmailTemplateName=""" & strEmailTemplatename & """ onClick=""loadDelete(this);""><i class=""fa fa-trash""></i> Delete template</button></div>"
					
						Response.Write "<div class=""col-12""><hr /></div>"
						
						  	Response.Write "<div class=""col-12 information-rows"">"
								
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-4""><b>Message&nbsp;No</b></div><div class=""col-lg-10 col-md-8"">" & lngEmailErrorMsgID & "</div></div>"
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-4""><b>Message</b></div><div class=""col-lg-10 col-md-8"">" & strEmailErrorMsg & "</div></div>"
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-4""><b>Message&nbsp;Friendly</b></div><div class=""col-lg-10 col-md-8"">" & strEmailErrorMsgFriendly & "</div></div>"
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-4""><b>Card&nbsp;Type</b></div><div class=""col-lg-10 col-md-8"">" & strCardType & "</div></div>"			
								Response.Write "<div class=""row""><div class=""col-lg-2 col-md-4""><b>Card&nbsp;Type&nbsp;Sub</b></div><div class=""col-lg-10 col-md-8"">" & strCardTypeSub & "</div></div>"
							
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
	Dim lngEmailErrorMsgID

	intModalID = Request.Form("ModalSaveID")

  		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSEmailErrorMsgSave"

			.Parameters.Append objCmd.CreateParameter("EmailErrorMsgID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmailErrorMsgNo", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmailErrorMsg", adLongVarChar, adParamInput, -1)
			.Parameters.Append objCmd.CreateParameter("EmailErrorMsgFriendly", adLongVarChar, adParamInput, -1)              
			.Parameters.Append objCmd.CreateParameter("CardType", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("CardTypeSub", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmailErrorMsgIDOutput", adInteger, adParamOutput)
			
			.Parameters("EmailErrorMsgID") = Request.Form("EmailErrorMsgID" & intModalID & "")
			.Parameters("EmailErrorMsgNo") = Request.Form("EmailErrorMsgNo" & intModalID & "")
			.Parameters("EmailErrorMsg") = Request.Form("EmailErrorMsg" & intModalID & "")					
			.Parameters("EmailErrorMsgFriendly") = Request.Form("EmailErrorMsgFriendly" & intModalID & "")
			.Parameters("CardType") = Request.Form("CardType" & intModalID & "")
			.Parameters("CardTypeSub") = Request.Form("CardTypeSub" & intModalID & "")
			.Parameters("UpdatedBy") = Session("UserID")
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		lngEmailErrorMsgID = objCmd.Parameters.Item("EmailErrorMsgIDOutput")
		Session("EmailErrorMsgID") = lngEmailErrorMsgID		
	
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

		objCon.Execute "DELETE tblCAPSEmailDetail WHERE EmailErrorMsgID= " & ID & ""

End Sub

Set objRS = Nothing
Set objCon = Nothing
%>
