
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=../ADOVBS.inc -->
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
Dim objCon2
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

Dim lngApplicationID
Dim strEmployeeID
Dim strTitle
Dim strFirstName
Dim strLastName
Dim strFormalFirstName
Dim strFormalLastName
Dim strFormalMiddleName
Dim strAddress1
Dim strAddress2
Dim strAddress3
Dim strAddress4
Dim strSuburb
Dim strState
Dim strPostCode
Dim strGroup
Dim dteDateReceived
Dim strStatus
Dim strReviewedBy
Dim dteDateReviewed
Dim lngCreditLimit
Dim strWorkPhone
Dim strMobilePhone
Dim strRank
Dim strCMSUserName
Dim strCMSMessage
Dim strCMSCheckStatus

Dim strCheckStatus
Dim strColour
Dim strFaFa
Dim strSubmitEnable
Dim strHelpMessage
Dim strAlert
Dim strValid
Dim strApplicantUserName

Dim strLastFourDigits
Dim strCardType
Dim strCardTypeSub
Dim strJustificationAE602
Dim strProChargeUserName
Dim strApplicationNameOnCard
Dim strLimitChangeTitle

Dim strTrainingMandatory
Dim strTrainingCourse

on error resume next
	'ProMaster Connection details
	Set objCon2 = Server.CreateObject("ADODB.Connection")
	Session("DBConnection2") = "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	objCon2.ConnectionTimeout=2
	objCon2.Open Session("DBConnection2")
	
	'response.write "state=" & objCon2.State
on error goto 0
    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objRS1 = Server.CreateObject("ADODB.Recordset")
    Set objRS2 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	

    Session("CurrentPage") = "CC/ApplicationsEmployee.asp"

	If IsNull(Session("ApplicationID")) OR Session("ApplicationID") = "" Then Session("ApplicationID")= 0

If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	Session("ApplicationID") = Request.QueryString("ApplicationID")
End If

If Not IsEmpty(Request.QueryString("ApplicationChecks")) Then
	Session("ApplicationChecks") = Request.QueryString("ApplicationChecks")
End If

'If the user (admin user) has selected another user's application then get the details in this screen based on the applicant, not the user logged in
If Not IsEmpty(Request.QueryString("ApplicationEmployeeID")) Then
	If Session("UserTypeID") > 9 Then
		Session("ApplicationEmployeeID") = Request.QueryString("ApplicationEmployeeID")
		Session("CMSUserApplication") = ""

	Else
		Session("ApplicationEmployeeID") = Session("EmployeeID")
	End If
Else
	'Session("ApplicationEmployeeID") = Session("EmployeeID")
End If

'Set the Application EmployeeID to the logged in user if no other use has been selected (user has come straight to this screen)
If IsNull(Session("ApplicationEmployeeID")) Or Session("ApplicationEmployeeID")= "" Then 
	Session("ApplicationEmployeeID") = Session("EmployeeID")
End If

If Not IsEmpty(Request.QueryString("AccountCMS")) Then
	Session("CMSUserApplication") = Request.QueryString("AccountCMS")
End If

'Check to see if the Name Change has been selected
If Not IsEmpty(Request.QueryString("NameChange")) Then
	Session("NameChange") = Request.QueryString("NameChange")
End If

'If the Save Name button has been clicked then call the procedure to save name details
If Not IsEmpty(Request.QueryString("NameChangeSave")) Then
	Call NameOnCardSave(Request.QueryString("NameChangeSave"))
End If

If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "Cancel" Then
		Call CancelApplication()
	End If
	
	If Request.QueryString("Action") = "Release" Then
		'Call ReleaseApplication()
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

	'Make sure there is an Application Check Type selected, otherwise make this the Overview, which will display the checklist on the right
	If IsEmpty(Session("ApplicationChecks")) OR IsNull(Session("ApplicationChecks")) Then 
		Session("ApplicationChecks") = "Overview"
	Else
		'If the user has been to the Card Limit Application screen, they may have the Cards ApplicationCheck so make this overview, otherwise it will display nothing
		If Session("ApplicationChecks") = "Cards" Then Session("ApplicationChecks") = "Overview"
	End If
	
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "SubmitApp" Then
			Call SubmitApplication()
		End If
	End If
	
	If isNull(Session("ApplicationID")) Or Session("ApplicationID") = "" Then 
		Session("ApplicationID") = 0
	End If
	
  Call LoadDetails()
  
  If IsNull(Session("CardType")) OR Session("CardType") = "" Then Session("CardType") = "DTC - Diners"
%>



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
    { self.location = 'HomeCC2.asp'; }
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

	
	$(function(){
    $('#login').popover({
       
        placement: 'bottom',
        title: 'Popover Form',
        html:true,
        content:  $('#myForm').html()
    }).on('click', function(){
      // had to put it within the on click action so it grabs the correct info on submit
      $('.btn-primary').click(function(){
       $('#result').after("form submitted by " + $('#email').val())
        $.post('Stalls.asp',  {
            email: $('#email').val(),
            name: $('#name').val(),
            gender: $('#gender').val()
        }, function(r){
          $('#pops').popover('hide')
          $('#result').html('resonse from server could be here' )
        })
      })
  })
})

$(function(){
    $('.pops').popover({
       
        placement: 'bottom',
        title: 'Enter Details and Click Save',
        html:true,
        content:  $('#myForm').html()
    }).on('click', function(){
      // had to put it within the on click action so it grabs the correct info on submit
      $('.btn-primary').click(function(){
       $('#result').after("form submitted by " + $('#email').val())
        $.post('Stalls.asp?Action=Save',  {
            email: $('#email').val(),
            name: $('#name').val(),
            phone: $('#phone').val()
        }, function(r){
          $('#pops').popover('hide')
          $('#result').html('resonse from server could be here' )
        })
      })
  })
})

$(document).ready(function() {
    $("a").click(function(event) {
        //alert(event.target.id);
		//myForm.popForm.ID='dfdf'//event.target.id
		//document.getElementById('PopID').value = 0
		//$('#PopID').val=90
		//$('#PopID').val( "hello world" );
		$('input[name="PopID"]').val(event.target.id);
		$('input[name="FieldName"]').val(event.target.name);
    });
	
	$('input').on('keyup', function() {
		add1.innerHTML=this.value.length + ' chars';
		//alert(this.value.length);
	});
});

function textLength(value){
	var maxLength=39;
	if(value.length > maxlength) alert(value.length); return false;
	return true;
	
	
}

function valCard() {
    //d = document.getElementByName("CardTypeSelect").value;
	//alert(frm.CardTypeSelect.value);
    //frm.CardType.value=frm.CardTypeSelect.value;
}

function OpenSs(cb) {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("CardTypeSelect");
	var result = e.options[e.selectedIndex].value;
	
	document.getElementById('CardType').value=result;
	
}

$('#CardTypeSelect').change(function(){
    alert($(this).val());
})
</script>
<script js>
$(function(){
    $('#myFormSubmit').click(function(e){
      e.preventDefault();
      $('#formResults').text($('#frm').serialize());
      
      $.post('ApplicationsEmployee.asp?Action=SubmitApp', 
         $('#frm').serialize(), 
      /*   function(data, status, xhr){
           // do something here with response;
         });
      */
    });
});

</script>
<script>

$(document).ready(function () {
    //Initialize tooltips
    $('.nav-tabs > li a[title]').tooltip();
    
    //Wizard
    $('a[data-toggle="tab"]').on('show.bs.tab', function (e) {

        var $target = $(e.target);
    
        if ($target.parent().hasClass('disabled')) {
            return false;
        }
    });

    $(".next-step").click(function (e) {

        var $active = $('.wizard .nav-tabs li.active');
        $active.next().removeClass('disabled');
        nextTab($active);

    });
    $(".prev-step").click(function (e) {

        var $active = $('.wizard .nav-tabs li.active');
        prevTab($active);

    });
});

function nextTab(elem) {
    $(elem).next().find('a[data-toggle="tab"]').click();
}
function prevTab(elem) {
    $(elem).prev().find('a[data-toggle="tab"]').click();
}

function loadDocE() {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
  //alert(this.responseText);
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("EmployeeIDST").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetCMSAccount.asp?EmpID=" + frm.EmpIDS.value + "&AccountCMS=" + frm.AccountCMS.value + "", true);
  xhttp.send();
}

function SelectEmp(varEmpID) {

	varEmpID.toString();
	
	if(varEmpID==undefined) {	
	}
	{
	self.location = "ApplicationsSubmit.asp?Action=Search&AccountCMS=" + varEmpID;
	//self.location = "ApplicationsSubmit.asp?Action=Search&EIDNo=" + varEmpID + "&AccountCMS=" + document.getElementById('AccountCMS').text;
	}
}

function ChangeNameOnCard() {
//Reloads the page after the user clicks Save Name Change
	self.location = "ApplicationsSubmit.asp?NameChangeSave=" + document.getElementById('NameOnCardTitle2').value + ' ' + document.getElementById('CMSUserName2').value;
	
}
</script>



<body >
<main class="main py-0">
      <div class="container">
<form action="ApplicationsSubmit.asp?Action=SubmitApp" method="POST" id="frm" name="frm" class="inline">
<!-- Modal -->
<div class="modal fade" id="exampleModalCenter" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">Credit Application Declaration Form</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>I <span style="font-weight:bold;color:red;"><%=Session("ApplicationEmployeeID") & " - " & strApplicantUserName%></span> DECLARE that all the application details are correct.</label><br>
            <label for="CardType">Card:</label>
            <input type="text" name="CardType" id="CardType" class="form-control input-md" value="<%=Session("CardType")%>">
			<label for="name">Name On Card:</label>
            <input type="NameONCardA" name="NameONCardA" id="NameONCardA" class="form-control input-md" value="<%=strApplicantUserName%>">
			<label for="name">Employee ID:</label>
            <input type="EmployeeIDA" name="EmployeeIDA" id="EmployeeIDA" class="form-control input-md" value="<%=Session("ApplicationEmployeeID")%>">
			<label for="CreditLimit">Credit Limit:</label>
            <input type="text" name="CreditLimit" id="CreditLimit" class="form-control input-md" value="$30,000">
			<label for="GroupA">Group:</label>
			<input type="text" name="GroupA" id="GroupA" class="form-control input-md" value="<%=strGroup%>">
			<label for="IP">IP:</label>
            <input type="text" name="IP" id="IP" class="form-control input-md" disabled value="<%=Request.ServerVariables("remote_addr")%>">
			<!--<label for="DNS">DNS:</label>
            <input type="text" name="DNS" id="DNS" class="form-control input-md" disabled value="<%=Request.ServerVariables("remote_host")%>">-->
			<label for="LocalAddress">Local Address:</label>
            <input type="text" name="LocalAddress" id="LocalAddress" class="form-control input-md" disabled value="<%=Request.ServerVariables("LOCAL_ADDR")%>">
			<label for="UserLoggedIn">User Logged On:</label>
            <input type="text" name="UserLoggedIn" id="UserLoggedIn" class="form-control input-md" disabled value="<%=Request.ServerVariables("AUTH_USER")%>">
			<label for="Signature">Signature:</label>
            <input type="text" name="Signature" id="Signature" class="form-control input-md" Style="font-family:Bradley Hand, cursive; font-size:28px;" value="<%=strApplicantUserName%>">
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
        <button type="submit" class="btn btn-primary" id="myFormSubmit" name="myFormSubmit" onClick="frm.Submit()"><i class="fa fa-check-circle" ></i> Submit Application to GCFO</button>
      </div>
    </div>
  </div>
</div>

<!--ProMaster User Account Search Modal START -->
	<div class="modal fade" id="EmployeeModal" tabindex="-1" role="dialog" aria-labelledby="exampleModalScrollableTitle" aria-hidden="true">
	  <div class="modal-dialog modal-dialog-scrollable" role="document">
		<div class="modal-content">
		  <div class="modal-header">
			<h5 class="modal-title" id="exampleModalScrollableTitle">ProMaster/CMS Account Search</h5>
			<button type="button" class="close" data-dismiss="modal" aria-label="Close">
			  <i class="bx bx-x"></i>
			</button>
		  </div>
		  <div class="modal-body">
			 <!-- -->
			 <label>Search for Employee</label><br>
				<label for="AccountCMS">CMS Account Name:</label>
				<input type="text" name="AccountCMS" id="AccountCMS" class="form-control input-md" onKeyUp="loadDocE();">
				<label for="email">Employee ID:</label>
				<input type="text" name="EmpIDS" id="EmpIDS" class="form-control input-md" onKeyUp="loadDocE();">

			  <div id="EmployeeIDST">
			  
			  </div>
		  </div>
		  <div class="modal-footer">
			<button type="button" class="btn btn-light-secondary" data-dismiss="modal">
			  <i class="bx bx-x d-block d-sm-none"></i>
			  <span class="d-none d-sm-block">Close</span>
			</button>
		   
		  </div>
		</div>
	</div>
</div>
<!--ProMaster User Account Search Modal END-->
<div class="modal fade bd-example-modal-lg modalWait" id="ModalWait" data-backdrop="static" data-keyboard="false" tabindex="-1">
    <div class="modal-dialog modal-sm">
        <div class="modal-content" style="width: 48px">
            <span class="fa fa-spinner fa-spin fa-5x"></span>
        </div>
    </div>
</div>


        <div class="row">
          
		  
		   <div class="col-md-12">
          
            <div class="my-information">
              <ul class="nav nav-tabs" id="myFiTab" role="tablist">
                <li class="nav-item" role="presentation">
                  <a class="nav-link active" id="my-cards-tab" data-toggle="tab" href="#my-cards" role="tab" aria-controls="my-cards" aria-selected="true">Postal Address Check Details</a> for <%=strApplicantUserName%>
                </li>
              </ul>
			  
			  
              
			  <div class="tab-content" id="myFiTabContent">
                <div class="tab-pane fade show active" id="my-cards" role="tabpanel" aria-labelledby="my-cards-tab">
			  <%
					'Select Case Session("ApplicationChecks") 
					
						'Case "Overview" 
						'	Call LoadMenu(2)
							'Response.write "</ul></div>"
						'Case "Entitlement" 
						'	Call LoadEntitlement
						'Case "Contact" 
							Call LoadContact
						'Case "CMS" 
						'	Call LoadCMS
						'Case "Name" 
						'	Call LoadName
						'Case "Training" 
						'	Call LoadTraining
					'End Select
					%>
					</div>
					</div>
                </div>
				<div class="py-3"> 
					<div class="col-md-12 text-right my-auto">								
						<button type="button" class="btn btn-primary btn-block" data-toggle="modal" data-target="#exampleModalCenter" onClick="OpenSs(this);" <%=strSubmitEnable%>><i class="fa fa-check-circle"></i>
						  Submit Application
						</button>
					</div>
				</div>
			</div>
        </div>

</form>

<!--</DIV>-->


</form>
</main>

    <!-- jQuery -->
    <script src="../js/jquery.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="../js/bootstrap.min.js"></script>

<script>
function modal(){
       $('.modalWait').modal('show');
       setTimeout(function () {
       	console.log('hejsan');
       	$('.modalWait').modal('hide');
       }, 3000);
    }
</script>
	
<!-- #Include file=CAPSFooter.asp -->
</body>
</html>
<%



Sub LoadDetails()

   'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM tblCAPSCDMCHistory WITH(NOLOCK) WHERE EmployeeID = '" & Session("ApplicationEmployeeID") & "'",objCon

		If Not objRS.EOF Then
		   
			'lngApplicationID = objRS("ApplicationID")
			strEmployeeID = objRS("EmployeeID")
			strTitle = objRS("Title")
			strFirstName = objRS("FirstName")
			strLastName  = objRS("Surname")
			strAddress1 = objRS("PostalAddress_Unit")
			strAddress2 = objRS("PostalAddress_ClientLocation")
			strAddress3 = objRS("PostalAddress_DeliveryLocation")
			'strAddress4 = objRS("Address4")
			strSuburb = objRS("Postaladdress_City")
			strState = objRS("Postaladdress_State")
			strPostCode = objRS("Postaladdress_PostCode")
			
			strWorkPhone = objRS("TelephoneNumber")
			strMobilePhone = objRS("MobileNumber")
			
			strRank = objRS("EmployeeType")
			'dteDateReceived = objRS("DateReceived")
			'strCheckStatus = objRS("Status")
			'strReviewedBy = objRS("ReviewedBy")
			'dteDateReviewed = objRS("DateReviewed")
			'If IsNull(objRS("CreditLimit")) or objRS("CreditLimit") = "" then
				lngCreditLimit = 30000
			'Else
			'	lngCreditLimit = objRS("CreditLimit") 
			'End If
			strGroup = objRS("GroupName")
			strApplicantUserName = objRS("FirstName") & " " & objRS("Surname")
			
			strFormalFirstName = objRS("FormalFirstName")
			strFormalLastName = objRS("FormalLastName")
			strFormalMiddleName = objRS("FormalMiddleName")
		Else
			Session("ApplicationID") = 0
			lngApplicationID = 0'objRS("ApplicationID")
			strEmployeeID = ""
			strTitle = ""
			strFirstName = ""
			strLastName  = ""
			strAddress1 = ""
			strAddress2 = ""
			strAddress3 = ""
			strAddress4 = ""
			strSuburb = ""
			strState = ""
			strPostCode = ""
			dteDateReceived = ""
			strStatus = ""
			strReviewedBy = ""
			dteDateReviewed = ""
			lngCreditLimit = 30000
			strWorkPhone = ""
			strMobilePhone = ""
			strRank = ""
			strGroup = ""
			strApplicantUserName = "**Employee not in HR Data**"
			strFormalFirstName = ""
			strFormalLastName = ""
			strFormalMiddleName = ""
	   End If

	objRS.Close
	
	'Then Load the Application details if there is an existing application (AE602 loaded or a started application)
	 'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM tblCAPSApplication WITH(NOLOCK) WHERE ApplicationID = " & Session("ApplicationID") & "",objCon

		If Not objRS.EOF Then
			strLastFourDigits = objRS("LastFourDigits")
			Session("CardType") = objRS("CardType") & " - " & objRS("CardTypeSub")
			strCardType = objRS("CardType")
			strCardTypeSub = objRS("CardTypeSub")
			strJustificationAE602 = objRS("Justification")
			strProChargeUserName = objRS("ProChargeUserName")
			
			If IsNull(objRS("NameOnCard")) Then
				strApplicationNameOnCard = objRS("NameOnCard")
			Else
				strApplicationNameOnCard = objRS("NameOnCard")
			End If
			
			strLimitChangeTitle = objRS("Title")
		Else
			strLastFourDigits = ""
			strCardType = ""
			strCardTypeSub = ""
			strJustificationAE602 = ""
			strProChargeUserName = ""
			strApplicationNameOnCard = ""
			strLimitChangeTitle = ""
		End If
	
	objRS.Close
	
End Sub


Public Sub LoadHeader
'Procedure to load the header which contains the application type
Dim strSQL
Dim strCardTypeAll
Dim strHeader
Dim strApplicationType

	strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE ApplicationID = '" & Session("ApplicationID") & "'"

	objRS.Open strSQL,objCon

		If isNull(objRS("CardTypeSub")) Then
			strHeader = ""
			strCardTypeAll = ""
		Else
			strHeader = objRS("CardTypeSub")
			strCardTypeAll  = objRS("CardType") & " " & objRS("CardTypeSub")
		End If
		
		'Get the Application Type Name
		If IsNull(objRS("ApplicationTypeName")) Then
			strApplicationType = ""
		Else
			strApplicationType = objRS("ApplicationTypeName")
		End If
		
		'Determine the image and title based on the card type
		If Trim(strHeader) = "Diners" Then
			strHeader = "<img src=""../images/icon_diners.png"" title=""" & strCardTypeAll & """> " & strCardTypeAll & " " & strApplicationType
		ElseIf strHeader = "ANZ" Then
			strHeader = "<img src=""../images/logo_ANZ.png"" Title=""" & strCardTypeAll & """> " & strCardTypeAll & " " & strApplicationType
		ElseIf strHeader = "Mastercard" Then
			strHeader = "<img src=""../images/logo_mc.png"" Title=""" & strCardTypeAll & """> " & strCardTypeAll & " " & strApplicationType
		Else
			strHeader = "<img src=""../images/logo_coa.png"" Title=""" & strCardTypeAll & """> " & strCardTypeAll & " " & strApplicationType
		End If

	objRS.Close
	
	'Write the results of the detail above to the page
	Response.Write "<h4>" & strHeader & "</h4>"
	
End Sub


Public Sub LoadMenu(intType)

Dim strIco
Dim strEntitlement
Dim strContact
Dim strCMS
Dim strName
Dim strTraining
Dim bolCompleteCheck
Dim strCol
Dim strEmployeeType
Dim strActive
Dim bolEntitleCheck

'Possible ICOs
'ico fa fa-check ico-green
'ico fa fa-times ico-red
'ico fa fa-lock ico-muted

'Possible Classes
'class="current"
'class="complete prev-step"
'class="active"
'class="warning"
'class="locked"

'The boolean variable to note whether all checks are successful for the Overview Summary
bolCompleteCheck = True
bolEntitleCheck = True

	'1. Start the Address check, although this appears after entitlement in the list (as entitlement requires CDMC Rank/Level as well)
	'Check the Employee doesn't already have any contact detail issues
	objRS.Open "SELECT TOP 1 [OutDinersAddress1],[OutDinersAddress2],[IsValidPostal],[EmployeeType] FROM tblCAPSCDMCHistory WITH(NoLock) WHERE EmployeeID = '" & Session("ApplicationEmployeeID") & "' AND [Deleted] = 'N' ORDER BY [DateUpdated] DESC",objCon

		If objRS.EOF Then
			strMessage = Session("ApplicationEmployeeID") & " is not on the Corporate Directory/in PMKeYS"
			strCheckStatus = "warning"
			strIco = "fa fa-fw fa-times color-red"
			
			strEmployeeType = ""
			
			bolCompleteCheck = False
		Else
			If IsNull(objRS("IsValidPostal")) Then
				strMessage = "Postal Address is not Valid. Click to view"
				strCheckStatus = "warning"
				strIco = "fa fa-fw fa-times color-red"
				
				bolCompleteCheck = False
			Else
				If objRS("IsValidPostal") = "Y" Then
					strMessage = "Postal Address is good"
					strCheckStatus = "complete"
					strIco = "fa fa-fw fa-check color-green"
				Else
					strMessage = "Postal Address is NOT Valid. Click to view"
					strCheckStatus = "warning"
					strIco = "fa fa-fw fa-times color-red"
					
					bolCompleteCheck = False
				End If
			End If
			
			'Get the Employee Type for use in the Entitlement section
			If IsNull(objRS("EmployeeType")) or objRS("EmployeeType") ="" then
				strEmployeeType = ""
			Else
				strEmployeeType = objRS("EmployeeType")
			End If
		End If
		
	objRS.Close
	
	'If Contact is currently selected then add this to the State
	'If Session("ApplicationChecks") = "Contact" Then strCheckStatus = strCheckStatus & " current" 
	If Session("ApplicationChecks") = "Contact" Then 
		strActive = "active" 'strCheckStatus = strCheckStatus & " current" 
	Else
		strActive = ""
	End If
	
	'strContact = "<li class=""" & strCheckStatus & """>" & _
    '               "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Contact"">Contact Details <i class=""" & strIco & """></i>" & _
    '               "<span class=""desc"">" & strMessage & "</span></a></li>"
	
	'strContact = "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Contact"" class=""block-link""><i class=""" & strIco & """></i> Contact Details</a><span class=""content"" style=""font-weight:normal; color:gray; font-size:12px;"">" & strMessage & "</span>"
	
	strContact = "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Contact"" class=""section-link " & strActive & """><div class=""status""><i class=""" & strIco & """></i></div>" & _
                "<div class=""content""><span class=""title"">Contact Details</span><span class=""description"">" & strMessage & "</span></div></a>"
				
	'2. Start the Check for an existing card (Entitlement check part 1 of 2) -- although it appears first in the menu list
	'Open a recordset to check the Employee doesn't already have an active card
	objRS.Open "SELECT [CardNumber],[Expiry],[NameOnCard] FROM tblCAPSCard WITH(NoLock) WHERE EmployeeID = '" & Session("ApplicationEmployeeID") & "' AND [Status] = '00'",objCon

		If Not objRS.EOF Then
			strMessage = Session("ApplicationEmployeeID") & " already has an active card"
			strCheckStatus = "warning"
			strIco = "fa fa-fw fa-times color-red"
			
			bolCompleteCheck = False
			bolEntitleCheck = False
		Else
			strMessage = "No existing card and employee type able to hold a card"
			strCheckStatus = "complete"
			strIco = "fa fa-fw fa-check color-green"
		End If
		
	objRS.Close
	
	'2.2 Check the Employee Type is able to hold a Card
	strEmployeeType = GetEntitlement("DTC",strEmployeeType)
	
	If strEmployeeType = "Y" Then
		'Check to see if the Existing Card is an error (previous check) to note an error overall
		If bolEntitleCheck = False Then
			'strMessage = strEmployeeType
			strCheckStatus = "warning"
			strIco = "fa fa-fw fa-times color-red"
		Else
			strMessage = "Entitled"
			strCheckStatus = "complete"
			strIco = "fa fa-fw fa-check color-green"
		End If
	Else
		strMessage = strEmployeeType
		strCheckStatus = "warning"
		strIco = "fa fa-fw fa-times color-red"
		
		bolCompleteCheck = False
	End If
	
	'If Entitlement is currently selected then add this to the State - strCheckStatus (displays a different background - as selected - in the left hand side list)
	'If Session("ApplicationChecks") = "Entitlement" Then strCheckStatus = strCheckStatus & " current" 
	If Session("ApplicationChecks") = "Entitlement" Then 
		strActive = "active" 'strCheckStatus = strCheckStatus & " current" 
	Else
		strActive = ""
	End If
	
	'strEntitlement = "<li class=""" & strCheckStatus & """><a href=""ApplicationsSubmit.asp?ApplicationChecks=Entitlement"" onClick=""modal();"">Entitlement <i class=""" & strIco & """></i><span class=""desc"">" & strMessage & "</span></a></li>"
	
	'strEntitlement = "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Entitlement"" class=""block-link""><i class=""" & strIco & """></i> Entitlement </a><span class=""content"" style=""font-weight:normal; color:gray; font-size:12px;"">" & strMessage & "</span>"
	
	strEntitlement = "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Entitlement"" class=""section-link " & strActive & """><div class=""status""><i class=""" & strIco & """></i></div>" & _
                "<div class=""content""><span class=""title"">Entitlement</span><span class=""description"">" & strMessage & "</span></div></a>"

'If a different UserID has been selected then do not search for the current users ProMaster Account, just use the one selected
If IsNull(Session("CMSUserApplication")) OR Session("CMSUserApplication")= "" Then

	'If there is no connection to ProMaster (Card Management System) then do not try to use the connection
	If objCon2.State = 1 Then
		'Open a recordset in the ProMaster (CMS) database to check the Employee has a CMS Account
		objRS.Open "SELECT [user_name],[employee_id] FROM procharge_user WITH(NoLock) WHERE employee_id = '" & Session("ApplicationEmployeeID") & "' AND [active_indicator ] = 'Y'",objCon2

			If objRS.EOF Then
				strMessage = Session("ApplicationEmployeeID") & " has no active CMS Account"
				strCheckStatus = "warning"
				strIco = "fa fa-fw fa-times color-red"
				strCMSUserName = ""
				
				bolCompleteCheck = False
			Else
				strMessage = "CMS Account for " & Session("ApplicationEmployeeID") & ": " & objRS("user_name")
				strCheckStatus = "complete"
				strIco = "fa fa-fw fa-check color-green"
				strCMSUserName = objRS("user_name")
			End If
			
		objRS.Close
	Else
		strMessage = "CMS database currently unavailable, please try again in 1 hour"
		strCheckStatus = "warning"
		strIco = "fa fa-fw fa-times color-red"
		strCMSUserName = ""
		
		bolCompleteCheck = False
	End If	
	
Else
	strMessage = "CMS Account for " & Session("ApplicationEmployeeID") & ": " & Session("CMSUserApplication")
	strCheckStatus = "complete"
	strIco = "fa fa-fw fa-check color-green"
	strCMSUserName = Session("CMSUserApplication")
End If

	'Set the CMS USer Message variable for use in the right-hand-side detailed message (below)
	strCMSMessage = strMessage
	strCMSCheckStatus = strCheckStatus
	
	'If CMS is currently selected then add this to the State
	If Session("ApplicationChecks") = "CMS" Then 
		strActive = "active" 'strCheckStatus = strCheckStatus & " current" 
	Else
		strActive = ""
	End If
	
	'strCMS =  "<li class=""" & strCheckStatus & """><a href=""ApplicationsSubmit.asp?ApplicationChecks=CMS"">CMS User Account <i class=""" & strIco & """></i><span class=""desc"">" & strMessage  & "</span></a></li>" 
	'strCMS = "<a href=""ApplicationsSubmit.asp?ApplicationChecks=CMS"" class=""block-link""><i class=""" & strIco & """></i> CMS User Account</a><span class=""content"" style=""font-weight:normal; color:gray; font-size:12px;"">" & strMessage & "</span>"
	
	strCMS = "<a href=""ApplicationsSubmit.asp?ApplicationChecks=CMS"" class=""section-link " & strActive & """><div class=""status""><i class=""" & strIco & """></i></div>" & _
                "<div class=""content""><span class=""title"">CMS User Account</span><span class=""description"">" & strMessage & "</span></div></a>"
				
	'Check the length of the Name on Card
	If strEmployeeID = "" Or IsNull(strEmployeeID) Then
		strMessage = Session("ApplicationEmployeeID") & " is not in the Corporate Directory! " & GetSystemAdmin("SystemAdmin")
			strCheckStatus = "warning"
			strIco = "fa fa-fw fa-times color-red"
			
			bolCompleteCheck = False
	Else
	
		If Not isNull(strTitle & " " & strFirstName & " " & strLastName) Then
			If Len(strTitle & " " & strFirstName & " " & strLastName) >21 Then
				strMessage = Session("ApplicationEmployeeID") & " name is too long"
				strCheckStatus = "warning"
				strIco = "fa fa-fw fa-times color-red"
						
				If Len(strFirstName & " " & strLastName) >21 Then
					strMessage = Session("ApplicationEmployeeID") & " name is too long"
					strCheckStatus = "warning"
					strIco = "fa fa-fw fa-times color-red"
					
					bolCompleteCheck = False
				Else
					strMessage = "The Name on Card for " & Session("ApplicationEmployeeID") & " is good!"
					strCheckStatus = "complete"
					strIco = "fa fa-fw fa-check color-green"
				
				End If
				
				
			Else
				strMessage = "The Name on Card for " & Session("ApplicationEmployeeID") & " is good!"
				strCheckStatus = "complete"
				strIco = "fa fa-fw fa-check color-green"
			End If
		Else
			strMessage = Session("ApplicationEmployeeID") & " has no name! " & GetSystemAdmin("SystemAdmin")
			strCheckStatus = "warning"
			strIco = "fa fa-fw fa-times color-red"
			
			bolCompleteCheck = False
		End If
	End If
	
	'If Name is currently selected then add this to the State
	'If Session("ApplicationChecks") = "Name" Then strCheckStatus = strCheckStatus & " current" 
	If Session("ApplicationChecks") = "Name" Then 
		strActive = "active" 'strCheckStatus = strCheckStatus & " current" 
	Else
		strActive = ""
	End If
	
	'strName = "<li class=""" & strCheckStatus & """><a href=""ApplicationsSubmit.asp?ApplicationChecks=Name"">Name <i class=""" & strIco & """></i><span class=""desc"">" & strMessage  & "</span></a></li>"
	'strName = "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Name"" class=""block-link""><i class=""" & strIco & """></i> Name</a><span class=""content"" style=""font-weight:normal; color:gray; font-size:12px;"">" & strMessage & "</span>"
	
	strName = "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Name"" class=""section-link " & strActive & """><div class=""status""><i class=""" & strIco & """></i></div>" & _
                "<div class=""content""><span class=""title"">Name</span><span class=""description"">" & strMessage & "</span></div></a>"
				
	'If Training is currently selected then add this to the State
	'If Session("ApplicationChecks") = "Training" Then strCheckStatus = strCheckStatus & " current" 
	If Session("ApplicationChecks") = "Training" Then 
		strActive = "active" 'strCheckStatus = strCheckStatus & " current" 
	Else
		strActive = ""
	End If
	
	''''**Add the Training Check HERE****-----
	'2. Start the Check for an existing card (Entitlement check part 1 of 2) -- although it appears first in the menu list
	
	'Get the Course Number Based on the Card type
	If strCardType = "DTC" Then
		strTrainingMandatory = GetSystemAdmin("TrainingMandatoryDTC")
	ElseIf strCardType = "DPC" Then
		strTrainingMandatory = GetSystemAdmin("TrainingMandatoryDPC")
	Else
		strTrainingMandatory = "N"
	End If
	
	'If the Training is NOT Mandatory then skip the check
	If strTrainingMandatory = "Y" Then
	
		'Get the Training Curse Number to check for
		If IsNull(strCardType) Then strCardType = ""
		
		strTrainingCourse = GetSystemAdmin(strCardType & "TrainingCourse")
		
		'Get the System setting for whether the course is mandatory based on the Card Type
		strHelpMessage = GetSystemAdmin("TrainingMessage")
		
		'Open a recordset to check the Employee has completed the Training
		objRS.Open "SELECT TOP 1 [CourseID],[CourseTitle],[EmployeeID],[CompletionDate] FROM tblCAPSTraining WITH(NoLock) WHERE EmployeeID = '" & Session("ApplicationEmployeeID") & "' AND CourseID = '" & strTrainingCourse & "'",objCon

			If objRS.EOF Then
				strMessage = Session("ApplicationEmployeeID") & " has not completed Mandatory Training for " & strCardType
				strCheckStatus = "warning"
				strIco = "fa fa-fw fa-times color-red"
				
				bolCompleteCheck = False
				bolEntitleCheck = False
			Else
				strMessage = "Training Course " & objRS("CourseID") & " completed!"
				strCheckStatus = "complete"
				strIco = "fa fa-fw fa-check color-green"
			End If
			
		objRS.Close
	Else
		strMessage = "Training not Mandatory for " & strCardType
		strCheckStatus = "complete"
		strIco = "fa fa-fw fa-check color-green"
	End If
	
	'strMessage = Session("ApplicationEmployeeID") & " has completed training! "
	'strIco = "fa fa-fw fa-check color-green"
	
	'strTraining = "<li class=""" & strCheckStatus & """><a href=""ApplicationsSubmit.asp?ApplicationChecks=Training"">Training <i class=""" & strIco & """></i><span class=""desc"">" & strMessage  & "</span></a></li>"
	'strTraining = "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Training"" class=""block-link""><i class=""" & strIco & """></i> Training</a><span class=""content"" style=""font-weight:normal; color:gray; font-size:12px;"">" & strMessage & "</span>"
	
	strTraining = "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Training"" class=""section-link " & strActive & """><div class=""status""><i class=""" & strIco & """></i></div>" & _
                "<div class=""content""><span class=""title"">Training</span><span class=""description"">" & strMessage & "</span></div></a>"
				
	'Set the Overview/Summary Values based on whether all checks are good or not
	If bolCompleteCheck = False Then
		strMessage = "There are Errors in your Application, in Red below"
		strCheckStatus = "warning"
		strIco = "fa fa-fw fa-times color-red"
		strCol = "Red"
		
		'''*****-----Temporarily commented out so that applications can be submitted for testing -- uncomment for Production ------*************
		strSubmitEnable = "DISABLED Title=""Correct the errors in your details (as noted above) in order to submit an application"""
	Else
		strMessage = "You are all good to Submit your application!"
		strCheckStatus = "complete"
		strIco = "fa fa-fw fa-check color-green"
		strCol = "Green"
	End If
	
	'Write the breadcrumb/title at the top depending on the side of the screen it appears
	If intType = 1 Then
	
		'Response.write "<ol class=""breadcrumb"" ><li class=""breadcrumb-item"" ><i class=""fa fa-address-card""></i> Application Summary</li></ol><ul>"
						                   
	Else
		Response.write "<span style=""color:" & strCol & "; font-weight:bold;"">" & strMessage & "</span>"
		'Response.write "<div class=""bs-vertical-wizard""><ol class=""breadcrumb"" ><li class=""breadcrumb-item"" ><i class=""fa fa-address-card""></i> Application Check Details - <span style=""color:" & strCol & "; font-weight:bold;"">" & strMessage & "</span></li></ol><ul>"
	End If
	
	If Session("ApplicationChecks") = "Overview" Then 
		strActive = "active" 'strCheckStatus = strCheckStatus & " current" 
	Else
		strActive = ""
	End If
	
	Response.Write "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Overview"" class=""section-link " & strActive & """><div class=""status""><i class=""" & strIco & """></i></div><div class=""content""><span class=""title"">Overview</span>" & _
                  "<span class=""description"">" & strMessage & "</span></div></a>"
				
	'Write details of all of the other checks (after the Summary/Overview)
	Response.Write	strEntitlement & strContact & strCMS & strName & strTraining
	
	
	'Set the Enabled variable for the Submit button based on the User Type (it has been set based on any errors above as part of this procedure)
	If Session("UserTypeID") < 10 Then
		'If the User is viewing their own details then they have full access
		If Session("ApplicationEmployeeID") <> Session("EmployeeID") Then
			strSubmitEnable = "DISABLED Title=""Your User Profile is Read Only"""
		End If
	End If
	
End Sub

Public Sub LoadEntitlement

strMessage = ""
strHelpMessage = ""

	'Check the Employee doesn't already have any contact detail issues
	objRS.Open "SELECT TOP 1 [EmployeeType] FROM tblCAPSCDMCHistory WITH(NoLock) WHERE EmployeeID = '" & Session("ApplicationEmployeeID") & "' AND [Deleted] = 'N' ORDER BY [DateUpdated] DESC",objCon

		If objRS.EOF Then
			strMessage = Session("ApplicationEmployeeID") & " has no Employment Type!? "
			strColour = "red;"
			strFaFa = "times"
			'strHelpMessage = "<div class=""input-group col-sm-3"" Style=""Font-style:italic; color:orange;"">" & GetSystemAdmin("SystemAdmin") & "</div>"
			strHelpMessage = GetSystemAdmin("SystemAdmin")
			strAlert = "danger"
			strValid = "is-invalid"
		Else
			If objRS("EmployeeType") = "Not Provided" Then
			
				strMessage = Session("ApplicationEmployeeID") & " has Employment Type - Not Provided"
				strColour = "red;"
				strFaFa = "times"
				'strHelpMessage = "<div class=""input-group col-sm-3"" Style=""Font-style:italic; color:orange;"">" & GetSystemAdmin("SystemAdmin") & "</div>"
				strHelpMessage = GetSystemAdmin("SystemAdmin")
				strAlert = "danger"
				strValid = "is-invalid"
			Else
				strMessage = objRS("EmployeeType")
				strColour = "green;"'#86c5f9
				strFaFa = "check"
				strHelpMessage = ""
				strAlert = "success"
				strValid = "is-valid"
			End If
			
		End If
	
	objRS.Close
	
	Response.write "<div class=""alert alert-" & strAlert & """ role=""alert"" style=""color:" & strColour  &";""><i class=""fa fa-" & strFaFa & " "" style=""color:" & strColour  &";""></i> Card Entitlement for " & strApplicantUserName & "</div>"
			  
	Response.Write "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
					"<div class=""form-group col-sm-8""><label for=""EmployeeType"">Employee Type</label><input type=""text""  class=""form-control " & strValid & """ id=""EmployeeType"" placeholder=""Employee Type"" value=""" & strMessage & """/>" & _
					"<div class=""invalid-feedback"">" & strHelpMessage & "</div></div>"
					
	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT [CardNumber],[Expiry],[NameOnCard] FROM tblCAPSCard WITH(NoLock) WHERE EmployeeID = '" & Session("ApplicationEmployeeID") & "' AND [Status] = '00'",objCon

		If Not objRS.EOF Then
			strMessage = Session("ApplicationEmployeeID") & " already has an active card"
			strColour = "red;"
			strFaFa = "times"
			strValid = "is-invalid"
		Else
			strMessage = "No existing card"
			strColour = "green;"
			strFaFa = "check"
			strValid = "is-valid"
		End If
		
	objRS.Close
	
					
	Response.Write "<div class=""form-group col-sm-8""><label for=""ExistingCard"">Existing Card</label><input type=""text""  class=""form-control " & strValid & """ id=""ExistingCard"" placeholder=""Existing Card"" value=""" & strMessage & """/>" & _
					"<div class=""invalid-feedback"">" & strHelpMessage & "</div></div></div>"
					
End Sub

Public Sub LoadCMS

Dim strCMSUsers
Dim strActive

strMessage = ""

	strColour = "green;"

	'The CMS User Check is already done as part of the left hand side Menu, so just use the results from that
	If strCMSCheckStatus = "warning" Then
	'strCMSMessage
		strMessage = strCMSMessage
		strColour = "red;"
		strFaFa = "times"
		strAlert = "danger"
		strValid = "is-invalid"
	Else
		strMessage = strCMSMessage
		strColour = "green;"
		strFaFa = "check"
		strAlert = "success"
		strValid = "is-valid"
	End If
	

	Response.write "<div class=""alert alert-" & strAlert & """ role=""alert"" style=""color:" & strColour  &";""><i class=""fa fa-" & strFaFa & " "" style=""color:" & strColour  &";""></i> Card CMS Account for " & strApplicantUserName & "</div>"
	
	
	Response.Write "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
					"<div class=""form-group col-sm-8""><label for=""CMSUserName"">ProMaster/CMS Account for: " & strApplicantUserName & "</label><input type=""text""  class=""form-control " & strValid & """ id=""CMSUserName"" placeholder=""CMS User Name"" value=""" & strMessage & """/>" & _
					"<div class=""invalid-feedback"">" & strHelpMessage & "</div></div></div>"

	'If there is no connection to ProMaster (Card Management System) then do not try to use the connection
	'If objCon2.State = 1 Then
		'Open a recordset in the ProMaster (CMS) database to check the Employee has a CMS Account
	'	objRS.Open "SELECT Top 100 [user_name],[employee_id] FROM procharge_user WITH(NoLock) WHERE [active_indicator ] = 'Y'",objCon2

	'		If objRS.EOF Then
	'			strCMSUsers = strCMSUsers & "<option " & strSelected & " value=""0"">No CMS Users</option>"
	'		End If
			
	'		Do Until objRS.EOF
	'		
	'			If strActive = "N" Or strActive = "" Then strSelected = " SELECTED "
	'			
	'			strCMSUsers = strCMSUsers & "<option " & strSelected & " value=""" & objRS("user_name") & """>" & objRS("user_name") & " - " & objRS("employee_id") & "</option>"
	'			
	''		objRS.Movenext
	'		Loop
			
	'	objRS.Close
	'Else
	'	strCMSUsers = strCMSUsers & "<option " & strSelected & " value=""0"">CMS database currently unavailable, please try again in 1 hour</option>"
	
	'End If	
					
	'Response.Write "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
	'				"<div class=""form-group col-sm-8""><label for=""CMSUserNameChange"">Employee Type</label><SELECT class=""form-control"" name=""CMSUserSelect"" id=""CMSUserSelect"" >" & strCMSUsers & "</Select>" & _
	'				"<div class=""invalid-feedback"">Change</div></div></div>"	
	
	Response.Write "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
					"<div class=""form-group col-sm-8""><label for=""CMSUserNameChange"">Selected CMS User Account:</label><input disabled type=""text"" class=""form-control name=""CMSUserSelect"" id=""CMSUserSelect"" value=""" & strCMSUserName & """ Title=""CMS User Account this card will be added to. Click to Change Account."">" & _
					"<div class=""invalid-feedback"">Change</div></div></div>"	
	
	
	Response.Write "<div class=""row"" style=""width:100%;""><div class=""form-group col-sm-4"">" & _
				"<button type=""button"" class=""btn btn-primary btn-block"" data-toggle=""modal"" data-target=""#EmployeeModal""><i class=""fa fa-check-search""></i>Search For CMS Account</button>" & _
				"</div></div>"
					
End Sub

Public Sub LoadName
'Procedure to load the name on card details

Dim intNameLength
Dim strNameChangeVar
Dim strChangeDetails
Dim strFormName
Dim strPrefName
Dim strTitleOptions

strMessage = ""
strHelpMessage = ""

	'Get the System Value for the Name on Card Length
	intNameLength = GetSystemAdmin("NameOnCardLength")
	
	'Get the Middle Name initial
	If Len(strFormalMiddleName)>1 Then strFormalMiddleName = Left(strFormalMiddleName,1)
	
	strFormName = strLimitChangeTitle & " " & strFormalFirstName & " " & strFormalMiddleName & " " & strFormalLastName
	strPrefName = strLimitChangeTitle & " " & strFirstName & " " & strLastName
	
	'Make sure there is a system default otherwise provide one
	If IsNull(intNameLength) or intNameLength = "" Then intNameLength = 21
	'Make sure it is a number
	If Not IsNumeric(intNameLength) Then intNameLength = 21
	
	strColour = "green;"
	
	'Check to see if the Employee ID is valid (it has already been loaded above)
	If strEmployeeID = "" Or IsNull(strEmployeeID) Then
		strMessage = "You are not on the Corporate Directory"
		strColour = "red"
		strFaFa = "times"
		strAlert = "danger"
		strHelpMessage = GetSystemAdmin("SystemAdmin")
		strValid = "is-invalid"
	Else
	
		'Check the length of the Name
		If Len(strFormName) > intNameLength Then
			If Len(strFormalFirstName & " " & strFormalLastName) > intNameLength Then
			
				strMessage = strFormalFirstName & " " & strFormalLastName
				strColour = "green"
				strFaFa = "check"
				strAlert = "success"
				strValid = "is-valid"
			Else
				strMessage = strLimitChangeTitle & " " & strFormalFirstName & " " & strFormalLastName & " name is too long. Only " & intNameLength & " characters allowed on card."
				strColour = "red"
				strFaFa = "times"
				strAlert = "danger"
				strValid = "is-invalid"
			End If
			
		Else
			strMessage = strLimitChangeTitle & " " & strFormalFirstName & " " & strFormalLastName
			strColour = "green"
			strFaFa = "check"
			strAlert = "success"
			strValid = "is-valid"
			
		End If
	End If
	
	'Get the title from the existing name on Card
	strLimitChangeTitle = GetTitleFromNameOnCard(strApplicationNameOnCard)

	'Write the title options for the Select list
	strTitleOptions = GetTitleList("Select","","Y",strLimitChangeTitle)
	
	Response.write "<div class=""alert alert-" & strAlert & """ role=""alert"" style=""color:" & strColour  &";""><i class=""fa fa-" & strFaFa & " "" style=""color:" & strColour  &";""></i> Name On Card: " & strApplicationNameOnCard & "</div>"
	
	'Response.write "<ol class=""breadcrumb"" Style=""Background-color:" & strColour & ";color:white;"">" & _
	'				"<li class=""breadcrumb-item active"" Style=""Background-color:" & strColour & ";color:white;font-size:16px;""><i class=""fa fa-address-card""></i> Card Name for " & strApplicantUserName & "</li>" & _
	'				"</ol>"
	
	'Response.Write "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
	'				"<div class=""input-group col-sm-8"">" & _
	'				"<div class=""form-group col-sm-10""><label for=""inputAddress"">Name on Your Card</label><input type=""text"" style=""width:100%;"" class=""form-control"" id=""inputAddress"" placeholder=""CMS User Name"" value=""" & strMessage & """ title=""" & strMessage & """/></div>" & _
	'				"</div><div class=""input-group col-sm-1""><i class=""fa fa-" & strFaFa & " "" style=""color:" & strColour  &";""></i></div>" & strHelpMessage
					
	Response.Write "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
					"<div class=""form-group col-sm-9""><label for=""CMSUserName"">Card Name for " & strApplicantUserName & " (Formal Name)</label>" & _
					"<div class=""form-row col-12""><div class=""form-group col-sm-3""><SELECT class=""form-control"" id=""NameOnCardTitle"" name=""NameOnCardTitle""><Option value=""none"">None</option>" & strTitleOptions & "</SELECT></div>" & _
					"<div class=""form-group col-sm-9""><input type=""text""  class=""form-control " & strValid & """ id=""CMSUserName"" placeholder=""Card Name for " & strApplicantUserName & """ value=""" & strMessage & """/></div></div>" & _
					"<div class=""invalid-feedback"">" & strHelpMessage & "</div></div></div>"
	
	'If the User has selected to change the name then write the preferred name field which can be edited
	If Session("NameChange") = "Y" Then
		strNameChangeVar = "N"
		strChangeDetails = "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
					"<div class=""form-group col-sm-9""><label for=""CMSUserName2"">Card Name for " & strApplicantUserName & " (Preferred Name)</label>" & _
					"<div class=""form-row col-12""><div class=""form-group col-sm-3""><SELECT class=""form-control"" id=""NameOnCardTitle2"" name=""NameOnCardTitle2""><Option value=""none"">None</option>" & strTitleOptions & "</SELECT></div>" & _
					"<div class=""form-group col-sm-9""><input type=""text"" maxlength=""" & intNameLength & """ class=""form-control"" id=""CMSUserName2"" placeholder=""Enter Card Name for " & strApplicantUserName & """ value=""" & strPrefName & """ onKeyUp=""GetTextKey(this)""/></div></div>" & _
					"<div id=""NameChange2Help""></div></div></div>"
					
		strChangeDetails = strChangeDetails & "<div class=""row"" style=""width:100%;""><div class=""form-group col-sm-4"">" & _
				"<button type=""button"" class=""btn btn-outline-primary btn-block"" onClick=""ChangeNameOnCard()""><i class=""fa fa-check""></i> Save Name</button>" & _
				"</div></div>"
	Else
		strNameChangeVar = "Y"
		strChangeDetails = ""
	End If
	
	Response.Write "<div class=""row"" style=""width:100%;""><div class=""form-group col-sm-4"">" & _
				"<button type=""button"" class=""btn btn-outline-secondary btn-block"" onClick=""self.location='ApplicationsSubmit.asp?NameChange=" & strNameChangeVar & "'""><i class=""fa fa-exchange-alt""></i> Change Name</button>" & _
				"</div></div>"
					
	Response.Write strChangeDetails
	
End Sub

Public Sub LoadNameOLD
'Procedure to load the name on card details

Dim intNameLength

strMessage = ""
strHelpMessage = ""

	'Get the System Value for the Name on Card Length
	intNameLength = GetSystemAdmin("NameOnCardLength")
	
	'Make sure there is a system default otherwise provide one
	If IsNull(intNameLength) or intNameLength = "" Then intNameLength = 21
	'Make sure it is a number
	If Not ISNumeric(intNameLength) Then intNameLength = 21
	
	strColour = "green;"
	
	If strEmployeeID = "" Or IsNull(strEmployeeID) Then
		strMessage = "You are not on the Corporate Directory"
		strColour = "red"
		strFaFa = "times"
		strAlert = "danger"
		strHelpMessage = GetSystemAdmin("SystemAdmin")
		strValid = "is-invalid"
	Else
		If Len(strTitle & " " & strFirstName & " " & strLastName) > intNameLength Then
			If Len(strFirstName & " " & strLastName) > intNameLength Then
				strMessage = strTitle & " " & strFirstName & " " & strLastName & " name is too long. Only " & intNameLength & " characters allowed on card."
				strColour = "red"
				strFaFa = "times"
				strAlert = "danger"
				strValid = "is-invalid"
			Else
				strMessage = strFirstName & " " & strLastName
				strColour = "green"
				strFaFa = "check"
				strAlert = "success"
				strValid = "is-valid"
			End If
			
		Else
			strMessage = strTitle & " " & strFirstName & " " & strLastName
			strColour = "green"
			strFaFa = "check"
			strAlert = "success"
			strValid = "is-valid"
		End If
	End If
	
	Response.write "<div class=""alert alert-" & strAlert & """ role=""alert"" style=""color:" & strColour  &";""><i class=""fa fa-" & strFaFa & " "" style=""color:" & strColour  &";""></i> Card Name for " & strApplicantUserName & "</div>"
					
	Response.Write "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
					"<div class=""form-group col-sm-8""><label for=""CMSUserName"">Card Name for " & strApplicantUserName & "</label><input type=""text""  class=""form-control " & strValid & """ id=""CMSUserName"" placeholder=""Card Name for " & strApplicantUserName & """ value=""" & strMessage & """/>" & _
					"<div class=""invalid-feedback"">" & strHelpMessage & "</div></div>"
	
End Sub

Public Sub LoadTraining

strMessage = ""

	'Get the System Value for the Name on Card Length
	strHelpMessage = GetSystemAdmin("TrainingMessage" & strCardType)
	
	strHelpMessage = "<a href=""https://campus.defence.gov.au/Saba/Web/Cloud"" target=""_new"">" & strHelpMessage & "</a>"
	
	'If the Training is NOT Mandatory then skip the check
	If strTrainingMandatory = "Y" Then
	
		'Description:	Loads Training Course details into page if applicable.
		objRS.Open "SELECT TOP 1 [CourseID],[CourseTitle],[EmployeeID],[CompletionDate],[FirstName],[LastName] FROM tblCAPSTraining WITH(NoLock) WHERE EmployeeID = '" & Session("ApplicationEmployeeID") & "' AND CourseID = '" & strTrainingCourse & "'",objCon

			If Not objRS.EOF Then
				strMessage = objRS("EmployeeID") & " " & objRS("FirstName") & " " & objRS("LastName") & " Completed the " & objRS("CourseTitle") & " on " & objRS("CompletionDate")
				strHelpMessage = objRS("EmployeeID") & " " & objRS("FirstName") & " " & objRS("LastName") & " Completed the " & objRS("CourseTitle") & " on " & objRS("CompletionDate")
				strColour = "green;"
				strAlert = "success"
				strFaFa = "check"
				strValid = "is-valid"
			Else
				strMessage = Session("ApplicationEmployeeID") & " has not completed the Mandatory " & strCardType & " Training "
				strColour = "red;"
				strAlert = "danger"
				strFaFa = "times"
				strValid = "is-invalid"
			End If
			
		objRS.Close
	Else
		strMessage = "Training not Mandatory for " & strCardType
		strHelpMessage = "Training not Mandatory for " & strCardType
		strColour = "green;"
		strAlert = "success"
		strFaFa = "check"
		strValid = "is-valid"
	End If
	
	Response.write "<div class=""alert alert-" & strAlert & """ role=""alert"" style=""color:" & strColour  &";""><i class=""fa fa-" & strFaFa & " "" style=""color:" & strColour  &";""></i> " & strCardType & " Card Training for " & strApplicantUserName & "</div>"
	
	Response.Write "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
					"<div class=""form-group col-sm-12""><label for=""Training"">Training</label><input type=""text""  class=""form-control " & strValid & """ id=""Training"" placeholder=""Training Complete"" value=""" & strMessage & """/>" & _
					"<div class=""invalid-feedback"">" & strHelpMessage & "</div></div>"
					
End Sub


Public Sub LoadContact

strMessage = ""
Dim strValidPost
Dim strPostMessage
Dim strValidAddress3
Dim strValidSuburb
Dim strValidState
Dim strValidPostCode
Dim strValidWorkPhone
Dim strValidMobilePhone

	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT TOP 1 [FirstName],[Surname],[PostalMessage],[IsValidPostal] FROM tblCAPSCDMCHistory WITH(NoLock) WHERE EmployeeID = '" & Session("ApplicationEmployeeID") & "' AND [Deleted] = 'N' ORDER BY [DateUpdated] DESC",objCon

		If Not objRS.EOF Then
			If objRS("IsValidPostal") = "Y" Then
				strMessage = Session("ApplicationEmployeeID") & " Corporate Directory Details"
				strColour = "green;"
				strAlert = "success"
				strFaFa = "check"
			Else
				strMessage = Session("ApplicationEmployeeID") & " incorrectly formatted address"
				strMessage = " Incorrectly formatted address."
				strColour = "red;"
				strAlert = "danger"
				strFaFa = "times"
			End If
			
			If IsNull(objRS("IsValidPostal")) Then
				strValidPost = ""
			Else
				strValidPost = objRS("IsValidPostal")
			End If
			
			If IsNull(objRS("PostalMessage")) Then
				strPostMessage = ""
			Else
				strPostMessage = objRS("PostalMessage")
			End If
			
		Else
			strMessage = Session("ApplicationEmployeeID") & " is not on the Corporate Directory"
			strColour = "red;"
			strAlert = "danger"
			strFaFa = "times"
		End If
		
	objRS.Close
	
	
	'Determine what fields to highlight based on the Postal message
	If strPostMessage = "Address has 3 lines. Only 2 allowed." Then
		strValidAddress3 = "is-invalid"
	End If
	
	If strPostMessage = "No Address" Then
		strValidAddress3 = "is-invalid"
		strValidSuburb = "is-invalid"
		strValidState = "is-invalid"
		strValidPostCode = "is-invalid"
		strValidWorkPhone = "is-invalid"
		strValidMobilePhone = "is-invalid"
	End If

	Response.write "<div class=""alert alert-" & strAlert & """ role=""alert"" style=""color:" & strColour  &";""><i class=""fa fa-" & strFaFa & " "" style=""color:" & strColour  &";""></i> Contact Details for " & strApplicantUserName & ". " & strMessage & " " & strValidPost & "</br> " & strPostMessage & "</div>"
	
	'Response.write "<ol class=""breadcrumb"" Style=""Background-color:" & strColour & ";color:white;"">" & _
	'				"<li class=""breadcrumb-item active"" Style=""Background-color:" & strColour & ";color:white;font-size:16px;""><i class=""fa fa-address-card""></i> Card Contact Details for " & strApplicantUserName & "</li>" & _
	'				"</ol>"
	
'Response.Write "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
	Response.Write	"<div class=""input-group col-sm-12"">" & _
					"<div class=""form-row col-12""><div class=""form-group col-md-6"">" & _
					"<label for=""EmployeeID"">EmployeeID</label><input type=""text"" style=""width:100%;"" class=""form-control"" id=""EmployeeID"" placeholder=""EmployeeID"" value=""" & Session("ApplicationEmployeeID") & """ title=""" & Session("ApplicationEmployeeID") & """/></div>" & _
					"<div class=""form-group col-sm-6"">" & _
					"<label for=""EmployeeID"">Card Type</label>" & _
					"<select class=""form-control"" placeholder=""Card Type"" style=""width:100%;"" name=""CardTypeSelect"" id=""CardTypeSelect"" onchange=""valCard();"">" & _
							"<option value=""DTC - Diners"">DTC - Diners</Option><option value=""DPC - ANZ"">DPC - ANZ</Option>" & _
							"<option value=""DTC - MasterCard"">DTC - Companion Mastercard</Option>" & _
							"</select></div>" & _
					"</div><div class=""form-row col-12"">" & _
					"<table width=""100%"">" & _
					"<tr><TH>Title</TH><td><input type=""text"" style=""width:100%;"" class=""form-control"" id=""Title"" value=""" & strTitle & """ title=""" & strTitle & """/></td></tr>" & _
					"<tr><TH>FirstName <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(strFirstName) & " chars</span></TH><td><input type=""text"" style=""width:100%;"" class=""form-control"" id=""FirstName"" value=""" & strFirstName & """ title=""" & strFirstName & """/></td></tr>" & _
					"<tr><TH>LastName <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(strLastName) & " chars</span></TH><td><input type=""text"" style=""width:100%;"" class=""form-control"" id=""LastName"" value=""" & strLastName & """ title=""" & strLastName & """/></td></tr>" & _
					"<tr><TH>Address1 <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;<div id=""add1"">" & Len(strAddress1) & "</div></span></TH><td><input onKeyUp=""textLength(this);"" type=""text"" style=""width:100%;"" class=""form-control " & strValidAddress3 & """ id=""Address1"" value=""" & strAddress1 & """ title=""" & strAddress1 & """/></td></tr>" & _
					"<tr><TH>Address2 <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(strAddress2) & " chars</span></TH><td><input type=""text"" style=""width:100%;"" class=""form-control " & strValidAddress3 & """ id=""Address2"" value=""" & strAddress2 & """ title=""" & strAddress2 & """/></td></tr>" & _
					"<tr><TH>Address3 <span style=""font-size:12px; font-weight:bold; color:black; float:right;"">&nbsp;&nbsp;" & Len(strAddress3) & " chars</span></TH><td><input type=""text"" style=""width:100%;"" class=""form-control " & strValidAddress3 & """ id=""Address3"" value=""" & strAddress3 & """ title=""" & strAddress3 & """/></td></tr>" & _
					"<tr><TH>Suburb <span style=""font-size:12px; font-weight:bold; color:black; text-align:right; float:right; "">&nbsp;&nbsp;" & Len(strSuburb) & " chars</span></TH><td><input type=""text"" style=""width:100%;"" class=""form-control " & strValidSuburb & """ id=""Suburb"" value=""" & strSuburb & """ title=""" & strSuburb & """/></td></tr>" & _
					"<tr><TH>State</TH><td><input type=""text"" style=""width:100%;"" class=""form-control " & strValidState & """ id=""State"" value=""" & strState & """ title=""" & strState & """/></td></tr>" & _
					"<tr><TH>PostCode</TH><td><input type=""text"" style=""width:100%;"" class=""form-control " & strValidPostCode & """ id=""PostCode"" value=""" & strPostCode & """ title=""" & strPostCode & """/></td></tr>" & _
					"<tr><TH>WorkPhone</TH><td><input type=""text"" style=""width:100%;"" class=""form-control " & strValidWorkPhone & """ id=""WorkPhone"" value=""" & strWorkPhone & """ title=""" & strWorkPhone & """/></td></tr>" & _
					"<tr><TH>MobilePhone</TH><td><input type=""text"" style=""width:100%;"" class=""form-control " & strValidMobilePhone & """ id=""MobilePhone"" value=""" & strMobilePhone & """ title=""" & strMobilePhone & """/></td></tr>" & _
					"</table>" & _
					"<div class=""alert alert-danger"">" & _
					"<strong>Note!</strong> Details cannot be changed here. All Address and Contact details MUST be updated in the <a href=""http://directory/dcd/"" target=""_new"">Corporate Directory</a>. Updates may take up to 3 days to display above." & _
					"</div>"			
	
End Sub


Public Sub LoadContact_Old

strMessage = ""
Dim strValidPost
Dim strPostMessage
Dim strValidAddress3
Dim strValidSuburb
Dim strValidState
Dim strValidPostCode
Dim strValidWorkPhone
Dim strValidMobilePhone

	'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT TOP 1 [FirstName],[Surname],[PostalMessage],[IsValidPostal] FROM tblCAPSCDMCHistory WITH(NoLock) WHERE EmployeeID = '" & Session("ApplicationEmployeeID") & "' AND [Deleted] = 'N' ORDER BY [DateUpdated] DESC",objCon

		If Not objRS.EOF Then
			If objRS("IsValidPostal") = "Y" Then
				strMessage = Session("ApplicationEmployeeID") & " Corporate Directory Details"
				strColour = "green;"
				strAlert = "success"
				strFaFa = "check"
			Else
				strMessage = Session("ApplicationEmployeeID") & " incorrectly formatted address"
				strMessage = " Incorrectly formatted address."
				strColour = "red;"
				strAlert = "danger"
				strFaFa = "times"
			End If
			
			If IsNull(objRS("IsValidPostal")) Then
				strValidPost = ""
			Else
				strValidPost = objRS("IsValidPostal")
			End If
			
			If IsNull(objRS("PostalMessage")) Then
				strPostMessage = ""
			Else
				strPostMessage = objRS("PostalMessage")
			End If
			
		Else
			strMessage = Session("ApplicationEmployeeID") & " is not on the Corporate Directory"
			strColour = "red;"
			strAlert = "danger"
			strFaFa = "times"
		End If
		
	objRS.Close
	
	
	'Determine what fields to highlight based on the Postal message
	If strPostMessage = "Address has 3 lines. Only 2 allowed." Then
		strValidAddress3 = "is-invalid"
	End If
	
	If strPostMessage = "No Address" Then
		strValidAddress3 = "is-invalid"
		strValidSuburb = "is-invalid"
		strValidState = "is-invalid"
		strValidPostCode = "is-invalid"
		strValidWorkPhone = "is-invalid"
		strValidMobilePhone = "is-invalid"
	End If

	Response.write "<div class=""alert alert-" & strAlert & """ role=""alert"" style=""color:" & strColour  &";""><i class=""fa fa-" & strFaFa & " "" style=""color:" & strColour  &";""></i> Contact Details for " & strApplicantUserName & ". " & strMessage & " " & strValidPost & "</br> " & strPostMessage & "</div>"
	
	'Response.write "<ol class=""breadcrumb"" Style=""Background-color:" & strColour & ";color:white;"">" & _
	'				"<li class=""breadcrumb-item active"" Style=""Background-color:" & strColour & ";color:white;font-size:16px;""><i class=""fa fa-address-card""></i> Card Contact Details for " & strApplicantUserName & "</li>" & _
	'				"</ol>"
	
'Response.Write "<form class=""form-inline""><div class=""row"" style=""width:100%;"">" & _
	Response.Write	"<div class=""input-group col-sm-12"">" & _
					"<div class=""form-row col-12""><div class=""form-group col-md-6"">" & _
					"<label for=""EmployeeID"">EmployeeID</label><input type=""text"" style=""width:100%;"" class=""form-control"" id=""EmployeeID"" placeholder=""EmployeeID"" value=""" & Session("ApplicationEmployeeID") & """ title=""" & Session("ApplicationEmployeeID") & """/></div>" & _
					"<div class=""form-group col-sm-6"">" & _
					"<label for=""EmployeeID"">Card Type</label>" & _
					"<select class=""form-control"" placeholder=""Card Type"" style=""width:100%;"" name=""CardTypeSelect"" id=""CardTypeSelect"" onchange=""valCard();"">" & _
							"<option value=""DTC - Diners"">DTC - Diners</Option><option value=""DPC - ANZ"">DPC - ANZ</Option>" & _
							"<option value=""DTC - MasterCard"">DTC - Companion Mastercard</Option>" & _
							"</select></div>" & _
					"</div><div class=""form-row col-12""><div class=""form-group col-md-2"">" & _
					"<label for=""Title"">Title</label><input type=""text"" style=""width:100%;"" class=""form-control"" id=""Title"" value=""" & strTitle & """ title=""" & strTitle & """/></div>" & _
					"<div class=""form-group col-md-5"">" & _
					"<label for=""FirstName"">FirstName <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(strFirstName) & " chars</span></label><input type=""text"" style=""width:100%;"" class=""form-control"" id=""FirstName"" value=""" & strFirstName & """ title=""" & strFirstName & """/></div>" & _
					"<div class=""form-group col-md-5"">" & _
					"<label for=""LastName"">LastName <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(strLastName) & " chars</span></label><input type=""text"" style=""width:100%;"" class=""form-control"" id=""LastName"" value=""" & strLastName & """ title=""" & strLastName & """/></div>" & _
					"</div><div class=""form-row col-12"">" & _
					"<div class=""form-group col-md-6"">" & _
					"<label for=""Address1"">Address1 <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(strAddress1) & " chars</span></label><input type=""text"" style=""width:100%;"" class=""form-control " & strValidAddress3 & """ id=""Address1"" value=""" & strAddress1 & """ title=""" & strAddress1 & """/></div>" & _
					"<div class=""form-group col-md-6"">" & _
					"<label for=""Address2"">Address2 <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(strAddress2) & " chars</span></label><input type=""text"" style=""width:100%;"" class=""form-control " & strValidAddress3 & """ id=""Address2"" value=""" & strAddress2 & """ title=""" & strAddress2 & """/></div>" & _
					"</div><div class=""form-row col-12"">" & _
					"<div class=""form-group col-md-6"">" & _
					"<label for=""Address3"">Address3 <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(strAddress3) & " chars</span></label><input type=""text"" style=""width:100%;"" class=""form-control " & strValidAddress3 & """ id=""Address3"" value=""" & strAddress3 & """ title=""" & strAddress3 & """/></div>" & _
					"<div class=""form-group col-md-6"">" & _
					"<label for=""Suburb"">Suburb <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(strSuburb) & " chars</span></label><input type=""text"" style=""width:100%;"" class=""form-control " & strValidSuburb & """ id=""Suburb"" value=""" & strSuburb & """ title=""" & strSuburb & """/></div>" & _
					"</div><div class=""form-row col-12"">" & _
					"<div class=""form-group col-md-6"">" & _
					"<label for=""State"">State</label><input type=""text"" style=""width:100%;"" class=""form-control " & strValidState & """ id=""State"" value=""" & strState & """ title=""" & strState & """/></div>" & _
					"<div class=""form-group col-md-6"">" & _
					"<label for=""PostCode"">PostCode</label><input type=""text"" style=""width:100%;"" class=""form-control " & strValidPostCode & """ id=""PostCode"" value=""" & strPostCode & """ title=""" & strPostCode & """/></div>" & _
					"</div><div class=""form-row col-12"">" & _
					"<div class=""form-group col-md-6"">" & _
					"<label for=""WorkPhone"">WorkPhone</label><input type=""text"" style=""width:100%;"" class=""form-control " & strValidWorkPhone & """ id=""WorkPhone"" value=""" & strWorkPhone & """ title=""" & strWorkPhone & """/></div>" & _
					"<div class=""form-group col-md-6"">" & _
					"<label for=""MobilePhone"">MobilePhone</label><input type=""text"" style=""width:100%;"" class=""form-control " & strValidMobilePhone & """ id=""MobilePhone"" value=""" & strMobilePhone & """ title=""" & strMobilePhone & """/></div>" & _
					"</div>" & _
					"<div class=""alert alert-danger"">" & _
					"<strong>Note!</strong> Details cannot be changed here. All Address and Contact details MUST be updated in the <a href=""http://directory/dcd/"" target=""_new"">Corporate Directory</a>. Updates may take up to 3 days to display above." & _
					"</div>"			
	
End Sub



Public Sub CancelApplication()

	strSQL = "UPDATE tblApplication SET Status = 'Cancelled' WHERE ApplicationID = " & Session("ApplicationID") & ""
	
	objCon.Execute strSQL
	
End Sub

Public Sub SubmitApplication()

Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSApplicationProcess"

			.Parameters.Append objCmd.CreateParameter("ApplicationID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("ReviewedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CreditLimit", adDouble, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CMSUserID", adVarChar, adParamInput, 50) 
			.Parameters.Append objCmd.CreateParameter("CDMCApplicationProcessIDOutput", adInteger, adParamOutput)
			
			.Parameters("ApplicationID") = Session("ApplicationID")
			.Parameters("ReviewedBy") = Session("UserID")
			.Parameters("CreditLimit") = Request.Form("CreditLimit")
			.Parameters("CMSUserID") = "AUTO"'Request.Form("CMSUserSelect")
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCApplicationProcessIDOutput") 
	 
		If intRecord = 0 Then
			strMessageIcon = "&nbsp;&nbsp;<img src=""../images/cross.png"" /> Application " & intRecord & " NOT approved! ERROR..."
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Application for " & strApplicantUserName & " NOT Submitted! An Error has occurred. See System Admin with Application ID: " & Session("ApplicationID") & " </div>"
			
			strMessageColour = "Red"
		Else
			strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" /> Application " & intRecord & " approved!"
			Response.Write "<div class=""alert alert-success"" role=""alert"">Application for " & strApplicantUserName & " Submitted!</div>"
			
			strMessageColour = "Black"
		End If
		
End Sub


Public Sub NameOnCardSave(strNameOnCard)

Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSNameOnCardSave"

			.Parameters.Append objCmd.CreateParameter("UniqueID", adVarChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("NameOnCard", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("TableName", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput) 
			.Parameters.Append objCmd.CreateParameter("NameChangeIDOutput", adInteger, adParamOutput)
			
			.Parameters("UniqueID") = Session("ApplicationID")
			.Parameters("NameOnCard") = strNameOnCard'Request.Form("CMSUserName2")
			.Parameters("TableName") = "tblCAPSApplication"
			.Parameters("UpdatedBy") = Session("UserID")
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("NameChangeIDOutput") 
	 
		If intRecord = 0 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Name On Card NOT Changed! Please see System Admin.</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Name on Card Changed To " & strNameOnCard & "</div>"
		End If
		
End Sub


Set objRS = Nothing
Set objCon = Nothing
%>
