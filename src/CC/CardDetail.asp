
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
Dim objRS
Dim objRS1
Dim strSelected
Dim x 
Dim strMessage
Dim strColour
Dim strViewButton

Dim strCMSUserName
Dim strCMSAccountHolder
Dim strCMSAHEID
Dim strCMSLocation
Dim strCMSAdminCentre
Dim strCMSSupervisor
Dim strEmployeeID
Dim strLastFour
Dim strCardNo
Dim strAHLocked
Dim strAHActive
Dim strAHPhone
Dim strAHEmail
Dim strCardNoFull

Dim strCMSExpiryDate
Dim strCMSCreditLimit
				
Dim strPMNameOnCard
Dim strPMCardStatus
Dim strPMAddress1
Dim strPMAddress2
Dim strPMAddress3
Dim strPMAddress4
Dim strPMStatePostCode
Dim strPMEmail
Dim strAddr4
Dim strMobile
Dim strWork
Dim strPMCardEID
Dim strCMSTxnLimit

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

'Open database connection
objCon.Open Session("DBConnection")
Session("AuditFor") = ""
'Collect the action results to save details
If Not IsEmpty(Request.QueryString("Task")) Then
	'Response.Write Request.QueryString("TaskID") & "," & Request.QueryString("Task") & "," & Request.QueryString("TaskLogID")
	Call SaveTaskLog(Request.QueryString("TaskID"),Request.QueryString("Task"),Request.QueryString("TaskLogID"))
End If

If Not IsEmpty(Request.QueryString("ViewButton")) Then
	strViewButton = Request.QueryString("ViewButton")
End If

'If the user has come from the CSTransactionsto screen then get the card Id from the card no
If Not IsEmpty(Request.QueryString("CardNo")) Then
	Session("CardID") = GetCardID(Request.QueryString("CardNo"))
End If

If Not IsEmpty(Request.QueryString("CardID")) Then
	Session("CardID") = Request.QueryString("CardID")
End If

If Not IsEmpty(Request.QueryString("UpdateCSContact")) Then
	'If the Update to CS button has been clicked then call the procedure to add contact details to the CS File (if they have chnaged)
	Call AddContactToCSFile(Request.QueryString("CardStatus"),Request.QueryString("EID"),Request.QueryString("CardType"),Request.QueryString("CardTypeSub"))
End If

If Not IsEmpty(Request.QueryString("Action")) Then

	If Request.QueryString("Action")= "NameChange" Then
		Call ChangeNameOnCard(Request.QueryString("AppStatus"),Request.QueryString("NewTitle"),Request.QueryString("NewFirstName"),Request.QueryString("NewSurname"),Request.QueryString("NewNOC"),Request.QueryString("NCAppID"),Request.QueryString("NCCardType2"))
	ElseIf Request.QueryString("Action")= "NotesChange" Then
		'Call SaveMessage
	ElseIf Request.QueryString("Action")= "TransactionChange" Then
		Call ChangeTransaction(Request.QueryString("TLCardID"),Request.QueryString("NewTransaction"),Request.QueryString("TLNameOnCard"))
	ElseIf Request.QueryString("Action")= "BrandingChange" Then
		Call ChangeBranding(Request.QueryString("NewBrand"),Request.QueryString("CardID"),Request.QueryString("BDEID"),Request.QueryString("BDEID"),Request.QueryString("ActionedBy"))
	Else
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

function LoadCAT(cb) {

	//Display the spinner modal
	$('#ModApp').modal("show");
	
	var empID = cb.getAttribute('data-EmpID');
	
	self.location='../Admin/CardAccountTransfer.asp?Link=CD&ViewButton=NewTransfer&SearchInput=' + empID;
}

function padLeadingZeros(num, size) {
    var s = num+"";
    while (s.length < size) s = "0" + s;
    return s;
}

function loadCSToDiners(varID,varCardNo) {

	//If(varCardNo)
	//const varCardNo= ['a'+ varCardNo];
	//alert(document.getElementById("CardNoMod").value)
	varCardNo=document.getElementById("CardNoMod").value
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
	document.getElementById("compareModalLabel").innerHTML = '<button type="button" class="btn btn-outline-secondary" title="Displaying CS To Diners. Click to View CS From Diners." onClick="loadCSFromDiners('+varID+',' + varCardNo +');">CS To Diners</button>'
	document.getElementById("CSToDinersDetail").innerHTML = '<img src="../images/Load.gif" style="vertical-align:middle;" /> '
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CSToDinersDetail").innerHTML = this.responseText;
    }
  };

  xhttp.open("GET", "../CC/AJAX/GetCSToDinersAudit2.asp?EmployeeID=" + varID + "&CardNo=" + varCardNo + "", true);
  xhttp.send();
}

function loadCMToNAB(varID,varCardNo) {

	varCardNo=document.getElementById("CardNoMod").value
	var xhttp = new XMLHttpRequest();
	xhttp.onreadystatechange = function() {
		document.getElementById("compareModalLabelNAB").innerHTML = '<button type="button" class="btn btn-outline-secondary" onClick="loadCSFromNAB('+varID+',' + varCardNo +');">CM To NAB</button>'
		document.getElementById("CMToNABDetail").innerHTML = '<img src="../images/Load.gif" style="vertical-align:middle;" /> ';
		document.getElementById("CMToNABDetail").innerHTML = this.responseText
		if (this.readyState == 4 && this.status == 200) {
			document.getElementById("CMToNABDetail").innerHTML = this.responseText;

		}
	};
	xhttp.open("GET", "../CC/AJAX/GetCMToNABAudit.asp?EmployeeID=" + varID + "&CardNo=" + varCardNo + "", true);
	xhttp.send();
	  
}

function loadCSFromNAB(varID,varCardNo) {

varCardNo=document.getElementById("CardNoMod").value
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
  document.getElementById("compareModalLabelNAB").innerHTML = '<button type="button" class="btn btn-outline-secondary" title="Displaying CS From NAB. Click to View CM To NAB." onClick="loadCMToNAB('+varID+',' + varCardNo +');">CS From NAB</button>'
  document.getElementById("CMToNABDetail").innerHTML = '<img src="../images/Load.gif" style="vertical-align:middle;" /> '

    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CMToNABDetail").innerHTML = this.responseText;
    }
  };
	
  xhttp.open("GET", "../CC/AJAX/GetCSFromNABAudit.asp?EmployeeID=" + varID + "&CardNo=" + varCardNo + "", true);
  xhttp.send();
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

function loadCSFromDiners(varID,varCardNo) {

//alert(varCardNo)
varCardNo=document.getElementById("CardNoMod").value
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
  document.getElementById("compareModalLabel").innerHTML = '<button type="button" class="btn btn-outline-secondary" title="Displaying CS To Diners. Click to View CS From Diners." onClick="loadCSToDiners('+varID+',' + varCardNo +');">CS From Diners</button>'
  document.getElementById("CSToDinersDetail").innerHTML = '<img src="../images/Load.gif" style="vertical-align:middle;" /> '
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CSToDinersDetail").innerHTML = this.responseText;
    }
  };
	
  xhttp.open("GET", "../CC/AJAX/GetCSFromDinersAudit2.asp?EmployeeID=" + varID + "&CardNo=" + varCardNo + "", true);
  xhttp.send();
}

function loadMessage(varID) {

  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
	
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("MessageModalText").innerHTML = this.responseText;
    }
  };
	
  xhttp.open("GET", "../CC/AJAX/GetMessage.asp?MessageID=" + varID + "", true);
  xhttp.send();
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
	
	var CardType2 = cb.getAttribute('data-NCCardType2');
	document.getElementById('NCCardType2').value=CardType2;
	
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

function OpenTL(cb) {
//Function to add the Transaction limit change detail to the modal
var varStatus

	var id = cb.getAttribute('data-TLCardID');
	document.getElementById('TLCardID').value=id;

	var NameOnCard = cb.getAttribute('data-TLNameOnCard');
	document.getElementById('TLNameOnCard').value=NameOnCard;
	
	var TLLimit = cb.getAttribute('data-TLTransactionLimit');
	document.getElementById('TLTransactionLimit').value=TLLimit;
	
	var TLCardType = cb.getAttribute('data-TLCardType');
	
	document.getElementById('TLCardType').innerHTML=TLCardType;
	
	var TLCardStatus = cb.getAttribute('data-TLCardStatus');
	if (TLCardStatus='00') { varStatus='<span class="badge badge-pill badge-success">' + TLCardStatus +'</span>'
	} else { varStatus='<span class="badge badge-pill badge-danger">' + TLCardStatus +'</span>'
	}
	document.getElementById('TLCardStatus').innerHTML=varStatus;
	//document.getElementById('TLCardStatus').innerHTML='Card Status: ' + TLCardStatus;
	
}

function OpenBD(cb) {
//Function to add the Branding change detail to the modal
	//var varBDStatus
	
	//var id = cb.getAttribute('data-BDCardID');
	//	document.getElementById('BDCardID').value=id;
	var BDLastName = cb.getAttribute('data-BDLastName');
		
	var BDFirstName = cb.getAttribute('data-BDFirstName');
		document.getElementById('BDFirstName').value=BDFirstName + ' ' + BDLastName;
	
	var BDBrand = cb.getAttribute('data-BDCurrentBranding');
		document.getElementById('BDCurrentBranding').value = BDBrand;
 
	var BDBrandNew = cb.getAttribute('data-BDBrandingNew');
		document.getElementById('BDBrandingNew').value = BDBrandNew;
	
}

function OpenCLModal(cb) {
	
  var id = cb.getAttribute('data-CLCLID');

  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ModalReviewLimitBody").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetLimitDetails.asp?CLID=" + id + "", true);
  xhttp.send();

}
function SaveBrandChange() {

	self.location = "CardDetail.asp?Action=BrandingChange&NewBrand="+document.getElementById("BDBrandingNew").value+"&CardID="+document.getElementById("BDCardID").value+"&BDEID="+document.getElementById("BDEID").value+"&ActionedBy="+document.getElementById("BDActionedBy").value;
}

function SaveNameChange() {

	self.location = "CardDetail.asp?Action=NameChange&NewNOC="+document.getElementById("NewNOC").value+"&AppStatus="+document.getElementById("NCAppNameStatus").value+"&NewFirstName="+document.getElementById("NewFirstName").value+"&NewTitle="+document.getElementById("NewTitle").value+"&NewSurname="+document.getElementById("NewSurname").value+"&NCCardType="+document.getElementById("NCCardType").value+"&NCAppID="+document.getElementById("NCAppID").value+"&NCCardType2="+document.getElementById("NCCardType2").value;
}

function SaveNotesChange() {

	self.location = "CardDetail.asp?Action=NotesChange&NewNotes="+document.getElementById("NTNewNotes").value+"&NotesID="+document.getElementById("NTNotesID").value;
}


function SaveTransactionChange() {

	self.location = "CardDetail.asp?Action=TransactionChange&NewTransaction="+document.getElementById("TLTransactionNew").value+"&TLCardID="+document.getElementById("TLCardID").value+"&TLNameOnCard="+document.getElementById("TLNameOnCard").value;
}

</script>


<!-- Modal -->
<div class="modal fade" id="ModApp" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
   <div class="loader">
        <div class="wrap">
            <div class="spinner"></div>
				<span class="loading-message">Loading...</h6>
        </div>
    </div>
</div>

<!--Loading Wait Spinner-->
	<div class="modal fade bd-example-modal-lg modalWait" id="ModalWait" data-backdrop="static" data-keyboard="false" tabindex="-1">
    <div class="modal-dialog modal-sm">
        <div class="modal-content" style="width: 88px">
            <span style="color:black;" class="spinner-border spinner-border-lg"></span>
        </div>
    </div>
</div>

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

<!-- CS To Diners Modal -->
<div class="modal fade" id="CSToDinersModal" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-large modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  CS To Diners - Displayed (click to View CS From Diners)
				  
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="CSToDinersDetail" height="100px">
               
			
               <img src="../images/Load.gif" style="vertical-align:middle;" />   
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
 <!-- End of CS To Diners Modal -->
 
 <!-- CS To NAB Modal -->
<div class="modal fade" id="CMToNABModal" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-large modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabelNAB">
                  
				  
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="CMToNABDetail" height="100px">
               
			
              <img src="../images/Load.gif" style="vertical-align:middle;" />
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
 <!-- End of CS To NAB Modal -->

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
					"<td><INPUT id=""NewNOC"" name=""NewNOC"" READONLY class=""form-control"" value=""""/><td></tr>" & _	
					"<tr><td><label for""NewTitle"" style=""font-weight:bold;"">New Title:</label></td>" & _
					"<td><INPUT id=""NewTitle"" name=""NewTitle"" class=""form-control"" value=""""/><td></tr>" & _
					"<tr><td><label for""NewFirstName""  style=""font-weight:bold;"">New First Name:</label></td>" & _
					"<td><INPUT id=""NewFirstName"" name=""NewFirstName"" class=""form-control"" value=""""/><td></tr>" & _
					"<tr><td><label for""NewSurname""  style=""font-weight:bold;"">New Surname:</label></td>" & _
					"<td><INPUT id=""NewSurname"" name=""NewSurname"" class=""form-control"" value=""""/><td></tr>" & _
					"<tr><td><input id=""NCAppID"" name=""NCAppID"" value="""" HIDDEN /></td><td><input id=""NCAppNameStatus"" name=""NCAppNameStatus"" value="""" HIDDEN /></td><td><input id=""NCCardType"" name=""NCCardType"" value="""" HIDDEN /><input id=""NCCardType2"" name=""NCCardType2"" value="""" HIDDEN /></td></tr>"
		   
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
	
	<!-- Change Transaction Limit Modal -->
    <div class="modal fade" id="TransactionModal" tabindex="-1" role="dialog" aria-labelledby="demoModalLabel" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="demoModalLabel">Transaction Limit Change</h5>
            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
              <span aria-hidden="true">&times;</span>
            </button>
          </div>
          <div class="modal-body">
		  <div class="col-md-12">
			<table class="table table-hover">
			<tr><td style="font-weight:bold;">Changing Transaction Limit for:</td><td><input id="TLNameOnCard" name = "TLNameOnCard" value="" style="border: 0;" class="form-control"/></td></tr>
           <%
		   
			Response.Write "<tr><td><div class=""panel-content row pl-3""><div style=""color:black; font-weight:bold;"">Card Type: </div><div class=""pl-3"" id=""TLCardType"" name=""TLCardType"" style=""color:black;""></div></div></td>" & _
					"<td><div class=""panel-content row pl-3""><div style=""color:black; font-weight:bold;"">Card Status: </div><div class=""pl-3"" id=""TLCardStatus"" name=""TLCardStatus"" style=""color:black;""></div></div><td></tr>" & _	
					"<tr><td><label for""TLTransactionLimit"" style=""font-weight:bold;"">Current Transaction Limit:</label></td>" & _
					"<td><INPUT id=""TLTransactionLimit"" name=""TLTransactionLimit"" class=""form-control"" value="""" style=""text-align:right;""/><td></tr>" & _	
					"<tr><td><label for""TLTransactionNew"" style=""font-weight:bold;"">New Transaction Limit:</label></td>" & _
					"<td><INPUT id=""TLTransactionNew"" name=""TLTransactionNew"" class=""form-control"" value="""" type=""number""/><td></tr>" & _
					"<tr><td><input id=""TLCardID"" name=""TLCardID"" value="""" HIDDEN /></td></tr>"
		   
		   %>
			</table>
		   </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
            <button type="button" class="btn btn-primary" onClick="SaveTransactionChange();"><i class="fa fa-check"></i> Save Changes</button>
          </div>
        </div>
      </div>
    </div>
	<!-- End Transaction Limit Modal -->
	
	
	<!-- Change Branding Modal -->
    <div class="modal fade" id="BrandingModal" tabindex="-1" role="dialog" aria-labelledby="demoModalLabel" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="demoModalLabel">Card Branding Change</h5>
            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
              <span aria-hidden="true">&times;</span>
            </button>
          </div>
          <div class="modal-body">
		  <div class="col-md-12">
			<table class="table table-hover">
			<tr><td style="font-weight:bold;">Changing Card Branding for </td><td><input id="BDFirstName" name="BDFirstName" value="" style="border: 0;" READONLY class="form-control"/></td></tr>
           <%
		   
			Response.Write "<thead><tr HIDDEN><td><input id=""BDCardID"" name=""BDCardID"" value="""" HIDDEN /><td><td><input id=""BDEID"" name=""BDEID"" value="""" HIDDEN /></td><td><input id=""BDActionedBy"" name=""BDActionedBy"" value="""" HIDDEN /></td></tr></thead><tbody><tr><td><label for""BDCurrentBranding"" style=""font-weight:bold;"">Current Card Branding:</label></td>" & _
					"<td><INPUT id=""BDCurrentBranding"" name=""BDCurrentBranding"" READONLY class=""form-control"" value=""""/><td></tr>" & _	
					"<tr><td><label for""BDBrandingNew"" style=""font-weight:bold;"">New Card Branding:</label></td>" & _
					"<td><INPUT id=""BDBrandingNew"" name=""BDBrandingNew"" READONLY class=""form-control"" value=""""/></td></tr></tbody>"
		   
		   %>
			</table>
			<%
				Response.Write "<center style=""font-weight: bold;""><span>Important Note</span></center><center><span>Submitting this request will cancel the existing card and issue a replacement card.<br>Ensure cardholder is aware of alternative payment methods that can be used while the new card is in transit.</span></center>"
			%>
		   </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
            <button type="button" class="btn btn-primary" onClick="SaveBrandingChange();"><i class="fa fa-check"></i> Save Changes</button>
          </div>
        </div>
      </div>
    </div>
	<!-- End Branding Modal -->
	
	<!-- Change Notes Modal -->
    <div class="modal fade" id="NotesModal" tabindex="-1" role="dialog" aria-labelledby="demoModalLabel" aria-hidden="true">
      <div class="modal-dialog modal-dialog-centered modal-dialog-scrollable">
        <div class="modal-content">
          <div class="modal-header">
            <h5 class="modal-title" id="demoModalLabel">Edit Notes</h5>
            <button type="button" class="close" data-dismiss="modal" aria-label="Close">
              <span aria-hidden="true">&times;</span>
            </button>
          </div>
          <div class="modal-body">
		  <div class="col-md-12" id="MessageModalText">
			
          
			
		   </div>
          </div>
          <div class="modal-footer">
            <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
            <button type="button" class="btn btn-primary" onClick="SaveNotesChange();"><i class="fa fa-check"></i> Save changes</button>
          </div>
        </div>
      </div>
    </div>
	<!-- End Change Notes Modal -->

	<!-- CDMC Modal -->
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
	
    <main class="main py-3">
      <div class="container">
		<form action="CardDetail.asp?Action=Save" method="POST" id="frm" name="frm">
        <div class="row">
          <div class="col-md-8">
  
              <div class="panel-header">
			  
				<div class="panel-header">
									 
						<%
						Call DisplayTableDetails()
						
						%>

				</div>

            </div>
			
			 <div class="col-md-4 sidebar">
		  
              <div class="panel panel-shadow panel-validation mb-3">
				  <div class="panel-header">
					<h4>Card Notes</h4>
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
                <h4>Card Change Summary <% Call LoadButtons() %></h4>
                
              </div>
			  <% Call LoadAudit() %>
              </div>
            </div>
          </div>
		  </form>
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
Dim strCardTypeSub

Dim strSQL
Dim strRelease
Dim strErrorsHeader
Dim strUpdate
'Dim strCMSUserName
'Dim strCMSAccountHolder
'Dim strCMSAHEID
Dim strErrors
Dim strCardNumber
Dim strCompanion
Dim strTransactionLimit
Dim strAccountNumber
Dim strUpdateToCSButton
Dim strNameOnCard
Dim strCDMCButton
Dim strTransferButton

Dim dteDateIssued
Dim dteDateSinceIssued
Dim strDaysColour
Dim strActiveFlag
Dim strPlasticID
Dim strReportGroup
Dim strPMLoadStatus

'If Session("EmployeeID") = "" OR ISNull(Session("EmployeeID")) Then
'	strSQL = "SELECT * FROM qryCAPSApplications WHERE CardID = '" & Session("CardID") & "'"
'Else
	strSQL = "SELECT * FROM qryCAPSCards WITH(NOLOCK) WHERE CardID = '" & Session("CardID") & "'"
'End If

objRS.Open strSQL,objCon

'Response.Write strSQL

	
    If Not objRS.EOF Then
		'If isNull(objRS(9)) Then
		'	dblEmpCont = 0
		'Else
		'	dblEmpCont = objRS(9)
		'End If
				
		If isNull(objRS("EmployeeID")) OR Trim(objRS("EmployeeID")) = "" Then
			strEmployeeID = "0"
			
			'Also reset the EmployeeID used for searches
			Session("EmployeeSearchID") = ""
			
			'''
			strTransferButton = ""'"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<button type=""button"" title=""Click to add Card to Account Transfer Screen " & trim(strCMSAHEID) & """ class=""btn btn-outline-secondary btn-sm"" data-EmpID=""" & strCMSAHEID & """ onClick=""LoadCAT(this)"";><i class=""fa fa-recycle""></i> CMS AH Transfer</button>"
		Else
			strEmployeeID = Trim(objRS("EmployeeID"))
			
			'Also set the EmployeeID used for searches
			Session("EmployeeSearchID") = objRS("EmployeeID")
			
			strCDMCButton = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<button type=""button"" title=""Click to view CDMC details for " & trim(strEmployeeID) & """ class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='CDMCDetail.asp?CardID=" & objrs("CardID") & "&EmployeeSearchID=" & objRS("EmployeeID") & "'"";><i class=""fa fa-binoculars""></i> View CDMC Screen</button>"
			strTransferButton = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<button type=""button"" title=""Click to add Card to Account Transfer Screen " & trim(strEmployeeID) & """ class=""btn btn-outline-secondary btn-sm"" data-EmpID=""" & objRS("EmployeeID") & """ onClick=""LoadCAT(this)"";><i class=""fa fa-recycle""></i> CMS AH Transfer</button>"
		End If
		
		If IsNull(objRS("CardNumberShort")) Then
			strCardNo = ""
			strLastFour = ""
		Else
			strCardNo = objRS("CardNumberShort")
			'strCardNo = MaskCard(objRS("CardNumber"))
			
			strLastFour = Right(strCardNo,4)
			
		End If	
		
		If IsNull(objRS("CardNumber")) Then
			strCardNoFull = ""
		Else
			strCardNoFull = TRIM(objRS("CardNumber"))
		End If

		If IsNull(objRS("NameOnCard")) Then
			strNameOnCard = ""
		Else
			strNameOnCard = Trim(objRS("NameOnCard"))
		End If
			
		'Call the procedure to load the CMS details for the selected EmployeeID
		'Call LoadCMSDetails()
		Call LoadCMSDetailsLocal()
		
		
		''''New for the Transfer button for the CMS Account Holder CTS accounts  ---- Feb 2025 --MG
		If strTransferButton="" Then
			strTransferButton = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<button type=""button"" title=""Click to add Card to Account Transfer Screen " & trim(strCMSAHEID) & """ class=""btn btn-outline-secondary btn-sm"" data-EmpID=""" & strCMSAHEID & """ onClick=""LoadCAT(this)"";><i class=""fa fa-recycle""></i> CMS AH Transfer</button>"
		End If
		
		
		If isNull(objRS("CardTypeSub")) Then
			strHeader = ""
			strCardType = ""
			strCardTypeSub = ""
		Else
			strHeader = objRS("CardTypeSub")
			strCardType  = objRS("CardType") & " " & objRS("CardTypeSub")
			strCardTypeSub = objRS("CardTypeSub")
		End If
				
		'Determine the image and title based on the card type
		If Trim(strHeader) = "Diners" Then
			strHeader = "<img src=""../images/icon_diners.png"" title=""" & strCardType & """> " & strCardType
		ElseIf strHeader = "ANZ" Then
			strHeader = "<img src=""../images/logo_ANZ.png"" Title=""" & strCardType & """> " & strCardType
		ElseIf strHeader = "Mastercard" Then
			strHeader = "<img src=""../images/logo_mc.png"" Title=""" & strCardType & """> " & strCardType
		Else
			strHeader = "<img src=""../images/logo_coa.png"" Title=""" & strCardType & """> " & strCardType
		End If
		
		If IsNull(objRS("Status")) Then
			'If strHeader = "ANZ" Then
			'	strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
			'Else
				strStatus = ""
			'End If
		Else
			'Determine the Status display based on the Card Type
			If strCardType = "DPC ANZ" Then
				If objRS("Status") = "" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
				ElseIf objRS("Status") = "L" OR objRS("Status") = "C" Then
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				ElseIf objRS("Status") = "T" Then
					strStatus = "<span class=""badge badge-pill badge-warning"">Temporary Hold</span>"
				Else
					strStatus = ""
				End If
			Else
				If objRS("Status") = "00" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
				ElseIf objRS("Status") = "01" OR objRS("Status") = "02" OR objRS("Status") = "03" Then
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				Else
					strStatus = ""
				End If
			End If
			
			If Left(strCardTypeSub,3) = "NAB" OR Left(strCardTypeSub,3) = "GLA" Then
				If objRS("Status") = "" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
				ElseIf objRS("Status") = "ZQ" OR objRS("Status") = "SF" OR objRS("Status") = "SS" OR objRS("Status") = "VX" Then
					'strProcessAction = "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-danger btn-sm"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";>Cancelled "& CStr(Mid(objRS("FileSeqNum"),7,2) & "/" & Mid(objRS("FileSeqNum"),5,2) & "/" & Left(objRS("FileSeqNum"),4)) & "</button>"
					strStatus = "<span class=""badge badge-pill badge-danger"">" & objRS("Status") & " " & objRS("BlockCodeDesc") & " - " & CStr(Mid(objRS("FileSeqNum"),7,2) & "/" & Mid(objRS("FileSeqNum"),5,2) & "/" & Left(objRS("FileSeqNum"),4)) & "</span>"					
				ElseIf objRS("Status") = "XS" Then
					strStatus = "<span class=""badge badge-pill badge-danger"">" & objRS("Status") & " " & objRS("BlockCodeDesc") & " - " & CStr(Mid(objRS("FileSeqNum"),7,2) & "/" & Mid(objRS("FileSeqNum"),5,2) & "/" & Left(objRS("FileSeqNum"),4)) & "</span>"					
				Else 
					strStatus = "<span class=""badge badge-pill badge-danger"">" & objRS("Status") & "</span>"
				End If	
			End If
			
			'If the Card is a DTC then display the Update CS button to add contact details to the CS File if they have changed
			strUpdateToCSButton = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<button type=""button"" title=""Click to add " & trim(strEmployeeID) & " to the CS File with Changed contact details from the CDMC file, if they are different to the Card contact details and are VALID"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='CardDetail.asp?UpdateCSContact=Y&CardID=" & objrs("CardID") & "&EID=" & objRS("EmployeeID") & "&CardStatus=" & objRS("Status") & "&CardType=" & objRS("CardType") & "&CardTypeSub=" & objRS("CardTypeSub") &"'"";><i class=""fa fa-cogs""></i> Update Details to CS</button>"
			
		End If
		
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
		
		If IsNull(objRS("CreditLimit")) Then
			strCreditLimit = ""
		Else
			If IsNumeric(objRS("CreditLimit")) Then
				If objRS("CreditLimit") > 0 Then
					If strCardType = "DPC ANZ" Then
						'strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
						strCreditLimit = FormatCurrency(objRS("CreditLimit")/100,0)

					Else
						strCreditLimit = FormatCurrency(objRS("CreditLimit")/100,0)
					End If
					
				Else
					strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
				End If
			Else
				strCreditLimit = objRS("CreditLimit")
			End If
			
		End If
		
		If left(strCardTypeSub,3) = "NAB" OR left(strCardTypeSub,3) = "GLA" Then
			If IsNumeric(objRS("CreditLimit")) Then
				strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
			Else
				strCreditLimit = 0
			End If
		End If
		
		'Get the transaction limit and format
		If IsNull(objRS("TransactionLimit")) Then
			strTransactionLimit = 0
		Else
			If IsNumeric(objRS("TransactionLimit")) Then
				If Left(strCardType,3) = "DTC" Then
				'If strCardType = "DTC Diners" Then
					'strCurrentLimit = FormatCurrency(objRS("TransactionLimit")/100,0)
					strTransactionLimit = FormatCurrency(objRS("TransactionLimit"),0)
				Else
				
					'strTransactionLimit = FormatCurrency(objRS("TransactionLimit")/100,0)
					strTransactionLimit = objRS("TransactionLimit")
					'strTransactionLimit=Len(strTransactionLimit)
					
					'''---UPDATE APRIL 2023 for new DPC change
					'-----If the card is a DPC Mastercard then display the transaction limit without alterting number (dividing by anything)
					If strCardType = "DPC Mastercard" Then
						'Leave Transaction limit as is
					Else
					
						If Len(strTransactionLimit)=7 Then
							strTransactionLimit = strTransactionLimit/100
						Else
							strTransactionLimit = strTransactionLimit/10
						End If
						
					End If
					
					'Changed March 2022 to being formatted as the data from ANZ can have decimal issues which are now fixed in the SQL import ANZ Cardlist procedure
					strTransactionLimit = FormatCurrency(strTransactionLimit,0)
					'strTransactionLimit=Len(strTransactionLimit)
				End If
			End If
		End If
		
		If Left(strCardTypeSub,3) = "NAB" OR Left(strCardTypeSub,3) = "GLA" Then
			If IsNull(objRS("NABTransLimit")) Then
				strTransactionLimit = 0
			Else
				strTransactionLimit = objRS("NABTransLimit")
			End If
			
			If IsNull(strTransactionLimit) = True Then
				strTransactionLimit = 0
			End If
		End If
		
		If IsNull(objRS("CardNumber")) Then
			strCardNumber = ""
		Else
			'strCardNumber = FormatCardNumber(objRS("CardNumber"))
			
			'''''If the User Type is not an admin then mask the card numbers
			If Session("UserTypeID") < 10 Then
				strCardNumber = MaskCard(objRS("CardNumber"))
			Else
				strCardNumber = FormatCardNumber(objRS("CardNumber"))
			End If		
		End If
		
		If IsNull(objRS("Companion")) Then
			strCompanion = ""
		Else
			'strCompanion = FormatCardNumber(objRS("Companion"))
			
			'''''If the User Type is not an admin then mask the card numbers
			If Session("UserTypeID") < 10 Then
				strCompanion = MaskCard(objRS("Companion"))
			Else
				strCompanion = FormatCardNumber(objRS("Companion"))
			End If
			
		End If
		
		If IsNull(objRS("AccountNumber")) Then
			strAccountNumber = ""
		Else
			'If strCardType = "DTCDiners" Then
				strAccountNumber = FormatCardNumber(objRS("AccountNumber"))
			'Else
			
			'End If
			
		End If
		
		If IsNull(objRS("DateIssued")) Then
			dteDateIssued = ""
		Else
			dteDateIssued = FormatDateTime(objRS("DateIssued"),vbShortDate)
		End If
		
		If IsNull(objRS("DateIssued")) Then
			dteDateSinceIssued = ""
		Else
			dteDateSinceIssued = DateDiff("d",objRS("DateIssued"),now())
			If dteDateSinceIssued > 80 Then
				strDaysColour = "Style=""color:red; font-weight:bold;"""
			ElseIf dteDateSinceIssued > 45 Then
				strDaysColour = "Style=""color:orange; font-weight:bold;"""
			Else
				strDaysColour = "Style=""color:black"""
			End If
		End If
		Dim strActiveTitle
		'Get the ActivationFlag details for each record
		If IsNull(objRS("ActivationFlag")) or objRS("ActivationFlag")="" Then
			strActiveFlag = ""
		Else
			strActiveFlag = objRS("ActivationFlag")
			
			If strActiveFlag = "Y" Then
				strActiveFlag = "<span class=""badge badge-pill badge-success"">Activated</span>"
			ElseIf strActiveFlag = "N" Then
				strActiveFlag = "<span class=""badge badge-pill badge-danger"">Not Activated</span>  <b>" & dteDateSinceIssued & "</b> Days since Card Issued"
				strActiveTitle = "Title="" " & dteDateSinceIssued & " Days since card was issued"" "
			Else
				strActiveFlag = ""
			End If
			
		End If
		
		'Get the Plastic ID value and return the Card Type (branded or unbranded)
		If IsNull(objRS("PlasticID")) Then
			dteDateSubmitted = ""
		Else
			Select Case objRS("PlasticID")
			
				Case "001" 
					strPlasticID = "<span class=""badge badge-pill badge-secondary"">Unbranded DTC</span>"
				Case "035"
					strPlasticID = "<span class=""badge badge-pill badge-secondary"">Branded DTC</span>"
				Case "048"
					strPlasticID = "<span class=""badge badge-pill badge-secondary"">Branded DPC</span>"
				Case "049"
					strPlasticID = "<span class=""badge badge-pill badge-secondary"">Unbranded DPC</span>"
				Case "V"
					strPlasticID = "<span class=""badge badge-pill badge-secondary"">Virtual</span>"
				Case "U" 
					strPlasticID = "<span class=""badge badge-pill badge-secondary"">Unbranded DTC</span>"
				Case "B"
					strPlasticID = "<span class=""badge badge-pill badge-secondary"">Branded DTC</span>"
				Case Else
					strPlasticID = ""
			End Select
			
		End If
		
		If IsNull(objRS("DateLoaded")) Then
			dteDateSubmitted = ""
		Else
			dteDateSubmitted = FormatDateTime(objRS("DateLoaded"),vbShortDate)
		End If
		
		If IsNull(objRS("PMLoadStatus")) OR objRS("PMLoadStatus") = "" Then
			strPMLoadStatus = "<span class=""badge badge-pill badge-danger"">No PMLoadStatus, contact System Admin</span>"
		Else
				Select CASE objRS("PMLoadStatus")
					CASE "Pending"
						strPMLoadStatus = "<span class=""badge badge-pill badge-success"">Pending</span>"
					CASE "Pending update"
						strPMLoadStatus = "<span class=""badge badge-pill badge-warning"">Pending Update</span>"
					CASE "Cancel"
						strPMLoadStatus = "<span class=""badge badge-pill badge-danger"">Pending Cancellation</span>"
					CASE "Exported"
						strPMLoadStatus = "<span class=""badge badge-pill badge-primary"">Loaded</span>"
					CASE "Pending New Card"
						strPMLoadStatus = "<span class=""badge badge-pill badge-primary"">Pending New Card</span>"
				End Select
		End If
		
	'Create the view contact details from CDMC button for further down the screen
	strUpdate = "<tr><th></th><td><button type=""button"" class=""btn btn-outline-secondary btn-sm"" data-toggle=""modal"" data-target=""#CDMCModal"" HREF=""#"" onClick=""loadCDMC(" & objRS("EmployeeID") & ")""><i class=""fa fa-eye""></i> View CDMC</button></td></tr>"
		
	Response.write "<div class=""panel-content row""><div class=""mb-3 col-md-4""><h4>" & strHeader & "</h4></div><div class=""mb-3 col-md-5"">" &  _
		"<div class=""btn-group btn-selector table-tabs-selector"" role=""group"" aria-label=""Basic example"">" &  _
		"<button type=""button"" data-target=""table-tabs"" data-type=""as-tabs"" class=""btn btn-outline-primary active"">" &  _
		"<i class=""fa fa-list""></i> View as Tabs</button>" &  _
		"<button type=""button"" data-target=""table-tabs"" data-type=""as-table"" class=""btn btn-outline-primary"">" &  _
		"<i class=""fa fa-table""></i> View as Table</button></div></div>" & strRelease & "</div>" &  _
		"<div class=""panel-content row""><div class=""mb-3 col-md-8""><h6>" & objRS("EmployeeID") & " - " & objRS("NameOnCard") & "</h6></div><div class=""mb-3 col-md-1""></div><div class=""mb-3 col-md-3""><button type=""button"" class=""btn btn-outline-primary btn-sm"" title=""Close and return to Card List with " & strNameOnCard & " selected"" onclick=""self.location='Cards.asp?EmployeeID=" & objRS("EmployeeID") & "'"";><i class=""fa fa-times""></i> <i class=""fa fa-user""></i> Close Employee</button></div></div>" & _
		"<div id=""table-tabs"" class=""as-tabs""><ul class=""nav nav-tabs"" id=""myFiTab"" role=""tablist""><li class=""nav-item"" role=""presentation"">" &  _
		"<a class=""nav-link active"" id=""overview-tab"" data-toggle=""tab"" href=""#overview"" role=""tab"" aria-controls=""overview"" aria-selected=""true"">Card Details</a>" &  _
		"</li><li class=""nav-item"" role=""presentation"">" &  _
		"<a class=""nav-link"" id=""card-details-tab"" data-toggle=""tab"" href=""#card-details"" role=""tab"" aria-controls=""card-details"" aria-selected=""false"">Contact Details</a>" &  _
		"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-limits-tab"" data-toggle=""tab"" href=""#my-limits"" role=""tab"" aria-controls=""my-limits"" aria-selected=""false"">Limits</a>" &  _
		"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-cms-tab"" data-toggle=""tab"" href=""#my-cms"" role=""tab"" aria-controls=""my-cms"" aria-selected=""false"">CMS Details</a>" &  _
		"</li><li class=""nav-item"" role=""presentation""><a class=""nav-link"" id=""my-workflow-tab"" data-toggle=""tab"" href=""#my-workflow"" role=""tab"" aria-controls=""my-workflow"" aria-selected=""false"">Workflow</a>" &  _
		"</li>" & strErrorsHeader & "</ul><div class=""tab-content panel panel-light p-3"" id=""myFiTabContent"">" &  _
		"<div class=""tab-pane fade show active"" id=""overview"" role=""tabpanel"" aria-labelledby=""overview-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>Card ID</th><td>" & objRS("CardID") & "</td></tr>" &  _
		"<tr><th>Card No.</th><td>" & strCardNumber & "</td></tr>" &  _
		"<tr><th>Status</th><td>" & strStatus & "</td></tr>" &  _
		"<tr><th>Date Loaded</th><td>" & objRS("DateLoaded") & "</td></tr>" &  _
		"<tr><th>Employee ID</th><td style=""font-weight:bold;"">" & objRS("EmployeeID") & "</td></tr>" 
		If Left(strCardTypeSub,3) = "NAB" Then
			Response.Write "<tr><th>Name On Card</th><td style=""font-weight:bold;""><a data-toggle=""modal"" data-target=""#NameOnCardModal"" data-NCCardType2=""" & objRS("CardType") & """ data-NCCardType=""" & objRS("CardTypeSub") & """ data-NCAppNameStatus=""" & objRS("Status") & """ data-NCAppID=""" & objRS("CardID") & """ data-NCAppName=""" & objRS("NameOnCard") & """ data-NCTitle=""" & objRS("Title") & """ data-NCFirstName=""" & objRS("FirstName") & " " & objRS("MiddleName") & """ data-NCSurname=""" & objRS("Surname") & """ onClick=""OpenNC(this);"">" & objRS("NameOnCard") & " <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(Trim(objRS("NameOnCard"))) & " chars</a></span> &nbsp;&nbsp;<span style=""font-size:11px; color:gray;"">Click to change</span></td></tr>" 
		Else
			Response.Write "<tr><th>Name On Card</th><td style=""font-weight:bold;"">" & objRS("NameOnCard") & " <span style=""font-size:12px; font-weight:bold; color:black;"">&nbsp;&nbsp;" & Len(Trim(objRS("NameOnCard"))) & " chars</a></span></td></tr>" 
		End If
		
		Response.Write "<tr><th>Title</th><td>" & objRS("Title") & "</td></tr>" & _
		"<tr><th>First Name(s)</th><td>" & objRS("FirstName") & " " & objRS("MiddleName") & "</td></tr>" &  _
		"<tr><th>Surname</th><td>" & objRS("Surname") & "</td></tr>" &  _
		"<tr><th>Expiry</th><td>" & objRS("Expiry") & "</td></tr>" &  _
		"<tr><th>Plastic ID</th><td>" & objRS("PlasticID") & " " & strPlasticID & "</td></tr>" &  _
		"<tr><th>Report Group</th><td>" & objRS("ReportGroup") & "</td></tr>"
		
		
		If Left(strCardTypeSub,3) = "NAB" OR Left(strCardTypeSub,3) = "GLA" Then
		
		Dim strCashLimitAmount
		Dim strNewSuffix
		Dim strFirstName
		Dim strLastName
		Dim strName
		
		If IsNull(objRS("CashLimitAmount")) = True Then		
			strCashLimitAmount = 0 			
		Else
			strCashLimitAmount = objRS("CashLimitAmount")
		End If
		
		If IsNull(objRS("Surname")) = True Then		
			strLastName = "Error" 			
		Else
			strLastName = objRS("Surname")
		End If
		
		If IsNull(objRS("FirstName")) = True Then		
			strFirstName = "Error" 			
		Else
			strFirstName = objRS("FirstName")
		End If
		

		
			Select CASE IsNull(objRS("CardSuffix"))
				CASE "LW"
					strNewSuffix = "NA"
				CASE "LU"
					strNewSuffix = "EK"
				CASE "EK"
					strNewSuffix = "LU"
				CASE "KH"	
					strNewSuffix = "EV"
				CASE "EV"
					strNewSuffix = "KH"
				Case Else
					strNewSuffix = ""
			End Select
		
			If objRS("CardSuffix") = "LW" Then
				'Response.Write 	"<tr><th>Cash Limit Amount</th><td>" & formatcurrency(strCashLimitAmount,0) & "</td></tr>" &  _
				'"<tr><th>Card Suffix</th><td>" & objRS("CardSuffix") & "</td></tr>"
				Response.Write "<tr><th>Card Suffix</th><td>" & objRS("CardSuffix") & "</td></tr>"
			Else
				'Response.Write 	"<tr><th>Cash Limit Amount</th><td>" & formatcurrency(strCashLimitAmount,0) & "</td></tr>" &  _
				'"<tr><th>Card Suffix</th><td>" & objRS("CardSuffix") & "</td></tr>"
				Response.Write "<tr><th>Card Suffix</th><td>" & objRS("CardSuffix") & "</td></tr>"
			End If	
				'"<tr><th>Card Suffix</th><td><a data-toggle=""modal"" data-target=""#BrandingModal"" data-BDCardID=""" & objRS("CardID") & """ data-BDCurrentBranding=""" & objRS("CardSuffix") & """ data-BDBrandingNew=" & strNewSuffix & " data-BDEID=" & strEmployeeID & " data-BDActionedBy=""" & Session("UserID") & """ data-BDFirstName=" & strFirstName & " data-BDLastName=" & strLastName & " data-BDCardStatus=""" & objRS("Status") & """ data-BDCardType=""" & objRS("CardType") & " " & objRS("CardTypeSub") & """ onClick=""OpenBD(this);"">" & objRS("CardSuffix") & "</a> &nbsp;&nbsp;<span style=""font-size:11px; color:gray;"">Click to change</span></td></tr>" & _ 

		Else	
			Response.Write 	"<tr><th>ATM Limit</th><td>" & objRS("ATMLimit") & "</td></tr>" &  _
			"<tr><th>OTC Limit</th><td>" & objRS("OTCLimit") & "</td></tr>"
			
		End If
		
		Response.write "</table></div>" &  _
		"<div class=""tab-pane fade"" id=""my-limits"" role=""tabpanel"" aria-labelledby=""my-limits-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>Permanent Credit Limit</th><td>" & strCreditLimit & " <span style=""font-size:11px; color:grey;"">($" & objRS("CreditLimitAmount") & " - unformatted - notify admin if these numbers differ)</span></td></tr>"
		
		Response.write"<tr><th>---- Credit Limit</th><td>" & strCardTypeSub & " --" & Left(strCardTypeSub,3) & "--</td></tr>"
		'For the NAB temp limit changes coming Feb 2024 - Tiff
		If Left(strCardTypeSub,3) = "NAB" OR Left(strCardTypeSub,3) = "GLA" Then
		Dim strActiveLimitAmount
		Dim strActiveMonthlyAppID
		Dim strActiveTXNCode
		Dim strActiveTXNAppID	
		Dim strDisplayActiveMonthlyDetails
		Dim strDisplayActiveTXNDetails
		Dim strNABLimitAmount
		Dim strNABLimitEndDate
		Dim strDisplayNABLimitDetails
		Dim strCLMonthlyAction
		Dim strCLTXNAction
		
		strActiveLimitAmount = objRS("ActiveCLAmount")
		strActiveMonthlyAppID = objRS("ActiveCLID")
		strActiveTXNCode = objRS("ActiveTXNCode")
		strActiveTXNAppID = objRS("ActiveTLID")
		strNABLimitAmount = objRS("NABTempCrLmt")
		strNABLimitEndDate = objRS("NABTempCrLmtExpd")
		strCLMonthlyAction = "<button type=""button"" class=""btn btn-primary btn-sm"" onClick=OpenCLModal(this); data-toggle=""modal"" data-target=""#ModalReviewLimit"" data-CLCLID=" & strActiveMonthlyAppID & " title=""View Limit Details""><i class=""fa fa-money-bill""></i></button>"
		strCLTXNAction = "<button type=""button"" class=""btn btn-primary btn-sm"" onClick=OpenCLModal(this); data-toggle=""modal"" data-target=""#ModalReviewLimit"" data-CLCLID=" & strActiveTXNAppID & " title=""View Limit Details""><i class=""fa fa-money-bill""></i></button>"
		
		If strActiveMonthlyAppID = 0 Then
			strDisplayActiveMonthlyDetails = ""
		Else
			strDisplayActiveMonthlyDetails = "<tr><th>Temporary Credit Limit</th><td>" & FormatCurrency(strActiveLimitAmount,0) & "&nbsp;&nbsp;&nbsp;&nbsp;" & strCLMonthlyAction & "<span style=""font-size:11px; color:grey;""> Click to view more details</span></td></tr>"
			'strDisplayActiveMonthlyDetails = strDisplayActiveMonthlyDetails & "<tr><th>Credit Limit ID</th><td>" & strActiveMonthlyAppID & "  " & strCLMonthlyAction & "</td></tr>"
		End If
		
		If strActiveTXNAppID = 0 Then
			strDisplayActiveTXNDetails = ""
		Else
			strDisplayActiveTXNDetails = "<tr><th>Temporary Trans. Code</th><td>" & strActiveTXNCode & "&nbsp;&nbsp;&nbsp;&nbsp;" & strCLTXNAction & "<span style=""font-size:11px; color:grey;""> Click to view more details</span></td></tr>"
			'strDisplayActiveTXNDetails = strDisplayActiveTXNDetails & "<tr><th>Transaction Limit ID</th><td>" & strActiveTXNAppID & "  " & strCLTXNAction & "</td></tr>"
		End If
		
		If IsNumeric(strNABLimitAmount) Then
		If strNABLimitAmount = 0 Then
			strDisplayNABLimitDetails = ""
		Else
			strDisplayNABLimitDetails = "<tr><th title=""This is the 'additional' amount currently available against the card"">NABConnect Limit Amount</th><td>" & FormatCurrency(strNABLimitAmount,0) & "</td></tr>"
			strDisplayNABLimitDetails = strDisplayNABLimitDetails & "<tr><th>NABConnect Limit End Date</th><td>" & strNABLimitEndDate & "</td></tr>"
		End If	
		End If
			Response.Write 	strDisplayActiveMonthlyDetails & "<tr><th>Permanent Trans. Code</th><td>" & objRS("TransLimitCode") & "</td></tr>" & strDisplayActiveTXNDetails & strDisplayNABLimitDetails

		Else
			Response.Write "<tr><th>Credit Limit</th><td>" & strCreditLimit & " <span style=""font-size:11px; color:grey;"">($" & objRS("CreditLimitAmount") & " - unformatted - notify admin if these numbers differ)</span></td></tr>" & _
						"<tr><th>Transaction Limit</th><td><a data-toggle=""modal"" data-target=""#TransactionModal"" data-TLCardID=""" & objRS("CardID") & """ data-TLTransactionLimit=""" & strTransactionLimit & """ data-TLNameOnCard=""" & objRS("NameOnCard") & """ data-TLCardStatus=""" & objRS("Status") & """ data-TLCardType=""" & objRS("CardType") & " " & objRS("CardTypeSub") & """ onClick=""OpenTL(this);"">" & formatcurrency(strTransactionLimit,0) & "</a> &nbsp;&nbsp;<span style=""font-size:11px; color:gray;"">Click to change</span></td></tr>"
		End If

		
		Response.Write "</table></div>" &  _
		"<div class=""tab-pane fade"" id=""card-details"" role=""tabpanel"" aria-labelledby=""card-details-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th title=""The Date " & Trim(objRS("NameOnCard")) & " last appeared on the HR/CDMC File, prior to 5 days countdown"">Left Defence?</th><td>" & GetDateLeftDefence(strEmployeeID) & " " & strCDMCButton & "</td></tr>" &  _
		"<tr><th>Address 1</th><td>" & objRS("Address1") & "</td></tr>" &  _
		"<tr><th>Address 2</th><td>" & objRS("Address2") & "</td></tr>" &  _
		"<tr><th>Address 3</th><td>" & objRS("Address3") & "</td></tr>" &  _
		"<tr><th>Suburb</th><td>" & objRS("Suburb") & "</td></tr>" &  _
		"<tr><th>State</th><td>" & objRS("State") & "</td></tr>" &  _
		"<tr><th>PostCode</th><td>" & objRS("PostCode") & "</td></tr>" &  _
		"<tr><th>Home Phone</th><td>" & objRS("HomePhone") & "</td></tr>" &  _
		"<tr><th>Work Phone</th><td>" & objRS("WorkPhone") & "</td></tr>" &  _
		"<tr><th>Mobile Phone</th><td>" & objRS("MobilePhone") & "</td></tr>" &  _
		"<tr><th>Email</th><td>" & objRS("Email") & " " & strUpdateToCSButton & "</td></tr>" & strUpdate & _
		"</table></div>" & _
		"<div class=""tab-pane fade"" id=""my-cms"" role=""tabpanel"" aria-labelledby=""cms-details-tab"">" &  _
		"<table class=""table"">" &  _
		"<tr><th>PM Load Status</th><td>" & strPMLoadStatus & "</td></tr>"
			
		Select CASE strPMCardStatus
			CASE "N"
				strPMCardStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
			CASE "C"
				strPMCardStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
		End Select
		
		If LTRIM(RTRIM(objRS("Email"))) <> LTRIM(RTRIM(strPMEmail)) Then
			strPMEmail = "<span class=""badge badge-pill badge-danger"">" & strPMEmail
		End If
			
		If LTRIM(RTRIM(objRS("NameOnCard"))) <> LTRIM(RTRIM(strPMNameOnCard)) Then
			strPMNameOnCard = "<span class=""badge badge-pill badge-danger"">" & strPMNameOnCard
		End If
		
		If LTRIM(RTRIM(objRS("Address1"))) <> LTRIM(RTRIM(strPMAddress1)) Then
			strPMAddress1 = "<span class=""badge badge-pill badge-danger"">" & strPMAddress1
		End If
		
		If LTRIM(RTRIM(objRS("Address2"))) <> LTRIM(RTRIM(strPMAddress2)) Then
			strPMAddress2 = "<span class=""badge badge-pill badge-danger"">" & strPMAddress2
		End If
		
		If LTRIM(RTRIM(objRS("Address3"))) <> LTRIM(RTRIM(strPMAddress3)) Then
			strPMAddress3 = "<span class=""badge badge-pill badge-danger"">" & strPMAddress3
		End If
		
		'--concat suburb, state, postcode to compare against cms address 4
		Dim strCapsAddr4
		strCapsAddr4 = (LTRIM(RTRIM(objRS("Suburb"))) & " " & LTRim(RTRIM(objRS("State"))) & " " & LTRIM(RTRIM(objRS("PostCode"))))
		'response.write strCapsAddr4
		If strCapsAddr4 <> LTRIM(RTRIM(strPMAddress4)) Then
			strPMAddress4 = "<span class=""badge badge-pill badge-danger"">" & strPMAddress4
		End If
		
		If LTRIM(RTRIM(objRS("WorkPhone"))) <> LTRIM(RTRIM(strWork)) Then
			strWork = "<span class=""badge badge-pill badge-danger"">" & strWork
		End If
			
		If LTRIM(RTRIM(objRS("MobilePhone"))) <> LTRIM(RTRIM(strMobile)) Then
			strMobile = "<span class=""badge badge-pill badge-danger"">" & strMobile
		End If
		
		If LTRIM(RTRIM(objRS("MobilePhone"))) <> LTRIM(RTRIM(strMobile)) Then
			strMobile = "<span class=""badge badge-pill badge-danger"">" & strMobile
		End If
		
		If IsNumeric(strCMSTxnLimit) Then
			strCMSTxnLimit = strCMSTxnLimit
		Else
			If IsNull(strCMSTxnLimit) Then
				strCMSTxnLimit = 0
			Else
				strCMSTxnLimit = strCMSTxnLimit
			End If
		End If
		
		If IsNumeric(strCMSCreditLimit) Then
			strCMSCreditLimit = strCMSCreditLimit
		Else
			If IsNull(strCMSCreditLimit) Then
				strCMSCreditLimit = 0
			Else
				strCMSCreditLimit = strCMSCreditLimit
			End If
		End If
		
		'response.write objRS("CreditLimitAmount")
		If LTRIM(RTRIM(strCMSCreditLimit)) <> LTRIM(RTRIM(objRS("CreditLimitAmount"))) Then
				strCMSCreditLimit = "<span class=""badge badge-pill badge-danger"">" & FormatCurrency(strCMSCreditLimit,0)
			Else
				strCMSCreditLimit = FormatCurrency(strCMSCreditLimit,0)
		End If

		If LTRIM(RTRIM(strCMSTxnLimit)) <> LTRIM(RTRIM(objRS("TransLimitAmount"))) Then
				strCMSTxnLimit = "<span class=""badge badge-pill badge-danger"">" & FormatCurrency(strCMSTxnLimit,0)
		Else
				strCMSTxnLimit = FormatCurrency(strCMSTxnLimit,0)
		End If
		
		Dim strOwnHolder 'checks if user is their own account holder in CMS based off EID
		If strCMSAHEID = strPMCardEID Then
			strOwnHolder = "Yes"
		Else
			strOwnHolder = "No"
		End If

		If strCMSUserName = "None" Then 'card not loaded to CMS yet
			Response.Write "<tr><th></th><td>If you believe an error is preventing this card being loaded to CMS,<br>please visit the ProMaster Reconciliation Screen</td></tr>"
		Else
			If strOwnHolder = "Yes" Then
				If objRS("CardSuffix") = "LW" Then
					Response.Write	"<tr><th><span style=""font-size:1.5em"">CMS Account Details</span></th><td></td></tr>" & _
									"<tr><th>Own Account Holder?</th><td>" & strOwnHolder & "</td></tr>" & _
									"<tr><th>CMS UserID</th><td>" & strCMSUserName & " " & strTransferButton & "</td></tr>" &  _
									"<tr><th>Active?</th><td>" & strAHActive & "</span></td></tr>" &  _
									"<tr><th>Locked?</th><td>" & strAHLocked & "</span></td></tr>" &  _
									"<tr><th><span style=""font-size:1.5em""><centre>CMS Card Details</centre></span></th><td></td></tr>" & _
									"<tr><th>Card Status</th><td>" & strPMCardStatus & "</span></td></tr>" &  _
									"<tr><th>Expiry Date</th><td>" & strCMSExpiryDate & "</span></td></tr>" &  _
									"<tr><th>Email</th><td>" & strPMEmail & "</span></td></tr>" &  _
									"<tr><th>Name on Card</th><td>" & strPMNameOnCard & "</span></td></tr>" &  _
									"<tr><th>Address 1</th><td>" & strPMAddress1 & "</span></td></tr>" &  _
									"<tr><th>Address 2</th><td>" & strPMAddress2 & "</span></td></tr>" &  _
									"<tr><th>Address 3</th><td>" & strPMAddress3 & "</span></td></tr>" &  _
									"<tr><th>Address 4</th><td>" & strPMAddress4 & "</span></td></tr>" &  _
									"<tr><th>Work Phone</th><td>" & strWork & "</span></td></tr>" &  _
									"<tr><th>Mobile Phone</th><td>" & strMobile & "</span></td></tr>" &  _
									"<tr><th>Credit Limit</th><td>" & strCMSCreditLimit & "</span></td></tr>"
									
				Else
					Response.Write "<tr><th><span style=""font-size:1.5em"">CMS Account Details</span></th><td></td></tr>" & _
									"<tr><th>Own Account Holder?</th><td>" & strOwnHolder & "</td></tr>" & _
									"<tr><th>CMS UserID</th><td>" & strCMSUserName & " " & strTransferButton & "</td></tr>" &  _
									"<tr><th>Active?</th><td>" & strAHActive & "</span></td></tr>" &  _
									"<tr><th>Locked?</th><td>" & strAHLocked & "</span></td></tr>" &  _
									"<tr><th><span style=""font-size:1.5em""><centre>CMS Card Details</centre></span></th><td></td></tr>" & _
									"<tr><th>Card Status</th><td>" & strPMCardStatus & "</span></td></tr>" &  _
									"<tr><th>Expiry Date</th><td>" & strCMSExpiryDate & "</span></td></tr>" &  _
									"<tr><th>Email</th><td>" & strPMEmail & "</span></td></tr>" &  _
									"<tr><th>Name on Card</th><td>" & strPMNameOnCard & "</span></td></tr>" &  _
									"<tr><th>Address 1</th><td>" & strPMAddress1 & "</span></td></tr>" &  _
									"<tr><th>Address 2</th><td>" & strPMAddress2 & "</span></td></tr>" &  _
									"<tr><th>Address 3</th><td>" & strPMAddress3 & "</span></td></tr>" &  _
									"<tr><th>Address 4</th><td>" & strPMAddress4 & "</span></td></tr>" &  _
									"<tr><th>Work Phone</th><td>" & strWork & "</span></td></tr>" &  _
									"<tr><th>Mobile Phone</th><td>" & strMobile & "</span></td></tr>" &  _
									"<tr><th>Credit Limit</th><td>" & strCMSCreditLimit & "</span></td></tr>" & _
									"<tr><th>Transaction Limit</th><td>" & strCMSTxnLimit & "</span></td></tr>"
				End If
			Else
				If objRS("CardSuffix") = "LW" Then
					Response.Write	"<tr><th><span style=""font-size:1.5em"">CMS Account Details</span></th><td></td></tr>" & _
									"<tr><th>Own Account Holder?</th><td>" & strOwnHolder & "</td></tr>" & _
									"<tr><th>Account Holder</th><td>" & strCMSAccountHolder & "</span></td></tr>" &  _
									"<tr><th>Account Holder EID</th><td>" & strCMSAHEID & "</span></td></tr>" &  _
									"<tr><th>CMS UserID</th><td>" & strCMSUserName & " " & strTransferButton & "</td></tr>" &  _
									"<tr><th>Active?</th><td>" & strAHActive & "</span></td></tr>" &  _
									"<tr><th>Locked?</th><td>" & strAHLocked & "</span></td></tr>" &  _
									"<tr><th>Phone</th><td>" & strAHPhone & "</span></td></tr>" &  _
									"<tr><th>Email</th><td>" & strAHEmail & "</span></td></tr>" &  _
									"<tr><th><span style=""font-size:1.5em"">CMS Card Details</span></th><td></td></tr>" & _
									"<tr><th>Card Status</th><td>" & strPMCardStatus & "</span></td></tr>" &  _
									"<tr><th>Expiry Date</th><td>" & strCMSExpiryDate & "</span></td></tr>" &  _
									"<tr><th>Email</th><td>" & strPMEmail & "</span></td></tr>" &  _
									"<tr><th>Name on Card</th><td>" & strPMNameOnCard & "</span></td></tr>" &  _
									"<tr><th>Address 1</th><td>" & strPMAddress1 & "</span></td></tr>" &  _
									"<tr><th>Address 2</th><td>" & strPMAddress2 & "</span></td></tr>" &  _
									"<tr><th>Address 3</th><td>" & strPMAddress3 & "</span></td></tr>" &  _
									"<tr><th>Address 4</th><td>" & strPMAddress4 & "</span></td></tr>" &  _
									"<tr><th>Work Phone</th><td>" & strWork & "</span></td></tr>" &  _
									"<tr><th>Mobile Phone</th><td>" & strMobile & "</span></td></tr>" &  _
									"<tr><th>Credit Limit</th><td>" & strCMSCreditLimit & "</span></td></tr>"
									
				Else
					Response.Write "<tr><th><span style=""font-size:1.5em"">CMS Account Details</span></th><td></td></tr>" & _
									"<tr><th>Own Account Holder?</th><td>" & strOwnHolder & "</td></tr>" & _
									"<tr><th>Account Holder</th><td>" & strCMSAccountHolder & "</span></td></tr>" &  _
									"<tr><th>Account Holder EID</th><td>" & strCMSAHEID & "</span></td></tr>" &  _
									"<tr><th>CMS UserID</th><td>" & strCMSUserName & " " & strTransferButton & "</td></tr>" &  _
									"<tr><th>Active?</th><td>" & strAHActive & "</span></td></tr>" &  _
									"<tr><th>Locked?</th><td>" & strAHLocked & "</span></td></tr>" &  _
									"<tr><th>Phone</th><td>" & strAHPhone & "</span></td></tr>" &  _
									"<tr><th>Email</th><td>" & strAHEmail & "</span></td></tr>" &  _
									"<tr><th><span style=""font-size:1.5em"">CMS Card Details</span></th><td></td></tr>" & _
									"<tr><th>Card Status</th><td>" & strPMCardStatus & "</span></td></tr>" &  _
									"<tr><th>Expiry Date</th><td>" & strCMSExpiryDate & "</span></td></tr>" &  _
									"<tr><th>Email</th><td>" & strPMEmail & "</span></td></tr>" &  _
									"<tr><th>Name on Card</th><td>" & strPMNameOnCard & "</span></td></tr>" &  _
									"<tr><th>Address 1</th><td>" & strPMAddress1 & "</span></td></tr>" &  _
									"<tr><th>Address 2</th><td>" & strPMAddress2 & "</span></td></tr>" &  _
									"<tr><th>Address 3</th><td>" & strPMAddress3 & "</span></td></tr>" &  _
									"<tr><th>Address 4</th><td>" & strPMAddress4 & "</span></td></tr>" &  _
									"<tr><th>Work Phone</th><td>" & strWork & "</span></td></tr>" &  _
									"<tr><th>Mobile Phone</th><td>" & strMobile & "</span></td></tr>" &  _
									"<tr><th>Credit Limit</th><td>" & strCMSCreditLimit & "</span></td></tr>" & _
									"<tr><th>Transaction Limit</th><td>" & strCMSTxnLimit & "</span></td></tr>"
				End If
			End If
		End If
			
			Response.Write "</table></div>" & _
			"<div class=""tab-pane fade"" id=""my-workflow"" role=""tabpanel"" aria-labelledby=""my-workflow-tab"">" &  _
			"<table class=""table"">" &  _
			"<tr><th>Date Submitted</th><td>" & objRS("DateLoaded") & "</td></tr>" &  _
			"<tr><th>Last CS Update</th><td>" & objRS("FileDateTime") & "</td></tr>" &  _
			"<tr><th>Date Issued</th><td Title=""" & dteDateSinceIssued & " Days since card was issued"">" & dteDateIssued & "</td></tr>" &  _
			"<tr><th>File Seq Num</th><td>" & objRS("FileSeqNum") & "</td></tr>" &  _
			"<tr><th>Notes</th><td>" & objRS("Notes") & "</td></tr>" &  _
			"</table>" &  _
			"</div>" & strErrors & "</div></div></div><input type=""hidden"" id=""CardNoMod"" name=""CardNoMod"" value=""" & objRS("CardNumber") & """>"

	End If
	
				
objRS.Close

Response.Write "<button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='Cards.asp'"";><i class=""fa fa-times""></i> Close</button>" & _
	"<button type=""button"" class=""btn btn-outline-primary btn-sm"" title=""Close and return to Card List with " & strNameOnCard & " selected"" onclick=""self.location='Cards.asp?EmployeeID=" & strEmployeeID & "'"";><i class=""fa fa-times""></i> Close Employee</button>"

End Sub


Public Sub LoadCMSDetails
'Procedure to load ProMaster details from PrpoMaster (live link) not yet completed....6th April 2021....
on error resume next

Dim objCon2
Dim strSQL
Dim strCMSFilePath

	'ProMaster Connection details
	Set objCon2 = Server.CreateObject("ADODB.Connection")
	
	'Get the File Path for the CMS UDL from System Parameters
	If IsEmpty(Session("DBConnection2")) Then
		strCMSFilePath = GetSystemAdmin("CMSServerFilePath")
		
		Session("DBConnection2") = "File Name=" & strCMSFilePath & ";"
		'Session("DBConnection2") = "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	End If
	
	'Session("DBConnection2") = "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	objCon2.ConnectionTimeout=2
	objCon2.Open Session("DBConnection2")
	
	'response.write "state=" & objCon2.State
on error goto 0

	
'If there is no connection to ProMaster (Card Management System) then do not try to use the connection
	If objCon2.State = 1 Then
		'Open a recordset in the ProMaster (CMS) database to check the Employee has a CMS Account

		
		strSQL = "SELECT procharge_user.user_name AS UserName, procharge_user.employee_id AS EmplID, card_account.account_ref_no AS AccountHolder, " & _
			"REVERSE(PARSENAME(REPLACE(REVERSE(decode_data.data_line), CHAR(9), '.'), 1)) AS [User], REVERSE(PARSENAME(REPLACE(REVERSE(decode_data.data_line), CHAR(9), '.'), 2)) AS CardType, " & _
            "REVERSE(PARSENAME(REPLACE(REVERSE(decode_data.data_line), CHAR(9), '.'), 3)) AS CardNo, REVERSE(PARSENAME(REPLACE(REVERSE(decode_data.data_line), CHAR(9), '.'), 4)) " & _
            "AS UserName2, card_account.name_on_card, card_account.address_line1, card_account.address_line2, card_account.address_line3, card_account.city, card_account.state, " & _
            "card_account.postal_code, card_account.card_status, card_account.expiry_date, card_account.monthly_credit_limit, card_account.auto_approval_ind, card_account.auto_approval_value, " & _
            "card_account.card_account_number, decode_data.user_name, procharge_user.email_address, payment_cards.card_status AS CardStatus " & _
		"FROM card_account INNER JOIN " & _
             "decode_data ON card_account.account_ref_no = REVERSE(PARSENAME(REPLACE(REVERSE(decode_data.data_line), CHAR(9), '.'), 4)) LEFT OUTER JOIN " & _
             "payment_cards ON payment_cards.card_type = card_account.card_type AND payment_cards.card_account_number = card_account.card_account_number LEFT OUTER JOIN " & _
              "procharge_user ON card_account.user_name = procharge_user.user_name " & _
		"WHERE procharge_user.employee_id= '" & strEmployeeID & "'  AND RIGHT(payment_cards.masked_pan, 4) = '" & strLastFour & "'"
		
		'response.write strSQL
		
		objRS1.Open strSQL,objCon2
		
			If objRS1.EOF Then
				'strMessage = Session("ApplicationEmployeeID") & " has no active CMS Account"
				strCMSUserName = "No data in CMS for " & strEmployeeID & " and card last 4: " & strLastFour
				strCMSAccountHolder = ""
				strCMSAHEID = ""'strEmployeeID""'Session("ApplicationEmployeeID")
				strCMSLocation = "No CMS details"
				strCMSAdminCentre = "Or CMS may be unavailable"
				strCMSSupervisor = ""
			Else
				'strMessage = "CMS Account for " & Session("ApplicationEmployeeID") & ": " & objRS("user_name")
				strCMSUserName = objRS1("UserName")
				strCMSAccountHolder = objRS1("AccountHolder")
				strCMSAHEID = objRS1("EmplID")
				strCMSLocation = objRS1("EmplID")
				strCMSAdminCentre = objRS1("EmplID")
				strCMSSupervisor = objRS1("EmplID")
				
				strPMNameOnCard = objRS1("name_on_card")
				strPMCardStatus = objRS1("CardStatus")
				strPMAddress1 = objRS1("address_line1")
				strPMAddress2 = objRS1("address_line2")
				strPMAddress3 = objRS1("address_line3")
				strPMStatePostCode = objRS1("city") & " " & objRS1("state") & " " & objRS1("postal_code")
				strPMEmail = objRS1("email_address")
				
			End If
			
		objRS1.Close
		
	Else
		strMessage = "CMS database currently unavailable, please try again in 1 hour"
		strCMSUserName = ""
	End If	

End Sub

Public Sub LoadCMSDetailsLocal
'Procedure to load ProMaster details from the CAPS database (rather than a live link to ProMaster - which is in procedure LoadCMSDetails() )
Dim strSQL

		
		strSQL = "SELECT * FROM qryCAPSProMasterAccountsUserDecode WHERE [CardAccountNumber] = '" & strCardNo & "'"
		
		objRS1.Open strSQL,objCon
		
			If objRS1.EOF Then
				strCMSUserName = "None"
				strCMSAccountHolder = "None"
				strCMSAHEID = Session("ApplicationEmployeeID")
				strCMSLocation = ""
				strCMSAdminCentre = ""
				strCMSSupervisor = ""
			
			Else
				strCMSUserName = objRS1("UserName")
				strCMSAccountHolder = objRS1("first_name") & " " & objRS1("surname")'objRS1("account_ref_no")
				strCMSAHEID = Trim(objRS1("employee_id"))
				strCMSLocation = objRS1("unit_id")
				strCMSAdminCentre = objRS1("admin_ctr") & " " & objRS1("admin_ctr_name")
				strCMSSupervisor = objRS1("Supervisor")
				
				strAHLocked = objRS1("locked")
				strAHActive = objRS1("active_indicator")

				strAHPhone = objRS1("Work_Phone")
				strWork = objRS1("cardholder work")
				strMobile = objRS1("cardholder mobile")
				strAHEmail = objRS1("email_address")
				strCMSExpiryDate = objRS1("expiry_date")
				strCMSCreditLimit = objRS1("monthly spend limit")
				strCMSTxnLimit = objRS1("transaction limit")
				strPMNameOnCard = objRS1("name on account")
				strPMCardStatus = objRS1("Card_Status")
				strPMAddress1 = objRS1("addr1")
				strPMAddress2 = objRS1("addr2")
				strPMAddress3 = objRS1("addr3")
				strPMAddress4 = objRS1("addr4") 
				'strPMStatePostCode = objRS1("city") & " " & objRS1("state") & " " & objRS1("postal_code")
				strPMEmail = objRS1("cardholder email")
				strPMCardEID = Trim(objRS1("cardholder eid"))
				
			End If
			
		objRS1.Close
	
End Sub




Public Sub LoadMessages()
'Procedure to load any messages relating to the card
Dim strSQL
Dim strPerson
Dim strMessage
Dim strEdit
Dim strDate

	strSQL = "SELECT * FROM qryCAPSMessage WITH(NOLOCK) WHERE [Object] = 'Card' AND [ObjectID] = '" & Session("CardID") & "' ORDER BY DateUpdated DESC"

	objRS.Open strSQL,objCon
	
    Do Until objRS.EOF
		
		If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
			strPerson = ""
		Else
			If objRS("MessageFrom") = Session("UserID") Then
				strPerson = "(You)"
				strEdit = " data-toggle=""modal"" data-target=""#NotesModal"" onClick=""loadMessage('" & objRS("MessageID") & "');"" "
			Else
				If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
					If objRS("MessageFrom") = 0 Then
						strPerson = "(Admin)"
					End If
				End If
				strEdit = " data-toggle=""modal"" data-target=""#NotesModal"" onClick=""loadMessage('" & objRS("MessageID") & "');"" "
			End If
		End If
		
		If IsNull(objRS("MessageDetail")) or objRS("MessageDetail")= "" Then
			strMessage=""
		Else
			strMessage = objRS("MessageDetail")
			strMessage = Replace(strMessage,chr(13),"</BR>")
		End If
		
		'Get the date for display
		If IsNull(objRS("DateUpdated") ) or objRS("DateUpdated")  = "" Then
			strDate = ""
		Else
			If IsDate(objRS("DateUpdated")) Then
				strDate = FormatDateTime(objRS("DateUpdated"),vbShortDate)
			Else
				strDate = objRS("DateUpdated")
			End If
		End If
		
		Response.write "<div class=""panel panel-light col-12""><div class=""panel-header"">" & _
			"<h6 " & strEdit & ">" & objRS("UserFrom") & " " & strPerson & " - " & strDate & "</h6><span class=""panel-subheader"">" & strMessage & "</span></div></div>"

		objRS.Movenext
	Loop
				
objRS.Close

End Sub


Public Sub LoadMessagesEdit(lngMessageID)
'Procedure to load any messages relating to the card
Dim strSQL
Dim strPerson
Dim strMessage
Dim strEdit
Dim strDate

	strSQL = "SELECT * FROM tblCAPSMessage WITH(NOLOCK) WHERE [MessageID] = " & lngMessageID & ""

	objRS.Open strSQL,objCon
	
    Do Until objRS.EOF
		
		If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
			strPerson = ""
		Else
			If objRS("MessageFrom") = Session("UserID") Then
				strPerson = "(You)"
				'strEdit = " data-toggle=""modal"" data-target=""#NotesModal"" "
			Else
				If IsNull(objRS("MessageFrom") ) or objRS("MessageFrom")  = "" Then
					If objRS("MessageFrom") = 0 Then
						strPerson = "(Admin)"
					End If
				End If
				'strEdit = " data-toggle=""modal"" data-target=""#NotesModal"" "
			End If
		End If
		
		If IsNull(objRS("MessageDetail")) or objRS("MessageDetail")= "" Then
			strMessage=""
		Else
			strMessage = objRS("MessageDetail")
			strMessage = Replace(strMessage,chr(13),"</BR>")
		End If
		
		'Get the date for display
		If IsNull(objRS("DateUpdated") ) or objRS("DateUpdated")  = "" Then
			strDate = ""
		Else
			If IsDate(strDate) Then
				strDate = FormateDateTime(objRS("DateUpdated"),vbShortDate)
			Else
				strDate = objRS("DateUpdated")
			End If
		End If
		
		
		Response.write "<div class=""panel panel-light col-12""><div class=""panel-header"">" & _
			"<h6 " & strEdit & ">" & objRS("UserFrom") & " " & strPerson & " " & strDate & "</h6>" & _
			"<textarea rows=""4"" id=""NTNewNotes"" name=""NTNewNotes"" class=""form-control input-md"" value=""" & strMessage & """ ></textarea>" & _
			"<input type=""text"" id=""NTNotesID"" name=""NTNotesID"" value=""" & lngMessageID & """></div></div>"

		objRS.Movenext
	Loop
				
objRS.Close

End Sub

Public Sub LoadAudit()
'Procedure to load any audit records relating to the card
Dim strSQL
Dim strPerson

	strSQL = "SELECT * FROM tblCAPSAuditLog WITH(NOLOCK) WHERE [CardID] = '" & Session("CardID") & "' ORDER By [ChangeDate] DESC"

	objRS.Open strSQL,objCon

	If objRS.EOF THEN
		Response.write "<div class=""panel col-12""><span class=""panel-subheader"" style=""color:grey; font-style:italic;"">No Card history details..</span></div>"
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
		
		Response.write "<div class=""panel col-12""><a href=""AuditLog.asp?CardID=" & objRS("CardID") & """><span class=""panel-subheader"">" & objRS("ChangeDate") & " - " & objRS("ChangeDetails") & "</span></a></div>"

		objRS.Movenext
	Loop
				
objRS.Close
	response.write "</div>"
	
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
strMessageTitle = "Card"

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
			.Parameters("Object") = "Card"
			.Parameters("ObjectID") = Session("CardID")
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

Public Sub LoadButtons()
Dim strSQL
	
	strSQL = "SELECT CardTypeSub FROM qryCAPSCards WITH(NOLOCK) WHERE CardID = '" & Session("CardID") & "'"
	
	objRS.Open strSQL,objCon
	
	If left(objRS("CardTypeSub"),3) = "NAB" Then
		Response.Write "<button type=""button"" class=""btn btn-outline-secondary btn-sm"" data-toggle=""modal"" data-target=""#CMToNABModal"" HREF=""#"" onClick=""loadCMToNAB(" & strEmployeeID & ",'" & strCardNoFull & "')""><i class=""fa fa-eye""></i> View Audit</button></td></tr>"	
	Else
		Response.Write "<button type=""button"" class=""btn btn-outline-secondary btn-sm"" data-toggle=""modal"" data-target=""#CSToDinersModal"" HREF=""#"" onClick=""loadCSToDiners(" & strEmployeeID & ",'" & strCardNoFull & "')""><i class=""fa fa-eye""></i> View Audit</button></td></tr>"	
	End If

objRS.close

End Sub

Public Function GetCardID(strCardNo)
'Procedure to return the CARD ID from teh Card Number psased in
Dim strSQL

	strSQL = "SELECT [CardID] FROM tblCAPSCard WITH(NOLOCK) WHERE [CardNumber] = '" & strCardNo & "' "

	objRS.Open strSQL,objCon

	If objRS.EOF THEN
		GetCardID = 0
	Else
		GetCardID = objRS("CardID")
	End If
		
				
objRS.Close

End Function


Public Sub ChangeNameOnCard(strAppProcess,strTitle,strFirstName,strSurname,strNewName,intID,strCardType)
'Procedure to run a stored procedure which updates the Cardholder Name on Card
Dim intRecord

	'Makes sure that the application is only on hold or awaiting review
	If strAppProcess = "00" or strAppProcess = "" or strAppProcess = "XS" Then
	
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
			.Parameters("TableName") = "tblCAPSCard"
			.Parameters("UpdatedBy") = Session("UserID")
			
			.ActiveConnection = objCon
			
			
					
		End With
					
		objCmd.Execute        
	
		'Return the result of the Save Function.
		 intRecord = objCmd.Parameters.Item("NameChangeIDOutput")  
		 
		If intRecord = -1 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Name On Card " & strTitle & " " & strFirstName & " " & strSurname & " NOT ADDED TO THE CS File as they are already on the CS File for today! <br><br> Remove their existing CS Record in the <a href=""../Admin/CSTransactionsTo.asp?CardType=" & strCardType & "&Link=AD"">CS File Details screen</a> or Change Name tomorrow.</div>"
		ElseIf intRecord = -2 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Application Name On Card NOT updated for " & strTitle & " " & strFirstName & " " & strSurname & " !</div>"
		ElseIf intRecord = 0 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Application Name On Card NOT updated for " & strTitle & " " & strFirstName & " " & strSurname & " !</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Card Name on Card updated to " & strTitle & " " & strFirstName & " " & strSurname & "!</div>"
		End If
	Else
	'If the application is not on hold or awaiting review then display error message
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Only Active or Temporary Lost/Stolen Cards can have a name change. Name On Card NOT updated for " & strTitle & " " & strFirstName & " " & strSurname & " as Card Status is: " & strAppProcess & "!</div>"
	End If
	
End Sub


Public Sub ChangeTransaction(lngCardID,lngLimit, strNameOnCard)
'Procedure to run a stored procedure which updates the Card Transaction Limit
Dim intRecord

	
	With objCmd
	
		.CommandType = 4
		.CommandText = "spCAPSTransactionLimitSave"
		
		.Parameters.Append objCmd.CreateParameter("CardID", adInteger,adParamInput)
		.Parameters.Append objCmd.CreateParameter("NewTransactionLimit", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("TransactionIDOutput", adInteger, adParamOutput)
		
		.Parameters("CardID") = lngCardID
		.Parameters("NewTransactionLimit") = lngLimit 'The new name on card
		.Parameters("UpdatedBy") = Session("UserID")
		
		.ActiveConnection = objCon
				
	End With
				
	objCmd.Execute        

	'Return the result of the Save Function.
	 intRecord = objCmd.Parameters.Item("TransactionIDOutput")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Transaction Limit NOT updated for " & strNameOnCard & " ! Card Not Found.</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">Card Transaction Limit updated to " & FormatCurrency(lngLimit,2) & "!</div>"
	End If
	
End Sub


Public Sub AddContactToCSFile(strCardStatus,strEmpIDCS,strCardType, strCardTypeSub)
'Procedure to run a stored procedure which checks the employee and adds then to the CS File if their contact details are valid AND they have changed (different between tblCAPSCDMCHistory - CDMC and tblCAPSCard - Card)
Dim intRecord

	

	'Makes sure that the application is only on hold or awaiting review
	
	If Left(strCardTypeSub,3) = "NAB" AND (strCardStatus = "" or strCardStatus = "XS") Then
	
		With objCmd
		
			.CommandType = 4
			.CommandText = "spCAPSCDMCProcessCSFileToNAB"
			
			.Parameters.Append objCmd.CreateParameter("UserID", adInteger,adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarchar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("CardType", adVarchar, adParamInput, 20)
			'.Parameters.Append objCmd.CreateParameter("CardTypeSub", adVarchar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("CDMCProcessCSFileToOutput", adVarchar, adParamOutput,200)
			
			.Parameters("UserID") = Session("UserID")
			.Parameters("EmployeeID") = strEmpIDCS
			.Parameters("CardType") = strCardType		
			'.Parameters("CardTypeSub") = strCardTypeSub					
			
			.ActiveConnection = objCon
					
		End With
					
		objCmd.Execute        
	
		'Return the result of the Save Function.
		 intRecord = objCmd.Parameters.Item("CDMCProcessCSFileToOutput")
		 
		

		If IsNumeric(intRecord) Then
			Response.Write "<div class=""alert alert-success"" role=""alert"">Contact Details added to the CS File for " & strEmpIDCS & "! Please check the CS File to confirm. " & intRecord  &"</div>"
		Else
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Card Contact Details NOT updated for " & strEmpIDCS & " ! " & intRecord & "</div>"
		End If
	
	ElseIf Left(strCardTypeSub,3) <> "NAB" AND strCardStatus = "00" Then
	
		With objCmd
		
			.CommandType = 4
			.CommandText = "spCAPSCDMCProcessCSFileTo"
			
			.Parameters.Append objCmd.CreateParameter("UserID", adInteger,adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarchar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("CardType", adVarchar, adParamInput, 20)			
			.Parameters.Append objCmd.CreateParameter("CDMCProcessCSFileToOutput", adVarchar, adParamOutput,200)
			
			.Parameters("UserID") = Session("UserID")
			.Parameters("EmployeeID") = strEmpIDCS
			.Parameters("CardType") = strCardType		
			
			.ActiveConnection = objCon
					
		End With
					
		objCmd.Execute        
	
		'Return the result of the Save Function.
		 intRecord = objCmd.Parameters.Item("CDMCProcessCSFileToOutput")
		 
		 

		If IsNumeric(intRecord) Then
			Response.Write "<div class=""alert alert-success"" role=""alert"">Contact Details added to the CS File for " & strEmpIDCS & "! Please check the CS File to confirm. " & intRecord  &"</div>"
		Else
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Card Contact Details NOT updated for " & strEmpIDCS & " ! " & intRecord & "</div>"
		End If
	Else
	'If the application is not on hold or awaiting review then display error message
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Only Active or Temporary Lost/Stolen Cards can have an Address update on the CS File. Contact details NOT updated for " & strEmpIDCS & "!</div>"
	End If
	
End Sub

Public Sub ChangeBranding(strBrandingNew,lngCardID,strEmpEID)
Dim intRecord
	
		With objCmd
		
			.CommandType = 4
			.CommandText = "spCAPSBrandingChange"
			
			.Parameters.Append objCmd.CreateParameter("CardEmployeeID", adVarChar,adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("CAPSCardID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("NewBranding", adVarchar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("ReviewedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("BrandingApplicationIDOutput", adInteger, adParamOutput)
			
			.Parameters("CardEmployeeID") = strEmpEID
			.Parameters("CAPSCardID") = lngCardID
			.Parameters("NewBranding") = strBrandingNew		
			.Parameters("ReviewedBy") = Session("UserID")					
			
		.ActiveConnection = objCon
					
	End With
					
		objCmd.Execute        
	
		'Return the result of the Branding Change Function.
		 intRecord = objCmd.Parameters.Item("BrandingApplicationIDOutput")

		If IsNumeric(intRecord) Then
			If intRecord = -1 Then
				Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR: Branding Change not actioned due to a system error. Please contact System Administrator and quote Error No.: " & intRecord & ", CardID: " & lngCardID & ", EmployeeID: " & strEmpEID & "</div>"
			ElseIf intRecord = -99 Then
				Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR Branding Change not released to NA file. Please contact System Administrator and quote Error No.: " & intRecord & ", CardID: " & lngCardID & ", EmployeeID: " & strEmpEID & "</div>"
			ElseIf intRecord = -88 Then
				Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR Branding Change requested but original card NOT cancelled. Please contact System Administrator and quote Error No.: " & intRecord & ", CardID: " & lngCardID & ", EmployeeID: " & strEmpEID & "</div>"
			Else
				Response.Write "<div class=""alert alert-success"" role=""alert"">Branding Change request submitted. Reference: ApplicationID " & intRecord & "</div>"
			End If
		End If
End Sub
	

Set objRS = Nothing
Set objCon = Nothing

%>