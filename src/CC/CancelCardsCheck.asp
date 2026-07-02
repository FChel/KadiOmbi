
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
Dim objRS
Dim objCmd

Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL
Public strW
Dim lngCardID
Dim strEmployeeID
Dim strTitle
Dim strFirstName
Dim strLastName
Dim strAddress1
Dim strAddress2
Dim strAddress3
Dim strAddress4
Dim strSUburb
Dim strState
Dim strPostCode
Dim dteDateReceived
Dim strStatus
Dim strReviewedBy
Dim dteDateReviewed
Dim lngCreditLimit
Dim strLastRun



    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	
	
	objRS.Open "SELECT Top 1 DateUpdated FROM tblCAPSCancelCardCheck",objCon
	
		If objRS.EOF Then
			strlastRun = ""
		Else
			strLastRun = objRS(0)
		End If
		
	objRS.Close

    Session("CurrentPage") = "CC/ApplicationsEmployeeHF.asp"

	If IsNull(Session("CardID")) OR Session("CardID") = "" Then Session("CardID")= 0

	If isNull(Session("CardID")) Or Session("CardID") = "" Then 
		Session("CardID") = 0
	End If
	
	If Not IsEmpty(Request.QueryString("UserView")) Then
		Session("UserView") = Request.QueryString("UserView")
	End If

	If Not IsEmpty(Request.QueryString("ViewButton")) Then
		Session("ViewButton") = Request.QueryString("ViewButton")
	
	Else
		Session("ViewButton") = "All"
	End If
	
	If Not IsEmpty(Request.QueryString("Action")) Then
		If Request.QueryString("Action") = "Check" Then
			Call CheckCards()
		End If
		
	End If
	
  Call LoadDetails()
  
  If IsNull(Session("CardType")) OR Session("CardType") = "" Then Session("CardType") = "DTC - Diners"
%>

<html>
<head>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

	<!--<link rel="stylesheet" type="text/css" href="../CAPSStyle.css">-->
	<!--<script src="../assets/node_modules/jquery/jquery-3.2.1.min.js"></script>-->
	  <!-- Custom fonts for this template-->
  <!--<link href="../vendor/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css">-->
<script LANGUAGE="javascript">

function SaveData(){
	var varSubmit = true
		frm.msgbox.value='Saving.......';
		frm.submit();
	//}
}

function loadCancel(cb) {

var id = cb.getAttribute('data-id');
var name = cb.getAttribute('data-name');
var cardno = cb.getAttribute('data-cardno');


	document.getElementById("ModalCancelTitle").innerHTML = 'Do you wish to cancel the selected card - ' + id + '?';
	document.getElementById("ModalCancel").style.display = "block";
	document.getElementById("CardNo").value = cardno;
	document.getElementById("CardHolder").value = name;
	document.getElementById("CardID").value = id;
	
}

function ProcessCheck() {
	
	var id = document.getElementById('CardID').value
	self.location = "CancelCardsCheck.asp?Action=Check";

}

function ExcelExport(sql) {
	
	alert(sql);
	window.open('ExcelExport.asp?sql=SQL');

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
      
      $.post('ApplicationsEmployeeHF.asp?Action=SubmitApp', 
         $('#frm').serialize(), 
      /*   function(data, status, xhr){
           // do something here with response;
         });
      */
    });
});

</script>
<style>

/* Bootstrap 4 text input with search icon */

.has-search .form-control {
    padding-left: 2.375rem;
}

.has-search .form-control-feedback {
    position: absolute;
    z-index: 2;
    display: block;
    width: 2.375rem;
    height: 2.375rem;
    line-height: 2.375rem;
    text-align: center;
    pointer-events: none;
    color: #aaa;
}
</style>
</head>
<body >
<main class="main py-3">
    <div class="container">	
			   
<!-- Confirm Cancel Card Modal -->
<div class="modal fade" id="ModalCancel" tabindex="-1" role="dialog" aria-labelledby="ModalCancel" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalCancelTitle" style="font-weight:bold;">Delete Card?</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
	  <div class="modal-body" id="ErrorDetailsMod">
			<h6 style="color:navy;">Card No:&nbsp;&nbsp;</h6><input type="text" class="form-control" id="CardNo" name="CardNo"><BR>
			<h6 style="color:navy;">Name:&nbsp;&nbsp;</h6><input type="text" class="form-control" id="CardHolder" name="CardHolder">
			<input type="text" class="form-control" id="CardID" name="CardID">
			
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
		<button type="button" class="btn btn-primary" onClick="CancelCard()" ><i class="fa fa-check"></i> Yes</button>
		<input type="hidden" id="NewStatus" name="NewStatus" value=""/>
      </div>
    </div>
  </div>
</div>
<!-- End Cancel Card Modal -->

<!-- End the first part of the Header Container -->
<div id='tbl-container'>

	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-10">
				<h2 class="text-left">Cancel Cards Check <% If Session("UserView") = "User" Then Response.write " for " & Session("UserName")%></h2>
			</div>
		</div>

          <div class="row py-2">
            <div class="col-md-5">
              <%Call LoadViewButtons()%>
            </div>
			
				<div class="row col-3" style="text-align:right;">
				<button type="button" class="btn btn-outline-secondary btn-sm" onclick="ProcessCheck();" Title="Click to Process the current CDMC File against the Current Diners Card list to add changes to the CS File To Diners"><i class="fa fa-cogs"></i> Run Cancel Card Check</button>

				</div>
				<div class="row col-4" style="text-align:right;"><h5>Last Run : <%=strLastRun%></h5>
					
				</div>
          </div>

      </section>
	  
	  
	 <section class="table py-2">
        <div class="container">
          <div class="row">
            <div class="col-12">
              <table class="table table-compact text-left">
                <thead>
                  <tr>
					<th scope="col">Card ID</th>
					<th scope="col">EID</th>
					<th scope="col">Name</th>
					<th scope="col">Card No.</th>
					<th scope="col">Card Type</th>
					<th scope="col">Card Status</th>
					<th scope="col">Countdown</th>
					<th scope="col">On CDMC</th>
					<th scope="col">Active Employee</th>
					<th scope="col" style="width:200px;">Error Status</th>
                  </tr>
                </thead>
                <tbody>
                 <%
        
				DisplayTableDetails()				
        
				%>	
                </tbody>
              </table>
            </div>
          </div>
        </div>
		
		
	  </section>
</div>


<!--</DIV>-->

</div>

</main>	

    <!-- jQuery -->
    <script src="../js/jquery.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="../js/bootstrap.min.js"></script>
	
<!-- #Include file=CAPSFooter.asp -->

</body>
</html>
<%

Public Sub DisplayTableDetails()
Dim y

Dim strAction
Dim strStatus
Dim strAddress
Dim dteDateSubmitted
Dim dteDateReviewed
Dim strSearch
Dim strRecordMessage
Dim strSort
Dim strWhere
Dim strName

Dim arrNames
Dim strFNameSearch
Dim strLNameSearch
Dim strCardNoSearch

Dim strCardNo
Dim strProcessAction
Dim strProcessStatus
Dim strCardType 
Dim strCardTypeSub
Dim strNameOnCard
Dim intCountDown
Dim strOnCDMC
Dim strActiveEmployee
Dim strErrorStatus
Dim strErrorPill
Dim strDoNotCancel
	
	If Session("ViewButton") = "Cancelled" Then
		strWhere = " AND [ProcessStatus] = 'Exported' "
	ElseIf Session("ViewButton") = "All" Then
		strWhere = " AND [ProcessStatus] = 'Cancelled' "
	Else
		'This catches ALL
		strWhere = ""
	End If

	If Session("ViewButton") = "All" Then		
		
		strSQL = "SELECT * FROM qryCAPSCancelCardCheck WITH(NOLOCK) Order By RemoveCountDown" ' WHERE " & strWhere & strSort
			
	Else
		'strSQL = "SELECT * FROM qryCAPSCards WITH(NOLOCK) WHERE CardType = 'DPC' AND EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
		strSQL = "SELECT * FROM qryCAPSCancelCardCheck WITH(NOLOCK) Order By RemoveCountDown" ' WHERE " & strWhere & strSort
	End If


	
'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = "for " & strSearch
Else
	strRecordMessage = "for " & Session("UserName") 
End If



objRS.Open strSQL,objCon

    y = 0
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center; color:red;"">No Cards " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	End If
    	
    Do until objRS.EOF 
	
		strAction = ""
		strStatus = ""
						
		'Set the Card Type for use below
		If IsNull(objRS("CardType")) or objRS("CardType") = "" Then
			strCardType = ""
		Else
			strCardType = objRS("CardType")
		End If	

		If IsNull(objRS("CardTypeSub")) or objRS("CardTypeSub") = "" Then
			strCardTypeSub = ""
		Else
			strCardTypeSub = objRS("CardTypeSub")
		End If	

		If IsNull(objRS("NameOnCard")) or objRS("NameOnCard") = "" Then
			strNameOnCard = ""
		Else
			strNameOnCard = objRS("NameOnCard")
		End If

		If IsNull(objRS("Status")) or objRS("Status") = "" Then
			strStatus = ""
		Else
			strStatus = objRS("Status")
		End If		
		
		If IsNull(objRS("CardNumber")) Then
			strCardNo = Null
		Else
			strCardNo = MaskCard(objRS("CardNumber"))
		End If

		If IsNull(objRS("RemoveCountDown")) Then
			intCountDown = ""
		Else
			intCountDown = objRS("RemoveCountDown")
		End If

		If IsNull(objRS("OnCDMCLast5")) Then
			strOnCDMC = ""
		Else
			strOnCDMC = objRS("OnCDMCLast5")
		End If

		If IsNull(objRS("ActiveEmployee")) Then
			strActiveEmployee = ""
		Else
			strActiveEmployee = objRS("ActiveEmployee")
		End If
		
		If IsNull(objRS("DoNotCancel")) Then
			strDoNotCancel = ""
		Else
			strDoNotCancel = objRS("DoNotCancel")
		End If
		
		strErrorStatus = "OK"

		If IsNumeric(intCountDown) Then
			If intCountDown < 1 AND strONCDMC <> "" Then
				strErrorStatus = "On CDMC"
			End If
			If intCountDown < 1 AND strActiveEmployee = "Y" Then
				strErrorStatus = strErrorStatus & " ; Employee should be inactive"
			End If
		End If

		If intCountDown = "" AND strONCDMC = "" AND strActiveEmployee = "" Then
			strErrorStatus = "NO CDMC History record"
		End If
		
		If strErrorStatus = "OK" Then	
			If intCountDown = "0" or  intCountDown = "" Then
				strErrorPill = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">Will be cancelled today</span>"
			Else
				strErrorPill = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">Will be cancelled in " & intCountDown & " Days</span>"
			End If
		Else	
			strErrorPill = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">" & strErrorStatus & "</span>"
		End If
		
		If strDoNotCancel = "Y" Then
		
			strErrorPill = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">DO NOT CANCEL</span>"
			
		End If
		
		If Session("ViewButton") = "All" Then
		
			Response.write "<TR><TD ><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD Style=""font-size:14px;""><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """>" & strNameOnCard & "</a></TD><TD >" & strCardNo & "</TD>" & _
					"<TD Style=""font-size:14px;""><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """>" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
					"<TD Style=""font-size:14px; text-align:center;"">" & strStatus & "</TD><TD Style=""font-size:14px; text-align:center;"">" & intCountDown & "</TD>" & _
					"<TD Style=""font-size:14px; text-align:center;"">" & strOnCDMC & "</TD><TD Style=""font-size:14px; text-align:center;"">" & strActiveEmployee & "</TD>" & _
					"<TD Style=""font-size:14px; text-align:center; font-color=red; "">" & strErrorPill & "</TD></TR>"
					y = y + 1
			Else
				If strErrorStatus <> "OK" Then
					Response.write "<TR><TD ><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
						"<TD Style=""font-size:14px;""><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """>" & strNameOnCard & "</a></TD><TD >" & strCardNo & "</TD>" & _
						"<TD Style=""font-size:14px;""><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """>" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
						"<TD Style=""font-size:14px; text-align:center;"">" & strStatus & "</TD><TD Style=""font-size:14px; text-align:center;"">" & intCountDown & "</TD>" & _
						"<TD Style=""font-size:14px; text-align:center;"">" & strOnCDMC & "</TD><TD Style=""font-size:14px; text-align:center;"">" & strActiveEmployee & "</TD>" & _
						"<TD Style=""font-size:14px; text-align:center; font-color=red; "">" & strErrorPill & "</TD></TR>"
					y = y + 1
				End If
			End If		
			
			
			
		objRS.movenext
	Loop
	
	
	Response.Write "<TR><TH colspan=""10"">Total</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;"">" & y & "</TH></TR>"
				
objRS.Close

End Sub


Sub LoadDetails()

   'Description:	Loads Position details into page if applicable.
	objRS.Open "SELECT * FROM tblCAPSCDMC WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "'",objCon

		If Not objRS.EOF Then
		   
			'lngCardID = objRS("CardID")
			strEmployeeID = objRS("EmployeeID")
			strTitle = objRS("Title")
			strFirstName = objRS("FirstName")
			strLastName  = objRS("Surname")
			strAddress1 = objRS("Addressline1")
			strAddress2 = objRS("Addressline2")
			strAddress3 = objRS("Addressline3")
			'strAddress4 = objRS("Address4")
			strSuburb = objRS("Postaladdress_City")
			strState = objRS("Postaladdress_State")
			strPostCode = objRS("Postaladdress_PostCode")
			'dteDateReceived = objRS("DateReceived")
			'strStatus = objRS("Status")
			'strReviewedBy = objRS("ReviewedBy")
			'dteDateReviewed = objRS("DateReviewed")
			'If IsNull(objRS("CreditLimit")) or objRS("CreditLimit") = "" then
				lngCreditLimit = 30000
			'Else
			'	lngCreditLimit = objRS("CreditLimit") 
			'End If
		Else
			Session("CardID") = 0
			lngCardID = 0'objRS("CardID")
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
	   End If

	objRS.Close
	
End Sub

Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(5)

If Session("ViewButton") = "Cancelled" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "Mastercard" Then
	arrButton(3) = "active"
ElseIf Session("ViewButton") = "ANZ" Then
	arrButton(4) = "active"
ElseIf Session("ViewButton") = "CTS" Then
	arrButton(5) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='CancelCardsCheck.asp?ViewButton=All';""><i class=""fa fa-folder""></i> To be Cancelled</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='CancelCardsCheck.asp?ViewButton=Cancelled';""><i class=""fa fa-plane""></i> Cards with Errors</button>" & _
				"</div>"

End Sub

Public Sub CheckCards()

'Procedure to check all cards that are to be cancelled	
		
			objCon.Execute "spCAPSRunCancelCardCheck " & Session("UserID") & ""
		
		   
			'objCmd.Execute		

			Response.Write "<div class=""alert alert-success"" role=""alert"">Cancel Card check completed.</div>"

	
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

Set objRS = Nothing
Set objCon = Nothing
%>
