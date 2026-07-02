
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

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	

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
		Session("CardTypeNameView") = ""
	End If
	
	If Not IsEmpty(Request.QueryString("StatusViewButton")) Then
		Session("StatusViewButton") = Request.QueryString("StatusViewButton")
	End If
	
	If Not IsEmpty(Request.QueryString("CardTypeNameView")) Then
		Session("CardTypeNameView") = Request.QueryString("CardTypeNameView")
		Session("ViewButton") = ""
		
	End If
	
	If Not IsEmpty(Request.QueryString("Action")) Then
		If Request.QueryString("Action") = "Cancel" Then
			Call CancelApplication()
		End If
		
		If Request.QueryString("Action") = "Release" Then
			Call ReleaseApplication()
		End If
	End If

  
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "SubmitApp" Then
			'Response.Write "CPID=" & Session("CarParkingID")
			'Session("CardID") = 0
			Call SubmitApplication()
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

function LoadDeletModal(cb) {

	var id = cb.getAttribute('data-CID');
	document.getElementById('DeleteCardID').value=id;

	var CardID = cb.getAttribute('data-CCardID');
	document.getElementById('DeleteCardID').value=CardID;
	
	var EmpID = cb.getAttribute('data-CEmpID');
	document.getElementById('DeleteEmpID').value=EmpID;

	var CardName = cb.getAttribute('data-CNameOnCard');
	document.getElementById('DeleteNameOnCard').value=CardName;
	
	var CardStat = cb.getAttribute('data-CStatus');
	document.getElementById('DeleteStatus').innerHTML=CardStat;

	var CardType = cb.getAttribute('data-CCardType');
	document.getElementById('DeleteType').value=CardType;
	document.getElementById('ModalDeleteHeader').innerHTML=CardType;
	
	var CardMessage = cb.getAttribute('data-CMessage');
	document.getElementById('ModalDeleteHeader').innerHTML=CardMessage;
	
	var CardNo = cb.getAttribute('data-CCardNo');
	document.getElementById('DeleteCardNo').value=CardNo;
	
	
	
}

function DeleteCard() {

	var CardID = document.getElementById('DeleteCardID').value;

	var CardName = document.getElementById('DeleteNameOnCard').value;
	
	var EmpID = document.getElementById('DeleteEmpID').value;
	
	//alert(CardID);
	
	self.location='Cards.asp?Link=CD&Action=Cancel&CardID=' + CardID + "&Name=" + CardName + "&EmployeeID=" + EmpID;
	//self.location='Cards.asp?Link=CD&Action=Cancel&CardID=" & objrs("CardID") & "&Name=" & Replace(objRS("FirstName"),"'","") & " " & Replace(objRS("Surname"),"'","") & "&EmployeeID=" & objRS("EmployeeID") & "'"";
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

function loadCSFromDiners(varID,varCardNo) {

//alert(varCardNo)
//varCardNo=document.getElementById("CardNoMod").value
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
  //document.getElementById("compareModalLabel").innerHTML = '<button type="button" class="btn btn-outline-secondary" title="Displaying CS To Diners. Click to View CS From Diners." onClick="loadCSToDiners('+varID+',' + varCardNo +');">CS From Diners</button>'
  document.getElementById("ModalCancelCSDetails").innerHTML = '<img src="../images/Load.gif" style="vertical-align:middle;" /> '
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ModalCancelCSDetails").innerHTML = this.responseText;
    }
  };
	
  xhttp.open("GET", "../CC/AJAX/GetCSFromDinersCancel.asp?EmployeeID=" + varID + "&CardNo=" + varCardNo + "", true);
  xhttp.send();
}


function OpenSs(cb) {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("CardTypeSelect");
	var result = e.options[e.selectedIndex].value;
	
	document.getElementById('CardType').value=result;
	
}

function LoadFiles() {

  //var id = cb.getAttribute('data-id');
  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ModalFileMessage").innerHTML = this.responseText;
    }
  };

  xhttp.open("GET", "AJAX/GetExcelFiles2.asp?FileStart=Cards", true);
  xhttp.send();
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
	  
	<!--Loading Wait Spinner-->
	<div class="modal fade bd-example-modal-lg modalWait" id="ModalWait" data-backdrop="static" data-keyboard="false" tabindex="-1">
    <div class="modal-dialog modal-sm">
        <div class="modal-content" style="width: 88px">
            <span style="color:black;" class="spinner-border spinner-border-lg"></span>
        </div>
    </div>
</div>
				<div id="wait" style="display: none;position: absolute;width: 350;height: 100;margin-left: 300;margin-top: 150;background-color: #FFFFFF; text-align: center; color:#333366; line-height:80px; vertical-align:middle; border: solid 1px #333366;">
               <img src="images/Load.gif" style="vertical-align:middle;" /> &nbsp;&nbsp;Please wait while loading...</div>


<!-- Cancelled Modal -->
<div class="modal fade" id="ModalCancelCS" tabindex="-1" role="dialog" aria-labelledby="ModalRelease" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalCancelCSTitle" style="font-weight:bold;">CS From Diners</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
	  <div class="modal-header">
	  <h6 style="color:navy;">Card: <%=Session("ApplicationName")%></h6>
	  </div>
      <div class="modal-body" id="ModalCancelCSDetails">
        
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
		
      </div>
    </div>
  </div>
</div>
<!-- End Select Batch Number Modal -->

			   
<!-- Confirm Delete Modal -->
<div class="modal fade" id="ModalDelete" tabindex="-1" role="dialog" aria-labelledby="ModalRelease" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalDeleteTitle" style="font-weight:bold;">Delete Card?</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
	  <div class="modal-header">
	  <h6 id="ModalDeleteHeader" style="color:navy;">Card:</h6>
	  </div>
      <div class="modal-body" id="ModalDeleteBody">
        
		<table class="table table-hover">
			<tr><td style="font-weight:bold;">Deleting Card for:</td><td><input id="DeleteNameOnCard" name = "DeleteNameOnCard" value="" style="border: 0;" class="form-control"/></td></tr>
			<tr><td style="font-weight:bold;">Employee ID:</td><td><input id="DeleteEmpID" name = "DeleteEmpID" value="" style="border: 0;" DISABLED class="form-control"/></td></tr>
           <tr><td style="font-weight:bold;">Card No:</td><td><input id="DeleteCardNo" name = "DeleteCardNo" value="" style="border: 0;" DISABLED class="form-control"/></td></tr>
		   <tr><td style="font-weight:bold;">Card Status:</td><td id="DeleteStatus"></td></tr>
		   <!--<tr><td style="font-weight:bold;">Card Status:</td><td><input id="DeleteStatus" name = "DeleteStatus" value="" style="border: 0;" DISABLED class="form-control" /></td></tr>-->
		   <tr><td style="font-weight:bold;">Card Type:</td><td><input id="DeleteType" name = "DeleteType" value="" style="border: 0;" DISABLED class="form-control" /></td></tr>
		   <tr><td><input id="DeleteCardID" name = "DeleteCardID" value="" HIDDEN /></td></tr>
			</table>
			
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
		<button type="button" class="btn btn-primary" onClick='DeleteCard();' ><i class="fa fa-check"></i> Yes</button>
		<input type="hidden" id="NewStatus" name="NewStatus" value=""/>
      </div>
    </div>
  </div>
</div>
<!-- End Select Batch Number Modal -->

<!-- Start Download Excel Modal -->
	<div class="modal fade" id="ModalExcel" tabindex="-1" role="dialog" aria-labelledby="ModalDeleteCenterTitle" aria-hidden="true">
		<div class="modal-dialog modal-dialog-centered" role="document">
		  	<div class="modal-content">
				<div class="modal-header">
			  		<h5 class="modal-title" id="ModalDeleteLongTitle" style="font-weight:bold;">Download Excel file</h5>
			  		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				</div>
				<div class="modal-body">			  
				  	<div class="row">
						<div class="col-md-12 mb-3">			
						  		<div id="ModalFileMessage">
								<span id="Progress" style="padding-left:20px; padding-bottom:20px;"><img src="../images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>
								</div>
						</div>
				  	</div>
				  	<div class="row">
					  	<div class="col-md-12 mb-3" style="text-align:right;">
						  	<button type="button" class="btn btn-secondary btn-sm" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
					  	</div>
				  	</div>
			  	</div>
			</div>	
		</div>
		<div class="modal-footer"></div>
	</div>
	<!-- End Download Excel Modal -->

<!-- End the first part of the Header Container -->
<div id='tbl-container'>
  <form action="Cards.asp?Action=Search" method="POST" id="frm" name="frm">
	<div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-8">
				<h2 class="text-left">All Cards <% If Session("UserView") = "User" Then Response.write " for " & Session("UserName")%></h2>
			</div>
			<div class="col-md-4 float-right">
				<button type="button" class="btn btn-outline-success" data-toggle="modal" data-target="#ModalExcel" onClick="LoadFiles();" title="Click to Export Current Card Search to Excel"><i class="fa fa-file-excel"></i> Export To Excel</button>
				<%Call LoadStatusButtons()%>
				<!--<button type="button" class="btn btn-primary float-right" onClick='window.location="ApplicationsSubmit.asp"'><i class="fa fa-plus"></i> New Application</button>-->
			</div>
			
		</div>

          <div class="row py-2">
            <div class="col-md-9">
              <%Call LoadViewButtons()%>
            </div>
			<div class="col-md-3">
				<div class="form-group has-search" title="Enter Search then press Enter key. (Will look for EID or First or Last Name). C: before search will find any cards with numbers entered. E: will find cards ending with digits entered. F: First name only. L: Last name only. N: Name On Card">
					<span class="fa fa-search form-control-feedback" onClick="frm.submit();"></span>
				 <input type="text" class="form-control" type="search" id="SearchInput" name="SearchInput" placeholder="Enter EID, First and/or Last Name" value="<%=Request.Form("SearchInput")%>"/>
				 </div>
			</div>
          </div>

      </section>
	  
	  
	 <section class="table py-2">
        <div class="container">
          <div class="row">
            <div class="col-12">
              <table class="table table-compact text-left">
                <thead style="text-align: center;">
                  <tr>
					<th scope="col" style="font-size:14px;">Card ID</th>
					<th scope="col" style="font-size:14px;">EID</th>
					<th scope="col" style="font-size:14px;">Name</th>
					<th scope="col" style="font-size:14px;">Card No.</th>
					<th scope="col" style="font-size:14px;">Card Type</th>
					<th scope="col" style="font-size:14px;">Address</th>
					<th scope="col" style="font-size:14px;">Card Status</th>
					<th scope="col" style="font-size:14px;">Date Issued</th>
					<th scope="col" style="font-size:14px;">PM Load Date</th>
					<th scope="col" style="font-size:14px;">Action</th>
					<th scope="col" style="width:20px; font-size:14px;">Process</th>
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
</form>
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
Dim strCardType 
Dim strDeleteMessage
Dim strCardTypeDelete
Dim strCardNoDelete
Dim strNameonCardMod

	strSearch = Trim(Request.Form("SearchInput"))
	
	'If an employee has been selected then re-select then after delete or release actions
	If Not IsEmpty(Request.QueryString("EmployeeID")) Then
		strSearch = Trim(Request.QueryString("EmployeeID"))
	End If
	
	If IsEmpty(Request.QueryString("SortType")) Then
		'strOrderType = "ASC"
	Else
		If Request.QueryString("SortType") = "ASC" Then
			strOrderType = "DESC"
		Else
			strOrderType = "ASC"
		End If
	End If
	
	If IsEmpty(Request.QueryString("Sort")) Then
		strSort = "ORDER BY [CardID] DESC"
	Else		
		strSort = " ORDER BY " & Request.QueryString("Sort") & " " & strOrderType
	End If
	
	'If Session("ViewButton") = "Diners" Then
	'	strWhere = " AND [CardTypeSub] = 'Diners' "
	'ElseIf Session("ViewButton") = "Mastercard" Then
	'	strWhere = " AND [CardTypeSub] = 'Mastercard' "
	'ElseIf Session("ViewButton") = "ANZ" Then
	'	strWhere = " AND [CardTypeSub] = 'ANZ' "
	'ElseIf Session("ViewButton") = "CTS" Then
	'	strWhere = " AND [CardTypeSub] = 'CTS' "
	'ElseIf Session("ViewButton") = "DINERS" Then
	'	strWhere = " AND [CardType] = 'DPC' "
	'Else
		'This catches ALL
	'	strWhere = ""
	'End If
	
	'Changed from teh above for the new list of the possible card types based on new NAB card types
	If Session("ViewButton") = "NABDTC" Then
		strWhere = " AND [CardType] = 'DTC' AND Left([CardTypeSub],3) = 'NAB' "
	ElseIf Session("ViewButton") = "NABDPC" Then
		strWhere = " AND [CardType] = 'DPC' AND Left([CardTypeSub],3) = 'NAB' "
	ElseIf Session("ViewButton") = "CTS" Then
		strWhere = " AND [CardTypeSub] = 'CTS' "
	Else
		'This catches ALL
		strWhere = ""
	End If

	'If a Card Type and Card Type Sub has been selected from teh drop down lst of all card types then build the where for card type and card type sub which overwrites the above list for a specific button
	If Not IsEmpty(Request.QueryString("CardTypeNameView")) Then
		strWhere = " AND [CardType] = '" & Left(Request.QueryString("CardTypeNameView"),3) & "' AND [CardTypeSub] = '" & Right(Request.QueryString("CardTypeNameView"),Len(Request.QueryString("CardTypeNameView"))-3) & "'"
	End If

	If Session("StatusViewButton") = "Active" Then
		strWhere = strWhere & " AND ([Status] = '00' OR [Status]= '' OR [Status]= 'T')"
	ElseIf Session("StatusViewButton") = "Cancelled" Then
		strWhere = strWhere & " AND ([Status] = '01' OR [Status]= '02' OR [Status]= '03' OR [Status]= 'L' OR [Status]= 'C' OR [Status]= 'S' OR [Status] = 'VX')"
	ElseIf Session("StatusViewButton") = "TempHold" Then
		strWhere = strWhere & " AND ([Status]= 'T' OR [Status] = 'XS' OR [Status] = 'SS')"
	Else
		'This catches ALL
		'strWhere = ""
	End If
	
If strSearch = "" OR ISNull(strSearch) Then
	If Session("UserView") = "All" Then
		strSQL = "SELECT TOP 500 * FROM qryCAPSCards WITH(NOLOCK) WHERE [CardID] > 0 " & strWhere & strSort
	Else
		strSQL = "SELECT TOP 500 * FROM qryCAPSCards WITH(NOLOCK) WHERE EmployeeID = '" & Trim(Session("EmployeeID")) & "'" & strWhere & strSort
	End If
	
	'Response.Write strSQL
	
Else
	If Session("UserView") = "All" Then

		'Replace all single inverted commas with 2 for SQL
		strSearch = Replace(strSearch,"'","''")
		
		'If the user has entered a search term with a space the assume this is a first and last name so search on that only
		If Instr(1,strSearch," ")>0 Then
			arrNames = Split(strSearch," ")
			strFNameSearch = arrNames(0)
			strLNameSearch = arrNames(1)
			
			'strWhere = " WHERE ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%')"
			strSQL = "SELECT * FROM qryCAPSCards WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%'))" & strWhere & strSort
		Else
		
		strSQL = "SELECT TOP 500 * FROM qryCAPSCards WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
		End If
	Else
		strSQL = "SELECT TOP 500 * FROM qryCAPSCards WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
	End If
End If
'Response.Write strSQL


	'If the user has entered the Card Number lookup then process this, otherwise perform the EmployeeID and Name searches
	If UCASE(Left(strSearch,2)) = "C:" Then
	
		strCardNoSearch = Right(strSearch,Len(strSearch)-2)
		
		'strWhere = strWhere & " AND [CardNumberShort] like '%" & strCardNoSearch & "%'"
		strSQL = "SELECT TOP 200 * FROM qryCAPSCards WITH(NOLOCK) WHERE [CardNumberShort] like '%" & strCardNoSearch & "%' " & strWhere & strSort
		
	End If
	
	'If the user has entered the End Card Number lookup then process this, otherwise perform the EmployeeID and Name searches
	If UCASE(Left(strSearch,2)) = "E:" Then
	
		strCardNoSearch = Right(strSearch,Len(strSearch)-2)
		
		'strWhere = strWhere & " AND [CardNumberShort] like '%" & strCardNoSearch & "%'"
		strSQL = "SELECT TOP 200 * FROM qryCAPSCards WITH(NOLOCK) WHERE [CardNumberShort] like '%" & strCardNoSearch & "' " & strWhere & strSort
		
	End If
	
	'If the user has entered the First Name lookup only (F: for CTS Accounts search) then process this, otherwise perform the EmployeeID and Name searches (above)
	If UCASE(Left(strSearch,2)) = "F:" Then
	
		strCardNoSearch = Right(strSearch,Len(strSearch)-2)
		
		'strWhere = strWhere & " AND [CardNumberShort] like '%" & strCardNoSearch & "%'"
		strSQL = "SELECT TOP 200 * FROM qryCAPSCards WITH(NOLOCK) WHERE [FirstName] like '%" & strCardNoSearch & "' " & strWhere & strSort
		
	End If
	
	'If the user has entered the Last Name (Surname) lookup only (L: for CTS Accounts search) then process this, otherwise perform the EmployeeID and Name searches (above)
	If UCASE(Left(strSearch,2)) = "L:" Then
	
		strCardNoSearch = Right(strSearch,Len(strSearch)-2)
		
		'strWhere = strWhere & " AND [CardNumberShort] like '%" & strCardNoSearch & "%'"
		strSQL = "SELECT TOP 200 * FROM qryCAPSCards WITH(NOLOCK) WHERE [Surname] like '%" & strCardNoSearch & "' " & strWhere & strSort
		
	End If
	
	'If the user has entered the Name on Card lookup only (L: for CTS Accounts search) then process this, otherwise perform the EmployeeID and Name searches (above)
	If UCASE(Left(strSearch,2)) = "N:" Then
	
		strCardNoSearch = Right(strSearch,Len(strSearch)-2)
		
		'strWhere = strWhere & " AND [CardNumberShort] like '%" & strCardNoSearch & "%'"
		strSQL = "SELECT TOP 200 * FROM qryCAPSCards WITH(NOLOCK) WHERE [NameOnCard] like '%" & strCardNoSearch & "' " & strWhere & strSort
		
	End If
	
	
'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = "for " & strSearch
Else
	strRecordMessage = "for " & Session("UserName") 
End If

Session("ExcelSearch") = strSQL

If not isEmpty(Session("ExcelSearch")) Then
	Session("ExcelSearch") = Replace(Session("ExcelSearch"),"TOP 500","")
	Session("ExcelSearch") = Replace(Session("ExcelSearch"),"TOP 200","")
	
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
		
		If IsNull(objRS("CardType")) or objRS("CardType") = "" Then
			strCardTypeDelete = ""
		Else
			If IsNull(objRS("CardTypeSub")) or objRS("CardTypeSub") = "" Then
				strCardTypeDelete = objRS("CardType")
			Else
				strCardTypeDelete = objRS("CardType") & " - " & objRS("CardTypeSub")
			End If
		End If
		
		If Left(objRS("CardTypeSub"),3) = "NAB" Then
			strCardTypeDelete = objRS("CardType") & " - NAB"
		End If
		
		
		If strCardTypeDelete = "DTC - Diners" Then
			strDeleteMessage = "<div class=""alert alert-danger"" role=""alert"">This Card is a Diners so the DTC and CMC will be cancelled if you continue!</div>"
		ElseIf strCardTypeDelete = "DPC - Diners" Then
			strDeleteMessage = "<div class=""alert alert-danger"" role=""alert"">This Card is a Diners DPC so the related DPC Mastercard will be cancelled if you continue!</div>"
		ElseIf strCardTypeDelete = "DPC - Mastercard" Then
			strDeleteMessage = "<div class=""alert alert-danger"" role=""alert"">This Card is a Diners DPC Mastercard so the DPC will be cancelled if you continue!</div>"
		ElseIf strCardTypeDelete = "DTC - NAB" Then
			strDeleteMessage = "<div class=""alert alert-danger"" role=""alert"">This Card is a NAB DTC so the DTC will be cancelled if you continue!</div>"
		ElseIf strCardTypeDelete = "DPC - NAB" Then
			strDeleteMessage = "<div class=""alert alert-danger"" role=""alert"">This Card is a NAB DPC so the DPC will be cancelled if you continue!</div>"
		Else
			strDeleteMessage = "<div class=""alert alert-warning"" role=""alert"">This Card is a Mastercard so ONLY the CMC will be cancelled if you continue!</div>"
		End If
		
		If IsNull(objRS("CardNumber")) Then
			strCardNo = ""
		Else
			strCardNo = MaskCard(RTrim(objRS("CardNumber")))
			
			If Session("UserTypeID") > 9 Then 
				strCardNoDelete = FormatCardNumber(objRS("CardNumber"))
			Else
				strCardNoDelete = MaskCard(objRS("CardNumber"))
			End If
		End If
		
		If Left(objRS("CardTypeSub"),3) = "NAB" Then
		
				If objRS("Status") = "" Then
				
					If IsNull(objRS("NameOnCard")) Then
						strNameonCardMod = ""
					Else
						strNameonCardMod = objRS("NameOnCard")
					End If

					strStatus = "<center><span class=""badge badge-pill badge-success"">Active</span></center>"
					strAction = "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-danger btn-sm"" data-toggle=""modal"" data-target=""#ModalDelete"" data-CStatus=""" & Replace(strStatus,"""","&quot;") & """ data-CCardID=""" & objRS("CardID") & """ data-CEmpID=""" & objRS("EmployeeID") & """ data-CNameOnCard=""" & Replace(strNameonCardMod,"'","") & """ data-CCardType=""" & objRS("CardType") & " - " & objRS("CardTypeSub") & """ data-CCardNo=""" & strCardNoDelete & """ data-CMessage=""" & Replace(strDeleteMessage,"""","&quot;") & """ onClick=""LoadDeletModal(this);""><i class=""fa fa-minus-circle""></i> Cancel</button>"	

				End If
				
		Else
			
			Select Case objRS("Status")
				
				Case  "Received"
					strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" Style=""font-size:13px;"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Release</button>"
					strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
				Case "Added To CS"
					
					strAction = "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CSToDiners.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
					
				Case "Submitted"
					strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardType=" & objRS("CardType") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"

					'strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Submitted to GCFO</button>"
					strStatus = "<center><span class=""badge badge-pill badge-success"">Submitted to GCFO</span></center>"
				Case "Cancelled"
					strAction = "Cancelled"
					'strAction = "Cancelled - " & FormatDateTime(objRS("DateUpdated"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
					'strStatus  = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
					strStatus = "<center><span class=""badge badge-pill badge-danger"">Cancelled</span></center>"
				Case "GCFO Approved"
					strAction = "<button type=""button"" title=""Approved by GCFO"" class=""btn btn-secondary btn-xs"" onclick=""self.location='ApplicationsEmployeeHF.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-check""></i>GCFO Approved</button>"
				
					'strStatus  = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Approved by GCFO</button>"
					strStatus = "<center><span class=""badge badge-pill badge-success"">Approved by GCFO</span></center>"
				Case "L"
					strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
					strStatus = "<span class=""badge badge-pill badge-danger"">Lost</span>"
				Case "C"
					strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				Case "S"
					strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
					strStatus = "<span class=""badge badge-pill badge-danger"">Stolen</span>"
				Case "T"
					strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
					strStatus = "<span class=""badge badge-pill badge-warning"">Temporary Hold</span>"
				Case "01"
					strAction = "<a href=""#"" onClick=""loadCSFromDiners(" & objrs("EmployeeID") & "," & objrs("CardNumber") & ")"" data-toggle=""modal"" data-target=""#ModalCancelCS"">Cancelled - " & objRS("FileDateTime") & "</a>"'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
					'strAction = "Cancelled - " & FormatDateTime(objRS("DateLoaded"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
					'strStatus  = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				Case "02"
					strAction = "Cancelled"
					'strAction = "Cancelled - " & FormatDateTime(objRS("DateUpdated"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
					'strStatus  = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				Case "03"
					strAction = "Cancelled"
					'strAction = "Cancelled - " & FormatDateTime(objRS("DateUpdated"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
					'strStatus  = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				
				Case Else
					
					'strAction = "Rejected"
					'strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Submitted</button>"
					strStatus = "<span class=""badge badge-pill badge-sm badge-success"">Active</span>"
					
					'Removed action from here as this is only DPC cards and they cannot be cancelled in the system
					If strCardType = "DPC" and objRS("CardTypeSub") = "ANZ" Then
						
					Else
						
						strAction = "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-danger btn-sm"" data-toggle=""modal"" data-target=""#ModalDelete"" data-CStatus=""" & Replace(strStatus,"""","&quot;") & """ data-CCardID=""" & objRS("CardID") & """ data-CEmpID=""" & objRS("EmployeeID") & """ data-CNameOnCard=""" & Replace(objRS("NameOnCard"),"'","") & """ data-CCardType=""" & objRS("CardType") & " - " & objRS("CardTypeSub") & """ data-CCardNo=""" & strCardNoDelete & """ data-CMessage=""" & Replace(strDeleteMessage,"""","&quot;") & """ onClick=""LoadDeletModal(this);""><i class=""fa fa-minus-circle""></i> Cancel</button>"	

					End If
					
				End Select
			
			
				
		End If

		strProcessAction = ""
		
		'Check the Process Status (which determines any actions being performed on them)
		Select Case objRS("ProcessStatus")
		
			Case "Added To CS"
				If Left(objRS("CardTypeSub"),3) <> "NAB" Then
					strProcessAction = "<button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='../Admin/CAPSAdmin/ExportCS.asp?Link=AD&CardID=" & objrs("CardID") & "'"";>View CM</button>"
				Else
					strProcessAction = "<button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""self.location='../Admin/CAPSAdmin/ExportCSNAB.asp?Link=AD&CardID=" & objrs("CardID") & "'"";>View CM</button>"
					'strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";>View Card</button>"
					'strAction = strAction + "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-danger btn-sm"" data-toggle=""modal"" data-target=""#ModalDelete"" data-CStatus=""" & Replace(strStatus,"""","&quot;") & """ data-CCardID=""" & objRS("CardID") & """ data-CEmpID=""" & objRS("EmployeeID") & """ data-CNameOnCard=""" & Replace(strNameonCardMod,"'","") & """ data-CCardType=""" & objRS("CardType") & " - " & objRS("CardTypeSub") & """ data-CCardNo=""" & strCardNoDelete & """ data-CMessage=""" & Replace(strDeleteMessage,"""","&quot;") & """ onClick=""LoadDeletModal(this);"">Cancel Requested</button>"	

				End If
			
			Case "Cancel Requested"
				If Left(objRS("CardTypeSub"),3) = "NAB" Then
					If objRS("Status") = "" Then
						strProcessAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportCSNAB.asp?Link=AD&CardID=" & objrs("CardID") & "'"";>View CM</button>"
						strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";>View Card</button><br><br>"
						'strAction = strAction + "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-danger btn-sm"" data-toggle=""modal"" data-target=""#ModalDelete"" data-CStatus=""" & Replace(strStatus,"""","&quot;") & """ data-CCardID=""" & objRS("CardID") & """ data-CEmpID=""" & objRS("EmployeeID") & """ data-CNameOnCard=""" & Replace(strNameonCardMod,"'","") & """ data-CCardType=""" & objRS("CardType") & " - " & objRS("CardTypeSub") & """ data-CCardNo=""" & strCardNoDelete & """ data-CMessage=""" & Replace(strDeleteMessage,"""","&quot;") & """ onClick=""LoadDeletModal(this);"">Cancel Requested</button>"	
						strStatus = "<center><span class=""badge badge-pill badge-warning"">Cancel Requested</span></center>"	
					End If
				End If
				
		End Select
		
		'Tiffany things
		If Left(objRS("CardTypeSub"),3) = "NAB" OR Left(objRS("CardTypeSub"),3) = "GLA" Then
			If objRS("Status") = "VX" Then
				strProcessAction = "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-danger btn-sm"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";>Cancelled "& CStr(Mid(objRS("FileSeqNum"),7,2) & "/" & Mid(objRS("FileSeqNum"),5,2) & "/" & Left(objRS("FileSeqNum"),4)) & "</button>"
				strAction = "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
				strStatus = "<center><span class=""badge badge-pill badge-danger"">" & objRS("BlockCodeDesc") & "</span></center>"
			End If
			If objRS("Status") = "ZQ" OR objRS("Status") = "SF" OR objRS("Status") = "SS" Then
				strProcessAction = ""
				strAction = "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i>View Card</button>"
				strStatus = "<center><span class=""badge badge-pill badge-danger"">" & objRS("BlockCodeDesc") & "</span></center>"
			End If
			If objRS("Status") = "XS" Then
				strProcessAction = "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-danger btn-sm"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";>Blocked "& CStr(Mid(objRS("FileSeqNum"),7,2) & "/" & Mid(objRS("FileSeqNum"),5,2) & "/" & Left(objRS("FileSeqNum"),4)) & "</button>"
				strAction = "<button type=""button"" Style=""font-size:13px;"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
				strStatus = "<center><span class=""badge badge-pill badge-danger"">" & objRS("BlockCodeDesc") & "</span></center>"
			End If
		End If
		
		strAddress = Trim(objRS("Address1")) & " " & Trim(objRS("Address2")) & " " & Trim(objRS("Suburb")) & " " & Trim(objRS("State")) & " " & Trim(objRS("PostCode"))
		
		If len(strAddress) > 15 Then strAddress = left(strAddress,15) & "..."
		
		'If IsNull(objRS("DateLoaded")) Then
		If IsNull(objRS("DateIssued")) Then
			dteDateSubmitted = ""
		Else
			dteDateSubmitted = FormatDateTime(objRS("DateIssued"),vbShortDate)
		End If
		
		If IsNull(objRS("PMLoadDate")) Then
			dteDateReviewed = ""
		Else
			dteDateReviewed = FormatDateTime(objRS("PMLoadDate"),vbShortDate)
		End If
		
		If IsNull(objRS("FirstName")) AND IsNull(objRS("Surname")) Then
			strName = ""
		Else
			strName = Trim(objRS("FirstName")) & " " & Trim(objRS("Surname"))
			If Len(strName)>15 Then strName = Left(strName,15)
		End If
		
		
		response.write "<TR><TD ><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """ Style=""font-size:13px;"">" & objRS(0) & "</a></TD><TD Style=""font-size:13px;"">" & objRS("EmployeeID") & "</a></TD>" & _
				"<TD Style=""font-size:13px;""><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """ Style=""font-size:13px;"">" & strName & "</a></TD><TD Style=""font-size:13px;"">" & strCardNo & "</TD>" & _
				"<TD Style=""font-size:13px;""><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """ Style=""font-size:13px;"">" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
				"<TD Style=""font-size:13px;"">" & strAddress & "</TD><TD >" & strStatus & "</TD>" & _
				"<TD Style=""font-size:13px; text-align:right;"">" & dteDateSubmitted & "</TD><TD Style=""font-size:13px; text-align:right;"">" & dteDateReviewed & "</TD>" & _
				"<TD Style=""font-size:12px;"">" & strAction & "</TD><TD Style=""font-size:13px;"">" & strProcessAction & "</TD></TR>"
				
		'response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & strAction & "</a></TD>" & _
		'		"<TD><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?CardID=" & objRS(0) & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?CardID=" & objRS(0) & """>" & objRS(2) & "</a></TD>" & _
		'		"<TD style=""text-align:center;"">" & objRS(3) & "</TD><TD style=""text-align:center;"">" & objRS(4) & "</TD>" & _
		'		"<TD style=""text-align:center;"">" & objRS(5) & "</TD><TD style=""text-align:center;"">" & objRS(6) & "</TD>" & _
		'		"<TD style=""text-align:center;"">" & objRS(7) & "</TD><TD style=""text-align:center;"">" & objRS(10) & "</TD>" & _
		'		"<TD style=""text-align:center;"">" & strStatus & "</TD><TD style=""text-align:center;"">" & objRS(14) & "</TD><TD style=""text-align:center;"">" & objRS(15) & "</TD></TR>"
			
			y = y + 1
			
		objRS.movenext
	Loop
	
	
	Response.Write "<TR><TH colspan=""10"">Total</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;"">" & y & " <input type=""HIDDEN"" id=""WhereClause"" name=""WhereClause"" value=""SQL=" & strSQL & """ ></TH></TR>"
				
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
Dim strDropDown
Dim strAppText
Dim strAppText2

Dim arrButton(7)

'If Session("ViewButton") = "Diners" Then
'	arrButton(2) = "active"
'ElseIf Session("ViewButton") = "Mastercard" Then
'	arrButton(3) = "active"
'ElseIf Session("ViewButton") = "ANZ" Then
'	arrButton(4) = "active"
'ElseIf Session("ViewButton") = "CTS" Then
'	arrButton(5) = "active"
'ElseIf Session("ViewButton") = "DINERS" Then
'	arrButton(6) = "active"
'Else
	'This catches ALL
'	arrButton(1) = "active"
'End If

If Session("ViewButton") = "NABDTC" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "NABDPC" Then
	arrButton(3) = "active"
ElseIf Session("ViewButton") = "CTS" Then
	arrButton(4) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

'''----------------BEGIN
'If Session("CardTypeNameView") <> "" Or Session("ApplicationStatusNameView") <> ""  Then
'Else

'		If Session("ViewButton") = "Review" Then
'			arrButton(2) = "active"
'		ElseIf Session("ViewButton") = "OnHold" Then
'			arrButton(3) = "active"
'		ElseIf Session("ViewButton") = "AwaitingIssue" Then
'			arrButton(4) = "active"
'		ElseIf Session("ViewButton") = "AddedToNA" Then
'			arrButton(5) = "active"
'		ElseIf Session("ViewButton") = "TempHold" Then
'			arrButton(6) = "active"
'			
'		Else
'			'This catches ALL
'			arrButton(1) = "active"
'		End If
	
'End If	


	If Session("CardTypeNameView") = "" Or IsNull(Session("CardTypeNameView")) Then
		strAppText = "Card Type"
	Else
		'Make the Button text the selected Application Type if one is selected
		If Session("ViewButton") = "" OR IsNull(Session("ViewButton")) Then
			strAppText = Left(Session("CardTypeNameView"),10)
			arrButton(7) = "active"
		End If
	End If

'Create the Application Type selection/filter
strDropDown = "<div class=""dropdown""><button style=""height:50px;"" class=""btn btn-outline-primary btn-sm dropdown-toggle " & arrButton(7) & """ type=""button"" id=""dropdownMenuButton"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false""><i class=""fa fa-credit-card""></i> " & strAppText & "</button>" & _
		"<div class=""dropdown-menu"" aria-labelledby=""dropdownMenuButton""><a class=""dropdown-item"" href=""Cards.asp?CardTypName="">All</a>"
		
		objRS.Open "SELECT [CardType],[CardTypeSub] FROM tblCAPSCard WITH(NOLOCK) GROUP BY [CardType],[CardTypeSub]",objCon

		Do Until objRS.EOF
		
			strDropDown = strDropDown & "<a class=""dropdown-item"" href=""Cards.asp?CardTypeNameView=" & objRS("CardType") & objRS("CardTypeSub") & """>" & objRS("CardType") & " - " & objRS("CardTypeSub") & "</a>"
		
		objRS.Movenext
		Loop
		
		objRS.Close

		strDropDown = strDropDown & "</div></div>"
	
	
	
	'''----------------END
	
	'Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
	'			"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='Cards.asp?ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
	'			"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='Cards.asp?ViewButton=Diners';""><i class=""fa fa-plane""></i> View DTC Diners</button>" & _
	'			"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='Cards.asp?ViewButton=Mastercard';""><i class=""fa fa-credit-card""></i> View DTC Mastercard</button>" & _
	'			"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='Cards.asp?ViewButton=ANZ';""><i class=""fa fa-dollar-sign""></i> View DPC ANZ</button>" & _
	'			"<button type=""button"" class=""btn btn-outline-primary " & arrButton(6) & """ onClick=""self.location.href='Cards.asp?ViewButton=DINERS';""><i class=""fa fa-dollar-sign""></i> View DPC DINERS</button>" & _
	'			"<button type=""button"" class=""btn btn-outline-primary " & arrButton(5) & """ onClick=""self.location.href='Cards.asp?ViewButton=CTS';""><i class=""fa fa-cogs""></i> View CTS</button>" & strDropDown & _
	'			"</div>" 
				
	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='Cards.asp?ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='Cards.asp?ViewButton=NABDTC';""><i class=""fa fa-plane""></i> View NAB DTC</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='Cards.asp?ViewButton=NABDPC';""><i class=""fa fa-credit-card""></i> View NAB DPC</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='Cards.asp?ViewButton=CTS';""><i class=""fa fa-cogs""></i> View CTS</button>" & strDropDown & _
				"</div>" 

End Sub

Public Sub LoadStatusButtons
'Load the the Status Selector buttons depending on what has been clicked
Dim strStatusButton

'Get the Status View button depending on what has been selected
If Session("StatusViewButton") = "Active" Then
	strStatusButton ="<button type=""button"" class=""btn btn-outline-success active"" onClick=""self.location.href='Cards.asp?StatusViewButton=Cancelled';"" title=""Click to view Cancelled Cards Only""><i class=""fa fa-check""></i> View Active</button>"
ElseIf Session("StatusViewButton") = "Cancelled" Then
	strStatusButton ="<button type=""button"" class=""btn btn-outline-danger active"" onClick=""self.location.href='Cards.asp?StatusViewButton=TempHold';"" title=""Click to View DPC Cards with a Temporary Hold""><i class=""fa fa-times""></i> View Cancelled</button>"
ElseIf Session("StatusViewButton") = "TempHold" Then
	strStatusButton ="<button type=""button"" class=""btn btn-outline-warning active"" onClick=""self.location.href='Cards.asp?StatusViewButton=All';"" title=""Click to View ALL Status Cards""><i class=""fa fa-hand-paper""></i> View Temp Hold</button>"
Else
	strStatusButton ="<button type=""button"" class=""btn btn-outline-secondary active"" onClick=""self.location.href='Cards.asp?StatusViewButton=Active';"" title=""Click to view Active Cards Only""><i class=""fa fa-asterisk""></i> View All Statuses</button>"
End If

	Response.Write 	"<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				strStatusButton & _
				"</div>"

End Sub


Public Sub CancelApplication()
'Procedure to Cancel the selected card

Dim lngCardIDCancel
Dim intRecord

	If Not IsEmpty(Request.QueryString("CardID")) Then
		lngCardIDCancel = Request.QueryString("CardID")
		
		'If the Card is a DPC/ANZ then change it to cancelled, otherwise call the save procedure to add it to the CS File
		
		If Request.QueryString("CardID") = "DPC" Then
		
			intRecord = CancelCardToCS(0,lngCardIDCancel,"")
		Else
		
			With objCmd

				.CommandType = 4
				.CommandText = "spCAPSCSToDinersCancelCard"'"spCAPSCancelCard"

				.Parameters.Append objCmd.CreateParameter("CSToDinersID", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("CardID", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("Notes", adVarChar, adParamInput, 100)
				.Parameters.Append objCmd.CreateParameter("CSToDinersIDDOutput", adInteger, adParamOutput)
				
				.Parameters("CSToDinersID") = 0
				.Parameters("CardID") = lngCardIDCancel
				.Parameters("UpdatedBy") = Session("UserID")
				.Parameters("Notes") = ""
				
				.ActiveConnection = objCon
				 
			End With
		   
			objCmd.Execute        
		  
			'Return the result of the Save Function.
			intRecord = objCmd.Parameters.Item("CSToDinersIDDOutput") 
		
		'End of Card Type Check
		End If
		
		'strSQL = "UPDATE tblCAPSCard SET Status = 'Cancelled' WHERE CardID = " & lngCardIDCancel & ""
		
		'objCon.Execute strSQL
		
		If intRecord = -1 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Card " & lngCardIDCancel & " for " & Request.QueryString("Name") & " NOT Cancelled. There is already a record on the CS File for this card today. Only one change can be added per day. Try again tomorrow or check the CS File.</div>"
		ElseIf intRecord = 0 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Card " & lngCardIDCancel & " for " & Request.QueryString("Name") & " NOT Cancelled. ERROR! See System Admin - CARD ID: " & lngCardIDCancel & "</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Card " & lngCardIDCancel & " for " & Request.QueryString("Name") & " Status Updated to CANCELLED!</div>"
		End If
		
	Else
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Card " & lngCardIDCancel & " for " & Request.QueryString("Name") & " NOT Cancelled. ERROR! See System Admin</div>"
	End If
	
End Sub

Public Sub SubmitApplication()

Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spCDMCToApplication"

			.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("CardType", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("CardTypeSub", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("CreditLimit", adDouble, adParamInput) 
			.Parameters.Append objCmd.CreateParameter("CDMCToCardIDOutput", adInteger, adParamOutput)
			
			.Parameters("EmployeeID") = Session("EmployeeID")
			.Parameters("CardType") = Left(Request.Form("CardType"),3)
			.Parameters("CreditLimit") = Request.Form("CreditLimit")
			.Parameters("CardTypeSub") = Right(Request.Form("CardType"),Len(Request.Form("CardType"))-6)
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCToCardIDOutput") 
	 
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" /> Application " & intRecord & " submitted to your GCFO for approval!"

		strMessageColour = "Black"
		
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
