
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
	Else
		Session("ViewButton") = "All"
	End If
	
	If Not IsEmpty(Request.QueryString("Action")) Then
		If Request.QueryString("Action") = "Cancel" Then
			Call CancelApplication()
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

function CancelCard() {
	
	var id = document.getElementById('CardID').value
	self.location = "ANZCancelCards.asp?Action=Cancel&CardID=" + id;

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
				<h2 class="text-left">Cancel ANZ Cards <% If Session("UserView") = "User" Then Response.write " for " & Session("UserName")%></h2>
			</div>
		</div>

          <div class="row py-2">
            <div class="col-md-9">
              <%Call LoadViewButtons()%>
            </div>
			<div class="col-md-3">
				<div class="form-group has-search">
					
				  
				 </div>
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
					<th scope="col">Address</th>
					<th scope="col">Card Status</th>
					<th scope="col">Date Issued</th>
					<th scope="col">PM Load Date</th>
					<th scope="col">Action</th>
					<th scope="col" style="width:20px;">Process</th>
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
		 <button type="button" class="btn btn-outline-secondary btn-sm" data-toggle="modal" data-target="#" onclick="ExcelExport('SQL')" ><i class="fa fa-file"></i>Export Results</button>
		
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
	
	If Session("ViewButton") = "Cancelled" Then
		strWhere = " AND [ProcessStatus] = 'Exported' "
	ElseIf Session("ViewButton") = "All" Then
		strWhere = " AND [ProcessStatus] = 'Cancelled' "
	Else
		'This catches ALL
		strWhere = ""
	End If
	
If strSearch = "" OR ISNull(strSearch) Then

	If Session("ViewButton") = "All" Then
		strSQL = "SELECT  * FROM qryCAPSCards WITH(NOLOCK) WHERE CardType = 'DPC' AND Status = '' AND [CardID] > 0 " & strWhere & strSort
	ElseIf  Session("ViewButton") = "Cancelled" Then
		strSQL = "SELECT  * FROM qryCAPSCards WITH(NOLOCK) WHERE CardType = 'DPC' AND Status = '' AND [CardID] > 0 " & strWhere & strSort
	Else
		strSQL = "SELECT  * FROM qryCAPSCards WITH(NOLOCK) WHERE CardType = 'DPC' AND Status = '' AND EmployeeID = '" & Trim(Session("EmployeeID")) & "'" & strWhere & strSort
	End If	
	
	Session("SQL") = strSQL
	
Else

	If Session("ViewButton") = "All" Then

		'Replace all single inverted commas with 2 for SQL
		strSearch = Replace(strSearch,"'","''")
		
		'If the user has entered a search term with a space the assume this is a first and last name so search on that only
		If Instr(1,strSearch," ")>0 Then
			arrNames = Split(strSearch," ")
			strFNameSearch = arrNames(0)
			strLNameSearch = arrNames(1)
		
			strSQL = "SELECT * FROM qryCAPSCards WITH(NOLOCK) WHERE CardType = 'DPC' AND (EmployeeID Like '%" & strSearch & "%' OR ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%'))" & strWhere & strSort
		Else
		
			strSQL = "SELECT * FROM qryCAPSCards WITH(NOLOCK) WHERE CardType = 'DPC' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')" & strWhere & strSort
		
		End If
	Else
		strSQL = "SELECT * FROM qryCAPSCards WITH(NOLOCK) WHERE CardType = 'DPC' AND EmployeeID = '" & Session("EmployeeID") & "' AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
	End If
	
End If

'If the user has entered the Card Number lookup then process this, otherwise perform the EmployeeID and Name searches
	If Left(strSearch,2) = "c:" Then
	
		strCardNoSearch = Right(strSearch,Len(strSearch)-2)
		
		'strWhere = strWhere & " AND [CardNumberShort] like '%" & strCardNoSearch & "%'"
		strSQL = "SELECT top 100 * FROM qryCAPSCards WITH(NOLOCK) WHERE CardType = 'DPC' AND [CardNumberShort] like '%" & strCardNoSearch & "%'"
		
	End If
	
'Build the message displayed at the bottom of the screen with the search details
If Session("UserView") = "All" Then
	strRecordMessage = "for " & strSearch
Else
	strRecordMessage = "for " & Session("UserName") 
End If
'response.write strSQL
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
		strProcessStatus = objRS("ProcessStatus")
				
		'Set the Card Type for use below
		If IsNull(objRS("CardType")) or objRS("CardType") = "" Then
			strCardType = ""
		Else
			strCardType = objRS("CardType")
		End If
		
		Select Case objRS("Status")
		
		Case  "Received"
			strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
		Case "Added To CS"

			strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CSToDiners.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
			
		Case "Submitted"
			strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "&CardType=" & objRS("CardType") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"

			'strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Submitted to GCFO</button>"
			strStatus = "<span class=""badge badge-pill badge-success"">Submitted to GCFO</span>"
		Case "Cancelled"
			strAction = "Cancelled"
			'strAction = "Cancelled - " & FormatDateTime(objRS("DateUpdated"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
			'strStatus  = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
			strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
		Case "GCFO Approved"
			strAction = "<button type=""button"" title=""Approved by GCFO"" class=""btn btn-secondary btn-xs"" onclick=""self.location='ApplicationsEmployeeHF.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-check""></i>GCFO Approved</button>"
		
			'strStatus  = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Approved by GCFO</button>"
			strStatus = "<span class=""badge badge-pill badge-success"">Approved by GCFO</span>"
		Case "L"
			strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
			strStatus = "<span class=""badge badge-pill badge-danger"">Lost</span>"
		Case "C"
			strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
			strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
		Case "S"
			strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CardDetail.asp?Link=CD&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
			strStatus = "<span class=""badge badge-pill badge-danger"">Stolen</span>"
		Case "01"
			strAction = "Cancelled - " & objRS("FileDateTime")'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
			'strAction = "Cancelled - " & FormatDateTime(objRS("DateLoaded"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
			'strStatus  = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
			strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
		Case "02"
			strAction = "Cancelled"
			'strAction = "Cancelled - " & FormatDateTime(objRS("DateUpdated"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
			'strStatus  = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
			strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
		
		Case Else
			'Removed action from here as this is only DPC cards and they cannot be cancelled in the system
			If strCardType = "DPC" AND strProcessStatus = "Cancelled" Then
				'strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ANZCancelCards.asp?Link=CD&Action=Cancel&CardID=" & objrs("CardID") & "&Name=" & objRS("FirstName") & " " & objRS("Surname") & "&EmployeeID=" & objRS("EmployeeID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				strAction = "<button class=""btn btn-danger btn-xs""  data-toggle=""modal"" data-target=""#ModalCancel"" data-cardno=""" & objRS("CardNumber") & """ data-id=""" & objRS("CardID") & """ data-name=""" & objRS("NameOnCard") & """ onClick=""loadCancel(this);""><i class=""fa fa-pen""></i> Cancel Card </button>"
			Else
				strAction = "<span class=""badge badge-pill badge-danger"">Sent to ANZ for Cancellation</span>"
			End If
			
			'strAction = "Rejected"
			'strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?CardID=" & objrs("CardID") & "'"";>Submitted</button>"
			strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
		End Select

		strProcessAction = ""
		
		'Check the Process Status (which determines any actions being performed on them)
		Select Case objRS("ProcessStatus")
		
			Case "Added To CS"
				strAction = "<b>Cancel Requested</b>"
				strProcessAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportCS.asp?Link=AD&CardID=" & objrs("CardID") & "'"";>View CS</button>"
			
		End Select
		
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
		
		If IsNull(objRS("CardNumber")) Then
			strCardNo = ""
		Else
			strCardNo = MaskCard(objRS("CardNumber"))
		End If
		
		If IsNull(objRS("FirstName")) AND IsNull(objRS("Surname")) Then
			strName = ""
		Else
			strName = Trim(objRS("FirstName")) & " " & Trim(objRS("Surname"))
			If Len(strName)>15 Then strName = Left(strName,15)
		End If
		

		
		response.write "<TR><TD ><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
				"<TD Style=""font-size:14px;""><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """>" & strName & "</a></TD><TD >" & strCardNo & "</TD>" & _
				"<TD Style=""font-size:14px;""><a Target=""_self"" HREF=""CardDetail.asp?Link=CD&CardID=" & objRS(0) & """>" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
				"<TD Style=""font-size:14px;"">" & strAddress & "</TD><TD >" & strStatus & "</TD>" & _
				"<TD Style=""font-size:14px; text-align:right;"">" & dteDateSubmitted & "</TD><TD Style=""font-size:14px; text-align:right;"">" & dteDateReviewed & "</TD>" & _
				"<TD Style=""font-size:14px;"">" & strAction & "</TD><TD Style=""font-size:14px;"">" & strProcessAction & "</TD></TR>"
				
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
				"<TH colspan=""3"" style=""text-align:center;"">" & y & "</TH></TR>"
				
objRS.Close

Response.Write strW
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
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='ANZCancelCards.asp?ViewButton=All';""><i class=""fa fa-folder""></i> To be Cancelled</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='ANZCancelCards.asp?ViewButton=Cancelled';""><i class=""fa fa-plane""></i> Cancelled Cards</button>" & _
				"</div>"

End Sub

Public Sub CancelApplication()
'Procedure to Cancel the selected card

Dim lngCardIDCancel
Dim intRecord
response.write lngCardIDCancel
	If Not IsEmpty(Request.QueryString("CardID")) Then
		lngCardIDCancel = Request.QueryString("CardID")
		
		'If the Card is a DPC/ANZ then change it to cancelled, otherwise call the save procedure to add it to the CS File
		
		If Request.QueryString("CardID") = "DPC" Then
		
			intRecord = CancelCardToCS(0,lngCardIDCancel,"")
		Else
		
			With objCmd

				.CommandType = 4
				.CommandText = "spCAPSCancelANZCard"
			
				.Parameters.Append objCmd.CreateParameter("CardID", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("CAPSCancelANZIDDOutput", adInteger, adParamOutput)
				
				.Parameters("CardID") = lngCardIDCancel
				.Parameters("UpdatedBy") = Session("UserID")
					
				.ActiveConnection = objCon
				 
			End With
			
			Response.Write "Cancel Application" & " " & lngCardIDCancel
		   
			objCmd.Execute        
		  
			'Return the result of the Save Function.
			intRecord = objCmd.Parameters.Item("CAPSCancelANZIDDOutput") 
		
		'End of Card Type Check
		End If
		
		'strSQL = "UPDATE tblCAPSCard SET Status = 'Cancelled' WHERE CardID = " & lngCardIDCancel & ""
		
		'objCon.Execute strSQL
		
		If intRecord = -2 Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Card " & lngCardIDCancel & " for " & Request.QueryString("Name") & " NOT Cancelled. Card has been flagged as not to be cancelled.</div>"
		Else
			Response.Write "<div class=""alert alert-success"" role=""alert"">Card " & lngCardIDCancel & " for " & Request.QueryString("Name") & " Status Updated to CANCELLED!</div>"
		End If
		
	Else
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Card " & lngCardIDCancel & " for " & Request.QueryString("Name") & " NOT Cancelled. ERROR! See System Admin</div>"
	End If
	
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
