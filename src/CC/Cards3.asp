<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")

'Description:	Data Entry for General Expenses data
'Author:		MG
'Date:			Janaury 2014

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

Dim lngApplicationID
Dim strEmployeeID
Dim strTitle
Dim strFirstName
Dim strLastName
Dim strAddress1
Dim strAddress2
Dim strAddress3
Dim strAddress4
Dim strSuburb
Dim strState
Dim strPostCode
Dim dteDateReceived
Dim strStatus
Dim strStatus2
Dim strCardType
Dim strReviewedBy
Dim dteDateReviewed
Dim lngCreditLimit
Dim arrState(8)
Dim lngCardID
Dim strStatusBG
Dim strView
Dim strButt1
Dim strButt2
Dim strButt3
Dim strButt4
Dim strButt5
Dim strButt6

Dim strTrans
Dim strUnaquitt
Dim strCardNumF

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objRS1 = Server.CreateObject("ADODB.Recordset")
    Set objRS2 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

	arrState(1) = "ACT"
	arrState(2) = "NSW"
	arrState(3) = "NT"
	arrState(4) = "QLD"
	arrState(5) = "VIC"
	arrState(6) = "SA"
	arrState(7) = "WA"
	arrState(8) = "TAS"
	
    objCon.Open Session("DBConnection")	

    Session("CurrentPage") = "CC/Cards3.asp"

	If IsNull(Session("CardID")) OR Session("CardID") = "" Then Session("CardID")= 0

If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	Session("ApplicationID") = Request.QueryString("ApplicationID")
End If

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	Session("EmployeeID") = Request.QueryString("EmployeeID")
	Session("CarParkingID") = 0
End If

If Not IsEmpty(Request.QueryString("CardID")) Then
	Session("CardID") = Request.QueryString("CardID")
End If

If Not IsEmpty(Request.QueryString("Filter")) Then
	Session("Filter") = Request.QueryString("Filter")
End If

If Not IsEmpty(Request.QueryString("Filter2")) Then
	Session("Filter2") = Request.QueryString("Filter2")
End If

If Not IsEmpty(Request.QueryString("View")) Then
	Session("View") = Request.QueryString("View")
End If

If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "Reject" Then
		Call RejectApplication()
	End If
	
	If Request.QueryString("Action") = "Release" Then
		Call ReleaseApplication()
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

  
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "New" Then
			'Response.Write "CPID=" & Session("CarParkingID")
			Session("CarParkingID") = 0
		End If
	End If
	
	If isNull(Session("CarParkingID")) Or Session("CarParkingID") = "" Then 
		Session("CarParkingID") = 0
	End If

	If Session("View") = "All" Then
		strView = " All"
	Else
		strView = Session("UserType") & " - " & Session("UserName")
	End If
	
  Call LoadDetails()
  call LoadPMData()
  
%>

<html>
<head>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

 <!-- Bootstrap Core CSS -->
    <link href="../css/bootstrap.min.css" rel="stylesheet">
	

<link rel="stylesheet" type="text/css" href="../CAPSStyle.css">
	  <!-- Custom fonts for this template-->
  <link href="../vendor/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css">
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

    $(function(){           
        if (!Modernizr.inputtypes.date) {
            $('input[type=date]').datepicker({
                  dateFormat : 'yy-mm-dd'
                }
             );
        }
    });
	
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
});
</script>
	
</head>
<body >


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
        <label>I DECLARE that all the application details are correct.</label><br>
            <label for="name">First Name:</label>
            <input type="text" name="FirstName" id="FirstName" class="form-control input-md">
			<label for="name">Last Name:</label>
            <input type="text" name="LastName" id="LastName" class="form-control input-md">
			<label for="email">Email:</label>
            <input type="email" name="email" id="email" class="form-control input-md">
            <label for="phone">Phone:</label>
            <input type="text" name="phone" id="phone" class="form-control input-md">
			<label for="phone">Group:</label>
            
			<select class="form-control" id="Group">
			  <option>Please select your group...</option>
			  <option>Army</option>
			  <option>Navy</option>
			  <option>Air Force</option>
			  <option>CIOG</option>
			  <option>CFOG</option>
			  <option>DSTO</option> 
			  <option>JOC</option>
			</select>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
        <button type="button" class="btn btn-primary">Save changes</button>
      </div>
    </div>
  </div>
</div>


<!--<div id='tbl-container'>-->

<div class="content-wrapper">
    <div class="container-fluid">
		<ol class="breadcrumb" Style="Background-color:#86c5f9;color:white;">
        <li class="breadcrumb-item" Style="Background-color:#86c5f9;color:white;font-size:16px;">
          <a href="../IndexCC2.asp" target="_parent">Home</a>
        </li>
        <li class="breadcrumb-item active" Style="Background-color:#86c5f9;color:white;font-size:16px;"><i class="fa fa-credit-card"></i> Defence Credit Cards for <%=strView%></li>
      </ol>
	  
	
	  <hr>
	</div>
	  
<div class="container-fluid">
<form action="Applications3.asp?Action=Save" method="POST" id="frm" name="frm" class="form-inline">

<div class="card-header" >&nbsp;&nbsp;&nbsp;
    <i class="fa fa-address-card"></i> Existing Cards for <%=strView%>
</div>
</form>
</div>
 
 
<form class="form-inline" action="ApplicationsEmployee.asp?Action=SubmitApp"> 
<div class="container-fluid">

<div class="row" style="width:100%;">
<div class="input-group col-sm-6"> <!--Background-color:#1F71DE-->
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Employee ID</span>
    <input type="text" class="form-control" id="EmployeeID" placeholder="Employee ID" name="EmployeeID" style="width:50%;" value="<%=strEmployeeID%>"> 
</div>
	<div class="input-group col-sm-5">
	<span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Card Type</span> 
    <select class="form-control" placeholder="Card Type" style="width:40%;" name="CardTypeSelect" id="CardTypeSelect" onchange="valCard();">
							<option value="DTC - Diners">DTC - Diners</Option><option value="DPC - ANZ">DPC - ANZ</Option>
							<option value="DTC - MasterCard">DTC - Companion Mastercard</Option>
							</select>
</div>
</div>
<div class="row" style="width:100%;">
<div class="input-group col-sm-2">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Title</span>
    <input type="text" class="form-control" id="Title" placeholder="Title" name="Title" style="width:90%;" value="<%=strTitle%>"> 
</div>
 <div class="input-group col-sm-4">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">First Name</span>
    <input type="text" class="form-control" id="FirstName" placeholder="First Name" name="FirstName" style="width:90%;" value="<%=strFirstName%>"> 
</div>
 <div class="input-group col-sm-4">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Last Name</span>
    <input type="text" class="form-control" id="LastName" placeholder="Last Name" name="EmployeeID" style="width:90%;" value="<%=strLastName%>"> 
</div>
</div>

<div class="row" style="width:100%;">
 <div class="input-group col-sm-6">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Address 1</span>
    <input type="text" class="form-control" DISABLED title="Can Only be Updated in the Corporate Directory" id="Address1" placeholder="Address 1" name="Address1" style="width:90%;" value="<%=strAddress1%>"> 
</div>
 <div class="input-group col-sm-5">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Address 2</span>
    <input type="text" class="form-control" DISABLED title="Can Only be Updated in the Corporate Directory" id="Address2" placeholder="Address 2" name="Address2" style="width:90%;" value="<%=strAddress2%>"> 
</div>
</div>

<div class="row" style="width:100%;">
 <div class="input-group col-sm-6">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Address 3</span>
    <input type="text" class="form-control" DISABLED title="Can Only be Updated in the Corporate Directory" id="Address3" placeholder="Address 3" name="Address3" style="width:90%;" value="<%=strAddress3%>"> 
</div>
 <div class="input-group col-sm-5">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Suburb</span>
    <input type="text" class="form-control" DISABLED title="Can Only be Updated in the Corporate Directory" id="Suburb" placeholder="Suburb" name="Suburb" style="width:90%;" value="<%=strSuburb%>"> 
</div>
</div>

<div class="row" style="width:100%;">
 <div class="input-group col-sm-6">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">State</span>
    <input type="text" class="form-control" DISABLED title="Can Only be Updated in the Corporate Directory" id="State" placeholder="State" name="State" style="width:90%;" value="<%=strState%>"> 
</div>
 <div class="input-group col-sm-5">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Post Code</span>
    <input type="text" class="form-control" DISABLED title="Can Only be Updated in the Corporate Directory" id="PostCode" placeholder="Post Code" name="PostCode" style="width:90%;" value="<%=strPostCode%>"> 
</div>
</div>

<div class="row" style="width:100%;">
<div class="input-group col-sm-2">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Status</span>
    <input type="text" class="form-control <%=strStatusBG%>" id="Status" placeholder="Status" name="Status" Title="Status - <%=strStatus%>" style="width:90%;" value="<%=strStatus2%>"> 
</div>
 <div class="input-group col-sm-4">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Credit Limit</span>
    <input type="text" class="form-control" id="FirstName" placeholder="First Name" name="FirstName" style="width:90%;" value="<%=lngCreditLimit%>"> 
</div>
 <div class="input-group col-sm-4">
    <span class="input-group-addon" style="width:20%;Background-color:#2394F2;color:white;">Date Received</span>
    <input type="text" class="form-control" id="DateReceived" placeholder="DateReceived" name="DateReceived" style="width:90%;" value="<%=dteDateReceived%>"> 
</div>

</div>

<div class="row" style="width:70%;">
 <div class="input-group col-sm-6">
    <span class="input-group-addon input-xs" style="width:20%;Background-color:#eb9234;color:white;">ProMaster Transactions (last 90 Days)</span>
    <input type="text" class="form-control input-xs" DISABLED title="ProMaster Transactions for the previous 90 Days" id="PMTrans" placeholder="None" name="State" style="width:90%;" value="<%=strTrans%>"> 
</div>
 <div class="input-group col-sm-5">
    <span class="input-group-addon xs" style="width:20%;Background-color:#eb9234;color:white;">Longest UnAquitted Transaction</span>
    <input type="text" class="form-control xs" DISABLED title="Unaquitted ProMaster transactions" id="PMUnaquitt" placeholder="None" name="PostCode" style="width:90%;" value="<%=strUnaquitt%>"> 
</div>
</div>

</div>
  </form>
 
</div>

<div class="container-fluid">
<!-- Button trigger modal -->
<button type="button" class="btn btn-primary" data-toggle="modal" data-target="#exampleModalCenter"><i class="fa fa-bullhorn"></i> Declaration</button>

<button type="button" class="btn btn-primary" onclick="CloseScreen()";><i class="fa fa-times"></i> Close</button>
<button type="button" class="btn btn-primary" onclick="SaveData()"; ><i class="fa fa-check"></i> Save</button>
<button type="button" class="btn btn-primary" onclick="self.location='Applications3.asp?Action=New'"; ><i class="fa fa-pencil"></i> New</button>
<%=strMessageIcon %>
  
<hr />
</div>

      <!-- Breadcrumbs-->
     
      <!-- Example DataTables Card-->
      <div class="card mb-3">
        <div class="card-header">
          <i class="fa fa-table"></i> Existing Cards for <% Response.write strView
		  
		  If Session("Filter") = "Active" Then
			strButt1 = "secondary"
			strButt2 = "info"
		  Else
			strButt1 = "info"
			strButt2 = "secondary"
		  End If
		  
		  If Session("Filter2") = "DPC" Then
			strButt3 = "secondary"
			strButt4 = "secondary"
			strButt5 = "secondary"
			strButt6 = "info"
		  ElseIf Session("Filter2") = "DTC" Then
			strButt3 = "secondary"
			strButt4 = "info"
			strButt5 = "secondary"
			strButt6 = "secondary"
		ElseIf Session("Filter2") = "DTCMC" Then
			strButt3 = "secondary"
			strButt4 = "secondary"
			strButt5 = "info"
			strButt6 = "secondary"
		Else
			strButt3 = "info"
			strButt4 = "secondary"
			strButt5 = "secondary"
			strButt6 = "secondary"
		  End If
		  
		  
		  Response.write  " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<div class=""btn-group""><button type=""button"" class=""btn btn-" & strButt1 & """ onclick=""self.location='Cards3.asp?Filter=All'"";><i class=""fa fa-filter""></i> View All </button>" & _
							" <button type=""button"" class=""btn btn-" & strButt2 & """ onclick=""self.location='Cards3.asp?Filter=Active'"";><i class=""fa fa-filter""></i> View Active</button></div>" & _
							" &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<div class=""btn-group""><button type=""button"" class=""btn btn-" & strButt3 & """ onclick=""self.location='Cards3.asp?Filter2=All'"";><i class=""fa fa-filter""></i> View All </button>" & _
							" <button type=""button"" class=""btn btn-" & strButt4 & """ onclick=""self.location='Cards3.asp?Filter2=DTC'"";><i class=""fa fa-filter""></i> View DTC Diners</button>" & _
							" <button type=""button"" class=""btn btn-" & strButt5 & """ onclick=""self.location='Cards3.asp?Filter2=DTCMC'"";><i class=""fa fa-filter""></i> View DTC MC</button>" & _
							" <button type=""button"" class=""btn btn-" & strButt6 & """ onclick=""self.location='Cards3.asp?Filter2=DPC'"";><i class=""fa fa-filter""></i> View DPC</button></div>" & _
							"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <button type=""button"" class=""btn btn-success float-right"" onclick=""window.open('CardsExcel.asp?GCFO')"";><i class=""fa fa-file-excel-o""></i> Export</button>"
		  
		  %>
		  
		  </div>
		  
		<div class="table-responsive">
            <table class="table table-bordered table-hover" id="dataTable" width="100%" cellspacing="0">
              <thead class="CAPS">
                <tr>
                  <th>Card ID</th>
				  <th Style="background-color:Green; color:white;">Action</th>
                  <th>EID</th>
                  <th>Card Type</th>
				  <th>Card Sub Type</th>
				  <th>Card Number</th>
                  <th>First Name(s)</th>
                  <th>Last Name</th>
                  <th>Address1</th>
				  <th>Suburb</th>
				  <th>Credit Limit</th>
				  <th>Status</th>
				  <th>Loaded Date</th>
				  <th>Expiry Date</th>
				
                </tr>
              </thead>
              <tbody class="CAPS2">
               
				<%
        
      DisplayTableDetails()
        
%>	

              </tbody>
            </table>
          </div>
        </div>
       
    </div>
    



<!--</DIV>-->
</form>


    <!-- jQuery -->
    <script src="../js/jquery.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="../js/bootstrap.min.js"></script>

	
	
</body>
</html>
<%

Public Sub DisplayTableDetails()
Dim y
Dim strAction, strStatus
Dim strWhere, strWhere1
Dim strExpiry
Dim strCreditLimit
Dim strCardNo
Dim strDateLoaded
If Session("Filter") = "Active" Then
	strWhere = " AND [Status] = '00'"
	strWhere1 = " AND [Status] = '00'"
Else
	strWhere = ""
	strWhere1 = ""
End If

If Session("Filter2") = "DPC" Then
	strWhere = strWhere & " AND [CardTypeSub] = 'ANZ'"
	strWhere1 = strWhere1 & " AND [CardTypeSub] = 'ANZ'"
ElseIf Session("Filter2") = "DTC" Then
	strWhere = strWhere & " AND [CardTypeSub] = 'Diners'"
	strWhere1 = strWhere1 & " AND [CardTypeSub] = 'Diners'"
ElseIf Session("Filter2") = "DTCMC" Then
	strWhere = strWhere & " AND [CardTypeSub] = 'Mastercard'"
	strWhere1 = strWhere1 & " AND [CardTypeSub] = 'Mastercard'"
Else
	'strWhere = ""
	'strWhere1 = ""
End If

If Len(strWhere) > 5 Then
	If Instr(1,strWhere,"WHERE") = 0 Then
		strWhere = " WHERE " & Right(strWhere,Len(strWhere)-5)
	Else

	End If
End If

If Session("View") = "All" Then
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	strSQL = "SELECT Top 1000 * FROM qryCards " & strWhere
Else
	strSQL = "SELECT Top 1000 * FROM qryCards WHERE EmployeeID = '" & Session("EmployeeID") & "' " & strWhere1
	'strSQL = "SELECT * FROM qryCards WHERE EmployeeID = " & Session("EmployeeID") & ""
End If

objRS.Open strSQL,objCon
    y = 0
    	
    Do until objRS.EOF 
		
		Select Case objRS("Status")
		
		Case  "Received"
			strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='Cards3.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='Cards3.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
		Case "Added To CS"

			strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CSToDiners.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
		
		Case "Cancelled"

			strAction = "<button type=""button"" title=""Cancelled by the Applicant"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CSToDiners.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-ban""></i> Cancelled</button>"
		
			'strStatus  = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ApplicationsEmployee.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
		
		Case  "Submitted"
			strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='Cards3.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='Cards3.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			
			'strStatus  = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='ApplicationsEmployee.asp?CardID=" & objrs("CardID") & "'"";>Submitted to GCFO</button>"
		Case Else
			'strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='Cards3.asp?Action=Reject&CadID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"

			strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='Cards3.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
		End Select

		If isNull(objRS(15)) Then
			strExpiry = ""
		Else
			strExpiry = Month(objRS(15)) & "/" & Year(objRS(15))
			If Len(strExpiry) = 6 Then strExpiry = "0" & strExpiry
		End If
		
		If isNull(objRS(12)) Then
			strCreditLimit = 0
		Else
			strCreditLimit = FormatCurrency(objRS(12),0)
		End If
		
		If IsNull(objRS("CardNumber")) Then
			strCardNo = ""
		Else
			strCardNo = objRS("CardNumber")
			strCardNumF = objRS("CardNumber")
			If len(strCardNo)>8 Then 
				If Mid(strCardNo,4,1) = "5" Then
					strCardNo = mid(strCardNo,4,2) & "****" & right(strCardNo,4)
				Else
					strCardNo = mid(strCardNo,6,2) & "****" & right(strCardNo,4)
				End If
			End If
		End If
		'Format the Status value and make it user friendly with a name instead of number
		If IsNull(objRS("Status")) Then
				strStatus = ""
			Else
				strStatus = objRS("Status")
				If strStatus = "00" Then
					strStatus = "Active"
					strStatus = "<span class=""label label-success"" Title=""Active""><i class=""fa fa-check""></i></span>"
				Else
					strStatus = "Cancelled"
					strStatus = "<span class=""label label-danger"" Title=""Cancelled""><i class=""fa fa-times""></i></span>"
				End If
			End If
		
		If isNull(objRS("DateLoaded")) Then
			strDateLoaded = 0
		Else
			strDateLoaded = FormatDateTime(objRS("DateLoaded"),vbShortDate)
		End If
		
		Response.Write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""Cards3.asp?CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & strAction & "</a></TD>" & _
				"<TD><a Target=""_self"" HREF=""Cards3.asp?CardID=" & objRS(0) & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""Cards3.asp?CardID=" & objRS(0) & """>" & objRS(2) & "</a></TD>" & _
				"<TD style=""text-align:center;"">" & objRS(3) & "</TD><TD style=""text-align:center;"" Title=""" & objRS("CardNumber") & """>" & strCardNo & "</TD>" & _
				"<TD style=""text-align:center;"">" & objRS(4) & "</TD><TD style=""text-align:center;"">" & objRS(5) & "</TD><TD style=""text-align:center;"">" & objRS(6) & "</TD>" & _
				"<TD style=""text-align:center;"">" & objRS(9) & "</TD><TD style=""text-align:center;"">" & strCreditLimit & "</TD><TD style=""text-align:center;"">" & strStatus & "</TD>" & _
				"<TD style=""text-align:center;"">" & strDateLoaded & "</TD><TD style=""text-align:center;"">" & strExpiry & "</TD></TR>"
				
				'intDaysTotal = intDaysTotal + objRS("Days")
				'intCarsTotal = intCarsTotal + objRS("Cars")
				'intEmpContTotal = intEmpContTotal + objRS("EmployeeContribution")
			
			y = y + 1
			
		objRS.movenext
	Loop
	
	
	response.write "<TR><TH colspan=""10"">Total</TH>" & _
				"<TH colspan=""4"" style=""text-align:center;"">" & y & "</TH></TR>"
				
objRS.Close

End Sub


Sub LoadDetails()

       'Description:	Loads Position details into page if applicable.
		objRS.Open "SELECT * FROM tblCAPSCard WHERE CardID = " & Session("CardID") & "",objCon

			If Not objRS.EOF Then
               
				lngCardID = objRS("CardID")
				strEmployeeID = objRS("EmployeeID")
				strTitle = objRS("Title")
				strFirstName = objRS("FirstName")
				strLastName  = objRS("Surname")
				strAddress1 = objRS("Address1")
				strAddress2 = objRS("Address2")
				strAddress3 = objRS("Address3")
				'strAddress4 = objRS("Address4")
				strSuburb = objRS("Suburb")
				strState = objRS("State")
				strPostCode = objRS("PostCode")
				'dteDateReceived = objRS("DateReceived")
				strStatus = objRS("Status")
				'strReviewedBy = objRS("ReviewedBy")
				'dteDateReviewed = objRS("DateReviewed")
				lngCreditLimit = objRS("CreditLimit")
				strCardType = objRS("CardType")
				
				If left(strCardType,6) = "Diners" Then
					'For Diners cards change the Status to the text rather than code (00)
					If strStatus  = "00" Then
						strStatus2 = "Active"
					Else
						strStatus2 = "Cancelled"
					End If
					
					'Format the Credit Limit for the Diners Cards
					If IsNull(lngCreditLimit) Then
						lngCreditLimit = 0
					Else
						lngCreditLimit = lngCreditLimit/100
						lngCreditLimit = FormatCurrency(lngCreditLimit,0)
					End If
				End If
				
				If strStatus2 = "Active" Then
					strStatusBG = "bg-success"
				Elseif strStatus2 = "Cancelled" Then
					strStatusBG = "bg-danger"
				Else
					strStatusBG = ""
				End If
				
    		Else
				Session("CardID") = 0
			  	lngCardID = 0'objRS("ApplicationID")
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
				'dteDateReceived = ""
				strStatus = ""
				'strReviewedBy = ""
				'dteDateReviewed = ""
				lngCreditLimit =0
           End If

		objRS.Close
	
End Sub

Sub LoadPMData
Dim strSQL
Dim objCon2


on error resume next

	Set objCon2 = Server.CreateObject("ADODB.Connection")
	Session("DBConnection2") = "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	objCon2.ConnectionTimeout=2
	objCon2.Open Session("DBConnection2")



	strSQL = "SELECT TOP 1 Count(dbo.statement_data.reference_number) AS CountOfreference_number, Count(dbo.statement_data.payment_card_id) AS CountOfpayment_card_id, Count(dbo.statement_data.card_account_number) AS CountOfcard_account_number, Count(dbo.statement_data.posting_date) AS CountOfposting_date, dbo.card_account.name_on_card, dbo.procharge_user.employee_id " & _
			"FROM dbo.statement_data WITH(NOLOCK) INNER JOIN (dbo.procharge_user INNER JOIN dbo.card_account ON dbo.procharge_user.user_name = dbo.card_account.user_name) ON dbo.statement_data.card_account_number = dbo.card_account.card_account_number " & _
			"GROUP BY dbo.card_account.name_on_card, dbo.procharge_user.employee_id, dbo.statement_data.card_account_number " & _
			"HAVING Count(dbo.statement_data.posting_date)='' AND dbo.statement_data.card_account_number='" & strCardNumF & "'"
	
	strSQL = "SELECT min(sd.create_date) AS OldestUnaquitted, min(reference_number) as Ref_Number,sd.card_account_number, sd.card_type, " & _
			" sum(case when sd.create_date > dateadd(day,-90,getdate()) then sd.amount else 0 end) as ninety_days " & _
			"FROM statement_data sd with (nolock) " & _
			"left join wf_instance_status w2 with (nolock)on sd.reference_number = w2.instance_id " & _
			"WHERE reference_number not like '%Bal%' and (w2.activity_id = 3 OR w2.activity_id = 4 OR w2.activity_id = 8) AND sd.card_account_number = '" & strCardNumF & "'" & _
			"GROUP BY sd.card_account_number, sd.card_type"
	
	'strSQL = "WITH NonAquitted (OldestUnaquitted,instance_id,card_account_number, card_type) AS " & _
	'	"(SELECT min(sd.create_date) AS OldestUnaquitted, min(reference_number) as Ref_Number,sd.card_account_number, sd.card_type " & _
	'	"FROM statement_data sd with (nolock) left join wf_instance_status w2 with (nolock)on sd.reference_number = w2.instance_id " & _
	'	"WHERE reference_number not like '%Bal%' and (w2.activity_id = 3 OR w2.activity_id = 4 OR w2.activity_id = 8) " & _
	'	"GROUP BY sd.card_account_number, sd.card_type) " & _
	'	" " & _
	'	"SELECT top 1 ca.card_type,  " & _
	'	"isnull(cap.value_string,'') as EID,  " & _
	'	"max(datepart(yyyy,dateadd(hour, 10, sd.create_date))) as year, " & _
	'	"max(datepart(mm,dateadd(hour, 10, sd.create_date))) as month, " & _
	'	"sum(sd.amount) as sum_month, " & _
	'	"max(sd.amount) as max_month, " & _
	'	"sum(case when sd.create_date > dateadd(day,-30,getdate()) then sd.amount else 0 end) as thirty_days, " & _
	'	"sum(case when sd.create_date > dateadd(day,-90,getdate()) then sd.amount else 0 end) as ninety_days, " & _
	'	"sum(case when sd.create_date > dateadd(month,-12,getdate()) then sd.amount else 0 end) as LastYear, " & _
	'	"count(*) as count_year, " & _
	'	"sum(case when sd.create_date > dateadd(day,-30,getdate()) then 1 else 0 end) as count_month, " & _
	'	"MIN(NonAquitted.OldestUnaquitted) as OldestNonAcquitted " & _
	'	"FROM statement_data sd with(nolock)  " & _
	'	"inner join card_account ca with (nolock) on sd.card_type = ca.card_type and sd.card_account_number = ca.card_account_number " & _
	'	"inner join currency_codes cc with (nolock) on sd.card_type = cc.card_type " & _
	'	"inner join wf_instance_status  w with (nolock)on sd.reference_number = w.instance_id " & _
	'	"inner join wf_activities wa with (nolock) on w.activity_id = wa.activity_id " & _
	'	"left join card_account_param cap with (nolock) on ca.card_type = cap.card_type and ca.card_account_number = cap.card_account_number and cap.value_label = 'Cardholder EID' " & _
	'	"left join NonAquitted with (nolock)on sd.card_account_number = NonAquitted.card_account_number and sd.card_type = NonAquitted.card_type " & _
	'	"WHERE cap.value_string = '" & Session("EmployeeID") & "' " & _
	'   "GROUP BY ca.card_type, isnull(cap.value_string,'')"
	
	
	'Description:	Loads Transaction Data from ProMaster if applicable.
	objRS.Open strSQL,objCon2

		If Not objRS.EOF Then
		   
			strTrans = objRS("ninety_days")
			strUnaquitt = objRS("OldestNonAcquitted")
			
		Else
			
			strTrans =""
			strUnaquitt=""
			
	   End If

	objRS.Close
	
	on error goto 0
	
End Sub

Sub SaveCarParking()
Dim strDeclar

	If isNull(Request.Form("Declaration")) Or Request.Form("Declaration") = "" Then
		strDeclar = "checked"
	Else	
		strDeclar = Request.Form("Declaration") 
	End If

  	With objCmd

			.CommandType = 4
			.CommandText = "spCarParkingSave"

			.Parameters.Append objCmd.CreateParameter("CarParkingID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CostCentreID", adInteger, adParamInput)                
			.Parameters.Append objCmd.CreateParameter("EmployeeID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CalculationMethodID", adInteger, adParamInput) 
			.Parameters.Append objCmd.CreateParameter("LocationID",adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Cars", adDouble, adParamInput) 
			.Parameters.Append objCmd.CreateParameter("Days",adDouble, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmployeeContribution", adDouble, adParamInput)
			.Parameters.Append objCmd.CreateParameter("ParkDate", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Declaration", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("Notes", adVarChar, adParamInput, 200)
			.Parameters.Append objCmd.CreateParameter("BenefitType", adVarChar, adParamInput, 20)
			
			.Parameters("CarParkingID") = Request.Form("CarParkingID")
			.Parameters("BudgetID") = Session("BudgetID")	
			.Parameters("VersionID") = Session("VersionID")						
			.Parameters("CostCentreID") = Session("CostCentreID")
			.Parameters("EmployeeID") = Request.Form("EmployeeID")
			.Parameters("CalculationMethodID") = 0'Request.Form("CalculationMethodID") 
			.Parameters("LocationID") = 0'Request.Form("LocationID")             
			.Parameters("Cars") = Request.Form("Cars") 
			.Parameters("Days") = Request.Form("Days") 
			.Parameters("EmployeeContribution") = Request.Form("EmployeeContribution") 
			.Parameters("ParkDate") = Request.Form("ParkDate") 
			.Parameters("Declaration") = strDeclar'Request.Form("Declaration") 
			.Parameters("Notes") = Request.Form("Notes") 
			.Parameters("BenefitType") = "Loan"
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
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
			.Parameters.Append objCmd.CreateParameter("CSToDinernIDOutput", adInteger, adParamOutput)
			
			.Parameters("ApplicationID") = Session("ApplicationID")
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CSToDinernIDOutput") 
	 
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" /> " & intRecord & " Added to CS"
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
