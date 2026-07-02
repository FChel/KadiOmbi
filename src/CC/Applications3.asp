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
Dim objCmd

Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL

Dim lngApplicationID
Dim strEmployeeID
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
Dim strReviewedBy
Dim dteDateReviewed
Dim lngCreditLimit
Dim arrState(8)
Dim strButt1
Dim strButt2

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
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

    Session("CurrentPage") = "CC/Applications3.asp"

	If IsNull(Session("ApplicationID")) OR Session("ApplicationID") = "" Then Session("ApplicationID")= 0

	If Not IsEmpty(Request.QueryString("StyleSheet")) Then
		Session("StyleSheet") = Request.QueryString("StyleSheet")
		
		Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""" & Request.QueryString("StyleSheet") & """>"
	Else
		If IsEmpty(Session("StyleSheet")) Then Session("StyleSheet") = "<link rel=""stylesheet"" type=""text/css"" href=""../CAPSStyle.css"">"
	End If

If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	Session("ApplicationID") = Request.QueryString("ApplicationID")
End If

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	Session("EmployeeID") = Request.QueryString("EmployeeID")
End If

If Not IsEmpty(Request.QueryString("Filter")) Then
	Session("Filter") = Request.QueryString("Filter")
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

  
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "New" Then
			
		End If
	End If
	
  Call LoadDetails()
  
%>

<html>
<head>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

 <!-- Bootstrap Core CSS -->
    <!--<link href="../css/bootstrap.min.css" rel="stylesheet">-->
	<%=Session("StyleSheet")%>
	
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
        <li class="breadcrumb-item active" Style="Background-color:#86c5f9;color:white;font-size:16px;"><i class="fa fa-address-card"></i> All Credit Card Applications </li>
      </ol>
	  <hr>
	  
	  
<form action="Applications3.asp?Action=Save" method="POST" id="frm" name="frm" class="form-inline">

<div class="card-header" style="color:white;font-size:16px;font-weight:bold;background-color:#A9A9A9;height:30px;">&nbsp;&nbsp;&nbsp;
          <i class="fa fa-address-card"></i> Existing Applications for All Employees</div>
<div class="table-responsive">
                                    <table class="table-sm">
<!--<table width="100%" border="0" CELLSPACING="1" CELLPADDING="1">-->
<tr>
<td Style="Background-color:#86c5f9;color:white;">EmployeeID </td><td colspan="2"><input type="select" class="form-control" placeholder="Employee ID" value="<%=strEmployeeID%>" ></td>
<td Style="Background-color:green;color:white;">Card Type</td><td colspan="2"><select class="form-control" placeholder="Card Type">
							<option value="DTC">DTC - Diners</Option><option value="DPC">DPC - ANZ</Option>
							<option value="MC">DTC - Companion Mastercard</Option>
							</select></td>
</tr>
<tr>
<td Style="Background-color:#86c5f9;color:white;">Last Name</td><td colspan="2"><input type="text" class="form-control" placeholder="Last Name" style="width:90%;" value="<%=strLastName%>" ></td>
<td Style="Background-color:#86c5f9;color:white;">First Name</td><td colspan="2"><input type="text" class="form-control" placeholder="First Name" style="width:90%;" value="<%=strFirstName%>" ></td>
</tr>
<tr>
<td class="input-group-addon" Style="Background-color:#86c5f9;color:white;">Address 1</td><td colspan="2"><input type="text" class="form-control" placeholder="Address1" style="width:90%;" value="<%=strAddress1%>" ></td>
<td  class="input-group-addon" Style="Background-color:#86c5f9;color:white;">Address 2</td><td colspan="2"><input type="text" class="form-control" placeholder="Address2" style="width:90%;" value="<%=strAddress2%>" ></td>
</tr>
<tr>
<td class="input-group-addon" Style="Background-color:#86c5f9;color:white;">Address 3</td><td colspan="2"><input type="text" class="form-control" placeholder="Address3" style="width:90%;" value="<%=strAddress3%>" ></td>
<td  class="input-group-addon" Style="Background-color:#86c5f9;color:white;">Address 4</td><td colspan="2"><input type="text" class="form-control" placeholder="Address4" style="width:90%;" value="<%=strAddress4%>" ></td>
</tr>
<tr>
<td class="input-group-addon" Style="Background-color:#86c5f9;color:white;">Suburb</td><td><input type="text" class="form-control" placeholder="Suburb" style="width:90%;" value="<%=strSuburb%>" ></td>
<td  class="input-group-addon" Style="Background-color:#86c5f9;color:white;">State</td><td>
				<select class="form-control" placeholder="State" style="width:90%;">
				<%
					For x = 1 to 8
						If strState = cstr(arrState(x)) Then
							strSelected = " SELECTED "
						Else
							strSelected = ""
						End If
						
						Response.write "<option " & strSelected & " value="" & arrState(x) & "">" & arrState(x) & "</Option>"
							
					Next
				%>
							</select>
</td>
<td  class="input-group-addon" Style="Background-color:#86c5f9;color:white;">Post Code</td><td><input type="text" class="form-control" placeholder="PostCode" style="width:90%;" value="<%=strPostCode%>" ></td>
</tr>

<tr>
<td class="input-group-addon" Style="Background-color:#86c5f9;color:white;">Status</td><td><input type="text" class="form-control" placeholder="Status" style="width:90%;" value="<%=strStatus%>" ></td>
<td class="input-group-addon" Style="Background-color:#86c5f9;color:white;">Credit Limit</td><td><input type="text" class="form-control" placeholder="CreditLimit" style="width:90%;" value="<%=lngCreditLimit%>" ></td>
<td  class="input-group-addon" Style="Background-color:#86c5f9;color:white;" >Date Received</td><td colspan="2"><input type="text" class="form-control" placeholder="Date Received" disabled style="width:90%;" value="<%=dteDateReceived%>" ></td>
</tr>
</table>
</div>
</div>
</div>
<hr>
<a name="Style" id="Style" href="Applications3.asp?StyleSheet=CCStyle.css">CC Style</a>&nbsp;&nbsp;

<a name="Style2" id="Style2" href="Applications3.asp?StyleSheet=../BERTStyle3.css">BERT Style 3</a>&nbsp;&nbsp;

<a name="Style3" id="Style3" href="Applications3.asp?StyleSheet=../BERTStyle.css">BERT Style</a>&nbsp;&nbsp;

<a name="Style4" id="Style4" href="Applications3.asp?StyleSheet=../adminlte.min.css">Admin Lte Style</a>&nbsp;&nbsp;

<a name="Style5" id="Style5" href="Applications3.asp?StyleSheet=../css/bootstrap.min.css">Bootstrap Style</a>&nbsp;&nbsp;

<a name="Style8" id="Style8" href="Applications3.asp?StyleSheet=../css/bootstrap.min2.css">BootstrapStyle2</a>&nbsp;&nbsp;

<a name="Style6" id="Style6" href="Applications3.asp?StyleSheet=../ElegantStyle.css">Elegant Style</a>&nbsp;&nbsp;

<a name="Style7" id="Style7" href="Applications3.asp?StyleSheet=../MonsterStyle.css">Monster Style</a>

<hr>
</form>


<!-- Button trigger modal -->
<button type="button" class="btn btn-primary" data-toggle="modal" data-target="#exampleModalCenter">
  <i class="fa fa-bullhorn"></i> Declaration
</button>

<button type="button" class="btn btn-primary" onclick="CloseScreen()";><i class="fa fa-times"></i> Close</button>
<button type="button" class="btn btn-primary" onclick="SaveData()"; ><i class="fa fa-check-circle"></i> Save</button>
<button type="button" class="btn btn-primary" onclick="self.location='Applications3.asp?Action=New'"; ><i class="fa fa-plus"></i> New</button>
<%=strMessageIcon %>
  
<hr />

      <!-- Breadcrumbs-->
     
      <!-- Example DataTables Card-->
      <div class="card mb-3">
        <div class="card-header" >
          <i class="fa fa-table"></i> Existing Applications for <% Response.write Session("UType") & " - " & Session("UserName")
		  
		  If Session("Filter") = "DCC" Then
			strButt1 = "secondary"
			strButt2 = "info"
		  Else
			strButt1 = "info"
			strButt2 = "secondary"
		  End If
		  
		  Response.write  " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<div class=""btn-group""><button type=""button"" class=""btn btn-" & strButt1 & """ onclick=""self.location='Applications3.asp?Filter=All'"";><i class=""fa fa-filter""></i> View All </button>" & _
							" <button type=""button"" class=""btn btn-" & strButt2 & """ onclick=""self.location='Applications3.asp?Filter=DCC'"";><i class=""fa fa-filter""></i> View Awaiting Approval</button></div>" & _
							"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <button type=""button"" class=""btn btn-success float-right"" onclick=""window.open('ApplicationsExcel.asp?DCC')"";><i class=""fa fa-file-excel-o""></i> Export</button>"
		  
		  
		  %></div>
        <div class="card-body">
          <div class="table-responsive">
            <table class="table table-bordered table-hover" id="dataTable" width="100%" cellspacing="0">
              <thead>
                <tr>
                  <th Style="background-color:#2394F2; color:white;">App ID</th>
				  <th Style="background-color:Green; color:white;">Action</th>
                  <th Style="background-color:#2394F2; color:white;">Card Type</th>
                  <th Style="background-color:#2394F2; color:white;">Card Sub Type</th>
				  <th Style="background-color:#2394F2; color:white;">EID</th>
                  <th Style="background-color:#2394F2; color:white;">Title</th>
                  <th Style="background-color:#2394F2; color:white;">First Name</th>
                  <th Style="background-color:#2394F2; color:white;">Last Name</th>
				  <th Style="background-color:#2394F2; color:white;">Address1</th>
				  <th Style="background-color:#2394F2; color:white;">Suburb</th>
				  <th Style="background-color:#2394F2; color:white;">Status</th>
				  <th Style="background-color:#2394F2; color:white;">Reviewed Date</th>
				  <th Style="background-color:#2394F2; color:white;">Reviewed By</th>
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
Dim dblEmpCont
Dim intDaysTotal
Dim intCarsTotal
Dim intEmpContTotal
Dim strAction, strStatus
Dim strWhere
Dim dteDateReviewed

If Session("Filter") = "DCC" Then
	strWhere = "WHERE [ReviewedBy] IS NULL OR [ReviewedBy] = ''"

End If

If Session("EmployeeID") = "" Then
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	strSQL = "SELECT * FROM qryApplications " & strWhere
Else
	strSQL = "SELECT * FROM qryApplications " & strWhere
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
End If

objRS.Open strSQL,objCon
    y = 0
    	
    Do until objRS.EOF 
		If isNull(objRS(9)) Then
			dblEmpCont = 0
		Else
			dblEmpCont = objRS(9)
		End If

		If isNull(objRS("DateReviewed")) OR objRS("DateReviewed") = "" Then
			dteDateReviewed = ""
		Else
			dteDateReviewed = FormatDateTime(objRS(16),vbShortDate)
		End If
		
		
		Select Case objRS("Status")
		
		Case  "Received"
			strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='Applications3.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications3.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
		Case "Added To CS"

			strAction = "<button type=""button"" class=""btn btn-secondary"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
		
		Case "Cancelled"

			strAction = "<button type=""button"" title=""Cancelled by the Applicant"" class=""btn btn-secondary"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-ban""></i> Cancelled</button>"
		
			strStatus  = "<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployee.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Cancelled By Applicant</button>"
		
		Case  "Submitted"
			strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='Applications3.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications3.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			
			strAction = "<button type=""button"" title=""Awaiting Approval by the GCFO"" class=""btn btn-secondary"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-bank""></i> Waiting GCFO Approval</button>"
		
		
			strStatus  = "<button type=""button"" class=""btn btn-success"" onclick=""self.location='ApplicationsEmployee.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted to GCFO</button>"
		
		Case "Approved"

			strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='Applications3.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications3.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"

			strStatus  = "<button type=""button"" class=""btn btn-secondary"" onclick=""self.location='Applications3GCFO.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Approved by GCFO</button>"

		Case "GCFO Approved"
		
			strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='Applications3.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications3.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			
			strStatus  = "<button type=""button"" class=""btn btn-secondary"" onclick=""self.location='Applications3.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Approved by GCFO</button>"
		
		
		Case Else
			strAction = "<button type=""button"" class=""btn btn-secondary"" onclick=""self.location='Applications3.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Rejected</button>"
			'strAction = "Rejected"
			strStatus  = "<button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications3.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Rejected</button>"
		End Select

		response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""Applications3.asp?ApplicationID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & strAction & "</a></TD>" & _
				"<TD><a Target=""_self"" HREF=""Applications3.asp?ApplicationID=" & objRS(0) & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""Applications3.asp?ApplicationID=" & objRS(0) & """>" & objRS(2) & "</a></TD>" & _
				"<TD style=""text-align:center;"">" & objRS(3) & "</TD><TD style=""text-align:center;"">" & objRS(4) & "</TD>" & _
				"<TD style=""text-align:center;"">" & objRS(5) & "</TD><TD style=""text-align:center;"">" & objRS(6) & "</TD>" & _
				"<TD style=""text-align:center;"">" & objRS(7) & "</TD><TD style=""text-align:center;"">" & objRS(10) & "</TD>" & _
				"<TD style=""text-align:center;"">" & strStatus & "</TD><TD style=""text-align:center;"">" & dteDateReviewed & "</TD><TD style=""text-align:center;"">" & objRS(15) & "</TD></TR>"
				
				'intDaysTotal = intDaysTotal + objRS("Days")
				'intCarsTotal = intCarsTotal + objRS("Cars")
				'intEmpContTotal = intEmpContTotal + objRS("EmployeeContribution")
			
			y = y + 1
			
		objRS.movenext
	Loop
	
	
	response.write "<TR><TH colspan=""4"">Total</TH>" & _
				"<TH colspan=""8"" style=""text-align:center;"">" & y & "</TH></TR>"
				
objRS.Close

End Sub


Sub LoadDetails()

       'Description:	Loads Position details into page if applicable.
		objRS.Open "SELECT * FROM tblApplication WHERE ApplicationID = " & Session("ApplicationID") & "",objCon

			If Not objRS.EOF Then
               
				lngApplicationID = objRS("ApplicationID")
				strEmployeeID = objRS("EmployeeID")
				strFirstName = objRS("FirstName")
				strLastName  = objRS("Surname")
				strAddress1 = objRS("Address1")
				strAddress2 = objRS("Address2")
				strAddress3 = objRS("Address3")
				'strAddress4 = objRS("Address4")
				strSuburb = objRS("Suburb")
				strState = objRS("State")
				strPostCode = objRS("PostCode")
				dteDateReceived = objRS("DateReceived")
				strStatus = objRS("Status")
				strReviewedBy = objRS("ReviewedBy")
				dteDateReviewed = objRS("DateReviewed")
				lngCreditLimit = objRS("CreditLimit")
    		Else
				Session("ApplicationID") = 0
			  	lngApplicationID = 0'objRS("ApplicationID")
				strEmployeeID = ""
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
				lngCreditLimit =0
           End If

		objRS.Close
	
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
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CSToDinernIDOutput", adInteger, adParamOutput)
			
			.Parameters("ApplicationID") = Session("ApplicationID")
			.Parameters("UpdatedBy") = Session("Logon")
			
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
 ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "JulieMay88@iinet.net.au"'"tiamichael@grapevine.net.au"'"noreply@bizbudg.com"'"service@bizbudg.com.au"
 ObjSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "wAiZpuJ8xeun"'"chicken"'"ir3land"
      
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

Set objRS = Nothing
Set objCon = Nothing
%>
