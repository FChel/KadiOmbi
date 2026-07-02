
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=../ADOVBS.inc -->
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
    Set objRS1 = Server.CreateObject("ADODB.Recordset")
    Set objRS2 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	

    Session("CurrentPage") = "CC/ApplicationsEmployeeHF.asp"

	If IsNull(Session("ApplicationID")) OR Session("ApplicationID") = "" Then Session("ApplicationID")= 0

If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	Session("ApplicationID") = Request.QueryString("ApplicationID")
End If

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	Session("EmployeeID") = Request.QueryString("EmployeeID")
	Session("CarParkingID") = 0
End If


If Not IsEmpty(Request.QueryString("TransactionType")) Then
	Session("TransactionType") = Request.QueryString("TransactionType")
End If
 Session("InputSheetID") = 1
 

If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "Cancel" Then
		Call CancelApplication()
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
		If Request.QueryString("Action") = "SubmitApp" Then
			'Response.Write "CPID=" & Session("CarParkingID")
			'Session("ApplicationID") = 0
			Call SubmitApplication()
		End If
	End If
	
	If isNull(Session("ApplicationID")) Or Session("ApplicationID") = "" Then 
		Session("ApplicationID") = 0
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
      
      $.post('ApplicationsEmployeeHF.asp?Action=SubmitApp', 
         $('#frm').serialize(), 
      /*   function(data, status, xhr){
           // do something here with response;
         });
      */
    });
});

</script>
	
</head>
<body >
<main class="main py-1">
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
					
<!-- End the first part of the Header Container -->

<form action="ApplicationsEmployeeHF.asp?Action=SubmitApp" method="POST" id="frm" name="frm" class="inline">


<div id='tbl-container'>

<div class="content-wrapper">
   	  
<div class="card-header" >&nbsp;&nbsp;&nbsp;
    <i class="fa fa-address-book"></i> Corporate Directory Details for <%=Session("EmployeeID") & " - " & Session("UserName")%> (Employee details used for card application)
</div>
</form>
  
<form class="form-inline" action="ApplicationsEmployeeHF.asp?Action=SubmitApp" id="frm2" name="frm2"> 
<div class="container-fluid">


	<div class="input-group col-sm-12">
		<div class="form-row col-12"><div class="form-group col-md-6">
			<label for="EmployeeID">EmployeeID</label><input type="text" style="width:100%;" class="form-control" id="EmployeeID" placeholder="EmployeeID" value="<%=Session("EmployeeID")%>"></div>
			<div class="form-group col-sm-6">
			<label for="EmployeeID">Card Type</label>
			<select class="form-control" placeholder="Card Type" style="width:100%;" name="CardTypeSelect" id="CardTypeSelect" onchange="valCard();">
					<option value="DTC - Diners">DTC - Diners</Option><option value="DPC - ANZ">DPC - ANZ</Option>
					<option value="DTC - MasterCard">DTC - Companion Mastercard</Option>
					</select></div>
			</div><div class="form-row col-12"><div class="form-group col-md-2">
			<label for="Title">Title</label><input type="text" style="width:100%;" class="form-control" id="Title" placeholder="Title" value="<%=strTitle%>"/></div>
			<div class="form-group col-md-5">
			<label for="FirstName">FirstName</label><input type="text" style="width:100%;" class="form-control" id="FirstName" placeholder="FirstName" value="<%=strFirstName %>" DISABLED/></div>
			<div class="form-group col-md-5">
			<label for="LastName">LastName</label><input type="text" style="width:100%;" class="form-control" id="LastName" placeholder="LastName" value="<%=strLastName%>" DISABLED /></div>
			</div><div class="form-row col-12">
			<div class="form-group col-md-6">
			<label for="Address1">Address1</label><input type="text" style="width:100%;" class="form-control" id="Address1" placeholder="Address1" value="<%=strAddress1%>" DISABLED title="Can Only be Updated in the Corporate Directory"/></div>
			<div class="form-group col-md-6">
			<label for="Address2">Address2</label><input type="text" style="width:100%;" class="form-control" id="Address2" placeholder="Address2" value="<%=strAddress2 %>" DISABLED title="Can Only be Updated in the Corporate Directory"/></div>
			</div><div class="form-row col-12">
			<div class="form-group col-md-6">
			<label for="Address3">Address3</label><input type="text" style="width:100%;" class="form-control" id="Address3" placeholder="Address3" value="<%=strAddress3 %>" DISABLED title="Can Only be Updated in the Corporate Directory"/></div>
			<div class="form-group col-md-6">
			<label for="Suburb">Suburb</label><input type="text" style="width:100%;" class="form-control" id="Suburb" placeholder="Suburb" value="<%=strSuburb%>" DISABLED title="Can Only be Updated in the Corporate Directory"/></div>
			</div><div class="form-row col-12">
			<div class="form-group col-md-6">
			<label for="State">State</label><input type="text" style="width:100%;" class="form-control" id="State" placeholder="State" value="<%=strState %>" DISABLED title="Can Only be Updated in the Corporate Directory"/></div>
			<div class="form-group col-md-6">
			<label for="PostCode">PostCode</label><input type="text" style="width:100%;" class="form-control" id="PostCode" placeholder="PostCode" value="<%=strPostCode %>" DISABLED title="Can Only be Updated in the Corporate Directory"/></div>
			</div><div class="form-row col-12">
			<div class="form-group col-md-6">
			<label for="WorkPhone">WorkPhone</label><input type="text" style="width:100%;" class="form-control" id="WorkPhone" placeholder="WorkPhone" value="<%=strWorkPhone %>" DISABLED title="Can Only be Updated in the Corporate Directory"/></div>
			<div class="form-group col-md-6">
			<label for="MobilePhone">MobilePhone</label><input type="text" style="width:100%;" class="form-control" id="MobilePhone" placeholder="MobilePhone" value="<%=strMobilePhone%>" DISABLED title="Can Only be Updated in the Corporate Directory"/></div>
			</div>
		</div>				

	</div>
  </form>
 
</div>
<br>
<div class="container-fluid">
	<div class="alert alert-danger">
	  <strong>Note!</strong> Details cannot be changed here. All Address and Contact details MUST be updated in the <a href="http://directory/dcd/" target="_new">Corporate Directory</a>. Updates may take up to 3 days to display above.
	</div>

	<hr>
</div>
</div>
</form>

<div class="container-fluid">
	<div class="row">
		<div class="col-md-8">
		<!-- Button trigger modal -->
		<button type="button" class="btn btn-primary" data-toggle="modal" data-target="#exampleModalCenter" onClick="OpenSs(this);"><i class="fa fa-check-circle"></i>
		  Submit Application
		</button>

		<button type="button" class="btn btn-primary" onclick="CloseScreen()";><i class="fa fa-times"></i> Close</button>
		<!--<button type="button" class="btn btn-primary" onclick="SaveData()"; ><img src="../images/tick.png" alt="" /> Save</button>
		<button type="button" class="btn btn-primary" onclick="self.location='ApplicationsEmployeeHF.asp?Action=New'"; ><img src="../images/add.png" alt="" /> New</button>-->
		<%=strMessageIcon %>
		</div>
	</div>
</div>
<hr />


    
	<div class="container-fluid">
	 <section class="table py-2">
        <h4 class="text-left">Existing Applications for <%=Session("UserName")%></h4>
        <div class="container">
          <div class="row">
            <div class="col-12">
              <table class="table">
                <thead>
                  <tr>
					<th scope="col">App ID</th>
					<th scope="col">EID</th>
					<th scope="col">Name</th>
					<th scope="col">Card Type</th>
					<th scope="col">Address</th>
					<th scope="col">Status</th>
					<th scope="col">Submitted</th>
					<th scope="col">Reviewed</th>
					<th scope="col">Action</th>
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
Dim dblEmpCont
Dim intDaysTotal
Dim intCarsTotal
Dim intEmpContTotal
Dim strAction
Dim strStatus
Dim strAddress
Dim dteDateSubmitted
Dim dteDateReviewed

If Session("EmployeeID") = "" OR ISNull(Session("EmployeeID")) Then
	strSQL = "SELECT * FROM qryCAPSApplications WHERE EmployeeID = '" & Session("EmployeeID") & "'"
Else
	strSQL = "SELECT * FROM qryCAPSApplications WHERE EmployeeID = '" & Session("EmployeeID") & "'"
End If

objRS.Open strSQL,objCon

    y = 0
	
	'Write a message in the list if there are no applications
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Applications for " & Session("UserName") & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	End If
    	
    Do until objRS.EOF 
		If isNull(objRS(9)) Then
			dblEmpCont = 0
		Else
			dblEmpCont = objRS(9)
		End If

		Select Case objRS("Status")
		
		Case  "Received"
			strAction = "<button type=""button"" class=""btn btn-primary btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
		Case "Added To CS"

			strAction = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
		
		Case "Submitted"
			strAction = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"

			strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted to GCFO</button>"
		Case "Cancelled"
			strAction = "Cancelled - " & FormatDateTime(objRS("DateUpdated"),vbShortDate)'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"

			strStatus  = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Cancelled By Applicant</button>"
		
		Case "GCFO Approved"
			strAction = "<button type=""button"" title=""Approved by GCFO"" class=""btn btn-secondary btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-check""></i>GCFO Approved</button>"
		
			strStatus  = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Approved by GCFO</button>"
		
		Case Else
			strAction = "<button type=""button"" class=""btn btn-danger btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
			'strAction = "Rejected"
			strStatus  = "<button type=""button"" class=""btn btn-success btn-sm"" onclick=""self.location='ApplicationsEmployeeHF.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted</button>"
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
		
		response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?ApplicationID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & objRS("EmployeeID") & "</a></TD>" & _
				"<TD><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?ApplicationID=" & objRS(0) & """>" & objRS("FirstName") & " " & objRS("Surname") & "</a></TD><TD><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?ApplicationID=" & objRS(0) & """>" & objRS("CardType") & " " & objRS("CardTypeSub") & "</a></TD>" & _
				"<TD style=""text-align:center;"">" & strAddress & "</TD><TD style=""text-align:center;"">" & strStatus & "</TD>" & _
				"<TD style=""text-align:center;"">" & dteDateSubmitted & "</TD><TD style=""text-align:center;"">" & dteDateReviewed & "</TD>" & _
				"<TD style=""text-align:center;"">" & strAction & "</TD></TR>"
				
		'response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?ApplicationID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & strAction & "</a></TD>" & _
		'		"<TD><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?ApplicationID=" & objRS(0) & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""ApplicationsEmployeeHF.asp?ApplicationID=" & objRS(0) & """>" & objRS(2) & "</a></TD>" & _
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

End Sub


Sub LoadDetails()

       'Description:	Loads Position details into page if applicable.
		objRS.Open "SELECT * FROM tblCAPSCDMC WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeID") & "'",objCon

			If Not objRS.EOF Then
               
				'lngApplicationID = objRS("ApplicationID")
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
           End If

		objRS.Close
	
End Sub

Public Sub CancelApplication()

	strSQL = "UPDATE tblApplication SET Status = 'Cancelled' WHERE ApplicationID = " & Session("ApplicationID") & ""
	
	objCon.Execute strSQL
	
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
			.Parameters.Append objCmd.CreateParameter("CDMCToApplicationIDOutput", adInteger, adParamOutput)
			
			.Parameters("EmployeeID") = Session("EmployeeID")
			.Parameters("CardType") = Left(Request.Form("CardType"),3)
			.Parameters("CreditLimit") = Request.Form("CreditLimit")
			.Parameters("CardTypeSub") = Right(Request.Form("CardType"),Len(Request.Form("CardType"))-6)
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCToApplicationIDOutput") 
	 
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
