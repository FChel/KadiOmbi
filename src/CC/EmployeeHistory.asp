<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=CAPSFunctions.asp -->
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
Dim objRS2

Dim strForeColour
Dim intMode
Dim dblTotal
Dim intFinYearPart1
Dim intFinYearPart2
Dim arrHeadings(5)
Dim strCostCentreName
Dim strVersionName
Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL

Dim intEmployeeID

Dim lngApplicationID
Dim strEmployeeID
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
Dim lngCDMCID

Dim intApplicationsCount
Dim intCSToCount
Dim intCSFromCount
Dim intCardsCount
Dim intCDMCCount
Dim intCMToCount

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objRS2 = Server.CreateObject("ADODB.Recordset")
	
    objCon.Open Session("DBConnection")	

	If IsNull(Session("CardID")) OR Session("CardID") = "" Then Session("CardID")= 0

	If Not IsEmpty(Request.QueryString("SearchAll")) Then
		Session("SearchAll") = Request.QueryString("SearchAll")
	End If

If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	Session("ApplicationID") = Request.QueryString("ApplicationID")
End If

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	Session("EmployeeID") = Request.QueryString("EmployeeID")
End If

If Not IsEmpty(Request.QueryString("EmployeeSearchID")) Then
	Session("EmployeeSearchID") = Request.QueryString("EmployeeSearchID")
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
		If Request.QueryString("Action") = "Search" Then
			 Session("EmployeeSearchID") = Request.QueryString("EmployeeID")
		End If
	End If

  Call LoadDetails()
  call GetCounts()
  
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

function AppSearch() {
   
    self.location = "EmployeeHistory.asp?Action=Search&EmployeeID2=" + frm.EmployeeIDSearch.value;
      
}

function SelectApp(varAppID) {
	if(varAppID==undefined) {
		alert(varAppID);
	}
	{
	self.location = "EmployeeHistory.asp?Action=Search&ApplicationID=" + varAppID;
	}
}

function SelectEmp(varEmpID) {
	if(varEmpID==undefined) {
		alert(varEmpID);
	}
	{
	//self.location = "EmployeeHistory.asp?Action=Search&EmployeeID=" + varEmpID;
	self.location = "EmployeeHistory.asp?Action=Search&EmployeeID=" + varEmpID + "&FileSeqNum=1"
	//self.location = "EmployeeHistory.asp?Action=Search&EmployeeID=" + varEmpID + "&FileSeqNum=" + document.getElementById('FileAnchor').text;
	}
}


function loadDoc() {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("EmpSear").innerHTML = this.responseText;
    }
  };
  //xhttp.open("GET", "GetEmployees.asp?EmpID=" + frm.EmpIDS.value + "&FName=" + frm.FirstName.value + "&LName=" + frm.LNamms.value + "", true);
  xhttp.open("GET", "AJAX/GetEmployees.asp?EmpID=" + frm.EmpIDS.value + "&FName=" + frm.FirstName.value + "&LName=" + frm.LNamms.value + "", true);
  xhttp.send();
}

function loadXML(varID) {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("XMLDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetXMLDetails.asp?ApplicationID=" + varID + "", true);
  xhttp.send();
}

</script>
	
<body >

<form action="EmployeeHistory.asp?Action=Save" method="POST" id="frm" name="frm">

<!-- Modal -->
<div class="modal fade" id="exampleModalCenter" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle"> Search for an Employee</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>Enter details the click Search...</label><br>
            <label for="FirstName">First Name:</label>
            <input type="text" name="FirstName" id="FirstName" class="form-control input-md">
			<label for="LNamms">Last Name:</label>
            <input type="text" name="LNamms" id="LNamms" class="form-control input-md">
			<label for="EmpIDS">Employee ID:</label>
            <input type="email" name="EmpIDS" id="EmpIDS" class="form-control input-md">
           
      </div>
	  <div id="EmpSear">
	  
	  </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
        <button type="button" class="btn btn-primary" onClick="loadDoc()">Search</button>
      </div>
	  <table><tr><td></td><td></td><td></td></tr></table>
    </div>
  </div>
</div>
</form>

<!-- XML APplicaiton Modal -->
<div class="modal fade" id="XMLModal" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-dialog-right modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  XML Application Detail
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="XMLDetail">
               
				  
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
 <!-- END XML Applicaiton Modal -->

    <main class="main py-5">
      <div class="container">
        <div class="row">
          <div class="col-md-6">
            <h4 class="py-2">Employee History: <span style="font-size:18px; color:grey; font-style:italic; font-weight:lighter;"><% Response.Write Session("EmployeeSearchID") & " - " & strFirstName & " " & strLastName%></span></h4>
          </div>
          <div class="col-md-6 text-right">
			<button type="button" class="btn btn-primary" data-toggle="modal" data-target="#exampleModalCenter"><i class="fa fa-search"></i> Search for Employee</button>
            <div class="btn-group btn-selector table-tabs-selector" role="group" aria-label="Basic example">
              <button type="button" data-target="table-tabs" data-type="as-tabs" class="btn btn-outline-primary active">
                <i class="fa fa-list"></i> View as Tabs
              </button>
              <button type="button" data-target="table-tabs" data-type="as-table" class="btn btn-outline-primary">
                <i class="fa fa-table"></i> View as Table
              </button>
            </div>
          </div>
        </div>
        <div class="row">
          <div class="col-12">
            <div id="table-tabs" class="as-tabs">
              <ul class="nav nav-tabs" id="myFiTab" role="tablist">
                <li class="nav-item" role="presentation">
                  <a class="nav-link active" id="applications-tab" data-toggle="tab" href="#applications" role="tab" aria-controls="applications" aria-selected="true">Applications <span class="counter"><%=intApplicationsCount%></span></a>
                </li>
				<li class="nav-item" role="presentation">
                  <a class="nav-link" id="cm-to-nab-tab" data-toggle="tab" href="#cm-to-nab" role="tab" aria-controls="cm-to-nab" aria-selected="false">CM To NAB <span class="counter"><%=intCMToCount%></span></a>
                </li>
                <li class="nav-item" role="presentation">
                  <a class="nav-link" id="cs-to-diners-tab" data-toggle="tab" href="#cs-to-diners" role="tab" aria-controls="cs-to-diners" aria-selected="false">CS To Diners <span class="counter"><%=intCSToCount%></span></a>
                </li>
                <li class="nav-item" role="presentation">
                  <a class="nav-link" id="cs-from-diners-tab" data-toggle="tab" href="#cs-from-diners" role="tab" aria-controls="cs-from-diner" aria-selected="false">CS From Diners <span class="counter"><%=intCSFromCount%></span></a>
                </li>
                <li class="nav-item" role="presentation">
                  <a class="nav-link" id="cards-tab" data-toggle="tab" href="#cards" role="tab" aria-controls="card" aria-selected="false">Cards <span class="counter"><%=intCardsCount%></span></a>
                </li>
                <li class="nav-item" role="presentation">
                  <a class="nav-link" id="cdmc-tab" data-toggle="tab" href="#cdmc" role="tab" aria-controls="cdmc" aria-selected="false">CDMC <span class="counter"><%=intCDMCCount%></span></a>
                </li>
              </ul>

              <div class="tab-content" id="myFiTabContent">
                <div class="tab-pane fade show active" id="applications" role="tabpanel" aria-labelledby="applications">
                  <h5 class="py-2">Applications</h5>
                  <table class="table">
                    <thead>
                      <tr>
						<th scope="col">App ID</th>
                        <th scope="col">Card Type / EID</th>
                        <th scope="col">Application Status</th>
                        <th scope="col">Name On Card</th>
						<th scope="col">Application Type</th>
                        <th scope="col">Date Submitted</th>
                        <th scope="col">Action</th>
                      </tr>
                    </thead>
                    <tbody>
					<% Call DisplayTableDetailsApps()%>
                      
                    </tbody>
                  </table>
                </div>

				   <div class="tab-pane fade" id="cm-to-nab" role="tabpanel" aria-labelledby="cm-to-nab">
                  <h5 class="py-2">CM To NAB</h5>
                  <table class="table">
                    <thead>
                      <tr>
                        <th scope="col">Card Type / EID</th>
                        <th scope="col">Card No / Name</th>
                        <th scope="col">Address</th>
                        <th scope="col">Date</th>
                        <th scope="col">Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      <% Call DisplayTableDetailsCMTo()%>
                    </tbody>
                  </table>
                </div>
				
                <div class="tab-pane fade" id="cs-to-diners" role="tabpanel" aria-labelledby="cs-to-diners">
                  <h5 class="py-2">CS To Diners</h5>
                  <table class="table">
                    <thead>
                      <tr>
                        <th scope="col">Card Type / EID</th>
                        <th scope="col">Card No / Name</th>
                        <th scope="col">Application ID / Address</th>
                        <th scope="col">Date</th>
                        <th scope="col">Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      <% Call DisplayTableDetailsCSTo()%>
                    </tbody>
                  </table>
                </div>

                <div class="tab-pane fade" id="cs-from-diners" role="tabpanel" aria-labelledby="cs-from-diners">
                  <h5 class="py-2">CS From Diners</h5>
                  <table class="table">
                    <thead>
                      <tr>
                        <th scope="col">Card Type / EID</th>
                        <th scope="col">Card No / Name</th>
                        <th scope="col">Application ID / Address</th>
                        <th scope="col">Date</th>
                        <th scope="col">Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      <% Call DisplayTableDetailsCSFrom()%>
                    </tbody>
                  </table>
                </div>

                <div class="tab-pane fade" id="cards" role="tabpanel" aria-labelledby="cards">
                  <h5 class="py-2">Cards</h5>
                  <table class="table">
                    <thead>
                      <tr>
						<th scope="col">Card ID</th>
                        <th scope="col">Card Type / EID</th>
                        <th scope="col">Status</th>
                        <th scope="col">Card Number</th>
						<th scope="col">Credit Limit</th>
                        <th scope="col">Date Issued</th>
						<th scope="col">Expiry Date</th>
                        <th scope="col">Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      <% Call DisplayTableDetailsCards()%>
                    </tbody>
                  </table>
                </div>

                <div class="tab-pane fade" id="cdmc" role="tabpanel" aria-labelledby="cdmc">
                  <h5 class="py-2">CDMC</h5>
                  <table class="table">
                    <thead>
                      <tr>
                        <th scope="col" style="font-size:13px;">EID</th>
                        <th scope="col" style="font-size:13px;">Name</th>
                        <th scope="col" style="font-size:13px;">Formatted Address</th>
						<th scope="col" style="font-size:13px;">Postal Message</th>
						<th scope="col" style="font-size:13px;">Valid Post Address?</th>
                        <th scope="col" style="font-size:13px;">Date Updated</th>
			<th scope="col" title="If the Employee was on the most recent CDMC file" style="font-size:13px;">Active Employee</th>
			<th scope="col" title="If the Employee has left Defence then the Date they were last on the HR file will appear below" style="font-size:13px;">Left Defence?</th>
                        <th scope="col" style="font-size:13px;">Action</th>
                      </tr>
                    </thead>
                    <tbody>
                      <% Call DisplayTableDetailsCDMC()%>
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </main>
	
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
Dim strAction, strStatus

EXIT SUB
If Session("SearchAll") = "" or IsNull(Session("SearchAll")) Then
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	strSQL = "SELECT * FROM qryCAPSCDMC WITH(NOLOCK)"
Else
	strSQL = "SELECT * FROM qryCAPSCDMC WITH(NOLOCK) WHERE ([Surname] Like '%" & Session("SearchAll") & "%' OR [FirstName] Like '%" & Session("SearchAll") & "%' OR [EmployeeID] Like '%" & Session("SearchAll") & "%')"
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

		Select Case objRS("Status")
		
		Case  "Received"
			strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='Cards3.asp?Action=Release&EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='Cards3.asp?Action=Reject&EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
		Case "Added To CS"

			strAction = "<button type=""button"" class=""btn btn-secondary"" onclick=""self.location='CSToDiners.asp?EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
		
		Case "Cancelled"

			strAction = "<button type=""button"" title=""Cancelled by the Applicant"" class=""btn btn-secondary"" onclick=""self.location='CSToDiners.asp?EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-ban""></i> Cancelled</button>"
		
			strStatus  = "<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployee.asp?EmployeeID=" & objrs("EmployeeID") & "'"";>Cancelled By Applicant</button>"
		
		Case  "Submitted"
			strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='Cards3.asp?Action=Release&EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='Cards3.asp?Action=Reject&EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			
			strStatus  = "<button type=""button"" class=""btn btn-success"" onclick=""self.location='ApplicationsEmployee.asp?EmployeeID=" & objrs("EmployeeID") & "'"";>Submitted to GCFO</button>"
		Case Else
			strAction = "<button type=""button"" class=""btn btn-danger"" onclick=""self.location='Cards3.asp?Action=Reject&EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			'strAction = "Rejected"
		End Select

		response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""Cards3.asp?EmployeeID=" & objRS(0) & """>" & objRS(0) & "</a></TD>" & _
				"<TD><a Target=""_self"" HREF=""Cards3.asp?EmployeeID=" & objRS(0) & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""Cards3.asp?EmployeeID=" & objRS(0) & """>" & objRS(2) & "</a></TD>" & _
				"<TD style=""text-align:center;"">" & objRS(3) & "</TD><TD style=""text-align:center;"">" & objRS(4) & "</TD>" & _
				"<TD style=""text-align:center;"">" & objRS(5) & "</TD><TD style=""text-align:center;"">" & objRS(6) & "</TD>" & _
				"<TD style=""text-align:center;"">" & objRS(7) & "</TD><TD style=""text-align:center;"">" & objRS(8) & "</TD>"  & _
				"<TD style=""text-align:center;"">" & objRS(9) & "</TD><TD style=""text-align:center;"">" & objRS(10) & "</TD></TR>" 
				'"<TD style=""text-align:center;"">" & strStatus & "</TD><TD style=""text-align:center;"">" & objRS(14) & "</TD><TD style=""text-align:center;"">" & objRS(15) & "</TD></TR>"
				
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
	objRS.Open "SELECT * FROM tblCAPSCDMC WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeSearchID") & "' ",objCon

		If Not objRS.EOF Then
		   
			lngCDMCID = objRS("CDMCID")
			strEmployeeID = objRS("EmployeeID")
			strFirstName = objRS("FirstName")
			strLastName  = objRS("Surname")
			strAddress1 = objRS("Addressline1")
			strAddress2 = objRS("Addressline2")
			strAddress3 = objRS("Addressline3")
			'strAddress4 = objRS("Address4")
			strSuburb = objRS("PostalAddress_City")
			strState = objRS("PostalAddress_State")
			strPostCode = objRS("PostalAddress_PostCode")
			'dteDateReceived = objRS("DateReceived")
			'strStatus = objRS("Status")
			'strReviewedBy = objRS("ReviewedBy")
			'dteDateReviewed = objRS("DateReviewed")
			'lngCreditLimit = objRS("CreditLimit")
		Else
			Session("ApplicationID") = 0
			lngCDMCID = 0'objRS("ApplicationID")
			strEmployeeID = ""
			strFirstName = "No Employee Selected"
			strLastName  = ""
			strAddress1 = ""
			strAddress2 = ""
			strAddress3 = ""
			strAddress4 = ""
			strSuburb = ""
			strState = ""
			strPostCode = ""
			'dteDateReceived = ""
			'strStatus = ""
			'strReviewedBy = ""
			'dteDateReviewed = ""
			'lngCreditLimit =0
	   End If

	objRS.Close

End Sub

Public Sub DisplayTableDetailsApps()
'Procedure to load all of the Application records for the selected employee

Dim strAction, strStatus
Dim strCardType
Dim strWhere
Dim dteDateUpdated
Dim strViewAppType
Dim strViewAppTitle
Dim strApplicationType
Dim strXMLButton

'If IsNull(Session("ApplicationID")) or Session("ApplicationID") = "" OR Session("ApplicationID") = 0 Then

'Else
'	strWhere = " AND ApplicationID = " & Session("ApplicationID")
'End If

If Session("SearchAll") = "" or IsNull(Session("SearchAll")) Then
	strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeSearchID") & "'" & strWhere & " ORDER BY ApplicationID DESC"
	'strSQL = "SELECT * FROM qryApplications"
Else
	'strSQL = "SELECT * FROM qryApplications WHERE ([Surname] Like '%" & Session("SearchAll") & "%' OR [FirstName] Like '%" & Session("SearchAll") & "%' OR [EmployeeID] Like '%" & Session("SearchAll") & "%')"
	strSQL = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeSearchID") & "'" & strWhere & " ORDER BY ApplicationID DESC"
End If


objRS.Open strSQL,objCon
	'If the selected employee has no applications then display a No Records record
	If objRS.Eof Then
	
		Response.Write "<tr><th scope=""row""><i>No Applications</i></th><td></td><td></td> " & _
				"<td></td><td></td></tr>"
	End If
	
    Do until objRS.EOF 
	
		If IsNull(objRS("DateUpdated")) Or objRS("DateUpdated") = "" Then
			dteDateUpdated = ""
		Else
			dteDateUpdated = objRS("DateUpdated")
		End If
		
		strViewAppType = "secondary"
		
		strXMLButton = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#XMLModal"" HREF=""#"" onClick=""loadXML(" & objRS("ApplicationID") & ")""><i class=""fa fa-eye""></i> View XML</button>"
		
		'Determine which buttons and actions to display based on the application status
		Select Case objRS("Status")
		
		Case  "Received"
			'strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='Applications3.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			'strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications3.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			
			strAction = "<a href=""self.location='Applications.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"" class=""btn btn-outline-primary btn-xs""><i class=""fa fa-key""></i> Release</a>"
			strAction = strAction & " <a href=""self.location='Applications.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"" class=""btn btn-outline-danger btn-xs""><i class=""fa fa-cogs""></i> Reject</a>"
			
		Case "Added To NA"

			'strAction = "<button type=""button"" class=""btn btn-secondary"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
			strAction = "<a href=""self.location='NAToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"" class=""btn btn-outline-secondary btn-xs""><i class=""fa fa-key""></i> View NA</a>"
			strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View App</button>"
		Case "Cancelled"

			'strAction = "<button type=""button"" title=""Cancelled by the Applicant"" class=""btn btn-secondary"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-ban""></i> Cancelled</button>"
		
			'strStatus  = "<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployee.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Cancelled By Applicant</button>"
			strAction = "<a href=""#"" class=""btn btn-outline-danger btn-sm""><i class=""fa fa-ban""></i> Cancelled</a>"
			
		Case  "Submitted"
			'strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='Applications3.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			'strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications3.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			
			strStatus  = "<button type=""button"" class=""btn btn-success"" onclick=""self.location='ApplicationsEmployee.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted to GCFO</button>"
			
			strAction = "<a href=""self.location='Applications.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "'"" class=""btn btn-outline-primary btn-xs""><i class=""fa fa-key""></i> Release</a>"
			strAction = strAction & " <a href=""self.location='Applications.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"" class=""btn btn-outline-danger btn-xs""><i class=""fa fa-cogs""></i> Reject</a>"
			
		Case Else
			'strAction = "<button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications3.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			'strAction = "Rejected"
			'strAction = " <a href=""self.location='Applications.asp?Action=Reject&ApplicationID=" & objrs("ApplicationID") & "'"" class=""btn btn-outline-danger btn-sm""><i class=""fa fa-cogs""></i> Reject</a>"
			strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View App</button>"
		End Select

		
		'Select the Action and Status buttons/pills based on the application status
			Select Case objRS("Status")
			
			Case  "Received"
				strAction = "<button type=""button"" class=""btn btn-outline-primary btn-xs"" onclick=""self.location='Applications.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> Release</button>"
				strAction = strAction & " <button type=""button"" class=""btn btn-outline-danger btn-xs"" onclick=""self.location='Applications.asp?Action=Reject&Link=AP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
				strAction = strAction & strXMLButton
			Case "Added To CS"

				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='CSToDiners.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
				strAction = strAction & strXMLButton
			Case "Added To NA"

				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportNA.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-file""></i> View NA</button>"	
				strAction = strAction & "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View App</button>"
				
				strStatus = "<span class=""badge badge-pill badge-info"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
			Case "Submitted"
				strAction = "<button type=""button"" class=""btn btn-outline-danger btn-xs"" onclick=""self.location='Applications.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"

				'strStatus  = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='Applications.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted to GCFO</button>"
				strStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">Submitted to GCFO</span>"
			Case "Deleted"
				'strAction = "Deleted - " & objRS("DateUpdated")'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				strAction = "<Span style=""font-size:12px;"">Deleted - " & dteDateUpdated & "</span>"'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				'strStatus  = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='Applications.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Cancelled By Applicant</button>"
				strStatus = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">Deleted</span>"
			Case "ASFIN Approved"
				strAction = "<button type=""button"" title=""Approved by GCFO"" class=""btn btn-secondary btn-xs"" onclick=""self.location='Applications.asp?Link=AP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-check""></i>GCFO Approved</button>"
			
				'strStatus  = "<button type=""button"" class=""btn btn-secondary btn-sm"" onclick=""self.location='Applications.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Approved by GCFO</button>"
				strStatus = "<span class=""badge badge-pill badge-success"">Approved by ASFIN</span>"
			Case  "Awaiting Review"
				'strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='Applications.asp?Action=Release&ApplicationID=" & objrs("ApplicationID") & "&EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Release</button>"
				'strAction = strAction & " <button type=""button"" class=""btn btn-outline-danger btn-xs"" onclick=""self.location='Applications.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-times""></i> Delete</button>"
				
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=AP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				
				'If the application is a Limit Change then open the Limit Change Submit screen, otherwise open the Normal submit screen
				If strApplicationType = "LimitChange" Then
					strAction = strAction & "<button type=""button"" class=""btn btn-outline-primary btn-xs"" onclick=""self.location='ApplicationsLimitSubmit.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				Else
					strAction = strAction & "<button type=""button"" class=""btn btn-outline-primary btn-xs"" onclick=""self.location='ApplicationsSubmit.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				End If
				
				strStatus = "<span class=""badge badge-pill badge-warning"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
			Case  "Awaiting issue"
				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='../Admin/CAPSAdmin/ExportNA.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-key""></i> View NA</button>"	
				strStatus = "<span class=""badge badge-pill badge-info"">" & objRS("Status") & "</span>"
			Case "On Hold"
				strStatus = "<span class=""badge badge-pill badge-secondary "" data-toggle=""modal"" data-target=""#StatusModal"" data-AppID=""" & objRS("ApplicationID") & """  data-AppName=""" & objRS("FirstName") & " " & objRS("Surname") & " - " & objRS("CardType") & " " & objRS("CardTypeSub") & " Application"" onClick=""OpenSs(this);"">" & objRS("Status") & "</span>"
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=AP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				
				'If the application is a Limit Change then open the Limit Change Submit screen, otherwise open the Normal submit screen
				If strApplicationType = "LimitChange" Then
				
					strAction = strAction & "<button type=""button"" class=""btn btn-outline-primary btn-xs"" onclick=""self.location='ApplicationsLimitSubmit.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				Else
					strAction = strAction & "<button type=""button"" class=""btn btn-outline-primary btn-xs"" onclick=""self.location='ApplicationsSubmit.asp?Action=Release&Link=AP&ApplicationID=" & objrs("ApplicationID") & "&ApplicationEmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-check""></i> Check</button>"
				End If
			Case  "Done"
				strAction = "<button type=""button"" class=""btn btn-outline-" & strViewAppType & " btn-xs"" onclick=""self.location='ApplicationDetail.asp?Link=AP&ApplicationID=" & objrs("ApplicationID") & "'""; title=""" & strViewAppTitle & """><i class=""fa fa-file""></i> View App</button>"
				strAction = strAction & strXMLButton
				strStatus = "<span class=""badge badge-pill badge-info"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
			Case "Rejected"
				'strAction = "Deleted - " & objRS("DateUpdated")'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
				
				strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='ApplicationDetail.asp?ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-file""></i> View App</button>"
				strAction = strAction & strXMLButton & "&nbsp;&nbsp;"
				strAction = strAction & "<Span style=""font-size:12px;"">Rejected - " & dteDateUpdated & "</span>&nbsp;&nbsp;"'<button type=""button"" class=""btn btn-danger"" onclick=""self.location='Applications.asp?Action=Cancel&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Deleted</button>"
								
				strStatus = "<span class=""badge badge-pill badge-danger"" style=""font-size:12px;"">Rejected</span>"
			Case Else
				'strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='Applications.asp?Action=Cancel&Link=AP&ApplicationID=" & objrs("ApplicationID") & "'"";><i class=""fa fa-minus-circle""></i> Cancel</button>"
				'strAction = "Rejected"
				'strStatus  = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='Applications.asp?ApplicationID=" & objrs("ApplicationID") & "'"";>Submitted</button>"
				strStatus = "<span class=""badge badge-pill badge-success"" style=""font-size:12px;"">" & objRS("Status") & "</span>"
			End Select
			
		
		If trim(objRS("CardTypeSub")) = "Diners" Then
			strCardType = "<img src=""img/Diners2.png"" height=""30px"" width=""50px""> "
		ElseIf objRS("CardTypeSub") = "ANZ" Then
			strCardType = "<img src=""img/ANZ2.png"" height=""30px"" width=""50px""> "
		Else
			strCardType = "<img src=""img/Mastercard2.png"" height=""30px"" width=""50px""> "
		End If
			
		Response.Write "<tr><td style=""font-size:12px;"">" & objRS("ApplicationID") & "</td><th scope=""row"">" & objRS("CardType") & " - " & objRS("CardTypeSub") & "</th><td style=""font-size:12px;"">" & strStatus & "</td><td style=""font-size:12px;"">" & objRS("NameOnCard") & "</td> " & _
			"<td style=""font-size:12px;"">" & objRS("ApplicationTypeName") & "</td><td style=""font-size:12px;"">" & FormatDateTime(objRS("DateSubmitted"),vbShortDate) & "</td><td>" & strAction & "</td></tr>"
			
			'Response.Write "<tr><th scope=""row"">" & objRS("CardType") & " - " & objRS("CardTypeSub") & "</th><td>" & objRS("Status") & "</td><td>" & objRS("ApplicationID") & "</td> " & _
			'	"<td>" & FormatDateTime(objRS("DateReceived"),vbShortDate) & "</td><td><a href=""#"" class=""btn btn-outline-primary btn-sm"">" & strAction & "</a><a href=""#"" class=""btn btn-outline-danger btn-sm"">Action</a></td></tr>"
				
		objRS.movenext
	Loop
	
objRS.Close

End Sub


Public Sub DisplayTableDetailsCards()
'Procedure to load all of the Card records for the selected employee

Dim strAction, strStatus
Dim strCardType
Dim strWhere
Dim strDateIssued
Dim strCardNo
Dim strCreditLimit

'If IsNull(Session("CardID")) or Session("CardID") = "" OR Session("CardID") = 0 Then

'Else
'	strWhere = " AND CardID = " & Session("CardID")
'End If

'This is the Field to avoid returning CTS accounts, unless they are specifically required. ---TO BE COMPLETED
strWhere = strWhere & " AND CardTypeSub <> 'CTS'"

If Session("SearchAll") = "" or IsNull(Session("SearchAll")) Then
	strSQL = "SELECT * FROM qryCAPSCards WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeSearchID") & "' " & strWhere & " ORDER BY Status,CardTypeSub,CardID DESC"
	'strSQL = "SELECT * FROM qryCards WHERE ([Surname] Like '%" & Session("SearchAll") & "%' OR [FirstName] Like '%" & Session("SearchAll") & "%' OR [EmployeeID] Like '%" & Session("SearchAll") & "%')"
	'strSQL = "SELECT * FROM qryCards"
Else
	strSQL = "SELECT * FROM qryCAPSCards WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeSearchID") & "' " & strWhere & " ORDER BY Status,CardTypeSub,CardID DESC"
	'strSQL = "SELECT * FROM qryCards WHERE ([Surname] Like '%" & Session("SearchAll") & "%' OR [FirstName] Like '%" & Session("SearchAll") & "%' OR [EmployeeID] Like '%" & Session("SearchAll") & "%')"
End If
'response.write strsql
objRS.Open strSQL,objCon

    'If the selected employee has no cards then display a No Records record
	If objRS.Eof Then
	
		Response.Write "<tr><th scope=""row""><i>No Cards</i></th><td></td><td></td> " & _
				"<td></td><td></td></tr>"
	End If
	
    Do until objRS.EOF 
		
		Select Case objRS("Status")
		
		Case  "Received"
			'strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='Cards3.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			'strAction = strAction & " <button type=""button"" class=""btn btn-danger"" onclick=""self.location='Cards3.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			strAction = "<a href=""self.location='Cards.asp?Action=Release&CardID=" & objrs("CardID") & "'"" class=""btn btn-outline-primary btn-sm""><i class=""fa fa-key""></i> Release</a>"
			
		Case "Added To CS"

			'strAction = "<button type=""button"" class=""btn btn-secondary"" onclick=""self.location='CSToDiners.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
			strAction = "<a href=""self.location='CSToDiners.asp?CardID=" & objrs("CardID") & "'"" class=""btn btn-outline-secondary btn-xs""><i class=""fa fa-key""></i> View CS</a>"
			
		Case "Cancelled"

			'strAction = "<button type=""button"" title=""Cancelled by the Applicant"" class=""btn btn-secondary"" onclick=""self.location='CSToDiners.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-ban""></i> Cancelled</button>"
		
			'strStatus  = "<button type=""button"" class=""btn btn-danger"" onclick=""self.location='ApplicationsEmployee.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
			
			strAction = "<a href=""#"" class=""btn btn-outline-danger btn-sm""><i class=""fa fa-ban""></i> Cancelled</a>"
			
		Case  "Submitted"
			strAction = "<button type=""button"" class=""btn btn-outline-primary btn-xs"" onclick=""self.location='Cards3.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-outline-danger btn-xs"" onclick=""self.location='Cards3.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			
			strStatus  = "<button type=""button"" class=""btn btn-outline-success btn-xs"" onclick=""self.location='ApplicationsEmployee.asp?CardID=" & objrs("CardID") & "'"";>Submitted to GCFO</button>"
		Case Else
			'strAction = "<button type=""button"" class=""btn btn-danger"" onclick=""self.location='Cards3.asp?Action=Reject&CadID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			'strAction = "Rejected"
			strAction = "<a href=""CardDetail.asp?CardID=" & objrs("CardID") & """ class=""btn btn-outline-secondary btn-xs""><i class=""fa fa-key""></i> Card Details</a>"
		End Select
		
		If isnull(objRS("DateIssued")) or objRS("DateIssued") = "" then
			strDateIssued = ""
		Else
			strDateIssued = FormatDateTime(objRS("DateIssued"),vbShortDate)
		End If
		
		If IsNull(objRS("CardNumber")) Then
			strCardNo = ""
		Else
			strCardNo = MaskCard(objRS("CardNumber"))
			
			
			'If mid(strCardNo,5,1)=0 Then 
			'	strCardNo = mid(strCardNo,6,2) & "****" & right(strCardNo,4)
			'Else
			'	strCardNo = mid(strCardNo,4,2) & "****" & right(strCardNo,4)
			'End If
			'If len(strCardNo)>8 Then strCardNo = mid(strCardNo,6,2) & "****" & right(strCardNo,4)
		End If
		
		'Get the Card type for Card Status
		If isNull(objRS("CardTypeSub")) Then
			strCardType = ""
		Else
			strCardType  = objRS("CardType") & " " & objRS("CardTypeSub")
		End If
		
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
		
			If objRS("Status") = "00" or objRS("Status") = "" Then
				strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
			Else
				strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
			End If
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
						If left(objRS("CardTypeSub"),3) = "NAB" Then
							strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
						Else
							strCreditLimit = FormatCurrency(objRS("CreditLimit")/100,0)
						End If
					End If
					
				Else
					strCreditLimit = FormatCurrency(objRS("CreditLimit"),0)
				End If
			Else
				strCreditLimit = objRS("CreditLimit")
			End If
			
		End If
		
		Response.Write "<tr><td style=""font-size:12px;"">" & objRS("CardID") & "</td><th scope=""row"">" & objRS("CardType") & " - " & objRS("CardTypeSub") & "</th><td style=""font-size:12px;"">" & strStatus & "</td><td style=""font-size:12px;"">" & strCardNo & " - " & objRS("NameOnCard") & "</td> " & _
			"<td style=""font-size:12px; text-align:center;"">" & strCreditLimit & "</td><td style=""font-size:12px; text-align:center;"">" & strDateIssued & "</td><td style=""font-size:12px; text-align:center;"">" & objRS("Expiry") & "</td><td style=""font-size:12px;"">" & strAction & "</td></tr>"
			
		objRS.movenext
	Loop
				
objRS.Close

End Sub


Public Sub DisplayTableDetailsCSFrom()
'Procedure to load all of the CS From Diners records for the selected employee
Dim strAction
Dim strCardType
Dim strEmpty
Dim strWhere
Dim strEmpty2

If IsNull(Session("ApplicationID")) or Session("ApplicationID") = "" OR Session("ApplicationID") = 0 Then

Else
	'strWhere = " AND ApplicationID = " & Session("ApplicationID")
End If
If Session("EmployeeID") = "" Then
	strSQL = "SELECT * FROM qryCAPSCSFromDiners WITH(NOLOCK) WHERE EIDNo = '" & Session("EmployeeSearchID") & "' " & strWhere & ""
Else
	strSQL = "SELECT * FROM qryCAPSCSFromDiners WITH(NOLOCK) WHERE EIDNo = '" & Session("EmployeeSearchID") & "' " & strWhere & ""
End If

objRS.Open strSQL,objCon
	'If the selected employee has no CS From then display a No Records record
	If objRS.Eof Then

		Response.Write "<tr><th scope=""row""><i>No CS From Diners records</i></th><td></td><td></td> " & _
				"<td></td><td></td></tr>"
	End If
	
    Do until objRS.EOF 
		
		If objRS("Status") = "Added To CS" then
			
			'strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='CSFromDiners.asp?Action=Remove&CSFromDinersID=" & objrs("CSToDinersID") & "'"";><i class=""fa fa-retro""></i> Remove</button>"
			strAction = "<a href=""self.location='CSFromDiners.asp?CSFromDinersID=" & objrs("CSToDinersID") & "'"" class=""btn btn-outline-secondary btn-sm""><i class=""fa fa-key""></i> View CS</a>"
		Else
			
			strAction = "Received From Diners"
		End If
		
		Response.Write "<tr><th scope=""row"">" & objRS("EIDNo") & "</th><td style=""font-size:12px;"">" & objRS("CardNo") & "</td><td style=""font-size:12px;"">" & objRS("CSFromDinersID") & " - " & objRS("NameOnCard") & "</td> " & _
			"<td style=""font-size:12px;"">" & objRS("FileDateTime") & "</td><td style=""font-size:12px;"">" & strAction & "</td></tr>"
			
		objRS.movenext
	Loop
	
objRS.Close

End Sub


Public Sub DisplayTableDetailsCMTo()
'Procedure to load all of the CM To NAB records for the selected employee
Dim strAction
Dim strCardType
Dim strWhere
Dim strEmpty2

If IsNull(Session("ApplicationID")) or Session("ApplicationID") = "" OR Session("ApplicationID") = 0 Then

Else
	'strWhere = " AND ApplicationID = " & Session("ApplicationID")
End If

If Session("EmployeeID") = "" Then
	strSQL = "SELECT * FROM tblCAPSNABCM WITH(NOLOCK) WHERE EIDNo = '" & Session("EmployeeSearchID") & "' " & strWhere & ""
Else
	strSQL = "SELECT * FROM tblCAPSNABCM WITH(NOLOCK) WHERE EIDNo = '" & Session("EmployeeSearchID") & "' " & strWhere & ""
End If

objRS.Open strSQL,objCon

	'If the selected employee has no CS To then display a No Records record
    If objRS.Eof Then

		Response.Write "<tr><th scope=""row""><i>No CM To NAB records</i></th><td></td><td></td> " & _
				"<td></td><td></td></tr>"
	End If
	
    Do until objRS.EOF 
		
		If objRS("Status") = "Added To CM" then
			
			strAction = "<a href=""self.location='ExportCSNAB.asp?CAPSNABCMID=" & objrs("CAPSNABCMID") & "'"" class=""btn btn-outline-secondary btn-sm""><i class=""fa fa-key""></i> View CM</a>"
		Else
			
			strAction = "Sent To NAB"
		End If
		
		Response.Write "<tr><th scope=""row"">" & CheckString(objRS("EIDNo")) & "</th><td style=""font-size:12px;"">" & MaskCard(CheckString(objRS("CardNumber"))) & "</td><td style=""font-size:12px;"">" & CheckString(objRS("CAPSNABCMID")) & " - " & CheckString(objRS("EmbossingName")) & "</td> " & _
			"<td style=""font-size:12px;"">" & CheckString(objRS("FileDateTime")) & "</td><td style=""font-size:12px;"">" & strAction & "</td></tr>"
		
		'Response.Write "<tr><th scope=""row"">" & objRS("EIDNo") & "</th><td style=""font-size:12px;"">" & MaskCard(objRS("CardNo")) & "</td><td style=""font-size:12px;"">" & objRS("CAPSNABCMID") & " - " & objRS("NameOnCard") & "</td> " & _
		'	"<td style=""font-size:12px;"">" & objRS("FileDateTime") & "</td><td style=""font-size:12px;"">" & strAction & "</td></tr>"
		
		
		objRS.movenext
	Loop
	
objRS.Close

End Sub


Public Sub DisplayTableDetailsCSTo()
'Procedure to load all of the CS To Diners records for the selected employee
Dim strAction
Dim strCardType
Dim strWhere
Dim strEmpty2

If IsNull(Session("ApplicationID")) or Session("ApplicationID") = "" OR Session("ApplicationID") = 0 Then

Else
	'strWhere = " AND ApplicationID = " & Session("ApplicationID")
End If

If Session("EmployeeID") = "" Then
	strSQL = "SELECT * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE EIDNo = '" & Session("EmployeeSearchID") & "' " & strWhere & ""
Else
	strSQL = "SELECT * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE EIDNo = '" & Session("EmployeeSearchID") & "' " & strWhere & ""
End If

objRS.Open strSQL,objCon

	'If the selected employee has no CS To then display a No Records record
    If objRS.Eof Then
	
		Response.Write "<tr><th scope=""row""><i>No CS To Diners records</i></th><td></td><td></td> " & _
				"<td></td><td></td></tr>"
	End If
	
    Do until objRS.EOF 
		
		If objRS("Status") = "Added To CS" then
			
			'strAction = "<button type=""button"" class=""btn btn-primary"" onclick=""self.location='CSToDiners.asp?Action=Remove&CSToDinersID=" & objrs("CSToDinersID") & "'"";><i class=""fa fa-retro""></i> Remove</button>"
			strAction = "<a href=""self.location='CSToDiners.asp?CSToDinersID=" & objrs("CSToDinersID") & "'"" class=""btn btn-outline-secondary btn-sm""><i class=""fa fa-key""></i> View CS</a>"
		Else
			
			strAction = "Sent To Diners"
		End If
		
		Response.Write "<tr><th scope=""row"">" & objRS("EIDNo") & "</th><td style=""font-size:12px;"">" & MaskCard(objRS("CardNo")) & "</td><td style=""font-size:12px;"">" & objRS("CSToDinersID") & " - " & objRS("NameOnCard") & "</td> " & _
			"<td style=""font-size:12px;"">" & objRS("FileDateTime") & "</td><td style=""font-size:12px;"">" & strAction & "</td></tr>"
			
		objRS.movenext
	Loop
	
objRS.Close

End Sub


Public Sub DisplayTableDetailsCDMC()
'Procedure to load all of the CS To Diners records for the selected employee
Dim strAction, strStatus
Dim strActive
Dim dteLeftDefence
Dim dteDateLeft
Dim strDaysAgo

If Session("EmployeeID") = "" Then
	strSQL = "SELECT * FROM qryCAPSCDMCHistory WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeSearchID") & "' AND Deleted = 'N' ORDER BY [DateUpdated] DESC"
	'strSQL = "SELECT * FROM qryCAPSCDMC"
Else
	'strSQL = "SELECT * FROM qryCAPSCDMC"
	strSQL = "SELECT * FROM qryCAPSCDMCHistory WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeSearchID") & "' AND Deleted = 'N' ORDER BY [DateUpdated] DESC"
End If

objRS.Open strSQL,objCon
   	
    Do until objRS.EOF 
	
		'strAction = "<a href=""self.location='CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & "'"" class=""btn btn-outline-secondary btn-sm""><i class=""fa fa-binoculars""></i> View CDMC</a>"		
		strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='CDMCDetail.asp?CDMCID=" & objrs("CDMCID") & "&EmployeeSearchID=" & objRS("EmployeeID") & "'"";><i class=""fa fa-binoculars""></i> View CDMC</button>"
		
		'If y = 0 Then
		'	Response.Write "<td>" & objRS("EmployeeID") & "</td><td style=""font-size:16px;"">" & objRS("FirstName") & " " & objRS("Surname") & "</td><td style=""font-size:16px;"">" & objRS("Addressline1") & " " & objRS("OutSuburb") & "</td><td>" & objRS("DateUpdated") & "</td><td></td><td></td><td></td><td></td><td></td><td></td><td></td></tr>"
			'Response.Write "<td>" & objRS("EmployeeID") & "</td><td style=""font-size:16px;"">" & objRS("FirstName") & " " & objRS("Surname") & "</td><td style=""font-size:16px;"">" & objRS("Address1") & " " & objRS("Suburb") & "</td><td>" & objRS("DateUpdated") & "</td><td>" & objRS("Status") & "</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>"
		'Else
		'	Response.Write "<tr><td>" & objRS("EmployeeID") & "</td><td style=""font-size:16px;"">" & objRS("FirstName") & " " & objRS("Surname") & "</td><td style=""font-size:16px;"">" & objRS("addressline1") & " " & objRS("OutSuburb") & "</td><td>" & objRS("DateUpdated") & "</td><td>" & objRS("Loaded") & "</td><td></td><td></td><td></td><td></td><td></td><td></td></tr>"
		'End If
		
		If IsNull(objRS("ActiveEmployee")) OR objRS("ActiveEmployee") = "" Then
			strActive = "Y"
		Else
			strActive = objRS("ActiveEmployee")
		End If

		IF strActive = "N" then
			If IsNull(objRS("LastUpdated")) or objRS("LastUpdated")="" Then
				dteDateLeft = "No Date"
			Else
				dteDateLeft = FormatDateTime(objRS("LastUpdated"),vbShortDate)
				strDaysAgo = "title=""" & DateDiff("D",objRS("LastUpdated"),Now()) & " Days ago"""
			End If

			dteLeftDefence = "<span class=""badge badge-pill badge-danger"">" & dteDateLeft & "</span>"
		Else
			dteLeftDefence = "<span class=""badge badge-pill badge-success"">Active</span>"
		End If

		Response.Write "<tr><th scope=""row"">" & objRS("EmployeeID") & "</th><td style=""font-size:12px;"">" & objRS("FirstName") & " " & objRS("Surname") & "</td><td style=""font-size:12px;"">" & objRS("OutDinersAddress1") & " - " & objRS("OutDinersAddress2") & "</td><td style=""font-size:12px;"">" & objRS("PostalMessage") & "</td> " & _
			"<td style=""font-size:12px;"">" & objRS("IsValidPostal") & "</td><td style=""font-size:12px;"">" & FormatDateTime(objRS("DateUpdated"),vbShortDate) & "</td><td style=""font-size:12px;"">" & strActive & "</td><td style=""font-size:12px;"" " & strDaysAgo & ">" & dteLeftDefence & "</td><td>" & strAction & "</td></tr>"
			
		objRS.movenext
	Loop
	
				
objRS.Close

End Sub

Public Sub GetCounts()
'Temporary procedure to get the count of each record type for the counter displays at the top of the tabs
Dim arrSQL(5)

arrSQL(0) = "SELECT * FROM qryCAPSApplications WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeSearchID") & "'"
arrSQL(1) = "SELECT * FROM tblCAPSCSToDiners WITH(NOLOCK) WHERE EIDNo = '" & Session("EmployeeSearchID") & "'"
arrSQL(2) = "SELECT * FROM qryCAPSCSFromDiners WITH(NOLOCK) WHERE EIDNo = '" & Session("EmployeeSearchID") & "' " 
arrSQL(3) = "SELECT * FROM qryCAPSCards WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeSearchID") & "' AND CardTypeSub <> 'CTS'"
arrSQL(4) = "SELECT * FROM qryCAPSCDMCHistory WITH(NOLOCK) WHERE EmployeeID = '" & Session("EmployeeSearchID") & "' "
arrSQL(5) = "SELECT * FROM tblCAPSNABCM WITH(NOLOCK) WHERE EIDNo = '" & Session("EmployeeSearchID") & "' "

intApplicationsCount = 0
intCSToCount = 0
intCSFromCount = 0
intCardsCount = 0
intCDMCCount = 0
intCMToCount = 0

For x = 0 to 5

	objRS2.Open arrSQL(x),objCon,3,1
	
		If NOT objRS2.EOF Then
			objRS2.Movelast
			
			SELECT Case x
				Case 0
					intApplicationsCount = objRS2.Recordcount
				Case 1
					intCSToCount = objRS2.Recordcount
				Case 2
					intCSFromCount = objRS2.Recordcount
				Case 3
					intCardsCount = objRS2.Recordcount
				Case 4
					intCDMCCount = objRS2.Recordcount
				Case 5
					intCMToCount = objRS2.Recordcount
					
			END Select
			
		End If
		
	objRS2.Close
	
Next



End Sub


Public Function GetAuditLog(strTable)

Dim intRecord, intRows

If Session("ApplicationID") = 0 Then
	strSQL = "SELECT * FROM tblCCAuditLog WHERE EmployeeID = '" & Session("EmployeeID") & "' AND TableName = '" & strTable & "'"
Else
	strSQL = "SELECT * FROM tblCCAuditLog WHERE EmployeeID = '" & Session("EmployeeID") & "' AND TableName = '" & strTable & "' AND ApplicationID = " & Session("ApplicationID")
End If

	objRS2.Open strSQL,objCon

		Do Until objRS2.Eof
		
			intRecord = intRecord + 1

			If intRecord < 4 then
			
				GetAuditLog = GetAuditLog & "<td>" & objRS2("ChangeDate") & "</td><td>" & objRS2("ValueAfter") & "</td><td></td>"
			End If
			
			objRS2.Movenext
		Loop

	objRS2.Close
	
	For intRows = 1 to 9 - intRecord
	
		GetAuditLog = GetAuditLog & "<td></td>"
		
	Next

	If Len(GetAuditLog) > 30  Then GetAuditLog = left(GetAuditLog,Len(GetAuditLog)-9)

End Function

Set objRS2 = Nothing
Set objRS = Nothing
Set objCon = Nothing
%>
