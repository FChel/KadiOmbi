<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=ADOVBS.inc -->
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
Dim arrState(8)
Dim lngCDMCID
Dim strAddressType(3)
Dim strSearchTerm

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

    Session("CurrentPage") = "CC/CDMC.asp"

	If IsNull(Session("EmployeeID")) OR Session("EmployeeID") = "" Then Session("EmployeeID")= 0

	If IsNull(Session("CDMCID")) OR Session("CDMCID") = "" Then Session("CDMCID")= 0
	
If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	Session("ApplicationID") = Request.QueryString("ApplicationID")
End If

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	Session("EmployeeID") = Request.QueryString("EmployeeID")
End If

If Not IsEmpty(Request.QueryString("CDMCID")) Then
	Session("CDMCID") = Request.QueryString("CDMCID")
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

If Not IsEmpty(Request.QueryString("AddressType")) Then
	strAddressType(0) = Request.QueryString("AddressType")
End If

If Not IsEmpty(Request.QueryString("SearchTerm")) Then
	strSearchTerm = Request.Form("UserSearch")
	Session("Filter") = "UserSearch"
End If

  Call LoadDetails()
  
%>


<script LANGUAGE="javascript">


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


function FormatNumber(num, decimalNum, bolLeadingZero, bolParens, bolCommas)
    /**********************************************************************
        IN:
            NUM - the number to format
            decimalNum - the number of decimal places to format the number to
            bolLeadingZero - true / false - display a leading zero for
                                            numbers between -1 and 1
            bolParens - true / false - use parenthesis around negative numbers
            bolCommas - put commas as number separators.
     
        RETVAL:
            The formatted number!
     **********************************************************************/ {
    if (isNaN(parseInt(num))) return "NaN";

    var tmpNum = num;
    var iSign = num < 0 ? -1 : 1;		// Get sign of number

    // Adjust number so only the specified number of numbers after
    // the decimal point are shown.
    tmpNum *= Math.pow(10, decimalNum);
    tmpNum = Math.round(Math.abs(tmpNum))
    tmpNum /= Math.pow(10, decimalNum);
    tmpNum *= iSign;					// Readjust for sign


    // Create a string object to do our formatting on
    var tmpNumStr = new String(tmpNum);

    // See if we need to strip out the leading zero or not.
    if (!bolLeadingZero && num < 1 && num > -1 && num != 0)
        if (num > 0)
            tmpNumStr = tmpNumStr.substring(1, tmpNumStr.length);
        else
            tmpNumStr = "-" + tmpNumStr.substring(2, tmpNumStr.length);

    // See if we need to put in the commas
    if (bolCommas && (num >= 1000 || num <= -1000)) {
        var iStart = tmpNumStr.indexOf(".");
        if (iStart < 0)
            iStart = tmpNumStr.length;

        iStart -= 3;
        while (iStart >= 1) {
            tmpNumStr = tmpNumStr.substring(0, iStart) + "," + tmpNumStr.substring(iStart, tmpNumStr.length)
            iStart -= 3;
        }
    }

    // See if we need to use parenthesis
    if (bolParens && num < 0)
        tmpNumStr = "(" + tmpNumStr.substring(1, tmpNumStr.length) + ")";

    return tmpNumStr;		// Return our formatted string!
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
 
 <main class="main py-3">
 <div class="container">
<form action="CDMC.asp?Action=UserSearch&SearchTerm=" method="POST" id="frm" name="frm" class="form-inline">
<div class="row" style="width:100%;">
	<div class="card col-12">
		<div class="card-header" >
			<i class="fa fa-address-book"></i> Defence Corporate Directory Details <%=strFirstName & " " & strLastName%> <input type="text" name="UserSearch" id="UserSearch" style="width:200px;" placeholder="..Enter name or Employee ID" value="<%=strSearchTerm%>"><button type="submit" style="height:26px;"><i class="fa fa-search fa-1.5x"></i></button>
			<% If Not IsNull(strSearchTerm) Then Response.Write "<span style=""color:red; font-weight:bold;""> &nbsp;" & strSearchTerm & "</span>" %>
		</div>
	</div>
</div>
</form>
  
<form class="form-inline" action="CDMC.asp"> 

<div class="row" style="width:100%;">
 <div class="input-group col-sm-12">
	<div class="form-row col-12"><div class="form-group col-md-6">
		<label for="EmployeeID">EmployeeID</label><input type="text" style="width:100%;" class="form-control" id="EmployeeID" placeholder="EmployeeID" value="<%=strTitle%>" title="<%=strTitle%>"/></div>
	<div class="form-group col-sm-6">
		<label for="EmployeeID">Card Type</label>
		<select class="form-control" placeholder="Card Type" style="width:100%;" name="CardTypeSelect" id="CardTypeSelect" onchange="valCard();">
			<option value="DTC - Diners">DTC - Diners</Option><option value="DPC - ANZ">DPC - ANZ</Option>
			<option value="DTC - MasterCard">DTC - Companion Mastercard</Option>
			</select></div>
	</div><div class="form-row col-12"><div class="form-group col-md-2">
		<label for="Title">Title</label><input type="text" style="width:100%;" class="form-control" id="Title" placeholder="Title" value="<%=strTitle%>" title="<%=strTitle%>"/></div>
	<div class="form-group col-md-5">
		<label for="FirstName">FirstName</label><input type="text" style="width:100%;" class="form-control" id="FirstName" placeholder="FirstName" value="<%=strFirstName%>" title="<%=strFirstName%>"/></div>
	<div class="form-group col-md-5">
		<label for="LastName">LastName</label><input type="text" style="width:100%;" class="form-control" id="LastName" placeholder="LastName" value="<%=strLastName%>" title="<%=strLastName%>"/></div>
	</div><div class="form-row col-12">
	<div class="form-group col-md-6">
		<label for="Address1">Address1</label><input type="text" style="width:100%;" class="form-control" id="Address1" placeholder="Address1" value="<%=strAddress1%>" title="<%=strAddress1%>"/></div>
	<div class="form-group col-md-6">
		<label for="Address2">Address2</label><input type="text" style="width:100%;" class="form-control" id="Address2" placeholder="Address2" value="<%=strAddress2%>" title="<%=strAddress2%>"/></div>
	</div><div class="form-row col-12">
	<div class="form-group col-md-6">
		<label for="Address3">Address3</label><input type="text" style="width:100%;" class="form-control" id="Address3" placeholder="Address3" value="<%=strAddress3%>" title="<%=strAddress3%>"/></div>
	<div class="form-group col-md-6">
		<label for="Suburb">Suburb</label><input type="text" style="width:100%;" class="form-control" id="Suburb" placeholder="Suburb" value="<%=strSuburb%>" title="<%=strSuburb%>"/></div>
	</div><div class="form-row col-12">
	<div class="form-group col-md-6">
		<label for="State">State</label><input type="text" style="width:100%;" class="form-control" id="State" placeholder="State" value="<%=strState%>" title="<%=strState%>"/></div>
	<div class="form-group col-md-6">
		<label for="PostCode">PostCode</label><input type="text" style="width:100%;" class="form-control" id="PostCode" placeholder="PostCode" value="<%=strPostCode%>" title="<%=strPostCode%>"/></div>
	</div><div class="form-row col-12">
	<div class="form-group col-md-6">
		<label for="WorkPhone">WorkPhone</label><input type="text" style="width:100%;" class="form-control" id="WorkPhone" placeholder="WorkPhone" value="<%=strWorkPhone%>" title="<%=strWorkPhone%>"/></div>
	<div class="form-group col-md-6">
		<label for="MobilePhone">MobilePhone</label><input type="text" style="width:100%;" class="form-control" id="MobilePhone" placeholder="MobilePhone" value="<%=strMobilePhone%>" title="<%=strMobilePhone%>"/></div>
	</div>
	<div class="alert alert-danger">
		<strong>Note!</strong> Details cannot be changed here. All Address and Contact details MUST be updated in the <a href="http://directory/dcd/" target="_new">Corporate Directory</a>. Updates may take up to 3 days to display above.
	</div>
</div>
  </form>
</div>
</form>
<hr>
  <div class="container-fluid">
<!-- Button trigger modal -->

<button type="button" class="btn btn-primary" onclick="CloseScreen()";><i class="fa fa-times"></i> Close</button>
<button type="button" class="btn btn-primary" onclick="SaveData()"; ><i class="fa fa-check"></i> Save</button>
<button type="button" class="btn btn-primary" onclick="self.location='CDMC.asp?Action=New'"; ><i class="fa fa-pencil"></i> New/Clear</button>
<%=strMessageIcon %>
  
<hr />

  
     			  
      <!-- Example DataTables Card-->
      <div class="card mb-3">
        <div class="card-header">
          <i class="fa fa-table"></i> CDMC Data as Loaded in the CDMC Admin Screen</div>
    
          <div class="table-responsive">
            <table class="table table-bordered table-hover table-compact" id="dataTable" width="100%" cellspacing="0">
              <thead class="CAPS3">
                <tr>
                  <th Style="background-color:#2394F2; color:white;">EID</th>
				  <th Style="background-color:#2394F2; color:white;">Group</th>
                  <th Style="background-color:#2394F2; color:white;">Title</th>
                  <th Style="background-color:#2394F2; color:white;">First Name</th>
                  <th Style="background-color:#2394F2; color:white;">Last Name</th>
				  <th Style="background-color:#2394F2; color:white;">Address lines</th>
				  <th Style="background-color:#2394F2; color:white;">Address Postal</th>
				  <th Style="background-color:#2394F2; color:white;">Address CAPS Out</th>
				  <th Style="background-color:#2394F2; color:white;">First Updated</th>
				  <th Style="background-color:#2394F2; color:white;">Last Updated</th>
				  <th Style="background-color:#2394F2; color:white;">Active</th>
				  
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
    



</DIV>
</form>
</DIV>
</main>
<!-- #Include file=CAPSFooter.asp -->
</body>
</html>
<%

Public Sub DisplayTableDetails()
Dim y
Dim strAction, strStatus
Dim strAddr1, strAddr2, strAddr3
Dim arrNames
Dim strFNameSearch
Dim strLNameSearch
Dim strWhere
Dim dteFirstUpdated
Dim dteLastUpdated

If Session("Filter") = "UserSearch" Then
	If IsNull(strSearchTerm) or IsEmpty(strSearchTerm) Then
	Else
		'If the user has entered a search term with a space the assume this is a first and last name so search on that only
		If Instr(1,strSearchTerm," ")>0 Then
			arrNames = Split(strSearchTerm," ")
			strFNameSearch = arrNames(0)
			strLNameSearch = arrNames(1)
			
			strWhere = " WHERE ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%')"
		Else
			strWhere = " WHERE ([FirstName] Like '%" & strSearchTerm & "%' OR [Surname] Like '%" & strSearchTerm & "%' OR [EmployeeID] Like '%" & strSearchTerm & "%')"
		End If
	End If
End If

If Session("EmployeeID") = "" Then
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	strSQL = "SELECT Top 100 * FROM qryCAPSCDMC WITH(NOLOCK) " & strWhere
Else
	strSQL = "SELECT Top 100 * FROM qryCAPSCDMC WITH(NOLOCK) " & strWhere
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
End If

objRS.Open strSQL,objCon
    y = 0
    	
    Do until objRS.EOF 
		If isNull(objRS("FirstUpdated")) Then
			dteFirstUpdated = ""
		Else
			dteFirstUpdated = FormatDateTime(objRS("FirstUpdated"),vbShortDate)
		End If
		
		If isNull(objRS("LastUpdated")) Then
			dteLastUpdated = ""
		Else
			dteLastUpdated = FormatDateTime(objRS("LastUpdated"),vbShortDate)
		End If

		'Get the Address details based on the address type selected
		strAddr1 = Left(objRS("Addressline1") & " " & objRS("Addressline2") & " " & objRS("Addressline3") & " " & objRS("Addressline4") & " " & objRS("Addressline5") & " " & objRS("Addressline6"),30) & "..."
		strAddr2 = Left(objRS("PostalAddress_Unit") & " " & objRS("PostalAddress_ClientLocation") & " " & objRS("PostalAddress_DeliveryLocation") & " " & objRS("Postaladdress_City") & " " & objRS("Postaladdress_State") & " " & objRS("Postaladdress_PostCode"),30) & "..."
		strAddr3 = Left(objRS("OutAddr1") & " " & objRS("OutAddr2") & " " & objRS("OutAddr3") & " " & objRS("OutSuburb") & " " & objRS("OutState") & " " & objRS("OutPostCode"),30) & "..."
					
		response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""CDMC.asp?CDMCID=" & objRS("CDMCID") & """>" & objRS("EmployeeID") & "</a></TD>" & _
				"<TD><a Target=""_self"" HREF=""CDMC.asp?CDMCID=" & objRS("CDMCID") & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""CDMC.asp?CDMCID=" & objRS("CDMCID") & """>" & objRS("Title") & "</a></TD>" & _
				"<TD style=""text-align:center;"">" & objRS("Firstname") & "</TD><TD style=""text-align:center;"">" & objRS("Surname") & "</TD>" & _
				"<TD style=""text-align:center;"">" & strAddr1 & "</TD><TD style=""text-align:center;"">" & strAddr2 & "</TD>" & _
				"<TD style=""text-align:center;"">" & strAddr3 & "</TD><TD style=""text-align:center;"">" & dteFirstUpdated & "</TD>"  & _
				"<TD style=""text-align:center;"">" & dteLastUpdated & "</TD><TD style=""text-align:center;"">" & objRS("Active") & "</TD></TR>" 
				'"<TD style=""text-align:center;"">" & strStatus & "</TD><TD style=""text-align:center;"">" & objRS(14) & "</TD><TD style=""text-align:center;"">" & objRS(15) & "</TD></TR>"
				
				'intDaysTotal = intDaysTotal + objRS("Days")
				'intCarsTotal = intCarsTotal + objRS("Cars")
				'intEmpContTotal = intEmpContTotal + objRS("EmployeeContribution")
			
			y = y + 1
			
		objRS.movenext
	Loop
	
	
	response.write "<TR><TH colspan=""8"">Total</TH>" & _
				"<TH colspan=""4"" style=""text-align:center;"">" & y & "</TH></TR>"
				
objRS.Close

End Sub


Sub LoadDetails()

Dim Addr1, Addr2, Addr3, Addr4, Addr5, Addr6

       'Description:	Loads Position details into page if applicable.
		objRS.Open "SELECT * FROM tblCAPSCDMC WHERE CDMCID = " & Session("CDMCID") & "",objCon

			If Not objRS.EOF Then
               
				'Get the Address details based on the address type selected
				If strAddressType(0) = "AddressLine" Then
				
					strAddress1 = objRS("Addressline1")
					strAddress2 = objRS("Addressline2")
					strAddress3 = objRS("Addressline3")
					strSuburb = objRS("Addressline4")
					strState = objRS("Addressline5")
					strPostCode = objRS("Addressline6")
				
				ElseIf strAddressType(0) = "Postal" Then
				
					strAddress1 = objRS("PostalAddress_Unit")
					strAddress2 = objRS("PostalAddress_ClientLocation")
					strAddress3 = objRS("PostalAddress_DeliveryLocation")
					strSuburb = objRS("Postaladdress_City")
					strState = objRS("Postaladdress_State")
					strPostCode = objRS("Postaladdress_PostCode")
				Else
					strAddress1 = objRS("OutAddr1")
					strAddress2 = objRS("OutAddr2")
					strAddress3 = objRS("OutAddr3")
					strSuburb = objRS("OutSuburb")
					strState = objRS("OutState")
					strPostCode = objRS("OutPostCode")
				End If
			
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
				
				strWorkPhone = objRS("TelephoneNumber")
				strMobilePhone = objRS("MobileNumber")
				
				strRank = objRS("ActualRankLvl")
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
		   End If
	   
	objRS.Close
	
End Sub


Sub SaveRecord(lngCDMCID,strGroupName,strDivisionName,strBranchName,strDepartmentName,strDepartmentNumber,strCostCentre,strEmployeeID,strEmployeeType, _
			strFirstname,strSurname,strTitle,strEmail_Address,strTelephoneNumber,strMobileNumber,strDateofBirth,strGender,strActualRankLvl,strSite,strUnit,strReportsTo,strDCD_PostalAddress, _
			straddressline1,straddressline2,straddressline3,straddressline4,straddressline5,straddressline6,strPostalAddress_Unit,strPostalAddress_ClientLocation, _
			strPostalAddress_DeliveryLocation,strPostalAddress_City,strPostalAddress_State,strPostalAddress_PostCode,strPostalAddress_Country,strDCDProtectedIdentity, _
			strIsValidPostal,strOutAddr1,strOutAddr2,strOutAddr3,strOutSuburb,strOutState,strOutPostCode,strPostalMessage,strhasChanged,strDCD_WorkAddress,strClientLocation, _
			strStreetAddress,strCity,strState,strPostCode,strFormalFirstName,strFormalLastName,strFormalMiddleName,strOutTitle,strOutDinersWorkPhone,strOutDinersMobilePhone, _
			strOutANZPhone,strOutDinersAddress1,strOutDinersAddress2,strRemoveCountdown,strFirstUpdated,strLastUpdated,strActive,strUpdatedBy,strDateUpdated,strFileID,strLoaded, x)

Dim intRecord

  	With objCmd
  	
  	    'If the procedure has akready run then don't create the parameter objects again (more than once)
  	    If x = 1 then
			.CommandType = 4
			.CommandText = "spCAPSCDMCSave"
			
			.Parameters.Append objCmd.CreateParameter("CDMCID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("GroupName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("DivisionName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("BranchName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("DepartmentName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("DepartmentNumber", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("CostCentre", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarWChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("EmployeeType", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("Firstname", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("Surname", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("Title", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("Email_Address", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("TelephoneNumber", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("MobileNumber", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("DateofBirth", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("Gender", adVarWChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("ActualRankLvl", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("Site", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("Unit", adVarWChar, adParamInput,200)
			.Parameters.Append objCmd.CreateParameter("ReportsTo", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("DCD_PostalAddress", adVarWChar, adParamInput,500)
			.Parameters.Append objCmd.CreateParameter("addressline1", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline2", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline3", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline4", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline5", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline6", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_Unit", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_ClientLocation", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_DeliveryLocation", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_City", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_State", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_PostCode", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_Country", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("ClientLocation", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("StreetAddress", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("City", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("State", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("PostCode", adVarWChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("DCDProtectedIdentity", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("IsValidPostal", adVarWChar, adParamInput,3)
			.Parameters.Append objCmd.CreateParameter("OutAddr1", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutAddr2", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutAddr3", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutSuburb", adVarWChar, adParamInput,22)
			.Parameters.Append objCmd.CreateParameter("OutState", adVarWChar, adParamInput,3)
			.Parameters.Append objCmd.CreateParameter("OutPostCode", adVarWChar, adParamInput,4)
			.Parameters.Append objCmd.CreateParameter("PostalMessage", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("hasChanged", adVarWChar, adParamInput,3)
			.Parameters.Append objCmd.CreateParameter("FormalFirstName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("FormalLastName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("FormalMiddleName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutTitle", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutDinersWorkPhone", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutDinersMobilePhone", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutANZPhone", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutDinersAddress1", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutDinersAddress2", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("RemoveCountdown", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("FirstUpdated", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("LastUpdated", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput,1)

			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("FileID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Loaded", adChar, adParamInput,1)    
			.Parameters.Append objCmd.CreateParameter("CDMCIDOutput", adInteger, adParamOutput)				
            
        End If
                 
			.Parameters("CDMCID") = lngCDMCID
			.Parameters("GroupName") = strGroupName
			.Parameters("DivisionName") = strDivisionName
			.Parameters("BranchName") = strBranchName
			.Parameters("DepartmentName") = strDepartmentName
			.Parameters("DepartmentNumber") = strDepartmentNumber
			.Parameters("CostCentre") = strCostCentre
			.Parameters("EmployeeID") = strEmployeeID
			.Parameters("EmployeeType") = strEmployeeType
			.Parameters("Firstname") = strFirstname
			.Parameters("Surname") = strSurname
			.Parameters("Title") = strTitle
			.Parameters("Email_Address") = strEmail_Address
			.Parameters("TelephoneNumber") = strTelephoneNumber
			.Parameters("MobileNumber") = strMobileNumber
			.Parameters("DateofBirth") = strDateofBirth
			.Parameters("Gender") = strGender
			.Parameters("ActualRankLvl") = strActualRankLvl
			.Parameters("Site") = strSite
			.Parameters("Unit") = strUnit
			.Parameters("ReportsTo") = strReportsTo
			.Parameters("DCD_PostalAddress") = strDCD_PostalAddress
			.Parameters("addressline1") = straddressline1
			.Parameters("addressline2") = straddressline2
			.Parameters("addressline3") = straddressline3
			.Parameters("addressline4") = straddressline4
			.Parameters("addressline5") = straddressline5
			.Parameters("addressline6") = straddressline6
			.Parameters("PostalAddress_Unit") = strPostalAddress_Unit
			.Parameters("PostalAddress_ClientLocation") = strPostalAddress_ClientLocation
			.Parameters("PostalAddress_DeliveryLocation") = strPostalAddress_DeliveryLocation
			.Parameters("PostalAddress_City") = strPostalAddress_City
			.Parameters("PostalAddress_State") = strPostalAddress_State
			.Parameters("PostalAddress_PostCode") = strPostalAddress_PostCode
			.Parameters("PostalAddress_Country") = strPostalAddress_Country
			.Parameters("ClientLocation") = strClientLocation
			.Parameters("StreetAddress") = strStreetAddress
			.Parameters("City") = strCity
			.Parameters("State") = strState
			.Parameters("PostCode") = strPostCode
			.Parameters("DCDProtectedIdentity") = strDCDProtectedIdentity
			.Parameters("IsValidPostal") = strIsValidPostal
			.Parameters("OutAddr1") = strOutAddr1
			.Parameters("OutAddr2") = strOutAddr2
			.Parameters("OutAddr3") = strOutAddr3
			.Parameters("OutSuburb") = strOutSuburb
			.Parameters("OutState") = strOutState
			.Parameters("OutPostCode") = strOutPostCode
			.Parameters("PostalMessage") = strPostalMessage
			.Parameters("hasChanged") = strhasChanged
			.Parameters("FormalFirstName") = strFormalFirstName
			.Parameters("FormalLastName") = strFormalLastName
			.Parameters("FormalMiddleName") = strFormalMiddleName
			.Parameters("OutTitle") = strOutTitle
			.Parameters("OutDinersWorkPhone") = strOutDinersWorkPhone
			.Parameters("OutDinersMobilePhone") = strOutDinersMobilePhone
			.Parameters("OutANZPhone") = strOutANZPhone
			.Parameters("OutDinersAddress1") = strOutDinersAddress1
			.Parameters("OutDinersAddress2") = strOutDinersAddress2
			.Parameters("RemoveCountdown") = strRemoveCountdown
			.Parameters("FirstUpdated") = NULL 'strFirstUpdated
			.Parameters("LastUpdated") = NULL 'strLastUpdated
			.Parameters("Active") = strActive
			.Parameters("UpdatedBy") = Session("UserID")
			'.Parameters("DateUpdated") = strDateUpdated

			.Parameters("FileID") = strFileID
			.Parameters("Loaded") = "N"'strLoaded
					   
			.ActiveConnection = objCon
			
		End With
			
		objCmd.Execute        
		
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCIDOutput")    
     		                  			     				     		     		
       'response.write  "exec spGeneralExpensesSave =0," & Session("BudgetID") & "," & Session("VersionID") & "," & CostCentreID & ",'GEXP'" & GLCode & "," & BM1 & "," & BM2 & "," & BM3 & "," & BM4 & "," & BM5 & "," & _
       '                     BM6 & "," & BM7 & "," & BM8 & "," & BM9 & "," & BM10 & "," & BM11 & "," & BM12 & "," & OY1 & "," & OY2 & "," & OY3 & ",'" & Comments & "','" & UpdatedBy & "'," & Session("ColumnLock")
End Sub

Set objRS = Nothing
Set objCon = Nothing
%>
