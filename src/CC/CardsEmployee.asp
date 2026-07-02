
<!-- #Include file=CAPSHeader.asp -->
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
Dim lngUserID
Dim strUserLogon
Dim strFName
Dim strLName
Dim strEmailAddress
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
Dim strCardNo
Dim strStatusBG
Dim strView
Dim strButt1
Dim strButt2
Dim strButt3
Dim strButt4
Dim strButt5
Dim strButt6
Dim arrCardtype(1,3)
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
	
	arrCardtype(1,1) = "DTC - Diners"
	arrCardtype(1,2) = "DTC - Companion Mastercard"
	arrCardtype(1,3) = "DPC - ANZ"
	
	arrCardtype(0,1) = "Diners DTC"
	arrCardtype(0,2) = "Diners Mastercard"
	arrCardtype(0,3) = "DPC - ANZ"
	
    objCon.Open Session("DBConnection")	

    Session("CurrentPage") = "CC/Cards3_2.asp"

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


If Not IsEmpty(Request.QueryString("SearchTerm")) Then
	strSearchTerm = Request.Form("UserSearch")
	Session("Filter") = "UserSearch"
	'Request.QueryString("Action")
	'Response.write "SearchTerm=" & Request.form("UserSearch") &  " " & Request.QueryString("Action")
End If
 
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "New" Then
			Session("CardID") = 0
		End If
	End If
	
	If Session("View") = "All" Then
		strView = " All"
	Else
		strView = Session("UserType") & " - " & Session("UserName")
	End If
	
  Call LoadDetails()
  
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
	
<body >
<main class="main py-1">
<div class="container">

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
	</div>
	

	<form action="ApplicationsSubmit.asp?Action=SubmitApp" method="POST" id="frm" name="frm" class="inline">
	
<TABLE class="table table-bordered">
	<tbody >

	<tr>
		<th align="left">Employee ID</th>
		<td><input style="text-align:left;width:90%" id="EmployeeID" placeholder="Employee ID" name="EmployeeID" maxlength="15" TABINDEX="1" value="<%=strEmployeeID%>"></td>
		<th align="left">Card No.</th>
		<td><input style="text-align:left;width:90%" id="CardNo" placeholder="Card Number" name="CardNo" TABINDEX="2" value="<%=strCardNo%>"></td>
		<th align="left">Card Type</th>
		<td>
		<select placeholder="Card Type" name="CardTypeSelect" id="CardTypeSelect" onchange="valCard();"><option value="">Card Type</option>
			<%
				For x = 1 to 3
			
					If cstr(arrCardtype(0,x)) = cstr(strCardType) Then
						strSelected = " SELECTED "
					Else
						strSelected = ""
					End If
					
					Response.Write "<option " & strSelected & " value="" & arrCardtype(0,x) & "">" & arrCardtype(1,x) & "</Option>"
				Next
			%>			
			</select>
	</tr>
	
	<tr>
		<th align="left">Title</th>
		<td><input style="text-align:left;width:90%" id="Title" placeholder="Title" name="Title" TABINDEX="3" value="<%=strTitle%>"></td>
		<th align="left">First Name</th>
		<td><input style="text-align:left;width:90%" id="FirstName" placeholder="First Name" name="FirstName" TABINDEX="4" value="<%=strFirstName%>"></td>
		<th align="left">Last Name</th>
		<td><input style="text-align:left;width:90%" id="LastName" placeholder="Last Name" name="LastName" TABINDEX="5" value="<%=strLastName%>"></td> 
	</tr>
	
	<tr>
		<th align="left">Address 1</th>
		<td colspan="2"><input style="text-align:left;width:90%;" id="Address1" placeholder="Address 1" name="Address1" TABINDEX="6" value="<%=strAddress1%>"></td>

		<th align="left">Address 2</th>
		<td colspan="2"><input style="text-align:left;width:90%;" id="Address2" placeholder="Address 2" name="Address2" TABINDEX="7" value="<%=strAddress2%>"></td> 
	</tr>

	<tr>
		<th align="left">Address 3</th>
		<td colspan="2"><input style="text-align:left;width:90%;" id="Address3" placeholder="Address 3" name="Address3" TABINDEX="8" value="<%=strAddress3%>"></td>
		
		<th align="left">Suburb</th>
		<td colspan="2"><input style="text-align:left;width:90%;" id="Suburb" placeholder="Suburb" name="Suburb" TABINDEX="9" value="<%=strSuburb%>"></td> 
	</tr>
	
	<tr>
		<th align="left">State</th>
		<td><input style="text-align:left;width:90%;" id="State" placeholder="State" name="State" TABINDEX="10" value="<%=strState%>"></td>
		<td style="border:0px;"></td>
		<td style="border:0px;"></td>
		<th align="left">Post Code</th>
		<td><input style="text-align:left;width:90%;" id="PostCode" placeholder="Post Code" name="PostCode" TABINDEX="11" value="<%=strPostCode%>"></td> 
	</tr>

	<tr>
		<th align="left">Status</th>
		<td><input style="text-align:left; width:90%;" class="<%=strStatusBG%>" id="Status" placeholder="Status" name="Status" TABINDEX="12" value="<%=strStatus%>"></td>
		<th align="left">Credit Limit</th>
		<td><input style="text-align:right; width:90%;" id="CreditLimit" placeholder="Credit Limit" name="CreditLimit" TABINDEX="13" value="<%=lngCreditLimit%>"></td>
		<th align="left">Date Received</th>
		<td><input style="text-align:left;width:90%" id="DateReceived" placeholder="Date Received" name="DateReceived" TABINDEX="14" value="<%=dteDateReceived%>"></td> 
	</tr>
</table>

</form>

<div class="container-fluid">
<form action="Cards3_2.asp?Action=UserSearch&SearchTerm=" method="POST" id="frm" name="frm" class="form-inline">

<div class="card-header" >&nbsp;&nbsp;&nbsp;
    <i class="fa fa-address-card"></i> Existing Cards for <%=strView%> &nbsp;<input type="text" name="UserSearch" id="UserSearch" style="width:200px;" placeholder="..Enter name or Employee ID" value="<%=strSearchTerm%>"><button type="submit" style="height:26px;"><i class="fa fa-search fa-1.5x"></i></button><% If Not IsNull(strSearchTerm) Then Response.Write "<span style=""color:red; font-weight:bold;""> &nbsp;" & strSearchTerm & "</span>" %>
</div>



</div>

<div class="container-fluid">
<!-- Button trigger modal -->
<button type="button" class="btn btn-primary" data-toggle="modal" data-target="#exampleModalCenter"><i class="fa fa-bullhorn"></i> Declaration</button>

<button type="button" class="btn btn-primary" onclick="CloseScreen()";><i class="fa fa-times"></i> Close</button>
<button type="button" class="btn btn-primary" onclick="SaveData()"; ><i class="fa fa-check"></i> Save</button>
<button type="button" class="btn btn-primary" onclick="self.location='CardsEmployee.asp?Action=New'"; ><i class="fa fa-pencil"></i> New</button>
<%=strMessageIcon %>
  
<hr />
</div>


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
		  
		  
		  Response.write  " &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<div class=""btn-group""><button type=""button"" class=""btn btn-" & strButt1 & """ onclick=""self.location='Cards3_2.asp?Filter=All'"";><i class=""fa fa-filter""></i> View All </button>" & _
							" <button type=""button"" class=""btn btn-" & strButt2 & """ onclick=""self.location='Cards3_2.asp?Filter=Active'"";><i class=""fa fa-filter""></i> View Active</button></div>" & _
							" &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<div class=""btn-group""><button type=""button"" class=""btn btn-" & strButt3 & """ onclick=""self.location='Cards3_2.asp?Filter2=All'"";><i class=""fa fa-filter""></i> View All </button>" & _
							" <button type=""button"" class=""btn btn-" & strButt4 & """ onclick=""self.location='Cards3_2.asp?Filter2=DTC'"";><i class=""fa fa-filter""></i> View DTC Diners</button>" & _
							" <button type=""button"" class=""btn btn-" & strButt5 & """ onclick=""self.location='Cards3_2.asp?Filter2=DTCMC'"";><i class=""fa fa-filter""></i> View DTC MC</button>" & _
							" <button type=""button"" class=""btn btn-" & strButt6 & """ onclick=""self.location='Cards3_2.asp?Filter2=DPC'"";><i class=""fa fa-filter""></i> View DPC</button></div>" & _
							"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <button type=""button"" class=""btn btn-success float-right"" onclick=""window.open('Cards3Excel.asp?SearchTerm=" & strSearchTerm & "')"";><i class=""fa fa-file-excel-o""></i> Export</button>"
		  
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
Dim strAction, strStatus
Dim strWhere, strWhere1
Dim strExpiry
Dim strCreditLimit
Dim strCardNo
Dim strDateLoaded
Dim strExpiryFormat
Dim strLimitFormat
Dim arrNames
Dim strFNameSearch
Dim strLNameSearch

If Session("Filter") = "UserSearch" Then
	If IsNull(strSearchTerm) or IsEmpty(strSearchTerm) Then
	Else
		'If the user has entered a search term with a space the assume this is a first and last name so search on that only
		If Instr(1,strSearchTerm," ")>0 Then
			arrNames = Split(strSearchTerm," ")
			strFNameSearch = arrNames(0)
			strLNameSearch = arrNames(1)
			
			strWhere = " AND ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%')"
		Else
			strWhere = " AND ([FirstName] Like '%" & strSearchTerm & "%' OR [Surname] Like '%" & strSearchTerm & "%' OR [EmployeeID] Like '%" & strSearchTerm & "%')"
		End If
	End If
End If

If Session("Filter") = "Active" Then
	strWhere = strWhere & " AND [Status] = '00'"
	strWhere1 = strWhere1 & " AND [Status] = '00'"
Else
	'strWhere = ""
	'strWhere1 = ""
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
			strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='Cards3_2.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='Cards3_2.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
		Case "Added To CS"

			strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CSToDiners.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View CS</button>"
		
		Case "Cancelled"

			strAction = "<button type=""button"" title=""Cancelled by the Applicant"" class=""btn btn-secondary btn-xs"" onclick=""self.location='CSToDiners.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-ban""></i> Cancelled</button>"
		
			'strStatus  = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='ApplicationsEmployee.asp?CardID=" & objrs("CardID") & "'"";>Cancelled By Applicant</button>"
		
		Case  "Submitted"
			strAction = "<button type=""button"" class=""btn btn-primary btn-xs"" onclick=""self.location='Cards3_2.asp?Action=Release&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> Release</button>"
			strAction = strAction & " <button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='Cards3_2.asp?Action=Reject&CardID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"
			
			'strStatus  = "<button type=""button"" class=""btn btn-success btn-xs"" onclick=""self.location='ApplicationsEmployee.asp?CardID=" & objrs("CardID") & "'"";>Submitted to GCFO</button>"
		Case Else
			'strAction = "<button type=""button"" class=""btn btn-danger btn-xs"" onclick=""self.location='Cards3_2.asp?Action=Reject&CadID=" & objrs("CardID") & "'"";><i class=""fa fa-cogs""></i> Reject</button>"

			strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='Cards3_2.asp?CardID=" & objrs("CardID") & "'"";><i class=""fa fa-key""></i> View Card</button>"
		End Select

		If isNull(objRS(15)) Then
			strExpiry = ""
		Else
			strExpiry = Month(objRS(15)) & "/" & Year(objRS(15))
			If Len(strExpiry) = 6 Then strExpiry = "0" & strExpiry
			
			'Format the Expiry Date value depending on whether it is within 3 months of expiry or expired already.
			strExpiryFormat = " color:black;"
			If DateDiff("m",now(),"01/" & strExpiry) < 3 Then
				strExpiryFormat = " color:orange; font-weight:bold;"
			End If
			If DateDiff("m", now(),"01/" & strExpiry) < 0 Then
				strExpiryFormat = " color:red; font-weight:bold;"
			End If
			
		End If
		
		If isNull(objRS("CreditLimit")) Then
			strCreditLimit = 0
		Else
			'Format the Credit Limit for the Diners Cards
			strCreditLimit = cdbl(objRS("CreditLimit"))/100
			strCreditLimit = FormatCurrency(strCreditLimit,0)
			
			'Set the Credit Limit format depending on the value
			If strCreditLimit > 30000 Then
				strLimitFormat = " color:green; font-weight:bold;"
			ElseIf strCreditLimit < 30000 Then
				strLimitFormat = " color:red; font-weight:bold;"
			Else
				strLimitFormat = " color:black;"
			End If
			'strCreditLimit = FormatCurrency(objRS(12),0)
		End If
		
		If IsNull(objRS("CardNumber")) Then
			strCardNo = ""
		Else
			strCardNo = objRS("CardNumber")
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
		
		Response.Write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""Cards3_2.asp?CardID=" & objRS(0) & """>" & objRS(0) & "</a></TD><TD>" & strAction & "</a></TD>" & _
				"<TD><a Target=""_self"" HREF=""Cards3_2.asp?CardID=" & objRS(0) & """>" & objRS(1) & "</a></TD><TD><a Target=""_self"" HREF=""Cards3_2.asp?CardID=" & objRS(0) & """>" & objRS(2) & "</a></TD>" & _
				"<TD style=""text-align:center;"">" & objRS(3) & "</TD><TD style=""text-align:center;"" Title=""" & objRS("CardNumber") & """>" & strCardNo & "</TD>" & _
				"<TD style=""text-align:center;"">" & objRS(4) & "</TD><TD style=""text-align:center;"">" & objRS(5) & "</TD><TD style=""text-align:center;"">" & objRS(6) & "</TD>" & _
				"<TD style=""text-align:center;"">" & objRS(9) & "</TD><TD style=""text-align:center; " & strLimitFormat & """>" & strCreditLimit & "</TD><TD style=""text-align:center;"">" & strStatus & "</TD>" & _
				"<TD style=""text-align:center;"">" & strDateLoaded & "</TD><TD style=""text-align:center; " & strExpiryFormat & """>" & strExpiry & "</TD></TR>"
				
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
		objRS.Open "SELECT * FROM tblCard WHERE CardID = " & Session("CardID") & "",objCon

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
				strCardNo = objRS("CardNumber")
				
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
				strCardNo = ""
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
