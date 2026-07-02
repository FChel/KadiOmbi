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
Dim strReviewedBy
Dim dteDateReviewed
Dim lngCreditLimit
Dim arrState(8)
Dim lngCDMCID
Dim strAddressType(3)
Dim strSearchTerm
Dim strWorkPhone
Dim strMobilePhone
Dim strRank
Dim strGroup

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
	Set objRS1 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")
	
    objCon.Open Session("DBConnection")	

    Session("CurrentPage") = "CC/CDMCList.asp"

	If IsNull(Session("EmployeeID")) OR Session("EmployeeID") = "" Then Session("EmployeeID")= 0

	If IsNull(Session("CDMCID")) OR Session("CDMCID") = "" Then Session("CDMCID")= 0

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	Session("EmployeeID") = Request.QueryString("EmployeeID")
End If

If Not IsEmpty(Request.QueryString("CDMCID")) Then
	Session("CDMCID") = Request.QueryString("CDMCID")
End If

If Not IsEmpty(Request.QueryString("ViewButton")) Then
		Session("ViewButton") = Request.QueryString("ViewButton")
	End If
	
'Execute Action
If Request.QueryString("Action") = "Save" Then

   If Session("StatusID") = 1 Then
        'Do not allow Read Only users to make changes
        If Session("UserTypeID") = 4 Then
            strMessage = "NOT SAVED. You are a READ ONLY User and cannot make any changes."
        Else
            'Call SaveCarParking()
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
	If Request.QueryString("SearchTerm") = "" Or IsNull(Request.QueryString("SearchTerm")) Then
		strSearchTerm = Request.Form("UserSearch")
	Else
		strSearchTerm = Trim(Request.QueryString("SearchTerm"))
	End If
	
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


function DeleteData(GEXPID) {
   
        if (window.confirm('Would you like to DELETE the selected record?') == true) {

            self.location = "Loans.asp?Action=Delete&GeneralExpenseID=" + GEXPID;
        }
        
}

function DontCancel(e) {

	var id = e.getAttribute('data-EmployeeID');
	var varSelected = e.checked
	
	var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("SaveResults").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/SaveDontCancel.asp?EmployeeID=" + id + "&Selected=" + varSelected + "", true);
  xhttp.send();
  
  //alert(varSelected)
}

</script>
	
</head>
<body >
 
 <main class="main py-3">
 <div class="container">

 <form action="CDMCList.asp?Action=UserSearch&SearchTerm=" method="POST" id="frm" name="frm">
 <section class="breadcrumbs py-4">
		<div class="row">
			<div class="col-md-8" id="SaveResults">
				<h4 class="text-left">Corporate Directory</h4>
			</div>
			<div class="col-md-4 text-right">
				<%Call LoadExtraButtons()%>
			</div>
		</div>
		<div class="row py-2">
			<div class="col-md-8">
              <%Call LoadViewButtons()%>
            </div>
			
			<div class="col-md-4">
				<div class="form-group has-search">
					<span class="fa fa-search form-control-feedback" onClick="frm.submit();"></span>
					<input type="text" class="form-control" type="search" id="UserSearch" name="UserSearch" placeholder="Search by Keyword" value="<%=strSearchTerm%>"/>
				 </div>
			</div>	
		</div>
      </section>
</form>
  
  <div class="row">
	<div class="col-12">
      <!-- CDMC DataTable-->
          <div class="table-responsive">
            <table class="table table-compact" id="dataTable" width="100%" cellspacing="0">
              <thead>
                <tr>
                  <th style="font-size:13px;">EID</th>
				  <th style="font-size:13px;">Group</th>
                  <th style="font-size:13px;">Title</th>
                  <th style="font-size:13px;">First Name</th>
                  <th style="font-size:13px;">Last Name</th>
				  <th style="font-size:13px;">Address Postal</th>
				  <th style="font-size:13px;">Address CAPS Out</th>
				  <th style="font-size:13px;" title="Date record first appeared on CDMC file">First Updated</th>
				  <th style="font-size:13px;" title="Date record last appeared on CDMC file">Last Updated</th>
				  <th style="font-size:13px;" title="Active Employee status. If Y, employee is active. If N employee is inactive." >Active</th>
				  <th style="font-size:13px;" title="CDMC History record status.  Y indicates the active record." >Deleted</th>
				  <th style="font-size:13px;" title="Days remaining before employee is removed from CDMC">Countdown</th>
				  <th style="font-size:13px;">Don't Cancel</th>
				  <th style="font-size:13px;">View</th>
                </tr>
              </thead>
              <tbody>
               
				<%
        
      DisplayTableDetails()
        
%>	

              </tbody>
            </table>
          </div>
       
		</div></div>

    
</form>
</div>
</div>
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
Dim strGroup
Dim strSearch
Dim strSort
Dim strSortArrow
Dim strDontCancel
Dim strTop
Dim strRowClass
Dim strActiveCDMC
Dim strCDMCHistoryDeleted

strSearch = Replace(Request.QueryString("SearchInput"), "'", "''")
	
	If IsEmpty(Request.QueryString("SortType")) Then
		'strOrderType = "ASC"
	Else
		If Request.QueryString("SortType") = "ASC" Then
			strOrderType = "DESC"
			'Set the variable to be used in the sort order Fontawesone image
			strSortArrow = "-down"
		Else
			strOrderType = "ASC"
			'Set the variable to be used in the sort order Fontawesone image
			strSortArrow = "-up"
		End If
	End If
	
	If IsEmpty(Request.QueryString("Sort")) Then
		strSort = ""
	Else		
		strSort = " ORDER BY " & Request.QueryString("Sort") & " " & strOrderType
	End If
	
	'Build the TOP Statement
	If Session("PageCombo") = "" Or IsNull(Session("PageCombo")) Then
		Session("PageCombo") = 50
	End If
	
	'If there is no sort then sort by the most recent submitted
	If IsNull(strSort) Or strSort = "" Then strSort = " ORDER BY DateIssued DESC"

	'Response.Write Session("ViewButton")
	
	If Session("ViewButton") = "Cancel" Then
		
		'strWhere = " AND Action = 'DontCancel' AND ([RemoveCountdown] Is NOT NULL AND [RemoveCountdown] <> '') "
		strWhere = " AND Action = 'DontCancel' AND Active = 'Y' "

	ElseIf Session("ViewButton") = "Countdown" Then
		strWhere = " AND [RemoveCountdown] < 5 "	
	Else
		'This catches ALL
		strWhere = " "
	End If
	
If Session("Filter") = "UserSearch" Then
	If IsNull(strSearchTerm) or IsEmpty(strSearchTerm) Then
	Else
		'If the user has entered a search term with a space the assume this is a first and last name so search on that only
		If Instr(1,strSearchTerm," ")>0 Then
			arrNames = Split(strSearchTerm," ")
			strFNameSearch = arrNames(0)
			strLNameSearch = arrNames(1)
			
			strWhere = " WHERE ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%')"
			'strWhere = " WHERE Deleted ='N' AND ([FirstName] Like '%" & strFNameSearch & "%' AND [Surname] Like '%" & strLNameSearch & "%')"
		Else
			strWhere = " WHERE ([FirstName] Like '%" & strSearchTerm & "%' OR [Surname] Like '%" & strSearchTerm & "%' OR [EmployeeID] Like '%" & strSearchTerm & "%')"
			'strWhere = " WHERE Deleted ='N' AND ([FirstName] Like '%" & strSearchTerm & "%' OR [Surname] Like '%" & strSearchTerm & "%' OR [EmployeeID] Like '%" & strSearchTerm & "%')"
		End If
	End If
End If

'Make sure there is WHERE in the where statement
If Instr(1,strWhere,"WHERE")=0 Then
	'strWhere = Replace(strWhere,"AND","WHERE")
	If Trim(Left(strWhere,4)) = "AND" Then
		strWhere = " WHERE " & Right(strWhere,len(strWhere)-4)
		
	End If
	
	strTop = " TOP 100 "
	
End If


If Session("EmployeeID") = "" Then
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	strSQL = "SELECT " & strTop  & " * FROM qryCAPSCDMCListAction WITH(NOLOCK) " & strWhere & " Order By RemoveCountDown DESC "
Else
	strSQL = "SELECT " & strTop  & " * FROM qryCAPSCDMCListAction WITH(NOLOCK) " & strWhere & " Order By RemoveCountDown DESC "
	'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
End If
'response.write strSQL

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
		'strAddr1 = Left(objRS("Addressline1") & " " & objRS("Addressline2") & " " & objRS("Addressline3") & " " & objRS("Addressline4") & " " & objRS("Addressline5") & " " & objRS("Addressline6"),20) & "..."
		strAddr2 = Left(objRS("PostalAddress_Unit") & " " & objRS("PostalAddress_ClientLocation") & " " & objRS("PostalAddress_DeliveryLocation") & " " & objRS("Postaladdress_City") & " " & objRS("Postaladdress_State") & " " & objRS("Postaladdress_PostCode"),20) & "..."
		strAddr3 = Left(objRS("OutAddr1") & " " & objRS("OutAddr2") & " " & objRS("OutAddr3") & " " & objRS("OutSuburb") & " " & objRS("OutState") & " " & objRS("OutPostCode"),20) & "..."
		
		'Get the Group name and truncate if long
		If IsNull(objRS("GroupName")) or objRS("GroupName") = "" Then
			strGroup = "" 
		Else
			If Len(objRS("GroupName")) > 10 Then
				strGroup = Left(objRS("GroupName"),9) & "..."
			Else
				strGroup = objRS("GroupName")
			End If
		End If
		'Get the Dont Cancel checkbox value
		If IsNull(objRS("Active")) or objRS("Active") = "" Then
			strDontCancel = "" 
		Else
			If objRS("Active") = "Y" Then
				strDontCancel = "checked"
			Else
				strDontCancel = ""
			End If
		End If
		
		
		'strAction = "EmployeeHistory.asp?Action=Search&EmployeeID=" & objRS("EmployeeID") & ""
		
		strActiveCDMC = CheckCDMCActive(objRS("EmployeeID"))
		
		If objRS("Deleted") = "N" Then
			strRowClass = "text-success"
		Else
			strRowClass = "text-warning"
		End If
		
		strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='EmployeeHistory.asp?Action=Search&EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> More</button>"
		'strAction = "<button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""self.location='EmployeeHistory.asp?Action=Search&EmployeeID=" & objrs("EmployeeID") & "'"";><i class=""fa fa-key""></i> View All</button>"
		
		If strActiveCDMC = "Y" AND objRS("Deleted") = "N" Then
			Response.Write "<TR><TD style=""font-size:13px;""><a Target=""_self"" HREF=""CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & "&EmployeeSearchID=" & objRS("EmployeeID") & """>" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:13px;""><a Target=""_self"" HREF=""CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & "&EmployeeSearchID=" & objRS("EmployeeID") & """>" & strGroup & "</a></TD><TD><a Target=""_self"" HREF=""CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & "&EmployeeSearchID=" & objRS("EmployeeID") & """>" & objRS("Title") & "</a></TD>" & _
					"<TD style=""font-size:13px;"">" & objRS("Firstname") & "</TD><TD style=""font-size:13px;"">" & objRS("Surname") & "</TD>" & _
					"<TD style=""font-size:13px;"">" & strAddr2 & "</TD>" & _
					"<TD style=""font-size:13px;"">" & strAddr3 & "</TD><TD style=""font-size:13px;"">" & dteFirstUpdated & "</TD>"  & _
					"<TD style=""font-size:13px;"">" & dteLastUpdated & "</TD><TD style=""font-size:13px; text-align:center;"">" & objRS("ActiveEmployee") & "</TD><TD style=""font-size:13px; text-align:center;"">" & objRS("Deleted") & "</TD><TD style=""font-size:13px; text-align:center;"">" & objRS("RemoveCountdown") & "</TD>" & _
					"<TD style=""text-align:center;""><input type=""checkbox"" class=""form-check-input"" onClick=""DontCancel(this);"" data-EmployeeID=""" & objRS("EmployeeID") & """ name=""chkPreserve" & objRS("CDMCID") & """ " & strDontCancel & " /></TD><TD style=""font-size:13px;"">" & strAction & "</TD></TR>" 
					'"<TD style=""text-align:center;"">" & strStatus & "</TD><TD style=""text-align:center;"">" & objRS(14) & "</TD><TD style=""text-align:center;"">" & objRS(15) & "</TD></TR>"
		End If

		If strActiveCDMC = "N" Then
		
			Response.Write "<TR class=""" & strRowClass & """><TD style=""font-size:13px;""><a Target=""_self"" HREF=""CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & "&EmployeeSearchID=" & objRS("EmployeeID") & """>" & objRS("EmployeeID") & "</a></TD>" & _
					"<TD style=""font-size:13px;""><a Target=""_self"" HREF=""CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & "&EmployeeSearchID=" & objRS("EmployeeID") & """>" & strGroup & "</a></TD><TD><a Target=""_self"" HREF=""CDMCDetail.asp?CDMCID=" & objRS("CDMCID") & "&EmployeeSearchID=" & objRS("EmployeeID") & """>" & objRS("Title") & "</a></TD>" & _
					"<TD style=""font-size:13px;"">" & objRS("Firstname") & "</TD><TD style=""font-size:13px;"">" & objRS("Surname") & "</TD>" & _
					"<TD style=""font-size:13px;"">" & strAddr2 & "</TD>" & _
					"<TD style=""font-size:13px;"">" & strAddr3 & "</TD><TD style=""font-size:13px;"">" & dteFirstUpdated & "</TD>"  & _
					"<TD style=""font-size:13px;"">" & dteLastUpdated & "</TD><TD style=""font-size:13px; text-align:center;"">" & objRS("ActiveEmployee") & " <i class=""fa fa-user-slash"" title=""Left Defence on: " & dteLastUpdated & """></i></TD><TD style=""font-size:13px; text-align:center;"">" & objRS("Deleted") & "</TD><TD style=""font-size:13px; text-align:center;"">" & objRS("RemoveCountdown") & "</TD>" & _
					"<TD style=""text-align:center;""><input type=""checkbox"" class=""form-check-input"" onClick=""DontCancel(this);"" data-EmployeeID=""" & objRS("EmployeeID") & """ name=""chkPreserve" & objRS("CDMCID") & """ " & strDontCancel & " /></TD><TD style=""font-size:13px;"">" & strAction & "</TD></TR>" 
					'"<TD style=""text-align:center;"">" & strStatus & "</TD><TD style=""text-align:center;"">" & objRS(14) & "</TD><TD style=""text-align:center;"">" & objRS(15) & "</TD></TR>"
		
		End If
			
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

Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(3)

If Session("ViewButton") = "Cancel" Then
	arrButton(2) = "active"
ElseIf Session("ViewButton") = "Countdown" Then
	arrButton(3) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='CDMCList.asp?Link=RP&ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='CDMCList.asp?Link=RP&ViewButton=Cancel';""><i class=""fa fa-running""></i> Don't Cancel</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='CDMCList.asp?Link=RP&ViewButton=Countdown';""><i class=""fa fa-stopwatch""></i> Countdown</button>" & _
				"</div>"

End Sub


Public Sub LoadExtraButtons
'Load the Excel Export and Card Type buttons depending on what has been clicked
Dim strButtonType
Dim strButtonNext


	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"" style=""float:left;"">" & _
		"<button type=""button"" class=""btn btn-outline-success"" onClick=""OpenExcelReport();""><i class=""fa fa-file-excel""></i> Export To Excel</button>" & _
		"</div>"
	
	Response.Write "<button class=""btn btn-primary"" onClick=""self.location.href='CDMCList.asp?Link=ED&Action=Save';""><i class=""fa fa-check""></i> Save Changes</button>"
				'"<i class=""fa fa-plane""></i> DTC Displayed</button>"
			  
		'"<button type=""button"" class=""btn btn-outline-success"" onClick=""self.location.href='../Admin/CAPSAdmin/DisplayDataset.asp?tbl=qryCAPSTrainingReport&W=" & strWhere & "&Top=100';""><i class=""fa fa-file-excel""></i> Export To Excel</button>" & _
		
End Sub

Public Function CheckCDMCActive(EmpID)

	objRS1.Open "SELECT * FROM tblCAPSCDMC WHERE EmployeeID = '" & EmpID & "'",objCon
	
		If objRS1.EOF Then
			CheckCDMCActive = "N"
		Else
			CheckCDMCActive = "Y"
		End If
		
	objRS1.Close

End Function



Set objRS = Nothing
Set objCon = Nothing
%>
