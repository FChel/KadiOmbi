<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
	'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")

'Description:	Staffing Classifications Default Admin screen for updating global salary values
'Author:		MG
'Date:			April 2016

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon
Dim strMessageColour

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngCostCentreID
Dim strTransactionType
Dim lngStaffingClassificationID
Dim strStaffingClassification
Dim strStaffingClassificationDesc
Dim dblSalary
Dim strDeleted
Dim strSalaryClass
Dim intSortOrder
Dim dblPerformancePay
Dim lngEmployeeID
Dim intSuperFundID
Dim strPositionNo
Dim intPositionID
Dim arrMonthName(12)
Dim arrBM(15)
Dim arrHeadings(5)
Dim intFinYearPart1
Dim intFinYearPart2

Dim arrYesNo(2)

	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"

If IsNull(Session("FirstMonth")) or Session("FirstMonth") = "" Then Session("FirstMonth") = "JAN"

For x = 0 to 5
	
	intFinYearPart1 = cint(Session("FinancialYear")) + (x - 1)
	intFinYearPart1 = Right(intFinYearPart1,2)
	intFinYearPart2 = cint(Session("FinancialYear")) + x
	intFinYearPart2 = Right(intFinYearPart2,2)

	'If the Financial Year starts in January then do not display the split/multiple years.
    If Session("FirstMonth") = "JAN" Then
        arrHeadings(x) = "20" & cstr(intFinYearPart1)
    Else
        arrHeadings(x) = cstr(intFinYearPart1) & "/" & cstr(intFinYearPart2)
    End If

Next

'3. Capture Querystring variables

    If Not IsEmpty(Request.QueryString("StaffingClassificationID")) Then
	    lngStaffingClassificationID = Request.QueryString("StaffingClassificationID")
    Else
	    lngStaffingClassificationID = 0
    End If
    
    If Not IsEmpty(Request.QueryString("CostCentreID")) Then		
        Session("CostCentreID") = Request.QueryString("CostCentreID")
        lngCostCentreID = clng(Request.QueryString("CostCentreID"))	
    Else
        lngCostCentreID = Session("CostCentreID")				
    End If	 
		
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		If Request.Form("SaveType") = "" Then
			SaveDetails()
		Else

			If Request.Form("SaveType") = 1 Then
				IncrementsSave()
			Else
				SaveDetails()
			End If
		End If
	End If
	
	'Execute Recalculate	
	If Request.QueryString("Action") = "recalculate" Then
		RecalcStaff()
	End If
	
	'Execute Recalculate	
	If Request.QueryString("Action") = "IncrementsSave" Then
		IncrementsSave()
	End If

	'Execute Delete 	
	If Request.QueryString("Action") = "Delete" Then
		DeleteRecord(Request.QueryString("StaffingClassification"))
	End If

	'Call the procedure to create the Month Names
    	Call GetMonthNames()

	'Load page details
	LoadDetails()
		
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0"/>
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
<script type="text/javascript" src="../formChek.js"></script>
<script type="text/javascript" language="javascript">
<!--
   function SaveData()
    {
        var varSubmit = true
        var varAlert =""  
        
        
        if((isNonnegativeInteger(frm.Salary.value)==false))
	    {            
		    varAlert += "Please enter Salary. Salary must be a numeric value.  \n \n";
		    document.getElementById('Salary').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('Salary').style.backgroundColor="ffffff";	
	    
       	if(isWhitespace(frm.TransactionType.value))
        {            
		    varAlert += "Please enter Transaction Type. \n \n";
		    document.getElementById('TransactionType').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('TransactionType').style.backgroundColor="ffffff";    
          	    
		    
	    if(isWhitespace(frm.SalaryClassification.value))
        {            
		    varAlert += "Please enter SalaryClassification. \n \n";
		    document.getElementById('SalaryClassification').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('SalaryClassification').style.backgroundColor="ffffff";    
	    
	    if(isWhitespace(frm.SortOrder.value))
        {            
		    frm.SortOrder.value=0;
	    }	
	    
	    if(isWhitespace(frm.PerformancePay.value))
        {            
		    frm.PerformancePay.value=0;
	    }
	       
	    if(frm.Deleted.value == 0 )
	    {
		    varAlert += "Please select a Deleted Status. \n \n";
		    document.getElementById('Deleted').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }	
	    else document.getElementById('Deleted').style.backgroundColor="ffffff";
	    	   		  
	   	
	  if(varSubmit == true)
	  {
	        frm.submit();
	  }
	  else
	  {
	    window.alert ("" + varAlert);	    
	  }
  }
  
function CCIDSearch()
{	
	self.location="StaffingClassificationsDefault.asp?CostCentreID=" + frm.CostCentreID.value
}

function StaffRecalc(varBud, varVer)
 {
 	if(window.confirm('Would you like to Recalculate ALL existing Staff values? \n \n For the currently selected Budget ' + varBud + ' and Version ' + varVer + '.')==true){
 	
 		document.getElementById('Progress').style.display = "inline";
        self.location='StaffingClassificationsDefault.asp?Action=recalculate';
	}
}

function SaveData2(varBud, varVer)
 {
 	if(window.confirm('Would you like to Update All Monthly Salary Incremeents for all staff? \n \n For the currently selected Budget ' + varBud + ' and Version ' + varVer + '.')==true){
 	
 		document.getElementById('Progress').style.display = "inline";
        //self.location='StaffingClassificationsDefault.asp?Action=IncrementsSave';
	frm.elements['SaveType'].value = '1';
	frm.submit();
	}
}
function DeleteData(){
	if( confirm("Delete the selected record?") )
	{
		self.location="StaffingClassificationsDefault.asp?Action=Delete&StaffingClassification="+frm.StaffingClassification.value;
		frm.elements['msgbox'].value = 'Deleting...';}
	}
//-->
</script>
</head>
<body>
<form action="StaffingClassificationsDefault.asp?Action=Save" method="POST" id="frm" name="frm">
<table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
<tr>
	    <th style="width:30%;height20px;">Salary Classification Administration</th>
	    <th style="width:70%"></th>
	</tr>
	<tr>
	    <td colspan="2">&nbsp;</td>
	</tr>	
		<tr>
        <th align="Left">&nbsp;Cost Centre</th>
		<td>
		    <select Style="Width:40%" tabindex="1" id="CostCentreID" name="CostCentreID" onchange="CCIDSearch()"><OPTION Value="0">Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblCostCentres WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon
    		
		    Do until objRS.EOF
			    If objRS("CostCentreID") = clng(lngCostCentreID) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CostCentreID") & """" & strSelected & ">" & objRS("CostCentreID") & " - " & objRS("CostCentreName") &"</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
	</tr>
	</table>
<br>

<table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
	<tr>
		<th align="Left" Width="20%">&nbsp;Salary Classification</th>
	    <td Width="30%">&nbsp;<input style="text-align:left;width:90%" id="SalaryClassification" name="SalaryClassification" maxlength="20" tabindex="1" value="<%=strStaffingClassification%>"></td>
	  
        <th align="left" Width="20%">&nbsp;Transaction Type</th>
		<td><select Style="Width:100%" tabindex="4" id="TransactionType" name="TransactionType"><option Value="0">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM tblTransactionType",objCon
		Do until objRS.EOF
			If objRS("TransactionType") = strTransactionType Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("TransactionType") & """" & strSelected & ">" & objRS("TransactionType") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>
	</tr>
	<tr>
		<th align="left">&nbsp;Salary</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="Salary" name="Salary" maxlength="50" tabindex="2" value="<%=dblSalary%>" /></td>
		<th align="Left" Width="20%">&nbsp;Sort Order</th>
	    <td Width="30%">&nbsp;<input style="text-align:left;width:90%" id="SortOrder" name="SortOrder" tabindex="4" value="<%=intSortOrder%>"></td>
	</tr>
	
	<tr>
	<th align="Left" Width="20%">&nbsp;Salary Classification Desciption</th>
	    <td Width="30%">&nbsp;<input style="text-align:left;width:90%" id="StaffingClassificationDesc" name="StaffingClassificationDesc" maxlength="20" tabindex="4" value="<%=strStaffingClassificationDesc%>"></td>
		<th align="Left" Width="20%">&nbsp;Performance Pay</th>
	    <td Width="30%">&nbsp;<input style="text-align:left;width:90%" id="PerformancePay" name="PerformancePay" tabindex="5" value="<%=dblPerformancePay%>"></td>
	</tr>
	<tr>
        <th align="left">&nbsp;Super Fund</th>		
		<td><select Style="Width:100%" tabindex="6" id="SuperFundID" name="SuperFundID"><option Value="0">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM tblSuperFund WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Active = 'Y'",objCon
		
		Do until objRS.EOF
			If objRS("SuperFundID") = intSuperFundID Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("SuperFundID") & """" & strSelected & ">" & objRS("SuperFundID") & " : " & objRS("SuperFundName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>	
	<th align="left">&nbsp;Deleted</th>		
		<td><select Style="Width:40%" tabindex="3" id="Deleted" name="Deleted"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 2
			If strDeleted = arrYesNo(x)Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		Next
		%>
		</select> </td>
	</tr>
	<tr>
	<th align="Left" Width="20%">&nbsp;Position No</th>
	    <td Width="30%">&nbsp;<input style="text-align:left;width:90%" id="PositionNo" name="PositionNo" maxlength="20" tabindex="7" value="<%=strPositionNo%>"></td>
		<th align="left">&nbsp;Position</th>		
		<td><select Style="Width:100%" tabindex="8" id="PositionID" name="PositionID"><option Value="0">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM tblPosition WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Active = 'Y'",objCon
		
		Do until objRS.EOF
			If objRS("PositionID") = intPositionID Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("PositionID") & """" & strSelected & ">" & objRS("PositionID") & " : " & objRS("PositionName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>	
	</tr>
	<tr>
	    <td Width="30%">&nbsp;<input Type=HIDDEN id="EmployeeID" name="EmployeeID" value="<%=lngEmployeeID%>"></td>
		<td Width="30%">&nbsp;<input Type=HIDDEN id="StaffClassificationID" name="StaffClassificationID" value="<%=lngStaffingClassificationID%>"></td>
		<td colspan="3" align="left">&nbsp;<input id="SaveType" name="SaveType" value=""></td>
	</tr>
	
</table>

<table width="60%" align="Left" border="1" cellspacing="1" cellpadding="1">
<tr>
	<th style="height:20px;" colspan="15">&nbsp;Monthly Salary Increments</th></tr>
	<tr><td colspan="15"><span style="color:gray; font-weight:bold;">Increments ALL Salary Values by the percentage in each month - not cummulative. e.g. <span style="color:gray; font-weight:bold; font-size:14;">1.2</span> represents 120% increase in Salary Classification value in that month ONLY for ALL Staff/Salaries</span></td>
</tr>
<tr>
<%
	Response.Write "<th style=""height:20px;"" colspan=""12"">Budget Year</th><th colspan=""3"" style=""background-color:#349761;"">Out Years</th></tr><tr>" & _
		"<th>" & arrMonthName(1) & "</th><th>" & arrMonthName(2) & "</th>" & _
                "<th>" & arrMonthName(3) & "</th><th>" & arrMonthName(4) & "</th>" & _
                "<th>" & arrMonthName(5) & "</th><th>" & arrMonthName(6) & "</th>" & _
                "<th>" & arrMonthName(7) & "</th><th>" & arrMonthName(8) & "</th>" & _
                "<th>" & arrMonthName(9) & "</th><th>" & arrMonthName(10) & "</th>" & _
                "<th>" & arrMonthName(11) & "</th><th>" & arrMonthName(12) & "</th>" & _
		"<th style=""background-color:#349761;"">" & arrHeadings(3) & "</th><th style=""background-color:#349761;"">" & arrHeadings(4) & "</th>" & _
		"<th style=""background-color:#349761;"">" & arrHeadings(5) & "</th>"

	'Write the Input fields with any values for the related variable
	Response.write "</tr><tr>"

	For x = 1 to 15
		If x < 13 Then
			Response.write "<td style=""width:40px; text-align:right;"">&nbsp;<input style=""width:35px; text-align:right;"" id=""BM" & x & """ name=""BM" & x & """ tabindex=""" & x + 6 & """ value=""" & arrBM(x) & """></td>"
		Else
			'Write the Out Year text fields
			Response.write "<td style=""width:40px; text-align:right;"">&nbsp;<input style=""width:35px; text-align:right;"" id=""OY" & x & """ name=""OY" & x & """ tabindex=""" & x + 6 & """ value=""" & arrBM(x) & """></td>"
		End If
	Next

%>
	
</tr>
</table>
<br>
<br>
<br>
<br>
<br>
<br>&nbsp;
<TABLE Width="1500px" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<TR>

    <td Width="100px"><button type="button" onclick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
    <td Width="100px"><button type="button" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td> 
    <td Width="150px"><button type="button" onclick="self.location='StaffingClassificationsDefault.asp?StaffingClassification=0'" )""><img src="../images/add.png" alt="" /> Clear/Add New </button></td>
   <!-- <td Width="100px"><button type="button" onclick="DeleteData()";><img src="../images/delete.png" alt="" /> Delete </button></td>-->
   <td Width="100px" title="Click to apply Salary Increments to ALL Employees"><button type="button" onclick="SaveData2('<%=Session("BudgetName")%>','<%=Session("VersionName")%>')";><img src="../images/table_save.png" alt="" /> Save All </button></td> 
    <td Width="150px"><button type="button" onClick="StaffRecalc('<%=Session("BudgetName")%>','<%=Session("VersionName")%>');" )"" Title="Click to recalculate ALL staff data with salaries below (if changed)"><img src="../images/calculator.png" alt="" /> Recalculate </button></td>
    <TD Width="200px"><span id="Progress" style="display:none"><img src="../images/progress.gif">  &nbsp;&nbsp;&nbsp; <b>Recalculating Staff...</b></span></TD>
        <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
    <TD class='locked' align="left" Width="800x" style="BORDER-RIGHT:0px"><INPUT style="Align:Left; font-weight:Bold; width:100%; text-align:left; color:<%=strMessageColour%>;" type="text" id="msgbox" name="msgbox" value="<%=strMessage%>"></TD>
</TR>
</TABLE>


<hr />
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr><td colspan="9">&nbsp;</td></tr>
	<tr>
		
		<th Width="10%">Salary Classification</th>
		<th Width="10%">Transaction Type</th>
	    <th Width="10%">Salary</th>
	    <th Width="10%">Sort Order</th>
	    <th Width="10%">Salary Classification Description</th>
	    <th Width="5%">Performance Pay</th>
		<th Width="5%">Deleted</th>
		<th Width="5%">Updated By</th>
		<th Width="15%">Date Updated</th>
	</tr>
<%
    objRS.Open "SELECT * FROM  tblStaffingClassifications WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & lngCostCentreID & " Order By SortOrder",objCon
    'Response.Write "SELECT * FROM  tblSalaryClassifications WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & " AND TransactionType = '" & strTransactionType & "' Order By SalaryClassification"
	Do until objRS.eof			
   	   Response.Write "<TR><TD><A Target=""_self"" HREF=""StaffingClassificationsDefault.asp?StaffingClassificationID=" & objRS("StaffingClassificationID") & """>&nbsp;" & objRS("StaffClassification") & "</TD>" & _
   	                    "<TD style=""text-align:center"">&nbsp;" & objRS("TransactionType") & "</TD><TD style=""text-align:right"">&nbsp;" & formatnumber(objRS("Salary"),0,0) & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("SortOrder") & "</TD>" & _ 
   	                    "<TD style=""text-align:center"">&nbsp;" & objRS("StaffClassificationDesc") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("PerformancePay") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("Deleted") & "</TD>" & _ 
   	                    "<TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("DateUpdated") & "</TD></TR>"
       objRS.movenext
	Loop
		
	objRS.Close
	
%>

</table>
</body>

</html>

<% 

Sub LoadDetails()

       'Description:	Loads details into page if applicable.
		objRS.Open "SELECT * FROM tblStaffingClassifications WHERE StaffingClassificationID = " & lngStaffingClassificationID & "",objCon
		'Response.Write "SELECT * FROM tblSalaryClassifications WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND SalaryClassification = '" & strSalaryClassification & "'"
		'Response.write "SELECT * FROM tblStaffingClassifications WHERE StaffingClassificationID = " & lngStaffingClassificationID & ""
			If Not objRS.EOF Then
			    strStaffingClassification = objRS("StaffClassification")
			    strTransactionType = objRS("TransactionType")
			    dblSalary = Round(objRS("Salary"))
			    strStaffingClassificationDesc = objRS("StaffClassificationDesc")
			    intSortOrder = objRS("SortOrder")
                dblPerformancePay = objRS("PerformancePay")
			    strDeleted = objRS("Deleted")
			    lngEmployeeID = objRS("EmployeeID")
			    intSuperFundID = objRS("SuperFundID")
			arrBM(1) = objRS("BM1Index")
			arrBM(2) = objRS("BM2Index")
			arrBM(3) = objRS("BM3Index")
			arrBM(4) = objRS("BM4Index")
			arrBM(5) = objRS("BM5Index")
			arrBM(6) = objRS("BM6Index")
			arrBM(7) = objRS("BM7Index")
			arrBM(8) = objRS("BM8Index")
			arrBM(9) = objRS("BM9Index")
			arrBM(10) = objRS("BM10Index")
			arrBM(11) = objRS("BM11Index")
			arrBM(12) = objRS("BM12Index")
			arrBM(13) = objRS("OY1Index")
			arrBM(14) = objRS("OY2Index")
			arrBM(15) = objRS("OY3Index")
			'strPositionNo = objRS("PositionNo")
			'intPositionID = objRS("PositionID")

    		Else
			   ' strSalaryClassification = ""
			    dblSalary = 0
		        strDeleted = "N"
	            lngEmployeeID = 0
	            intSuperFundID = 0
			strStaffingClassificationDesc = ""
			    intSortOrder = 0
                	dblPerformancePay = 0
			arrBM(1) = ""
			arrBM(2) = ""
			arrBM(3) = ""
			arrBM(4) = ""
			arrBM(5) = ""
			arrBM(6) = ""
			arrBM(7) = ""
			arrBM(8) = ""
			arrBM(9) = ""
			arrBM(10) = ""
			arrBM(11) = ""
			arrBM(12) = ""
			arrBM(13) = ""
			arrBM(14) = ""
			arrBM(15) = ""
			strPositionNo = ""
			intPositionID = 0
           End If

		objRS.Close
	
End Sub

Sub SaveDetails()

       response.Write Request.Form("StaffClassificationID")

	       			Response.Write "UPDATE tblStaffingClassifications SET BM1Index = " & Request.Form("BM1") & ",BM2Index = " & Request.Form("BM2") & ",BM3Index = " & Request.Form("BM3") & ",BM4Index = " & Request.Form("BM4") & ",BM5Index = " & Request.Form("BM5") & ",BM6Index = " & Request.Form("BM6") & ",BM7Index = " & Request.Form("BM7") & ",BM8Index = " & Request.Form("BM8") & ",BM9Index = " & Request.Form("BM9") & ",BM10Index = " & Request.Form("BM10") & ",BM11Index = " & Request.Form("BM11") & ",BM12Index = " & Request.Form("BM12") & ",OY1Index = " & Request.Form("OY13") & ",OY2Index = " & Request.Form("OY14") & ",OY3Index = " & Request.Form("OY15") & " WHERE StaffingClassificationID = " & Request.Form("StaffClassificationID") & ""

			objCon.Execute "UPDATE tblStaffingClassifications SET BM1Index = " & Request.Form("BM1") & ",BM2Index = " & Request.Form("BM2") & ",BM3Index = " & Request.Form("BM3") & ",BM4Index = " & Request.Form("BM4") & ",BM5Index = " & Request.Form("BM5") & ",BM6Index = " & Request.Form("BM6") & ",BM7Index = " & Request.Form("BM7") & ",BM8Index = " & Request.Form("BM8") & ",BM9Index = " & Request.Form("BM9") & ",BM10Index = " & Request.Form("BM10") & ",BM11Index = " & Request.Form("BM11") & ",BM12Index = " & Request.Form("BM12") & ",OY1Index = " & Request.Form("OY13") & ",OY2Index = " & Request.Form("OY14") & ",OY3Index = " & Request.Form("OY15") & " WHERE StaffingClassificationID = " & Request.Form("StaffClassificationID") & ""


            'response.Write "Save=" & clng(Session("BudgetID")) & ", " & clng(Session("VersionID")) & ", " & Request.Form("TransactionType") & "," & Request.Form("StaffingClassificationDesc") & "," & Request.Form("Salary") & "," & Request.Form("SalaryClassification") & "," & Request.Form("EmployeeID") & "," & Request.Form("SortOrder")  & "," & Request.Form("SuperFundID") & "," & Request.Form("PerformancePay") & "," & Session("UserID")
	        'Return the result of the Save Function.
     	    strMessage = "RECORD SAVED."
			strMessage = UCASE(strMessage)
			strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
            strMessageColour = "Black"
     	    'strTransactionType = Request.Form("TransactionType")
		
End Sub	


Sub RecalcStaff()

    		objCon.Execute "spBuildStaffDataBudgetDataByBA " & Session("BudgetID") & "," & Session("VersionID") & "," & Session("BusinessAreaID") & ",'ESTA'"
			objCon.Execute "spBuildStaffDataBudgetDataByBA " & Session("BudgetID") & "," & Session("VersionID") & "," & Session("BusinessAreaID") & ",'NSTA'"
               
			'Return the result of the Save Function.
     	    strMessage = "Staff Values have been recalculated!"
			strMessage = UCASE(strMessage)
			strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
            strMessageColour = "Black"
	
End Sub	


Sub IncrementsSave()

	
		 With objCmd
                .CommandType = 4
                .CommandText = "spStaffingClassificationsSaveIndex"
                
		        .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("CostCentreID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM1Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM2Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM3Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM4Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM5Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM6Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM7Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM8Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM9Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM10Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM11Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM12Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY1Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY2Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY3Index", adDouble, adParamInput)
          	
		        .Parameters("BudgetID") = clng(Session("BudgetID"))
   		        .Parameters("VersionID") = clng(Session("VersionID"))
   		        .Parameters("CostCentreID") = Request.Form("CostCentreID")
		        .Parameters("BM1Index") = Request.Form("BM1")
		        .Parameters("BM2Index") = Request.Form("BM2")
		        .Parameters("BM3Index") = Request.Form("BM3")
		        .Parameters("BM4Index") = Request.Form("BM4")
		        .Parameters("BM5Index") = Request.Form("BM5")
		        .Parameters("BM6Index") = Request.Form("BM6")
		        .Parameters("BM7Index") = Request.Form("BM7")
		        .Parameters("BM8Index") = Request.Form("BM8")
		        .Parameters("BM9Index") = Request.Form("BM9")
		        .Parameters("BM10Index") = Request.Form("BM10")
		        .Parameters("BM11Index") = Request.Form("BM11")
		        .Parameters("BM12Index") = Request.Form("BM12")
		        .Parameters("OY1Index") = Request.Form("OY13")
		        .Parameters("OY2Index") = Request.Form("OY14")
		        .Parameters("OY3Index") = Request.Form("OY15") 
		        
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
            'response.Write "Save=" & clng(Session("BudgetID")) & ", " & clng(Session("VersionID")) & ", " & Request.Form("TransactionType") & "," & Request.Form("StaffingClassificationDesc") & "," & Request.Form("Salary") & "," & Request.Form("SalaryClassification") & "," & Request.Form("EmployeeID") & "," & Request.Form("SortOrder") & "," & Request.Form("SuperFundID") & "," & Request.Form("PerformancePay") & "," & Session("UserID")
	        'Return the result of the Save Function.
     	    strMessage = "ALL RECORDS HAVE BEEN UPDATED."			 		    	
			strMessage = UCASE(strMessage)
			strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
            strMessageColour = "Black"
			 
End Sub	

Public Sub DeleteRecord(strSalaryClassification)
'Procedure to delete Salary Classification records if no values exist against it.

Dim intAllow
Dim intCostCentreID

	'First Check to see if any values exist for the GL Code
    objRS.Open "SELECT [StaffClassification],Sum([BM1]+[BM2]+[BM3]+[BM4]+[BM5]+[BM6]+[BM7]+[BM8]+[BM9]+[BM10]+[BM11]+[BM12]) AS AllTotal, [CostCentreID] FROM tblStaffData With(NoLock) WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND StaffClassification = '" & strSalaryClassification & "' GROUP BY [StaffClassification],[CostCentreID]" ,objCon,adOpenStatic,adLockReadOnly
                	
        If objRS.EOF Then
        
            intAllow = 1
        Else
            intCostCentreID = objRS("CostCentreID")
            If objRS("AllTotal") = 0 Then
                intAllow = 1
            End If
        End If
                	
    objRS.Close
    
    
    'If there are no values for the account code then it can be deleted
	If intAllow = 1 Then
	
        objCon.Execute("DELETE FROM tblSalaryClassifications WHERE [BudgetID] = " & Session("BudgetID") & " AND SalaryClassification = '" & strSalaryClassification & "' AND VersionID = " & Session("VersionID") & "")
        'objCon.Execute("DELETE FROM tblGLCodes WHERE [BudgetID] = " & Session("BudgetID") & " AND GLCode = " & intGLCode & "")
        strMessage = "DELETED Salary Classification = " & strSalaryClassification 
		strMessage = UCASE(strMessage)
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
        strMessageColour = "Black"
	Else
        strMessage = "NOT DELETED!. Salary Classification: " & strSalaryClassification & " has values entered against it in this Budget, so CANNOT BE DELETED unless all values are removed...(CCID=" & intCostCentreID & ")"
		strMessage = UCASE(strMessage)
		strMessageIcon = "<img src=""../images/warning.gif"" />"
        strMessageColour = "Red"
    End If

End Sub


Public Sub GetMonthNames()
'This is a procedure to get the order of Month names to be used as titles for Month Columns
Dim intFirstMonth

    'set the First Month name to an integer
    intFirstMonth = Month("21-" & Session("FirstMonth") & "-2012")
    'intFirstMonth = intFirstMonth -1
    arrMonthName(0) = intFirstMonth
    For x = 1 to 12
    
        arrMonthName(x) = Left(MonthName(intFirstMonth + x - 1),3)'intFirstMonth + x
  '      arrMonthName(x) = intFirstMonth'MonthName(intFirstMonth)
        
        'Once the count goes over 12 then go back to 1 to fill the remaining months
        If intFirstMonth + x - 1 > 11 Then 
            If intFirstMonth > 6 Then
                intFirstMonth = 2 - intFirstMonth
            Else
                intFirstMonth = (x - 1) * - 1
            End If
        End If
        
  '      intFirstMonth = x
        
    Next
    
End Sub

Set objRS = Nothing
Set objCon = Nothing


%>
