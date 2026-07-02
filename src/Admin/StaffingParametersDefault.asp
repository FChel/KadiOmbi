<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	Staffing Parameters Default Admin Screen
'Author:		MG
'Date:			Janauray 2014

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

Dim strSQL 
Dim dblWTotal
Dim strCalcList
Dim strTransactionName

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'Declare and set default arrays
Dim arrYesNo(4,1)

	arrYesNo(1,0) = "Y"
	arrYesNo(2,0) = "N"
	arrYesNo(3,0) = "E"
	arrYesNo(4,0) = "U"

    arrYesNo(1,1) = "Yes"
	arrYesNo(2,1) = "No"
	arrYesNo(3,1) = "Editable GL"
	arrYesNo(4,1) = "Uneditable GL"
	
'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim strStaffingParameterID
Dim lngCostCentreID
Dim strTransactionType
Dim dblRate
Dim lngDRGLCode
Dim lngCRGLCode
Dim intCalculationID
Dim strDescription
Dim strEditable
Dim arrCalculation(5,1)
Dim strCalcDetail

    'Set the Caluclation values for the Drop-down
    arrCalculation(1,0) = "Salary"
    arrCalculation(2,0) = "Superannuation (Employee specific in Salary Schedule, not Rate value above)"
	arrCalculation(3,0) = "Rate Based (Salary Only)"
	arrCalculation(4,0) = "Performance Pay (Employee specific in Salary Schedule, not Rate value above)"
	arrCalculation(5,0) = "Rate Based (Salary AND Superannuation)"

	arrCalculation(1,1) = "Formula [MonthValue = (Salary*FTE/WorkDaysInTheYear) * WorkDaysInTheMonth]"
    arrCalculation(2,1) = "Formula [MonthValue = TotalSalary * (WorkDaysInTheMonth/WorkDaysInTheYear) * (SuperFundRate/100)]"
	arrCalculation(3,1) = "Formula [MonthValue = TotalSalary * (WorkDaysInTheMonth/WorkDaysInTheYear) * (Rate/100)]"
	arrCalculation(4,1) = "Formula [MonthValue = TotalSalary * (WorkDaysInTheMonth/WorkDaysInTheYear) * (PerformancePayRateForEmployee/100)]"
	arrCalculation(5,1) = "Formula [MonthValue = (TotalSalary + Superannuation) * (WorkDaysInTheMonth/WorkDaysInTheYear) * (Rate/100)]"
	
If Not IsEmpty(Request.QueryString("CostCentreID")) Then		
    Session("CostCentreID") = Request.QueryString("CostCentreID")
    lngCostCentreID = clng(Request.QueryString("CostCentreID"))	
Else
    lngCostCentreID = Session("CostCentreID")				
End If	

If Not IsEmpty(Request.QueryString("ParameterID")) Then		
   strStaffingParameterID = cstr(Request.QueryString("ParameterID"))	
Else
   strStaffingParameterID = ""
End If  

If Not IsEmpty(Request.QueryString("TransactionType")) Then		
	Session("TransactionType") = cstr(Request.QueryString("TransactionType"))
	strTransactionType = Session("TransactionType")
Else
	strTransactionType = Session("TransactionType")
End If 

If strTransactionName = "ESTA" Then
	strTransactionName = "EXISTING STAFFING PARAMETERS ADMINISTRATION"
Else
	strTransactionName = "NEW STAFFING PARAMETERS ADMINISTRATION"
End If

'Execute save 	
If Request.QueryString("Action") = "Save" Then
    SaveDetails()
End If

'Execute save all 	 
If Request.QueryString("Action") = "SaveAll" Then
    SaveAllDetails()
End If

'Execute delete 	
If Request.QueryString("Action") = "Delete" Then
    deleteRecord(Request.QueryString("StaffingParameterID"))
End If

'Execute Recalculate procedure 	
If Request.QueryString("Action") = "recalculate" Then
	RecalcStaff()
End If
	
'Load page details
if IsEmpty(Request.QueryString("Show")) then
    LoadDetails()
End if	
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
	<script src="../formChek.js">
	</script>
<script LANGUAGE="javascript">
<!--
function SaveData()
{	    
    var varSubmit = true						
    var varAlert="";	   
     
    if(isFloat(document.getElementById('Rate').value)==false)
	{
       varAlert += "Rate must be greater than 0. \n \n";
       document.getElementById('Rate').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('Rate').style.backgroundColor="ffffff"; 
    
    if(isInteger(document.getElementById('CalculationID').value)==false)
	{
       varAlert += "Calculation must be greater than 0. \n \n";
       document.getElementById('CalculationID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('CalculationID').style.backgroundColor="ffffff";   
    
    if(isWhitespace(frm.StaffingParameterID.value))
        {            
		    varAlert += "Please enter Staffing Parameter. \n \n";
		    document.getElementById('StaffingParameterID').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	else document.getElementById('StaffingParameterID').style.backgroundColor="ffffff";    
	    
   
    if(isWhitespace(frm.Description.value))
        {            
		    varAlert += "Please enter Description. \n \n";
		    document.getElementById('Description').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	else document.getElementById('Description').style.backgroundColor="ffffff";  
	  
	//if(frm.DRGLCode.value=="0")
	//{
     //  varAlert += "Please select a DR GL Code. \n \n";
    //   document.getElementById('DRGLCode').style.backgroundColor="ff8080";
    //   varSubmit = false;
    //}   
    //else document.getElementById('DRGLCode').style.backgroundColor="ffffff";  
    
    //if(frm.CRGLCode.value=="0")
	//{
    //   varAlert += "Please select a CR GL Code. \n \n";
    //   document.getElementById('CRGLCode').style.backgroundColor="ff8080";
    //   varSubmit = false;
    //}   
    //else document.getElementById('CRGLCode').style.backgroundColor="ffffff";
            
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}


function CCIDSearch()
{	        
	self.location="StaffingParametersDefault.asp?Show=True&CostCentreID=" + frm.CostCentreID.value	
}

function DeleteData() {
 	if(window.confirm('Would you like to DELETE '  + frm.StaffingParameterID.value + '?')==true){
	self.location="StaffingParametersDefault.asp?Action=Delete&StaffingParameterID=" + frm.StaffingParameterID.value
	}
}

function SaveAllData()
    {
        document.getElementById('Progress').style.display = "inline";
        self.location='StaffingParametersDefault.asp?Action=SaveAll';
    }
function ChangeCalc()
    {
        //document.getElementById('CalcDetail').value = document.getElementById('CalculationID').options[document.getElementById('CalculationID').selectedIndex].text;
        document.getElementById('CalcDetail').value = document.getElementById('CalculationID').value;
}

function StaffRecalc(varBud, varVer)
 {
 	if(window.confirm('Would you like to Recalculate ALL existing Staff values? \n \n For the currently selected Budget ' + varBud + ' and Version ' + varVer + '.')==true){
 	
 		document.getElementById('Progress').style.display = "inline";
        self.location='StaffingParametersDefault.asp?Action=recalculate';
	}
}
//-->
</script>
</head>
<body>
<form action="StaffingParametersDefault.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
	    <th style="width:30%;height:20px;"><%=strTransactionName%></th>
	    <th style="width:70%"></th>
	</tr>
	<tr>
	    <td colspan="2">&nbsp;</td>
	</tr>	
	</TABLE>
	<BR>
	<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
        <th align="left" Width="20%">&nbsp;Staffing Parameter</th>
		<td align="left" Width="30%">&nbsp;
        <input id="StaffingParameterID" name="StaffingParameterID" Style="text-align:left;" tabindex="1" value="<%=strStaffingParameterID%>"/></td>
	    <th align="left" Width="20%">&nbsp;Description</th>
		<td align="left" Width="30%">&nbsp;<input style="text-align:left; width:100%" tabindex="2" id="Description" name="Description" type="text" value="<%=strDescription%>"/></td>
	</tr>
	<tr>
        
	    <th align="left" Width="20%">&nbsp;Rate</th>
		<td align="left" Width="30%">&nbsp;<input style="Width:95%; text-align:left;"  id="Rate" name="Rate" tabindex="4" type="text" value="<%=dblRate%>"/></td>
		<th align="left">&nbsp;DR&nbsp;GL Code</th>		
		<td><select Style="Width:95%;" tabindex="5" id="Select1" name="DRGLCode"><option Value="0">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM tblGLCodes WITH(NOLOCK) WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon,0,1
		
		Do until objRS.EOF
			If objRS("GLCOde") = lngDRGLCode Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("GLCode") & """" & strSelected & ">" & objRS("GLCode") & " : " & objRS("GLCodeName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>	
	</tr>
	<tr>
	<th align="left" Width="20%">&nbsp;Calculation Type</th>
		<td align="left" Width="30%"><select style="text-align:left;" tabindex="7" id="CalculationID" name="CalculationID" onchange="ChangeCalc();"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 5
			If intCalculationID = x Then
				strSelected = " SELECTED " 
				strCalcDetail = arrCalculation(x,1)
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & x & """" & strSelected & ">" & arrCalculation(x,0) & "</OPTION>"
		Next
		%>
		</select> </td>
        
	<th align="left">&nbsp;CR&nbsp;GL Code</th>		
		<td><select style="Width:95%;" tabindex="6" id="CRGLCode" name="CRGLCode"><option Value="0">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM tblGLCodes WITH(NOLOCK) WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon,0,1
		
		Do until objRS.EOF
			If objRS("GLCode") = lngCRGLCode Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("GLCode") & """" & strSelected & ">" & objRS("GLCode") & " : " & objRS("GLCodeName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>	
	</tr>
	<tr>
      <td colspan="2"><span style="color:gray; font-weight:bold">Selected Calculation Type Detail: </span>
      <select style="text-align:left; Width:100%; color:gray;" DISABLED id="CalcDetail" name="CalcDetail"><option Value="0"></option>
		<%
		For x = 1 to 5
			If intCalculationID = x Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & x & """" & strSelected & ">" & arrCalculation(x,1) & "</OPTION>"
		Next
		%>
		</select></td>
	  <th align="left">&nbsp;Editable</th>		
		<td><select style="Width:40%" tabindex="8" id="Editable" name="Editable">
		<%
		For x = 1 to 4
			If strEditable = arrYesNo(x,0)Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & arrYesNo(x,0) & """" & strSelected & ">" & arrYesNo(x,0) & " - " & arrYesNo(x,1) & "</OPTION>"
		Next
		%>
		</select> </td>
	</tr>
		
</table>
<br>
<div class="buttons">
<TABLE Width="1500px" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<TR>

    <td Width="100px"><button type="button" onclick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
    <td Width="100px"><button type="button" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td> 
    <td Width="150px"><button type="button" onclick="self.location='StaffingParametersDefault.asp?StaffingParameterID=0'" )""><img src="../images/add.png" alt="" /> Clear/Add New </button></td>
    <td Width="100px"><button type="button" onclick="DeleteData()";><img src="../images/delete.png" alt="" /> Delete </button></td>
  <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
    <TD class='locked' align="left" Width="800x" style="BORDER-RIGHT:0px"><INPUT style="Align:Left; font-weight:Bold; width:100%; text-align:left; color:<%=strMessageColour%>;" type="text" id="msgbox" name="msgbox" value="<%=strMessage%>"></TD>
</TABLE>
</div>

<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr><td colspan="7">&nbsp;<span style="color:gray; font-weight:bold;">Note:</span><span style="color:gray;"> Changes to Salary Parameter details does not update existing Staff records. The Recalculate button (above) must be clicked to recalculate existing records.</span></td></tr>
	<tr>
	  	<th align="center">Parameter</th>	
		<th align="center">Description</th>				
		<th align="center">Rate</th>
	    <th align="center">DR GL Code</th>
	    <th align="center">CR GL Code</th>
	    <th align="center">Calculation Type</th>
	    <th align="center">Editable</th>						
	</tr>
	
<%
    
       
    dblWTotal = 0
    
        strSQL =  "SELECT * FROM tblStaffingParametersDefault WITH(NOLOCK) WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND TransactionType = '" & Session("TRansactionType") & "' ORDER BY TransactionType, StaffingParameterID"
       
    objRS.Open strSQL,objCon,0,1
    
		Do until objRS.eof
		
		    If IsNull(objRS("CalculationID")) Then
		        strCalcList = ""
		    Else
		        IF objRS("CalculationID") > 0 AND objRS("CalculationID") < 6 Then
		            strCalcList = left(arrCalculation(objRS("CalculationID"),0),25)
		        Else
		            strCalcList = objRS("CalculationID")
		        End If
		    End If
		    
		    Response.Write "<TR><TD><A Target='_self' HREF='StaffingParametersDefault.asp?ParameterID=" & objRS("StaffingParameterID") & "&TransactionType=" & objRS("TransactionType") & "'>&nbsp;" & objRS("StaffingParameterID") & "</TD>" & _
		                    "<TD>&nbsp;" & objRS("Description") & "</TD><TD style=""text-align:right"">&nbsp;" & formatnumber(objRS("Rate"),3,0) & "</TD><TD style=""text-align:center"">" & objRS("DRGLCode") & "&nbsp;</TD>" & _
		                    "<TD style=""text-align:center"">" & objRS("CRGLCode") & "</TD><TD style=""text-align:center"">" & strCalcList & "</TD><TD style=""text-align:center"">" & objRS("Editable") & "</TD></TR>"
			objRS.movenext
		Loop
			
	objRS.Close
	
	    If IsNull(dblWTotal) Then 
	        dblWTotal = 0
	    End If 
	    
	    Response.Write "<TR><TH Colspan=""7"">&nbsp;</TH>"

%>
</table>
</body>

</html>

<% 

Sub LoadDetails()

'Description:	Loads Staffing Parameter details into page if applicable.
		
		objRS.Open "SELECT * FROM tblStaffingParametersDefault WITH(NOLOCK) WHERE StaffingParameterID = '" & strStaffingParameterID & "' AND BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND TransactionType = '" & strTransactionType & "'",objCon,0,1
	
		If Not objRS.EOF Then
		
		    strStaffingParameterID = objRS("StaffingParameterID")
            strTransactionType = objRS("TransactionType")
            dblRate = objRS("Rate")
            lngDRGLCode = objRS("DRGLCode")
            lngCRGLCode = objRS("CRGLCode")
            intCalculationID = objRS("CalculationID")
            strDescription = objRS("Description")
            'strCCCategory = objRS("CostCentreCategory")
            strEditable = objRS("Editable")
                   						
		Else		                
		    'Do nothing                         
		End if

		objRS.Close	
End Sub

Sub deleteRecord(strStaffingParameterID)
    objCon.Execute "Delete from tblStaffingParametersDefault where StaffingParameterID = '" & strStaffingParameterID & "' AND BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & ""   
    'response.Write "Delete from tblStaffingParametersDefault where StaffingParameterID = '" & strStaffingParameterID & "' AND BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & ""  

       strMessage = strStaffingParameterID & " Staffing Parameter DELETED"	
End Sub

Sub SaveDetails()            

    objCon.Execute "spStaffingParametersDefaultSave '" & Request.Form("StaffingParameterID") & "'," & Session("BudgetID") & "," & Session("VersionID") & ",'" & strTransactionType & "'," & Request.Form("Rate") & "," & Request.Form("DRGLCode") & "," & Request.Form("CRGLCode") & "," & Request.Form("CalculationID") & ",'" & Request.Form("Description") & "'," & Request.Form("Editable") & "," & Session("UserID") & ""
    'response.Write "spStaffingParametersDefaultSave '" & Request.Form("StaffingParameterID") & "'," & Session("BudgetID") & "," & Session("VersionID") & ",'" & Request.Form("TransactionType") & "'," & Request.Form("Rate") & "," & Request.Form("DRGLCode") & "," & Request.Form("CRGLCode") & "," & Request.Form("CalculationID") & ",'" & Request.Form("Description") & "'," & Request.Form("Editable") & "," & Session("UserID") & ""
    strMessage = "RECORD SAVED."
                strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
                strMessageColour = "Black"										
     				
End Sub	

Sub SaveAllDetails()            

    objCon.Execute "spApplyDefaultStaffingParameters " & Session("BudgetID") & "," & Session("VersionID") & "," & Session("UserID") & ""
				strMessage = "RECORD SAVED."
                strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
                strMessageColour = "Black"									
     				
End Sub	

Sub RecalcStaff()

	 With objCmd
            .CommandType = 4
            .CommandText = "spStaffDataSalaryUpdate"
            
			.Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
      	
			.Parameters("BudgetID") = clng(Session("BudgetID"))
			.Parameters("VersionID") = clng(Session("VersionID"))
			
            .ActiveConnection = objCon
            
        End With
            
        objCmd.Execute          
           
		'Return the result of the Save Function.
 	    strMessage = "Staff Values have been recalculated!"
	
End Sub	

Function MediumDate (str)
	
	'Function to change all date formats to medium date to avoid American storage challenge!
	
	Dim aDay
	Dim aMonth
	Dim aYear
	
		aDay = 	(Left((str),InStr(1,(str),"/")-1))
		aMonth = Mid(str,(InStr(1,(str),"/")+1),2)
	
	If Right(aMonth,1) = "/" Then
		aMonth = Left(aMonth,1)
	End If
	
		aMonth = MonthName(aMonth)
		aYear = Year(str)
	
	If Len(aDay) = 1 Then aDay = "0" & aDay
	
		MediumDate = aDay & "-" & aMonth & "-" & aYear
		
End Function

Set objRS = Nothing
Set objCon = Nothing


%>
