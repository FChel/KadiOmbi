<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
'Description:	Business Area Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			November 2007

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim strBusinessAreaName 
Dim lngDefaultCostCentre
Dim lngRecoveryBusinessAreaID
Dim lngRecoveryGLCode
Dim lngFundingWeightingMethodID
Dim lngOHWeightingMethodID
Dim lngOHGLCode
Dim strOverHead
Dim strActive
Dim strPrimaryBA
Dim strFundingBA
Dim strCurrency
Dim strBusinessAreaCode
Dim intPLReportID
Dim intBSReportID
Dim intCFReportID
Dim strBusinessAreaNameL2
Dim dblOY1Index
Dim dblOY2Index
Dim strCCCeilingsOn
Dim strACCeilingsOn
Dim strProjCeilingsOn
Dim strContingencyBA
Dim intManagerID
Dim lngReallocationLimit
Dim lngVirementLimit

'Declare and set default arrays

Dim arrYesNo(2)

	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"
	
'3. Capture Querystring variables

    If Not IsEmpty(Request.QueryString("BusinessAreaID")) Then
	   Session("BusinessAreaID") = Request.QueryString("BusinessAreaID")
	End If
		
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Load page details
	LoadDetails()
		
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0"/>
<link rel="stylesheet" type="text/css" href="../BERTStyle.css"/>
<script type="text/javascript" src="../formChek.js"></script>
<script type="text/javascript" src="../ButtonRollOver.js"></script>
<script type="text/javascript" language="javascript">
<!--
   function SaveData()
    {
        var varSubmit = true
        var varAlert =""  
        
        if((isNonnegativeInteger(frm.BusinessAreaID.value)==false) || (frm.BusinessAreaID.value == 0))     
        {            
		    //varAlert += "Please enter Vote ID. Vote ID must be a numeric value. \n \n";
		    //document.getElementById('BusinessAreaID').style.backgroundColor="ff8080";
		    //varSubmit = false;
		    document.getElementById('BusinessAreaID').value = 0
	    }			
	    else document.getElementById('BusinessAreaID').style.backgroundColor="ffffff";	    	    
		    
	    if(isWhitespace(frm.BusinessAreaName.value))
        {            
		    varAlert += "Please enter Vote Name. \n \n";
		    document.getElementById('BusinessAreaName').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('BusinessAreaName').style.backgroundColor = "ffffff";

	    if (isWhitespace(frm.BusinessAreaNameL2.value)) {
	        varAlert += "Please enter Vote Name Swahili. \n \n";
	        document.getElementById('BusinessAreaNameL2').style.backgroundColor = "ff8080";
	        varSubmit = false;
	    }
	    else document.getElementById('BusinessAreaNameL2').style.backgroundColor = "ffffff";	  
	    
	    if(frm.Active.value == 0 )
	    {
		    varAlert += "Please select a Active Status. \n \n";
		    document.getElementById('Active').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }
		else document.getElementById('Active').style.backgroundColor = "ffffff";

		if (frm.CCCeilingsOn.value == 0) {
		    varAlert += "Please set Cost Centre Ceiling Active Status. \n \n";
		    document.getElementById('CCCeilingsOn').style.backgroundColor = "ff8080";
		    varSubmit = false;
		}
		else document.getElementById('CCCeilingsOn').style.backgroundColor = "ffffff";

		if (frm.ACCeilingsOn.value == 0) {
		    varAlert += "Please set Account Class Ceiling Active Status. \n \n";
		    document.getElementById('ACCeilingsOn').style.backgroundColor = "ff8080";
		    varSubmit = false;
		}
		else document.getElementById('ACCeilingsOn').style.backgroundColor = "ffffff";

		if (frm.ProjCeilingsOn.value == 0) {
		    varAlert += "Please set Project Ceiling Active Status. \n \n";
		    document.getElementById('ProjCeilingsOn').style.backgroundColor = "ff8080";
		    varSubmit = false;
		}
		else document.getElementById('ProjCeilingsOn').style.backgroundColor = "ffffff";
	    
	    if (isWhitespace(frm.BusinessAreaCode.value)) {
	        varAlert += "Please enter Vote Code. \n \n";
	        document.getElementById('BusinessAreaCode').style.backgroundColor = "ff8080";
	        varSubmit = false;
	    }
	    else document.getElementById('BusinessAreaCode').style.backgroundColor = "ffffff";
	    
	   if ((isFloat(frm.OY1Index.value) == false))
	    {            
		    varAlert += "Please enter Indexation for Projection Year + 1. It must be a numeric value.  \n \n";
		    document.getElementById('OY1Index').style.backgroundColor = "ff8080";
		    varSubmit = false;
	    }			
	   else document.getElementById('OY2Index').style.backgroundColor = "ffffff";
	   
	   if((isFloat(frm.OY2Index.value)==false))
       {
	       varAlert += "Please enter Indexation for Projection Year + 2. It must be a numeric value.  \n \n";
	       document.getElementById('OY2Index').style.backgroundColor = "ff8080";
	       varSubmit = false;
	   }
	   else document.getElementById('OY2Index').style.backgroundColor = "ffffff";

	   if (frm.ContingencyBA.value == 0) {
		    varAlert += "Please set Contingency Business Area Status. \n \n";
		    document.getElementById('ContingencyBA').style.backgroundColor = "ff8080";
		    varSubmit = false;
		}
	   else document.getElementById('ContingencyBA').style.backgroundColor = "ffffff";

	   if (frm.ManagerID.value == 0) {
		    varAlert += "Please select a Business Area Manager. \n \n";
		    document.getElementById('ManagerID').style.backgroundColor = "ff8080";
		    varSubmit = false;
		}
	   else document.getElementById('ManagerID').style.backgroundColor = "ffffff";

	    if ((isFloat(frm.ReallocationLimit.value) == false))
	    {            
		    varAlert += "Please enter Reallocation Limit. It must be a numeric value.  \n \n";
		    document.getElementById('ReallocationLimit').style.backgroundColor = "ff8080";
		    varSubmit = false;
	    }			
	   else document.getElementById('ReallocationLimit').style.backgroundColor = "ffffff";

	    if ((isFloat(frm.VirementLimit.value) == false))
	    {            
		    varAlert += "Please enter Virement Limit. It must be a numeric value.  \n \n";
		    document.getElementById('VirementLimit').style.backgroundColor = "ff8080";
		    varSubmit = false;
	    }			
	   else document.getElementById('VirementLimit').style.backgroundColor = "ffffff";
	   	
	  if(varSubmit == true)
	  {
	        frm.submit();
	  }
	  else
	  {
	    window.alert ("" + varAlert);	    
	  }
	  
    }

   function DeleteData()
    {
 	if(window.confirm('Confirm delete')==true){
	self.location="BusinessArea.asp?Action=Delete"
	}
}

function BusinessAreaIDSearch()
{	
	self.location="BusinessArea.asp?BusinessAreaID=" + frm.BusinessAreaID.value;

}

//-->
</script>
</head>
<body>
<h3><%=Session("BAName")%> Administration</h3>
<form action="BusinessArea.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Vote ID</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left;width:90%" id="BusinessAreaID" name="BusinessAreaID" tabindex="1" value="<%=Session("BusinessAreaID")%>" onblur="BusinessAreaIDSearch()"/></td>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Vote Name Eng</th>
	    <td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left;width:90%" id="BusinessAreaName" name="BusinessAreaName" maxlength="50" tabindex="2" value="<%=strBusinessAreaName%>" /></td>
	</tr>

	
	<tr>
		<th height="20px" style="text-align:left">&nbsp;Default CostCentre</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" READONLY id="DefaultCostCentre" name="DefaultCostCentre" tabindex="3" value="<%=lngDefaultCostCentre%>" /></td>
			<th style="text-align:left">&nbsp;Vote Name SWA</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="BusinessAreaNameL2" name="BusinessAreaNameL2" maxlength="100" tabindex="4" value="<%=strBusinessAreaNameL2%>" /></td>	
	</tr>
	
	
	<tr>
	
	<th height="20px" style="text-align:left">&nbsp;Budget Class Ceilings On</th>		
		<td><select Style="Width:40%" tabindex="5" id="PrimaryBA" name="PrimaryBA"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 2
			If strPrimaryBA = arrYesNo(x)Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		Next
		%>
		</select> </td>
		<th style="text-align:left">&nbsp;Vote Code</th><td>&nbsp;<input style="text-align:left;width:90%" id="BusinessAreaCode" name="BusinessAreaCode" maxlength="3" tabindex="6" value="<%=strBusinessAreaCode%>" /></td>
	</tr>
	<tr>
	    <th height="20px" style="text-align:left">&nbsp;Year + 1 Indexation</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" tabindex="7" id="OY1Index" name="OY1Index" value="<%=dblOY1Index%>" /></td>
	    <th style="text-align:left">&nbsp;Year + 2 Indexation</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" tabindex="8" id="OY2Index" name="OY2Index" value="<%=dblOY2Index%>" /></td>
	</tr>
	<tr>
	 <th height="20px" style="text-align:left">&nbsp;Active</th>		
		<td><select Style="Width:40%" tabindex="9" id="Active" name="Active"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 2
			If strActive = arrYesNo(x)Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		Next
		%>
		</select> </td>
        <th height="20px" style="text-align:left">&nbsp;Cost Centre Ceilings On</th>		
		<td><select Style="Width:40%" tabindex="10" id="CCCeilingsOn" name="CCCeilingsOn"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 2
			If strCCCeilingsOn = arrYesNo(x)Then
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
		  <th height="20px" style="text-align:left">&nbsp;Account Class Ceilings On</th>		
		<td><select Style="Width:40%" tabindex="11" id="ACCeilingsOn" name="ACCeilingsOn"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 2
			If strACCeilingsOn = arrYesNo(x)Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		Next
		%>
		</select> </td>
         <th height="20px" style="text-align:left">&nbsp;Project Ceilings On</th>		
		<td><select Style="Width:40%" tabindex="12" id="ProjCeilingsOn" name="ProjCeilingsOn"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 2
			If strProjCeilingsOn = arrYesNo(x)Then
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
		<th height="20px" style="text-align:left">&nbsp;Manager</th>		
		<td><select Style="Width:40%" tabindex="13" id="ManagerID" name="ManagerID"><option Value=0>Please Select....</option>
		<%
	
		    objRS.Open "SELECT * FROM tblUsers WHERE Active = 'Y'",objCon
    	          
		    Do until objRS.EOF
			    If clng(objRS("UserID")) = clng(intManagerID) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("UserID") & """" & strSelected & ">" & objRS("FName") & " " & objRS("LName") &"</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
		%>
		</select> </td>
         <th height="20px" style="text-align:left">&nbsp;Contingency Business Area</th>		
		<td><select Style="Width:40%" tabindex="14" id="ContingencyBA" name="ContingencyBA"><option Value="0">Please Select....</option>
		<%
		For x = 1 to 2
			If strContingencyBA = arrYesNo(x)Then
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
		<th height="20px" style="text-align:left">&nbsp;Reallocation Limit</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="ReallocationLimit" name="ReallocationLimit" tabindex="15" value="<%=lngReallocationLimit%>" /></td>
		<th style="text-align:left">&nbsp;Virement Limit</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="VirementLimit" name="VirementLimit" maxlength="100" tabindex="16" value="<%=lngVirementLimit%>" /></td>	
	</tr>
	
</table>
<br/>
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="10" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="11" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="12" onClick="self.location='BusinessArea.asp?BusinessAreaID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp;</button></td>
        <TD class='locked' Width="50px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
    <TD class='locked' style="Width:1000px;BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
	
</table>
<hr />
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">

	<tr>
		<th Height="20px">Vote</th>
        <th Height="20px">Vote ID</th>
	    <th>Vote Name ENG</th>
		<th>Vote Name SWA</th>
		<th>Active</th>
		<th>Updated By</th>
		<th>Date Updated</th>
	</tr>
    <tr><td colspan="9">&nbsp;</td></tr>
<%
    objRS.Open "SELECT * FROM  qryBusinessArea WHERE BudgetID = " & clng (Session("BudgetID")) & " Order By BusinessAreaCode"

	Do until objRS.eof			
   	   Response.Write "<TR><TD style=""text-align:center""><A Target=""_self"" HREF=""BusinessArea.asp?BusinessAreaID=" & objRS("BusinessAreaID") & """>&nbsp;" & objRS("BusinessAreaCode") & "</A></TD><TD Style=""text-align:center"">" & objRS("BusinessAreaID") & "</TD><TD>&nbsp;" & objRS("BusinessAreaName") & "</TD><TD>&nbsp;" & objRS("BusinessAreaNameL2") & "</TD><TD style=""text-align:center"">" & objRS("Active") & "</TD><TD style=""text-align:center"">" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
       objRS.movenext
	Loop
		
	objRS.Close
	
%>

</table>
</body>

</html>

<% 

Sub LoadDetails()

'Description:	Loads Caller's details into page if applicable.
		
		objRS.Open "SELECT * FROM tblBusinessArea WHERE BusinessAreaID = " & clng(Session("BusinessAreaID")) & " AND BudgetID = " & clng(Session("BudgetID")) & "",objCon
		
			If Not objRS.EOF Then
				
			  Session("BusinessAreaID") = objRS("BusinessAreaID")
			  strBusinessAreaName = objRS("BusinessAreaName")
              strBusinessAreaNameL2 = objRS("BusinessAreaNameL2")
              lngDefaultCostCentre = objRS("DefaultCostCentre")
              lngRecoveryBusinessAreaID = objRS("RecoveryBusinessAreaID")
              lngFundingWeightingMethodID = objRS("FundingWeightingMethodID")
              lngOHWeightingMethodID = objRS("OHWeightingMethodID")
              lngOHGLCode = objRS("OverheadGLCode")
              lngRecoveryGLCode = objRS("RecoveryGLCode")
              strOverHead = objRS("OverheadBusinessArea")
              strPrimaryBA = objRS("PrimaryBusinessArea")
              strFundingBA = objRS("FundingBusinessArea")
              strBusinessAreaCode = objRS("BusinessAreaCode")
              intPLReportID = objRS("PLReportID")
              dblOY1Index = objRS("OY1Index")
              dblOY2Index = objRS("OY2Index")
              strCCCeilingsOn = objRS("CostCentreCeilingsOn")
              strACCeilingsOn = objRS("AccountClassCeilingsOn")
              strProjCeilingsOn = objRS("ProjectCeilingsOn")
              strActive = objRS("Active")
              strCurrency = objRS("Currency")	
			  intManagerID = objRS("ManagerID")	
			  strContingencyBA = objRS("ContingencyBA")	
			  lngReallocationLimit = objRS("ReallocationLimit")
			  lngVirementLimit = objRS("VirementLimit")
			Else
			  
			  strBusinessAreaName = ""
              strBusinessAreaNameL2 = ""
              lngDefaultCostCentre = 0
              lngOHGLCode = 0
              lngRecoveryGLCode = 0
              lngRecoveryBusinessAreaID = 0
              lngFundingWeightingMethodID = 0
              lngOHWeightingMethodID = 0
              strBusinessAreaCode = ""
              strFundingBA = ""
              strPrimaryBA = ""
              intPLReportID = ""
              dblOY1Index = 0
              dblOY2Index = 0
              strOverhead =  "N"
              strActive = 0	
              strPrimaryBA = 0
              strCCCeilingsOn = ""
              strACCeilingsOn = ""
              strProjCeilingsOn = ""

		    End If

		objRS.Close
	

End Sub

Sub SaveDetails()

Dim strCeilingStatus

       'response.Write(Request.Form("VersionID"))

        objRS.Open "SELECT Approved FROM tblBACeilingLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = 1000 AND Level1ID = " & Session("BusinessAreaID") & "",objCon

            If Not objRS.EOF Then
                strCeilingStatus = objRS(0)
            Else
                strCeilingStatus = "off"
            End If

        objRS.Close

        IF Request.Form("PrimaryBA") = "N" AND strCeilingStatus = "on " Then
                strMessage = "<FONT Color=""Red""><B>BUDGET CLASS CEILINGS CANNOT BE REMOVED IF VOTE EXPENDITURE CEILINGS HAVE BEEN APPROVED.</B></FONT>"
                strMessageIcon = "<img src=""../images/warning.gif"" />"
             
        Else
		    With objCmd
                .CommandType = 4
                .CommandText = "spBusinessAreaSave"
                
                .Parameters.Append objCmd.CreateParameter("BusinessAreaID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Currency", adChar, adParamInput, 3)
                .Parameters.Append objCmd.CreateParameter("BusinessAreaName", adVarChar, adParamInput, 100)
                .Parameters.Append objCmd.CreateParameter("BusinessAreaNameL2", adVarChar, adParamInput, 100)
                .Parameters.Append objCmd.CreateParameter("DefaultCostCentre", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OverheadBusinessArea", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("RecoveryBusinessAreaID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OverheadGLCode", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("RecoveryGLCode", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("FundingWeightingMethodID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OHWeightingMethodID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("PrimaryBusinessArea", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("BusinessAreaCode", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("FundingBusinessArea", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("PLReportID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BSReportID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("CFReportID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY1Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY2Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("CCCeilingsOn", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("ACCeilingsOn", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("ProjCeilingsOn", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput, 1)
			 	.Parameters.Append objCmd.CreateParameter("ManagerID", adInteger, adParamInput)		
				.Parameters.Append objCmd.CreateParameter("ReallocationLimit", adDouble, adParamInput)
				.Parameters.Append objCmd.CreateParameter("VirementLimit", adDouble, adParamInput)
				.Parameters.Append objCmd.CreateParameter("ContingencyBA", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
               ' .Parameters.Append objCmd.CreateParameter("VersionIDOutput", adInteger, adParamOutput)
              
				.Parameters("BusinessAreaID") = Request.Form("BusinessAreaID")		
				.Parameters("BudgetID") = Session("BudgetID")
                .Parameters("VersionID") = Session("VersionID")
				.Parameters("Currency") = "TZS"	
                .Parameters("BusinessAreaName") = Request.Form("BusinessAreaName")
                .Parameters("BusinessAreaNameL2") = Request.Form("BusinessAreaNameL2")
                .Parameters("DefaultCostCentre") = Request.Form("DefaultCostCentre") 
                .Parameters("OverheadBusinessArea") = NULL
                .Parameters("RecoveryBusinessAreaID") = NULL
                .Parameters("OverheadGLCode") = NULL
                .Parameters("RecoveryGLCode") = NULL             
                .Parameters("FundingWeightingMethodID") = 0
                .Parameters("OHWeightingMethodID") = NULL
                .Parameters("PrimaryBusinessArea") = Request.Form("PrimaryBA")
                .Parameters("BusinessAreaCode") = Request.Form("BusinessAreaCode")
                .Parameters("FundingBusinessArea") = NULL
                .Parameters("PLReportID") = 1
                .Parameters("BSReportID") = 0
                .Parameters("CFReportID") = 0
                .Parameters("OY1Index") = Request.Form("OY1Index")
                .Parameters("OY2Index") = Request.Form("OY2Index")
            
                .Parameters("CCCeilingsOn") = Request.Form("CCCeilingsOn")
                .Parameters("ACCeilingsOn") = Request.Form("ACCeilingsOn")
                .Parameters("ProjCeilingsOn") = Request.Form("ProjCeilingsOn")
                .Parameters("Active") = Request.Form("Active")
				.Parameters("ManagerID") = Request.Form("ManagerID")
				.Parameters("ReallocationLimit") = Request.Form("ReallocationLimit")
				.Parameters("VirementLimit") = Request.Form("VirementLimit")
				.Parameters("ContingencyBA") = Request.Form("ContingencyBA")
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
			'Return the result of the Save Function.
     		strMessage = "<B>RECORD SAVED.</B>"
            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
	    End If
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
