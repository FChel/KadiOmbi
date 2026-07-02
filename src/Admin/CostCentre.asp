<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

    Session("CurrentPage") = "Admin/CostCentre.asp"

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
'Description:	Student Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			August 2004

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim objRS1
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngCostCentreID
Dim strCostCentreName
Dim strCostCentreNameL2
Dim strCostCentreType
Dim intCostObjectTypeID
Dim lngParentCostCentreID
Dim lngBusinessAreaID
Dim strProjectCode
Dim strProjectName
Dim strSourceCode
Dim strSourceName
Dim strDivisionCode
Dim strDivisionName
Dim strLocalCurrency
Dim strCapitalisationPeriod
Dim intWIPUsefulLife
Dim lngWIPGLCode
Dim lngAssetGLCode
Dim lngDepnGLCode
Dim strCCCategory
Dim intLGAID
Dim intStateElectorateID
Dim intStatisticalDivisionID
Dim strFrontLineCostObject
Dim strBlockedCC
Dim lngAlternateCC
Dim strRecurrent
Dim strPASP
Dim strActive
Dim strCCClass
Dim strCeilingType
Dim strVote
Dim intInputSheetID

'Declare and set default arrays

Dim arrActive(2)
	
	arrActive(1) = "Y"
	arrActive(2) = "N"
	

 	'3. Capture Querystring variables	
	If Not IsEmpty(Request.QueryString("CostCentreID")) Then		
		Session("CostCentreID") = Request.QueryString("CostCentreID")
		lngCostCentreID = clng(Request.QueryString("CostCentreID"))					
	End If		
	
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

    If IsNumeric(Session("CostCentreID")) = False THen
        Session("CostCentreID")  = 0
        lngCostCentreID = 0
    End If

	'Load page details
	LoadDetails()

     objRS.Open "SELECT BusinessAreaCode FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon

        If Not objRS.EOF Then
            strVote = objRS("BusinessAreaCode")
        End If

    objRS.Close
		
			
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
	<script src="../formChek.js">
	</script>
	<script src="../ButtonRollOver.js">
	</script>

<script LANGUAGE="javascript">
<!--
function SaveData()
{	    
    var varSubmit = true
    var varAlert = "";

	if ((isNonnegativeInteger(frm.CostCentreID.value) == false) || (frm.CostCentreID.value == 0)) {
        varAlert += "Please enter Cost Centre ID. Cost Centre ID must be a numeric value. \n \n";
        document.getElementById('CostCentreID').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('CostCentreID').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.ProjectCode.value)) {
        varAlert += "Cost Centre / Sub Cost Centre Code Cannot Be Blank. \n \n";
        document.getElementById('ProjectCode').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('ProjectCode').style.backgroundColor = "ffffff";          
    
    if(isWhitespace(frm.CostCentreName.value))
	{
       varAlert += "Name Eng Cannot Be Blank. \n \n";
       document.getElementById('CostCentreName').style.backgroundColor="ff8080";
       varSubmit = false;
    }
   else document.getElementById('CostCentreName').style.backgroundColor = "ffffff";

   if (isWhitespace(frm.CostCentreNameL2.value)) {
       varAlert += "Name Swa Cannot Be Blank. \n \n";
       document.getElementById('CostCentreNameL2').style.backgroundColor = "ff8080";
       varSubmit = false;
   }
   else document.getElementById('CostCentreNameL2').style.backgroundColor = "ffffff";

  if ((isNonnegativeInteger(frm.InputSheetID.value) == false) || (frm.InputSheetID.value == 0)) {
        varAlert += "Please enter Input Sheet ID. Input Sheet ID must be a numeric value. \n \n";
        document.getElementById('InputSheetID').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('InputSheetID').style.backgroundColor = "ffffff";


   if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}


function CostCentreIDSearch()
{	
	self.location="CostCentre.asp?CostCentreID=" + frm.CostCentreID.value
}
//-->
</script>
</head>
<body>
<h3>Cost Centre Administration Screen</h3>
<form action="CostCentre.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
        <th style="text-align:left; height:20px; width:20%;">&nbsp;Cost Centre Code</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; width:50%" id="ProjectCode" name="ProjectCode" maxlength="4" TABINDEX="1" value="<%=strProjectCode%>">		
	 <th style="text-align:left; height:20px; width:20%;">&nbsp;Cost Centre ID</th>	
	<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; width:99%"  id="CostCentreID" name="CostCentreID" maxlength="50" TABINDEX="2" value="<%=lngCostCentreID%>"></td>		
	</tr>
    <tr>	 
   <th style="text-align:left; height:20px; width:20%;">&nbsp;Cost Centre Name Eng</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; width:99%" id="CostCentreName" name="CostCentreName" maxlength="50" TABINDEX="3" value="<%=strCostCentreName%>"></td>
	
		<th align="left">&nbsp;Cost Centre Name Swa</th>
		<td>&nbsp;<input style="text-align:left;width:99%" id="CostCentreNameL2" name="CostCentreNameL2" maxlength="200" TABINDEX="4" value="<%=strCostCentreNameL2%>"></td>
	    
    </tr>
	<tr>
        <th align="left">&nbsp;Cost Object Type</th>
		<td>
	       <select Style="Width:80%" tabindex="5" id="CostObjectTypeID" name="CostObjectTypeID">
	 <% 	
		objRS.Open "SELECT * FROM tblCostObjectTypes WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon
		
		Do until objRS.EOF
			If objRS("CostObjectTypeID") = intCostObjectTypeID Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("CostObjectTypeID") & """" & strSelected & ">" & objRS("CostObjectTypeName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
	%>
	       </select>
	    </td>
	    
	    <th align="left">&nbsp;Parent Cost Centre</th>
		<td><input style="text-align:left" DISABLED style="width:100%" id="ParentCostCentreID" name="ParentCostCentreID" TABINDEX="6" value="<%=lngParentCostCentreID%>">
	
	    </td>
	</tr>


<tr>  
	  
	    <th align="left">&nbsp;Active</th>
		<td>
	       <select Style="Width:40%" tabindex="12" id="Active" name="Active">
	        <%
		        For x = 1 to 2
			        If arrActive(x) = cstr(strActive) Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrActive(x) & """" & strSelected & ">" & arrActive(x) & "</OPTION>"
		        Next
	        %>
	       </select>
	    </td><td colspan="2"></td>
</tr>
<tr>
    <th style="text-align:left">&nbsp;Cost Centre Class</th><td colspan="3">&nbsp;<input style="text-align:left; width:100%" id="CCClass" name="CCClass" maxlength="200" TABINDEX="13" value="<%=strCCClass%>"></td>
</tr>
<tr>
    <th style="text-align:left">&nbsp;Ceiling Type</th><td colspan="3">&nbsp;<input style="text-align:left; width:100%" id="CeilingType" name="CeilingType" maxlength="200" TABINDEX="15" value="<%=strCeilingType%>"></td>
</tr>
<tr>
    <th style="text-align:left">&nbsp;Input Sheet</th><td colspan="3">&nbsp;<input style="text-align:left; width:100%" id="InputSheetID" name="InputSheetID" maxlength="200" TABINDEX="15" value="<%=intInputSheetID%>"></td>
</tr>
</table>
<br>
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="21" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="22" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="23" onClick="self.location='CostCentre.asp?CostCentreID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
        <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr>
</form>
<H3>List of ALL Cost Centres</H3>
<!--<H3>List of Cost Centres for Vote : <FONT Color="red"><%Response.Write Session("Vote") %></FONT></H3>-->
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
		<th width="10%" height="20px" align="center">Cost Centre Code</th>
		<th width="10%" height="20px" align="center">Cost Centre ID</th>		
		<th width="25" align="center">Cost Centre Name Eng</th>
        <th width="25%" align="center">Cost Centre Name Swa</th>
       	<th width="20%" align="center">Cost Centre Type</th>	
	<th width="5%" align="center">Input Sheet ID</th>
		<th width="5%"align="center">Active</th>
	</tr>
	
<%
    
    If Session("BusinessAreaID") = 1000 Then
        objRS.Open "SELECT * FROM qryCostCentresList WHERE BudgetID = " & Session("BudgetID") & " AND CostCentreID > 0 Order By CostCentreID ASC",objCon
	Else
        objRS.Open "SELECT * FROM qryCostCentresList WHERE BudgetID = " & Session("BudgetID") & " AND Left(CostCentreID,4) = " & Session("BusinessAreaID") & " Order By CostCentreID ASC",objCon
    End If
    
    	Do until objRS.eof
			Response.Write "<TR><TD><A Target=""_self"" HREF=""CostCentre.asp?CostCentreID=" & objRS("CostCentreID") & """><B>&nbsp;" & objRS("ProgramCode") & "</A></TD><TD>&nbsp;" & objRS("CostCentreID") & "</B></TD><TD>&nbsp;" & objRS("CostCentreName") & "</B></TD>" & _
				"<TD>&nbsp;" & objRS("CostCentreNameL2") & "</B></TD><TD style=""text-align:center"">" & objRS("CostObjectTypeName") & "</TD><TD>&nbsp;" & objRS("InputSheetID") & "</B></TD><TD style=""text-align:center"">" & objRS("Active") & "</TD></TR>"
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
		
		objRS.Open "SELECT * FROM tblCostCentres WHERE BudgetID = " & Session("BudgetID") & " AND CostCentreID = " & clng(Session("CostCentreID")) & "",objCon							
		If Not objRS.EOF Then
		
            lngBusinessAreaID = Left(objRS("CostCentreID"),4)
		    lngCostCentreID = objRS("CostCentreID")
            strCostCentreName = objRS("CostCentreName")
            strCostCentreNameL2 = objRS("CostCentreNameL2")
            strCostCentreType = objRS("CostCentreType")
            strCCCategory = objRS("CostCentreCategory")
            intCostObjectTypeID = objRS("CostObjectTypeID")
            lngParentCostCentreID = objRS("ParentCostCentreID")
            strProjectCode = objRS("ProgramCode")
            strProjectName = objRS("ProgramName")
            strSourceCode = objRS("SourceCode")
            strSourceName = objRS("SourceName")
            strDivisionCode = objRS("DivisionCode")
            strDivisionName = objRS("DivisionName")
            strLocalCurrency = objRS("LocalCurrency")
            strCapitalisationPeriod = objRS("CapitalisationPeriod")
            intWIPUsefulLife = objRS("WIPUsefulLife")
            lngWIPGLCode = objRS("WIPGLCode")
            lngAssetGLCode = objRS("AssetGLCode")
            lngDepnGLCode = objRS("DepnGLCode")
            intLGAID = objRS("LGAID")
            intStateElectorateID = objRS("StateElectorateID")
            intStatisticalDivisionID = objRS("StatisticalDivisionID")
            strFrontLineCostObject = objRS("FrontLineCostObject")
            strRecurrent = objRS("Recurrent")
            strBlockedCC = objRS("BlockedCC")
            lngAlternateCC = objRS("AlternateCC")
            strPASP = objRS("PASP")
            strActive = objRS("Active")	
            strCCClass = objRS("CostCentreClass")
            strCeilingType = objRS("CeilingType")
            intInputSheetID = objRS("InputSheetID")
     
		End if

		objRS.Close	
End Sub

Sub SaveDetails()

Dim strExists   
Dim intCostCentreID

   '13th July 2016 MG - Disabled check in Business Area Table for Cost Centre Name
    'objRS1.Open "SELECT * FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaCode = '" & Left(Request.Form("ProjectCode") ,3) & "'",objCon
    
    '    If Not objRS1.EOF Then		

		    With objCmd
                .CommandType = 4
                .CommandText = "spCostCentreSave"
                  
                .Parameters.Append objCmd.CreateParameter("CostCentreID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
		        .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("CostCentreName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("CostCentreNameL2", adVarChar, adParamInput, 200)
                .Parameters.Append objCmd.CreateParameter("CostCentreType", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("CostCentreCategory", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("CostObjectTypeID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("ParentCostCentreID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("ProjectCode", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("ProjectName", adVarChar, adParamInput, 100)
                .Parameters.Append objCmd.CreateParameter("SourceCode", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("SourceName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("DivisionCode", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("DivisionName", adVarChar, adParamInput, 100)
                .Parameters.Append objCmd.CreateParameter("LocalCurrency", adVarChar, adParamInput, 3)
                .Parameters.Append objCmd.CreateParameter("CapitalisationPeriod", adVarChar, adParamInput, 5)
                .Parameters.Append objCmd.CreateParameter("WIPUsefulLife", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("WIPGLCode", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("AssetGLCode", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("DepnGLCode", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Recurrent", adVarChar, adParamInput, 1)
                
                .Parameters.Append objCmd.CreateParameter("LGAID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("StateElectorateID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("StatisticalDivisionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("FrontLineCostObject", adVarChar, adParamInput, 1)
                
                .Parameters.Append objCmd.CreateParameter("BlockedCC", adVarChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("AlternateCC", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("PASP", adVarChar, adParamInput, 1)
                     
                .Parameters.Append objCmd.CreateParameter("Active", adVarChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("InputSheetID", adInteger, adParamInput)
                
                .Parameters.Append objCmd.CreateParameter("CostCentreClass", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("CeilingType", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)   
        
                .Parameters.Append objCmd.CreateParameter("Response", adVarChar, adParamOutput, 100)                                                                 
              
                If IsNumeric(Request.Form("CostCentreID")) = False Then
                    intCostCentreID = 0
                Else
                    intCostCentreID = Request.Form("CostCentreID")
                End IF
		        .Parameters("CostCentreID") = intCostCentreID

		        .Parameters("BudgetID") = Session("BudgetID")
		        .Parameters("VersionID") = Session("VersionID")	
                .Parameters("CostCentreName") = Request.Form("CostCentreName")
                .Parameters("CostCentreNameL2") = Request.Form("CostCentreNameL2")
                .Parameters("CostCentreType") = NULL
                .Parameters("CostCentreCategory") = NULL
                .Parameters("CostObjectTypeID") = Request.Form("CostObjectTypeID")
                .Parameters("ParentCostCentreID") = NULL
                .Parameters("ProjectCode") = Request.Form("ProjectCode")      
                .Parameters("ProjectName") = NULL     
                .Parameters("SourceCode") = NULL   
                .Parameters("SourceName") = NULL    
                .Parameters("DivisionName") = NULL
                .Parameters("DivisionCode") = NULL    
                .Parameters("LocalCurrency") = "TZS"      
                .Parameters("CapitalisationPeriod") = NULL
                .Parameters("WIPUsefulLife") = NULL
                .Parameters("WIPGLCode") = NULL
                .Parameters("AssetGLCode") = NULL
                .Parameters("DepnGLCode") = NULL
                .Parameters("Recurrent") = ""
                
                .Parameters("LGAID") = NULL  
                .Parameters("StateElectorateID") = NULL 
                .Parameters("StatisticalDivisionID") = NULL  
                .Parameters("FrontLineCostObject") = NULL
                
                .Parameters("BlockedCC") = "N"
                .Parameters("AlternateCC") = NULL  
                .Parameters("PASP") = NULL                       
                .Parameters("Active") = Request.Form("Active") 
                .Parameters("InputSheetID") = Request.Form("InputSheetID") 
                .Parameters("CostCentreClass") = Request.Form("CCClass") 
                .Parameters("CeilingType") = Request.Form("CeilingType")     
                .Parameters("UpdatedBy") = Session("UserID")
                         
                .ActiveConnection = objCon
                
            End With
            If strExists <> "Y" Then    
                objCmd.Execute          
                strMessage = objCmd.Parameters.Item("Response")      	
                If strMessage = "OK" Then		     				
     		        strMessage = "<B>RECORD SAVED.</B>"
                    strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"							
     		        Session("CostCentreID") =  intCostCentreID '  Request.Form("CostCentreID")
                Else
                    strMessageIcon = "<img src=""../images/warning.gif"" />"
                    strMessage = "<B><FONT Color=""Red"">Cost Centre ALREADY EXISTS " & strMessage & ".</FONT></B>"
                End If
            Else

            End If
        'Else
        '    strMessage = "<B><FONT Color=""Red"">VOTE IS INVALID.</FONT></B>"
        '    strMessageIcon = "<img src=""../images/warning.gif"" />"
        'End If

    'objRS1.Close
					
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
