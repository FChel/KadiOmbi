<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	Fund Administration Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			March 2008

If IsEmpty(Session("PerformanceIndicatorID")) Then Session("PerformanceIndicatorID") = 0

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim arrHeadings(5)
Dim intFinYearPart1
Dim intFinYearPart2
'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngStrategicObjectiveID
Dim lngPerformanceIndicatorID
Dim strStrategicObjective

Dim strIndicatorName
Dim strIndicatorDesc
Dim strDate
Dim strBaseIndicator
Dim dblYear1
Dim dblYear2
Dim dblYear3
Dim dblYear4
Dim dblYear5
Dim strMDG
Dim strM
Dim strP
Dim strR
Dim strIndicatorSource


'Declare and set default arrays

Dim arrActive(2)
	
	arrActive(1) = "Y"
	arrActive(2) = "N"
		
	'3. Capture Querystring variables	
	If Not IsEmpty(Request.QueryString("StrategicObjectiveID")) Then		
		Session("StrategicObjectiveID") = Request.QueryString("StrategicObjectiveID")
		lngStrategicObjectiveID = Request.QueryString("StrategicObjectiveID")					
	End If
    
    If Not IsEmpty(Request.QueryString("PerformanceIndicatorID")) Then		
		Session("PerformanceIndicatorID") = Request.QueryString("PerformanceIndicatorID")
		lngPerformanceIndicatorID = Request.QueryString("PerformanceIndicatorID")
        				
	End If			
	
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Load page details
	LoadDetails()

    'Set Headings
For x = 0 to 4
	
	intFinYearPart1 = cint(Session("FinancialYear")) + (x - 2)
	intFinYearPart1 = Right(intFinYearPart1,2)
	intFinYearPart2 = cint(Session("FinancialYear")) + x - 1
	intFinYearPart2 = Right(intFinYearPart2,2)

	arrHeadings(x) = cstr(intFinYearPart1) & "/" & cstr(intFinYearPart2)

Next
		
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
    var varAlert="";	

	if(isWhitespace(frm.StrategicObjectiveID.value) || frm.StrategicObjectiveID.value=="0")
	{
       varAlert += "Fund ID Cannot Be Blank. \n \n";
       document.getElementById('StrategicObjectiveID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('StrategicObjectiveID').style.backgroundColor="ffffff";
    
    
    if(isWhitespace(frm.FundName.value))
	{
       varAlert += "Fund Name Cannot Be Blank. \n \n";
       document.getElementById('FundName').style.backgroundColor="ff8080";
       varSubmit = false;
    }
    else document.getElementById('FundName').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.FundNameL2.value)) {
        varAlert += "Fund Name L2 Cannot Be Blank. \n \n";
        document.getElementById('FundNameL2').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('FundNameL2').style.backgroundColor = "ffffff";

    if(isWhitespace(frm.Parent.value))
	{
       varAlert += "Parent Cannot Be Blank. \n \n";
       document.getElementById('Parent').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('Parent').style.backgroundColor="ffffff";	
					
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}


function StrategicObjectiveIDSearch()
{	
	self.location="StrategicObjectives.asp?StrategicObjectiveID=" + frm.StrategicObjectiveID.value
}
//-->
</script>
</head>
<body>
    <h3>Objectives Screen</h3>
<form action="StrategicObjectives.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
	    <th style="width:20%;height:20px;">Strategic Objectives</th>
	    <th style="width:30%"></th>
	    <th style="width:20%"></th>
	    <th style="width:30%"></th>
	</tr>
	<tr>
	    <td colspan="4">&nbsp;</td>
	</tr>
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Objective ID</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left" style="width:50%" id="StrategicObjectiveID" name="StrategicObjectiveID" maxlength="50" TABINDEX="1" onblur="StrategicObjectiveIDSearch()" value="<%=lngStrategicObjectiveID%>"></td>		
    <td colspan="2"></td>
	</tr>
	<tr>
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Objective</th>
		<td style="text-align:left; height:20px; width:30%;" colspan="3">&nbsp;<input style="text-align:left; width:98%" id="StrategicObjective" name="StrategicObjective" maxlength="150" TABINDEX="2" value="<%=strStrategicObjective%>"></td>
	 
	</tr>		
	
    <tr><th colspan="4"  style="height:20px;">&nbsp;</th></tr>
</table>
<br>
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="parent.location='../SetUpFrameset.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
		<td class='locked' Width="100px"><button type="button" tabindex="19" onclick="DeleteData()";><img src="../images/cross.png" alt="" /> Delete </button></td>
        <td Width="300px"><font Color="Gray"><b><input style="font-weight:bold; color:red; width:80%" type="text" id="msgbox" name="msgbox" value="<%=strMessage%>"></b></font></td>
	</tr>
</table>
<hr>
</form>
<h3>Results Framework</h3>
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
    <tr>
        <th colspan="2" Height="20px"></th><th colspan="2">Baseline</th><th colspan="5">Indicator Target Values (AS Per SP)</th><th colspan="4">Classifications</th><th></th>
    </tr>
	<tr>
		<th align="center" style="height:20px;width:10%;">Indicator Name</th>	
	    <th align="center" width="20%">Indicator Desc</th>
	    <th align="center" width="7.5%">Date</th>		
	 	<th align="center" width="7.5%">Indicator</th>
        <th align="center" width="7.5%"><%=arrHeadings(0) %></th>
        <th align="center" width="7.5%"><%=arrHeadings(1) %></th>
        <th align="center" width="7.5%"><%=arrHeadings(2) %></th>
        <th align="center" width="7.5%"><%=arrHeadings(3) %></th>
        <th align="center" width="7.5%"><%=arrHeadings(4) %></th>
		<th align="center" width="5%">MDG</th>
        <th align="center" width="5%">M</th>
        <th align="center" width="5%">P</th>
        <th align="center" width="5%">R</th>
        <th align="center" width="5%">Source</th>
	</tr>
    <tr>
        <td style="background-color:FFFFFF;"><input style="text-align:left; width:100%; background-color:FFFFFF;" id="IndicatorName" name="IndicatorName" maxlength="50" TABINDEX="3" value="<%=strIndicatorName%>" /></td>
        <td><input style="text-align:left; width:100%;" id="IndicatorDesc" name="IndicatorDesc" maxlength="50" TABINDEX="3" value="<%=strIndicatorDesc%>" /></td>
        <td><input style="text-align:center; width:100%;" id="Date" name="Date" maxlength="50" TABINDEX="3" value="<%=strDate%>" /></td>
        <td><input style="text-align:center; width:100%;" id="BaseIndicator" name="BaseIndicator" maxlength="50" TABINDEX="3" value="<%=strBaseIndicator%>" /></td>
        <td><input style="text-align:right; width:100%;" id="Year1" name="Year1" maxlength="50" TABINDEX="3" value="<%=dblYear1%>" /></td>
        <td><input style="text-align:right; width:100%;" id="Year2" name="Year2" maxlength="50" TABINDEX="3" value="<%=dblYear2%>" /></td>
        <td><input style="text-align:right; width:100%;" id="Year3" name="Year3" maxlength="50" TABINDEX="3" value="<%=dblYear3%>" /></td>
        <td><input style="text-align:right; width:100%;" id="Year4" name="Year4" maxlength="50" TABINDEX="3" value="<%=dblYear4%>" /></td>
        <td><input style="text-align:right; width:100%;" id="Year5" name="Year5" maxlength="50" TABINDEX="3" value="<%=dblYear5%>" /></td>
        <td><input style="text-align:center; width:100%;" id="MDG" name="MDG" maxlength="50" TABINDEX="3" value="<%=strMDG%>" /></td>
        <td><input style="text-align:center; width:100%;" id="M" name="M" maxlength="50" TABINDEX="3" value="<%=strM%>" /></td>
        <td><input style="text-align:center; width:100%;" id="P" name="P" maxlength="50" TABINDEX="3" value="<%=strP%>" /></td>
        <td><input style="text-align:center; width:100%;" id="R" name="R" maxlength="50" TABINDEX="3" value="<%=strR%>" /></td>
        <td><input style="text-align:center; width:100%;" id="IndicatorSource" name="IndicatorSource" maxlength="50" TABINDEX="3" value="<%=strIndicatorSource%>" /></td>

    </tr>
    <tr><td colspan="14">&nbsp;</td></tr>
	
<%
    objRS.Open "SELECT * FROM tblPerformanceIndicators WHERE BudgetID = " & Session("BudgetID") & " AND StrategicObjectiveID = " & Session("StrategicObjectiveID") & " Order By PerformanceIndicatorID ASC",objCon
		Do until objRS.eof
			Response.Write "<TR><TD style=""background-color:e9e9e9;height:20px""><A Target=""_self"" HREF=""StrategicObjectives.asp?PerformanceIndicatorID=" & objRS("PerformanceIndicatorID") & """><B>&nbsp;" & objRS("PerformanceIndicatorName") & "</B></A></TD><TD style=""background-color:e9e9e9;"">&nbsp;" & objRS("PerformanceIndicatorDesc") & "</B></TD><TD style=""text-align:center;background-color:e9e9e9;"">" & objRS("BaselineDate") & "</TD><TD style=""text-align:center;background-color:e9e9e9;"">" & objRS("BaseIndicator") & "</TD><TD style=""text-align:right;background-color:e9e9e9;"">" & objRS("Year1") & "</TD><TD style=""text-align:right;background-color:e9e9e9;"">" & objRS("Year2") & "</TD><TD style=""text-align:right;background-color:e9e9e9;"">" & objRS("Year3") & "</TD><TD style=""text-align:right;background-color:e9e9e9;"">" & objRS("Year4") & "</TD><TD style=""text-align:right;background-color:e9e9e9;"">" & objRS("Year5") & "</TD><TD style=""text-align:center;background-color:e9e9e9;"">" & objRS("MDG") & "</TD><TD style=""text-align:center;background-color:e9e9e9;"">" & objRS("M") & "</TD><TD style=""text-align:center;background-color:e9e9e9;"">" & objRS("P") & "</TD><TD style=""text-align:center;background-color:e9e9e9;"">" & objRS("R") & "</TD><TD style=""text-align:center;background-color:e9e9e9;"">" & objRS("IndicatorSource") & "</TD></TR>"
			objRS.movenext
		Loop
			
	objRS.Close

%>
</table>
    <br />
    
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="parent.location='../SetUpFrameset.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
		<td class='locked' Width="100px"><button type="button" tabindex="19" onclick="DeleteData()";><img src="../images/cross.png" alt="" /> Delete </button></td>
        <td Width="300px"><font Color="Gray"><b><input style="font-weight:bold; color:red; width:80%" type="text" id="msgbox" name="msgbox" value="<%=strMessage%>"></b></font></td>
	</tr>
</table>
    <hr>
</body>

</html>

<% 

Sub LoadDetails()

'Description:	Loads Caller's details into page if applicable.
		
		objRS.Open "SELECT * FROM tblStrategicObjectives WHERE BudgetID = " & Session("BudgetID") & " AND StrategicObjectiveID = " & Session("StrategicObjectiveID") & "",objCon							
		
        If Not objRS.EOF Then
		    
            lngStrategicObjectiveID = objRS("StrategicObjectiveCode")
            strStrategicObjective = objRS("StrategicObjective")
         									
		End if

		objRS.Close	

        objRS.Open "SELECT * FROM tblPerformanceIndicators WHERE BudgetID = " & Session("BudgetID") & " AND PerformanceIndicatorID = " & Session("PerformanceIndicatorID") & "",objCon							
		'Response.Write "SELECT * FROM tblPerformanceIndicators WHERE BudgetID = " & Session("BudgetID") & " AND PerformanceIndicatorID = " & Session("PerformanceIndicatorID") & ""
        If Not objRS.EOF Then
		    
            strIndicatorName = objRS("PerformanceIndicatorName")
            strIndicatorDesc = objRS("PerformanceIndicatorDesc")
            strDate = objRS("BaselineDate")
            strBaseIndicator = objRS("BaseIndicator")
            dblYear1 = objRS("Year1")
            dblYear2 = objRS("Year2")
            dblYear3 = objRS("Year3")
            dblYear4 = objRS("Year4")
            dblYear5 = objRS("Year5")
            strMDG = objRS("MDG")
            strM = objRS("M")
            strP = objRS("P")
            strR = objRS("R")
            strIndicatorSource = objRS("IndicatorSource")
         									
		End if

		objRS.Close
    	
End Sub

Sub SaveDetails()	
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spFundSave"
                  
                .Parameters.Append objCmd.CreateParameter("StrategicObjectiveID", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("FundName", adVarChar, adParamInput, 150)
                .Parameters.Append objCmd.CreateParameter("FundDesc", adLongVarChar, adParamInput, -1)
                .Parameters.Append objCmd.CreateParameter("FundNameL2", adVarChar, adParamInput, 150)
                .Parameters.Append objCmd.CreateParameter("FundDescL2", adLongVarChar, adParamInput, -1)
                .Parameters.Append objCmd.CreateParameter("FundNotes", adLongVarChar, adParamInput, -1)
                .Parameters.Append objCmd.CreateParameter("Parent", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("Active", adVarChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                                                                    
                
                .Parameters("StrategicObjectiveID") = Request.Form("StrategicObjectiveID")
				.Parameters("BudgetID") = Session("BudgetID")	
			    .Parameters("FundName") = Request.Form("FundName")
                .Parameters("FundDesc") = Request.Form("FundDesc")
                .Parameters("FundNameL2") = Request.Form("FundNameL2")
                .Parameters("FundDescL2") = Request.Form("FundDescL2")
                .Parameters("FundNotes") = ""
                .Parameters("Parent") = Request.Form("Parent")     
                .Parameters("Active") = Request.Form("Active")               
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessage = "Fund record saved !"									
     		Session("StrategicObjectiveID") =  Request.Form("StrategicObjectiveID")
					
End Sub	

Sub DeleteRecord(FundSourceID,Status)
    
    If Status = "I" Then
        objCon.Execute "DELETE FROM tblFundSources WHERE BudgetID = " & Session("BudgetID") & " AND FundSourceID = '" & FundSourceID & "'"   
        strMessage = "Record deleted."
        strFundSourceID = ""
    Else
        strMessage = "ERP sourced records cannot be deleted!"
    End If
End Sub

Sub LoadERPData()
    
    objCon.Execute "spLoadERPFundSources " & Session("BudgetID") & ",'N'," & Session("UserID") & ""
     
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
