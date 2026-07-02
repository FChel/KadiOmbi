<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
'Description:	Financial Year Admin Screen
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

Dim intFinancialYearID
Dim strFinancialYearName
Dim intDaysInYear
Dim lngFiscalYear
Dim dteStartDate
Dim dteEndDate
Dim strComments
	
	'3. Capture Querystring variables
	
	If Not IsEmpty(Request.QueryString("FinancialYearID")) Then
		
		Session("FinancialYearID") = Request.QueryString("FinancialYearID")
		intFinancialYearID = clng(Request.QueryString("FinancialYearID"))	
				
	End If	
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Load page details
	LoadDetails()
		
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
	<script src="../formChek.js">
	</script>
	<script src="../ButtonRollOver.js">
	</script>
	<script src="../calender.js">
	</script>
<script LANGUAGE="javascript">
<!--
function SaveData()
{	    
    var varSubmit = true						
    var varAlert="";	    
    
    if(isWhitespace(frm.FinancialYearName.value))
	{
       varAlert += "Financial Year Name Cannot Be Blank. \n \n";
       document.getElementById('FinancialYearName').style.backgroundColor = "ff8080";
       varSubmit = false;
    }   
    else document.getElementById('FinancialYearName').style.backgroundColor="ffffff";
   
    if (isWhitespace(frm.StartDate.value)) {
        varAlert += "Start Date Cannot Be Blank. \n \n";
        document.getElementById('StartDate').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('StartDate').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.EndDate.value)) {
        varAlert += "End Date Cannot Be Blank. \n \n";
        document.getElementById('EndDate').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('EndDate').style.backgroundColor = "ffffff";

    if ((isNonnegativeInteger(frm.DaysInYear.value) == false) || (frm.DaysInYear.value == 0)) {
        varAlert += "Please enter Days in Year. Days in Year must be a numeric value. \n \n";
        document.getElementById('DaysInYear').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('DaysInYear').style.backgroundColor = "ffffff";

    if ((isNonnegativeInteger(frm.FiscalYear.value) == false) || (frm.FiscalYear.value == 0)) {
        varAlert += "Please enter Days in Year. Days in Year must be a numeric value. \n \n";
        document.getElementById('FiscalYear').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('FiscalYear').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.FinancialYearName.value)) {
        varAlert += "Financial Year Name Cannot Be Blank. \n \n";
        document.getElementById('FinancialYearName').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('FinancialYearName').style.backgroundColor = "ffffff";
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}

//-->
</script>
</head>
<body>
<h3>Financial Year Administration</h3>
<form action="FinancialYears.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>

	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Financial Year Name</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left" style="width:90%" id="FinancialYearName" name="FinancialYearName" maxlength="50" TABINDEX="1" value="<%=strFinancialYearName%>"></td><td colspan="2"></td>
		
	</tr>
	
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Start Date</th>
		<td colspan="3">&nbsp;<input style="text-align:left" style="width:10%" readonly id="StartDate" name="StartDate" maxlength="50" TABINDEX="3" value="<%=dteStartDate%>">&nbsp; &nbsp;<a href="javascript:StartDate.popup();"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Click Here to pick up the date"></a>
		&nbsp;<a href="javascript:clearField('StartDate');"><img src="../Images/rubber.gif" border="0" alt="Click here to clear the date field"></a></td>
	</tr>
	
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;End Date</th>
		<td colspan="3">&nbsp;<input style="text-align:left" style="width:10%" readonly id="EndDate" name="EndDate" maxlength="50" TABINDEX="4" value="<%=dteEndDate%>">&nbsp; &nbsp;<a href="javascript:EndDate.popup();"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Click Here to pick up the date"></a>
		&nbsp;<a href="javascript:clearField('EndDate');"><img src="../Images/rubber.gif" border="0" alt="Click here to clear the date field"></a></td>
	</tr>
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Days in Year</th>
		<td>&nbsp;<input style="text-align:left" style="width:90%" id="DaysInYear" name="DaysInYear" maxlength="50" TABINDEX="2" value="<%=intDaysInYear%>"></td><td colspan="2"></td>
		
	</tr>
   
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Fiscal Year</th>
		<td>&nbsp;<input style="text-align:left" style="width:90%" id="FiscalYear" name="FiscalYear" maxlength="50" TABINDEX="2" value="<%=lngFiscalYear%>"></td><td colspan="2"></td>
		
	</tr>
   
	<tr>
		<td colspan="4" Align="left">&nbsp;</td>
	</tr>
	<tr>
		<th colspan="4" style="text-align:left;height:20px">&nbsp;Comments</th>
	</tr>
	<tr>
	    <td colspan="4"><TEXTAREA rows=4 cols=190 id=Comments name=Comments tabindex="5"><%=strComments%>
	
</TEXTAREA></td>
	</tr>
</table>
<br>

<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='FinancialYears.asp?FinancialYearID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
        <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>


<script LANGUAGE="javascript">
var StartDate;
var EndDate;
StartDate = new calendar(document.forms(0).elements['StartDate']);
EndDate = new calendar(document.forms(0).elements['EndDate']);
</script>
<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
    <tr>
        <td colspan="8">&nbsp;</td>
    </tr>
	<tr>
		<th>Financial Year Name</th>		
		<th>Start Date</th>
		<th>End Date</th>
		<th>Fiscal Year</th>
        <th>Days In Year</th>
		<th>Updated By</th>
		<th>Date Updated</th>
	</tr>
	
<%

 objRS.Open "SELECT * FROM tblFinancialYears Order By FinancialYearID ASC",objCon
		Do until objRS.eof
			Response.Write "<TR><TD><A Target=""_self"" HREF=""FinancialYears.asp?FinancialYearID=" & objRS("FinancialYearID") & """>&nbsp;" & objRS("FinancialYearName") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("StartDate") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("EndDate") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("ERPFiscalYear") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("DaysInYear") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("DateUpdated") & "</TD></TR>"
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
		
		objRS.Open "SELECT * FROM tblFinancialYears WHERE FinancialYearID = " & clng(Session("FinancialYearID")) & "",objCon
						
			If Not objRS.EOF Then
				
				strFinancialYearName = objRS("FinancialYearName")
				lngFiscalYear = objRS("ERPFiscalYear")
				intDaysInYear = objRS("DaysInYear")
				dteStartDate = objRS("StartDate")
				dteEndDate = objRS("EndDate")
			    strComments = objRS("Comments")
							
			
			Else
				
				intDaysInYear = ""
				lngFiscalYear = ""
				strFinancialYearName = ""
				dteStartDate = ""
				dteEndDate = ""
			    strComments = ""
							
				
			End If

		objRS.Close
	

End Sub

Sub SaveDetails()

		If Not IsEmpty(Request.Form("FinancialYearName")) Then
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spFinancialYearSave"
                
                .Parameters.Append objCmd.CreateParameter("FinancialYearID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("FinancialYearName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("StartDate", adDate)
                .Parameters.Append objCmd.CreateParameter("EndDate", adDate)
                .Parameters.Append objCmd.CreateParameter("DaysInYear", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("FiscalYear", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Comments", adLongVarChar, adParamInput, -1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)  
                .Parameters.Append objCmd.CreateParameter("ClientID", adInteger)  
              
				.Parameters("FinancialYearID") = Session("FinancialYearID")	
				.Parameters("FinancialYearName") = Request.Form("FinancialYearName")			
                .Parameters("StartDate") = Request.Form("StartDate")
                .Parameters("EndDate") = Request.Form("EndDate") 
                .Parameters("DaysInYear") = Request.Form("DaysInYear")   
                .Parameters("FiscalYear") = Request.Form("FiscalYear")               
                .Parameters("Comments") = Request.Form("Comments")
                .Parameters("UpdatedBy") = Session("UserID")
                .Parameters("ClientID") = 1'Session("ClientID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
			'Return the result of the Save Function.
     		 strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
             strMessage = "<B>RECORD SAVED.</B>"
			
						
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
