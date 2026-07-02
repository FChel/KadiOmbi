<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")

If IsEmpty(Session("WarrantReleassID")) Then Session("WarrantReleaseID") = 0
If IsEmpty(Session("Month")) Then Session("Month") = ""
If IsEmpty(Session("Ceiling")) Then Session("Ceiling") = ""
 
'Description:	Projects Administration Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			March 2008

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon
Dim strVote

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngWarrantReleaseID
Dim strWarrantReleaseRef
Dim lngFundsReleased
Dim dteReleaseDate
Dim strStatus

'Declare and set default arrays

Dim arrStatus(2)
	
	arrStatus(1) = "Y"
	arrStatus(2) = "N"
		
	'3. Capture Querystring variables	
	If Not IsEmpty(Request.QueryString("WarrantReleaseID")) Then		
		Session("WarrantReleaseID") = Request.QueryString("WarrantReleaseID")
		lngWarrantReleaseID = Request.QueryString("WarrantReleaseID")					
	End If	
    
    '3. Capture Querystring variables	
	If Not IsEmpty(Request.QueryString("Month")) Then		
		Session("Month") = Request.QueryString("Month")
	End If	

	If Not IsEmpty(Request.QueryString("CeilingClass")) Then		
		Session("CeilingClass") = Request.QueryString("CeilingClass")
	End If
    
    objRS.Open "SELECT BusinessAreaCode FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon

        If Not objRS.EOF Then
            strVote = objRS("BusinessAreaCode")
        End If

    objRS.Close
	
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

	var str1 = frm.FundsReleased.value
	var amount1 = str1.replace(/,/g, "");
    
    if(isWhitespace(frm.WarrantReleaseRef.value))
	{
       varAlert += "Reference Cannot Be Blank. \n \n";
       document.getElementById('WarrantReleaseRef').style.backgroundColor="ff8080";
       varSubmit = false;
    }
    else document.getElementById('WarrantReleaseRef').style.backgroundColor = "ffffff";  

	var str1 = frm.FundsReleased.value
	var amount1 = str1.replace(/,/g, "");
	
	if(isSignedInteger(amount1)==false)
    {
	   varAlert += "Funds Released cannot be blank. \n \n";
	   document.getElementById('FundsReleased').style.backgroundColor = "ff8080";

	   varSubmit = false;
	}
	else document.getElementById('FundsReleased').style.backgroundColor = "ffffff";

	if (frm.CeilingClass.value == 0) {
		    varAlert += "Please select Ceiling Class. \n \n";
		    document.getElementById('CeilingClass').style.backgroundColor = "ff8080";
		    varSubmit = false;
		}
		else document.getElementById('CeilingClass').style.backgroundColor = "ffffff";
		
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}

function FormatNum() {
    var mynum = frm.FundsReleased.value;
    mynum = mynum.replace(/,/g, "");
    mynum = FormatNumber(mynum,0,0,0,1);
    document.getElementById('FundsReleased').value = mynum;
    
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

//-->
</script>
</head>
<body>
<h3>Warrant Release Administration Screen</h3>
<form action="WarrantReleases.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=50% ALIGN=Left BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
		<th style="text-align:left; height:20px; width:30%;">&nbsp;Month</th>
		<td style="text-align:left; height:20px; width:70%;">
		    <select Style="Width:100%" tabindex="20" id="MonthID" name="MonthID" onchange="self.location='WarrantReleases.asp?Month=' + frm.MonthID.value;"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblMonths Order By SortOrder",objCon
    		
		    Do until objRS.EOF
			    If cstr(objRS("MonthName")) = cstr(Session("Month")) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("MonthName") & """" & strSelected & ">" & objRS("MonthName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
    
	</tr>
	 <tr>
	<th style="text-align:left; height:20px; width:20%;">&nbsp;Ceiling Class</th><td style="text-align:left; height:20px; width:100%;">
	 <select Style="Width:100%" tabindex="5" id="CeilingClass" name="CeilingClass" onchange="self.location='WarrantReleases.asp?CeilingClass=' + frm.CeilingClass.value;">
        <option value="0">Please Select...</option>
	    <%
		objRS.Open "SELECT SUM(Level2Ceiling),CeilingLevelID,CeilingLevelName FROM qryBACeilingLevel2 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " Group By CeilingLevelID,CeilingLevelName",objCon,0,1
		
		Do until objRS.EOF
			If objRS("CeilingLevelID") = cstr(Session("CeilingClass")) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("CeilingLevelID") & """" & strSelected & ">" & objRS("CeilingLevelID") & " : " & objRS("CeilingLevelName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
	       </select>
	    </td>
	</tr>
    <tr>
	    <td style="text-align:left; height:20px" colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Warrant Release ID</th>
		<td style="text-align:left; height:20px; width:30%; background-color:FFFFFF;">&nbsp;<input style="text-align:left; width:98%; background-color:FFFFFF;" READONLY id="WarrantReleaseID" name="WarrantReleaseID" TABINDEX="1" value="<%=lngWarrantReleaseID%>"></td>		

	</tr>
	<tr>
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Warrant Release Reference</th>
		<td style="text-align:left; height:20px; width:30%; background-color:FFFFFF;">&nbsp;<input style="text-align:left; width:98%; background-color:FFFFFF;" id="WarrantReleaseRef" name="WarrantReleaseRef" maxlength="50" TABINDEX="2" value="<%=strWarrantReleaseRef%>"></td>
	</tr>
	<tr>
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Funds Released</th>
		<td style="text-align:left; height:20px; width:30%; background-color:FFFFFF;">&nbsp;<input style="text-align:left; width:98%; background-color:FFFFFF;" id="FundsReleased" name="FundsReleased" maxlength="50" TABINDEX="2" value="<%=formatnumber(lngFundsReleased,0)%>" onchange="FormatNum()"></td>
	</tr>
	<tr>
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Release Date</th>
		<td style="background-color:FFFFFF;"><input style="text-align:left; width:55%; background-color:FFFFFF;" readonly id="ReleaseDate" name="ReleaseDate" maxlength="50" TABINDEX="6" value="<%=dteReleaseDate%>">&nbsp; &nbsp;<a href="javascript:ReleaseDate.popup();"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Click here to select the date"></a>
		&nbsp;<a href="javascript:clearField('ReleaseDate');"><img src="../Images/rubber.gif" border="0" alt="Click here to clear the date field"></a></td>
       
   </tr>
	<tr><th style="text-align:left; height:20px; width:20%; ">&nbsp;Status</th><td style="text-align:left; height:20px; width:30%; background-color:FFFFFF;">
	 <select Style="Width:40%; background-color:FFFFFF;" tabindex="5" id="Status" name="Status">
	        <%
		        For x = 1 to 2
			        If arrStatus(x) = cstr(strStatus) Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrStatus(x) & """" & strSelected & ">" & arrStatus(x) & "</OPTION>"
		        Next
	        %>
	       </select>
	
	</tr>		
    <tr>
	    <td style="text-align:left; height:20px" colspan="2">&nbsp;</td>
	</tr>
</table>
<script LANGUAGE="javascript">
var ReleaseDate;
ReleaseDate = new calendar(document.forms(0).elements['ReleaseDate']);

</script>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>
<br>

<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='WarrantReleases.asp?WarrantReleaseID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr>
</form>
<H3>List of Warrant Releases</H3>
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
		<th style="text-align:center; height:20px" width="10%">Release ID</th>	
	    <th style="text-align:center; height:20px" width="25%">Release Reference</th>
	    <th style="text-align:center; height:20px" width="25%">Release Date</th>
		<th style="text-align:center; height:20px" width="10%">Funds Released</th>			
	 	<th style="text-align:center; height:20px" width="10%">Status</th>
	</tr>
	
<%
        objRS.Open "SELECT * FROM tblWarrantReleases WHERE BudgetID = " & Session("BudgetID") & " AND Month = '" & Session("Month") & "' Order By ReleaseDate ASC",objCon
 
		Do until objRS.eof
			Response.Write "<TR><TD><A Target=""_self"" HREF=""WarrantReleases.asp?CeilingClass=" & objRS("Ceiling") & "&WarrantReleaseID=" & objRS("WarrantReleaseID") & """><B>&nbsp;" & objRS("WarrantReleaseID") & "</B></A></TD><TD>&nbsp;" & objRS("WarrantReleaseReference") & "</B></TD><TD>&nbsp;" & objRS("ReleaseDate") & "</B></TD><TD style=""text-align:center"">" & formatnumber(objRS("FundsReleased"),0) & "</TD><TD style=""text-align:center"">" & objRS("Status") & "</TD></TR>"
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
	
	objRS.Open "SELECT * FROM tblWarrantReleases WHERE BudgetID = " & Session("BudgetID") & " AND WarrantReleaseID = " & Session("WarrantReleaseID") & "",objCon							
		
		If Not objRS.EOF Then
		    lngWarrantReleaseID = objRS("WarrantReleaseID")
            strWarrantReleaseRef = objRS("WarrantReleaseReference")
			dteReleaseDate = objRS("ReleaseDate")
            strStatus = objRS("Status")	
			lngFundsReleased = objRS("FundsReleased")	
		Else
			lngWarrantReleaseID=""
			strWarrantReleaseRef = ""
			dteReleaseDate = ""
			lngFundsReleased = 0								
		End if

		objRS.Close	
End Sub

Sub SaveDetails()	

Dim Vote

	
		 With objCmd
                .CommandType = 4
                .CommandText = "spWarrantReleaseSave"
                  
                .Parameters.Append objCmd.CreateParameter("WarrantReleaseID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("WarrantReleaseRef", adVarChar, adParamInput, 50)
				.Parameters.Append objCmd.CreateParameter("ReleaseDate", adDate, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Ceiling", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("FundsReleased", adDouble, adParamInput)
				.Parameters.Append objCmd.CreateParameter("Month", adVarChar, adParamInput, 3)
             	.Parameters.Append objCmd.CreateParameter("Comments", adLongVarChar, adParamInput, -1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
				.Parameters.Append objCmd.CreateParameter("WarrantReleaseIDOutput", adInteger, adParamOutput)                                                            
                
				If IsNull(Request.Form("WarrantReleaseID")) or Request.Form("WarrantReleaseID") = "" Then
                	.Parameters("WarrantReleaseID") = 0
				Else
					.Parameters("WarrantReleaseID") = Request.Form("WarrantReleaseID")
				End If
				.Parameters("BudgetID") = Session("BudgetID")
				.Parameters("VersionID") = Session("VersionID")		
			    .Parameters("WarrantReleaseRef") = Request.Form("WarrantReleaseRef")
				.Parameters("ReleaseDate") = Request.Form("ReleaseDate")
                .Parameters("Ceiling") = Request.Form("CeilingClass")
                .Parameters("FundsReleased") = Request.Form("FundsReleased")
                .Parameters("Month") = Request.Form("MonthID")  
				.Parameters("Comments") = ""'Request.Form("Comments")             
                .Parameters("UpdatedBy") = Session("UserID")            
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessage = "<B>RECORD SAVED.</B>"
            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"								
     		Session("WarrantReleaseID") =  objCmd.Parameters.Item("WarrantReleaseIDOutput") 
		
					
End Sub	

Sub DeleteRecord(ProjectID,Status)
    
    If Status = "I" Then
        objCon.Execute "DELETE FROM tblProjects WHERE BudgetID = " & Session("BudgetID") & " AND ProjectID = '" & ProjectID & "'"   
        strMessage = "Record deleted."
        strProjectID = ""
    Else
        strMessage = "ERP sourced records cannot be deleted!"
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
