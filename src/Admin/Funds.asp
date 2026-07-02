<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
'Description:	Fund Source Administration Screen
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

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngFundID
Dim strFundName
Dim strFundDesc
Dim strFundNameL2
Dim strFundDescL2
Dim strFundNotes
Dim strParent
Dim strActive

'Declare and set default arrays

Dim arrActive(2)
	
	arrActive(1) = "Y"
	arrActive(2) = "N"
		
	'3. Capture Querystring variables	
	If Not IsEmpty(Request.QueryString("FundID")) Then		
		Session("FundID") = Request.QueryString("FundID")
		lngFundID = Request.QueryString("FundID")					
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

<script LANGUAGE="javascript">
<!--
function SaveData()
{	    
    var varSubmit = true						
    var varAlert="";	

	if(isWhitespace(frm.FundID.value) || frm.FundID.value=="0")
	{
       varAlert += "Fund ID Cannot Be Blank. \n \n";
       document.getElementById('FundID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('FundID').style.backgroundColor="ffffff";
    
    
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

		
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}


function FundIDSearch()
{	
	self.location="Funds.asp?FundID=" + frm.FundID.value
}
//-->
</script>
</head>
<body>
<h3>Fund Source Administration Screen</h3>
<form action="Funds.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Fund Source ID</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left" style="width:50%" id="FundID" name="FundID" maxlength="3" TABINDEX="1" onblur="FundIDSearch()" value="<%=lngFundID%>"></td>		
    <td colspan="2"></td>
	</tr>
	<tr>
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Fund Source Name Eng</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; width:98%" id="FundName" name="FundName" maxlength="50" TABINDEX="2" value="<%=strFundName%>"></td>
	
	    <th  style="text-align:left; height:20px; width:20%;">&nbsp;Fund Source Name Swa</th>
		<td  style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; width:98%" id="FundNameL2" name="FundNameL2" maxlength="150" TABINDEX="3" value="<%=strFundNameL2%>"></td>
       
   </tr>
	<tr><th  style="text-align:left; height:20px; width:20%;">&nbsp;Active</th><td style="text-align:left; height:20px; width:30%;">
	 <select Style="Width:40%" tabindex="5" id="Active" name="Active">
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
	    <td style="text-align:left; height:20px" colspan="4">&nbsp;</td>
	</tr>
</table>
<br>
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='Funds.asp?FundID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
		<th style="text-align:center; height:20px" width="20%">Fund Source ID</th>	
	    <th align="center" width="25%">Fund Source Name Eng</th>
	    <th align="center" width="25%">Fund Source Name Swa</th>		
	 		<th align="center" width="10%">Active</th>
	</tr>
	
<%
    objRS.Open "SELECT * FROM tblFunds WHERE BudgetID = " & Session("BudgetID") & " Order By FundID ASC",objCon
		Do until objRS.eof
			Response.Write "<TR><TD><A Target=""_self"" HREF=""Funds.asp?FundID=" & objRS("FundID") & """><B>&nbsp;" & objRS("FundID") & "</B></A></TD><TD>&nbsp;" & objRS("FundName") & "</B></TD><TD>&nbsp;" & objRS("FundNameL2") & "</B></TD><TD style=""text-align:center"">" & objRS("Active") & "</TD></TR>"
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
		
		objRS.Open "SELECT * FROM tblFunds WHERE BudgetID = " & Session("BudgetID") & " AND FundID = '" & Session("FundID") & "'",objCon							
		If Not objRS.EOF Then
		    lngFundID = objRS("FundID")
            strFundName = objRS("FundName")
         
            strFundNameL2 = objRS("FundNameL2")
       
            strFundNotes = objRS("FundNotes")
            strActive = objRS("Active")										
		End if

		objRS.Close	
End Sub

Sub SaveDetails()	
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spFundSave"
                  
                .Parameters.Append objCmd.CreateParameter("FundID", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("FundName", adVarChar, adParamInput, 150)
                .Parameters.Append objCmd.CreateParameter("FundNameL2", adVarChar, adParamInput, 150)
                .Parameters.Append objCmd.CreateParameter("FundNotes", adLongVarChar, adParamInput, -1)
                .Parameters.Append objCmd.CreateParameter("Active", adVarChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                                                                    
                
                .Parameters("FundID") = Request.Form("FundID")
				.Parameters("BudgetID") = Session("BudgetID")	
			    .Parameters("FundName") = Request.Form("FundName")
                .Parameters("FundNameL2") = Request.Form("FundNameL2")
                .Parameters("FundNotes") = ""
                .Parameters("Active") = Request.Form("Active")               
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessageIcon = "<img src=""../images/saveticksmall.jpg"" />"
            strMessage = "<B>RECORD SAVED.</B>"								
     		Session("FundID") =  Request.Form("FundID")
					
End Sub	

Sub DeleteRecord(FundID,Status)
    
    If Status = "I" Then
        objCon.Execute "DELETE FROM tblFunds WHERE BudgetID = " & Session("BudgetID") & " AND FundID = '" & FundID & "'"   
        strMessage = "Record deleted."
        strFundID = ""
    Else
        strMessage = "ERP sourced records cannot be deleted!"
    End If
End Sub

Sub LoadERPData()
    
    objCon.Execute "spLoadERPFunds " & Session("BudgetID") & ",'N'," & Session("UserID") & ""
     
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
