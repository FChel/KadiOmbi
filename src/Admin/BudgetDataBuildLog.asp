<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	GL Codes Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			November 2007

'Declare default variables

Server.ScriptTimeout = 6000
Session("CurrentPage") = "Admin/BudgetDataBuildLog.asp"

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

If Not IsEmpty(Request.QueryString("TransactionType")) Then
	Session("TransactionType") = Request.QueryString("TransactionType")
End If

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngGLCode
Dim strGLCodeName
Dim strGLCodeDesc
Dim strGLCodeType
Dim lngActualMapping
Dim strPrepayment
Dim strBalanceSheet
Dim strActive
Dim arrYesNo(2)
Dim strRebuildType

arrYesNo(1) = "N"
arrYesNo(2) = "Y"
  
    objCon.Execute "spBudgetDataBuildLogInsertAll " & Session("BudgetID") & "," & Session("VersionID") & ",'" & Session("TransactionType") & "'," & Session("UserID") & ""
    
	
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If
	
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
	<script src="../formChek.js">
	</script>
	<script src="../ButtonRollOver.js">
	</script>
<script type="text/javascript" language="javascript">

function SaveData(x){
	
var varSubmit = true
var varAlert =""

		if (isWhitespace(frm.TransactionType.value) || frm.TransactionType.value == "") {
        varAlert += "Please Select Dataset to Rebuild. \n \n";
        document.getElementById('TransactionType').style.backgroundColor = "ff8080";
        varSubmit = false;
    	}
    	else document.getElementById('TransactionType').style.backgroundColor = "ffffff";  
	   	  			
	  if(varSubmit == true){
	    		 document.getElementById('Progress').style.display = "inline";
        document.getElementById('Progress1').style.display = "inline";
		 self.location='BudgetDataBuildLog.asp?Action=Save&RebuildType=' + frm.RebuildType.value + '&BusinessAreaID=' + x
		}
	  
	  else{
		window.alert ("" + varAlert);
		}
}

       
    
</script>

</head>
<form action="BudgetDataBuildLog.asp?Action=Save" method="POST" id="frm" name="frm">
<body>
<h3>Recalculate Input Sheets</h3>
<TABLE WIDTH=50% BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TH Style="Width:20%; Height:25px">Data Set</TH>
		<Td Width=30%><select id="TransactionType" name="TransactionType" Style="Width:90%"  onchange="self.location='BudgetDataBuildLog.asp?TransactionType=' + frm.TransactionType.value"><option value="">Please select...</option>
<%	

	objRS.Open "SELECT * FROM tblTransactionType",objCon
		
	Do until objRS.EOF
		If objRS("TransactionType") = cstr(Session("TransactionType")) Then
			strSelected = " SELECTED "
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("TransactionType") & """" & strSelected & ">" & objRS("TransactionType") & " - " & objRS("TransactionTypeDesc") & "</OPTION>"
		objRS.Movenext
	Loop
	
    objRS.Close
	
%>
</select>
</Td>
</tr>
 <tr>
	<th style="height:25px; width:15%; text-align:center ">&nbsp;Indexation Rebuild</th>
		<td width="15%">
		    <select Style="Width:15%" tabindex="1" id="RebuildType" name="RebuildType" >
		    <%
		        For x = 1 to 2
			        If arrYesNo(x) = strRebuildType Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		        Next
	        %>
	      </select>
	    </td>
     </tr>
</tr>
<TR>
<td Colspan="2">&nbsp;<span id="Progress1" style="display:none"><img src=../Images/progress.gif />  &nbsp;&nbsp;&nbsp; <b><FONT Face="Arial">RECALCULATING INPUT SHEETS...</FONT></b></span></td>
</TR>

</TABLE>
</FORM>
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
    <tr>
        <td colspan="10">&nbsp;</td>
    </tr>
	<tr>
		<th Style="Height:20px; Width:5%;">Build</th>
		<th Width="55%">Data Set</th>
        <th Width="12.5%">Data Set</th>
		<th Width="12.5%">Date Last Run</th>
		<th Width="10%">Run By</th>
	</tr>
	
<% 

    objRS.Open "SELECT * FROM qryBudgetDataBuildLog WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND TransactionType = '" & Session("TransactionType") & "' AND BusinessAreaID = " & Session("BusinessAreaID") & " Order By BusinessAreaID",objCon
	'Response.Write "SELECT * FROM qryBudgetDataBuildLog WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND TransactionType = '" & Session("TransactionType") & "' Order By BusinessAreaID"
    	Do until objRS.eof
			Response.Write "<TR><TD Style=""Height:25px; Text-Align:Center;"">  <img src=""../images/wrench.png"" alt=""Build Data Set"" onclick=""SaveData('" & objRS("BusinessAreaID") & "')""/></TD><TD style=""text-align:left"">&nbsp;<B>" & objRS("BusinessAreaCode") & " - " & objRS("BusinessAreaName") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("TransactionType") & "</B></TD><TD style=""text-align:center"">&nbsp;" & objRS("DateLastRun") & "</B></TD><TD style=""text-align:center"">&nbsp;" & objRS("RunBy") & "</TD></TR>"
			
           
            objRS.movenext
		Loop
			
	objRS.Close

%>
</table>
<br>
<hr>
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
		<td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="parent.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>  
		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	    <td><span id="Progress" style="display:none"><img src=../Images/progress.gif />  &nbsp;&nbsp;&nbsp; <b><FONT Face="Arial">RECALCULATING INPUT SHEETS...</FONT></b></span></td>
	</tr>
</table>


</form>
</body>
</html>

<% 


Sub SaveDetails()
	
Dim strSQL
	'objCon.Execute strSQL
	
	 objRS.Open "SELECT * FROM tblBudgetDataBuildLog WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Request.QueryString("BusinessAreaID") & " AND TransactionType = '" & Session("TransactionType") & "'",objCon
     'Response.Write "SELECT * FROM tblBudgetDataBuildLog WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Request.QueryString("BusinessAreaID") & " AND TransactionType = '" & Session("TransactionType") & "'"
        If Not objRS.EOF Then
	        strSQL = objRS("SQL")
	       	
     
       With objCmd
                .CommandType = 4
                .CommandTimeout = 72000
                .CommandText = strSQL               
               
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)  
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BusinessAreaID", adInteger, adParamInput)  
                .Parameters.Append objCmd.CreateParameter("TransactionType", adVarChar, adParamInput,4)
                .Parameters.Append objCmd.CreateParameter("RebuildType", adVarChar, adParamInput,4)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)                          
               	   		   
    		    .Parameters("BudgetID") = Session("BudgetID")
    		    .Parameters("VersionID") = Session("VersionID")
                .Parameters("BusinessAreaID") = Request.QueryString("BusinessAreaID")
                .Parameters("TransactionType") = Session("TransactionType")
                .Parameters("RebuildType") = Request.QueryString("RebuildType")
                .Parameters("UpdatedBy") = Session("UserID")
    		            					 				 				 				 
				.ActiveConnection = objCon
				
			    objCmd.Execute	
			    		    
           End With  
           
           	strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
            strMessage = "<B>INPUT SHEETS HAVE BEEN RECALCULATED.</B>"

     	End If
     	
     	
     objRS.Close
     
     objCon.Execute "UPDATE tblBudgetDataBuildLog SET DateLastRun = GetDate(), RunBy = " & Session("UserID") & " WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Request.QueryString("BusinessAreaID") & ""
     	
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
