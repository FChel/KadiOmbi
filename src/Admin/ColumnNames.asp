<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
If IsEmpty(Session("ColumnID")) Then Session("ColumnID") = 0
If IsEmpty(Session("InputSheetID")) Then Session("InputSheetID") = 100
If IsEmpty(Session("Level1ID")) Then Session("Level1ID") = 0

 
'Description:	Segments Administration Screen
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
Dim strSegmentNo

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim intColumnID
Dim intInputSheetID
Dim strColumnName
Dim intColumnWidth
Dim arrWidthType(2)
Dim strWidthType
Dim strIsTotal
Dim strVisible
Dim arrFormat(4)
Dim strColour
Dim arrYesNo(2)
Dim strFormat

'Declare and set default arrays
	
	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"

    arrWidthType(1) = "px"
	arrWidthType(2) = "%"

    arrFormat(1) = "int"
    arrFormat(2) = "dec2"
    arrFormat(3) = "dec3"
    arrFormat(4) = "dec4"
		
	'3. Capture Querystring variables	
	If Not IsEmpty(Request.QueryString("Level1ID")) Then		
		Session("Level1ID") = Request.QueryString("Level1ID")
	End If	
    
    If Not IsEmpty(Request.QueryString("InputSheetID")) Then	
     	Session("InputSheetID") = Request.QueryString("InputSheetID")
	End If
    
      If Not IsEmpty(Request.QueryString("ColumnID")) Then	
     	Session("ColumnID") = Request.QueryString("ColumnID")
	End If		

   	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Execute copy	
	If Request.QueryString("Action") = "Copy" Then
		CopyColumns()
	End If

    If Request.QueryString("Action") = "Delete" Then
		DeleteRecord Session("ColumnID"), Session("Level1ID"), Session("InputSheetID")
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
    var varAlert = "";

    if ((isNonnegativeInteger(frm.InputSheetID.value) == false) || (frm.InputSheetID.value == 0)) {
        varAlert += "Please select Input Sheet ID. \n \n";
        document.getElementById('InputSheet').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('InputSheetID').style.backgroundColor = "ffffff";

    if ((isNonnegativeInteger(frm.ColumnID.value) == false) || (frm.ColumnID.value == -1)) {
        varAlert += "Please enter Column ID. Column ID must be a numeric value. \n \n";
        document.getElementById('ColumnID').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('ColumnID').style.backgroundColor = "ffffff";

    if ((isNonnegativeInteger(frm.ColumnWidth.value) == false)) {
        varAlert += "Please enter Column Width. Column Width must be a numeric value. \n \n";
        document.getElementById('ColumnWidth').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('ColumnWidth').style.backgroundColor = "ffffff";		  
    
    if(isWhitespace(frm.ColumnName.value))
	{
       varAlert += "Column Name Cannot Be Blank. \n \n";
       document.getElementById('ColumnName').style.backgroundColor="ff8080";
       varSubmit = false;
    }
    else document.getElementById('ColumnName').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.Colour.value)) {
        varAlert += "Colour Cannot Be Blank. \n \n";
        document.getElementById('Colour').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('Colour').style.backgroundColor = "ffffff";

		
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}


function ChangeAS()
{	
	self.location="ColumnNames.asp?Level1ID=" + frm.Level1ID.value
}

function ChangeIS() {
    self.location = "ColumnNames.asp?InputSheetID=" + frm.InputSheetID.value
}

function DeleteData(GEXPID) {

    if (window.confirm('Would you like to DELETE the selected record?') == true) {

        self.location = "ColumnNames.asp?Action=Delete&ColumnID=" + GEXPID;
    }

}

function CopyColumns()
{
	
    if(window.confirm('Would you like to applt Default Columns?')==true){
	self.location="ColumnNames.asp?Action=Copy"
	}
     
}


//-->
</script>
</head>
<body>
<h3>Input Sheet Configuration Screen</h3>
<form action="ColumnNames.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
        <th style="text-align:left; height:20px; width:20%;">&nbsp;Input Sheet ID</th>
		<td style="text-align:left; height:20px; width:30%;">
		    <select Style="Width:100%" tabindex="20" id="InputSheetID" name="InputSheetID" onchange="ChangeIS()"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblFinancialStatements WHERE BudgetID = " & Session("BudgetID") & "",objCon
    		
		    Do until objRS.EOF
			    If clng(objRS("FinancialStatementID")) = cint(Session("InputSheetID")) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("FinancialStatementID") & """" & strSelected & ">" & objRS("FinancialStatementID") & " - " & objRS("FinancialStatementName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Account Class</th>
		<td style="text-align:left; height:20px; width:30%;">
		    <select Style="Width:100%" tabindex="20" id="Level1ID" name="Level1ID" onchange="ChangeAS()"><OPTION Value=0>DEFAULT</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblReportLayoutLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportID = " & Session("InputSheetID") & "",objCon
    		
		    Do until objRS.EOF
			    If clng(objRS("Level1ID")) = clng(Session("Level1ID")) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("Level1ID") & """" & strSelected & ">" & objRS("Level1ID") & " - " & objRS("Level1Name") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
        
	</tr>
    <tr><td colspan="4">&nbsp;</td></tr>	
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Column ID</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left" style="width:50%" id="ColumnID" name="ColumnID" maxlength="9" TABINDEX="1" onblur="SegmentIDSearch()" value="<%=intColumnID%>"></td>		
        <th style="text-align:left; height:20px; width:20%;">&nbsp;Column Name</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left" style="width:50%" id="ColumnName" name="ColumnName" maxlength="50" TABINDEX="1" onblur="SegmentIDSearch()" value="<%=strColumnName%>"></td>
	</tr>
	<tr>    
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Column Width</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; width:98%" id="ColumnWidth" name="ColumnWidth" maxlength="50" TABINDEX="2" value="<%=intColumnWidth%>"></td>
	
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Colour</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; width:98%" id="Colour" name="Colour" maxlength="150" TABINDEX="3" value="<%=strColour%>"></td>
       
   </tr>
	<tr><th style="text-align:left; height:20px; width:20%;">&nbsp;Width Type</th><td style="text-align:left; height:20px; width:30%;">
	 <select Style="Width:40%" tabindex="5" id="WidthType" name="WidthType">
	        <%
		        For x = 1 to 2
			        If arrWidthType(x) = cstr(strWidthType) Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrWidthType(x) & """" & strSelected & ">" & arrWidthType(x) & "</OPTION>"
		        Next
	        %>
	       </select>
	    </td><td colspan="2"></td>
	</tr>
    <tr><th style="text-align:left; height:20px; width:20%;">&nbsp;Visible</th><td style="text-align:left; height:20px; width:30%;">
	 <select Style="Width:40%" tabindex="5" id="Visible" name="Visible">
	        <%
		        For x = 1 to 2
			        If arrYesNo(x) = cstr(strVisible) Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		        Next
	        %>
	       </select>
	    </td><td colspan="2"></td>
	</tr>
    <tr><th style="text-align:left; height:20px; width:20%;">&nbsp;Is Total</th><td style="text-align:left; height:20px; width:30%;">
	 <select Style="Width:40%" tabindex="5" id="IsTotal" name="IsTotal">
	        <%
		        For x = 1 to 2
			        If arrYesNo(x) = cstr(strIsTotal) Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		        Next
	        %>
	       </select>
	    </td><td colspan="2"></td>
	</tr>
      <tr><th style="text-align:left; height:20px; width:20%;">&nbsp;Format</th><td style="text-align:left; height:20px; width:30%;">
	 <select Style="Width:40%" tabindex="5" id="Format" name="Format">
	        <%
		        For x = 1 to 4
			        If arrFormat(x) = cstr(strFormat) Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrFormat(x) & """" & strSelected & ">" & arrFormat(x) & "</OPTION>"
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
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="CopyColumns()";><img src="../images/table_add.png" alt="" /> Apply Default </button></td>  
		<td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='ColumnNames.asp?ColumnID=-1';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
	    <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr>
</form>
<H3>List of Column Names</FONT></H3>
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
        <th style="text-align:center; height:20px" width="5%">Delete</th>	
		<th style="text-align:center; height:20px" width="12.5%">Column ID</th>	
	    <th style="text-align:center; height:20px" width="25%">Column Name</th>
	    <th style="text-align:center; height:20px" width="12.5%">Column Width</th>		
	 	<th style="text-align:center; height:20px" width="12.5%">Is Total</th>
        <th style="text-align:center; height:20px" width="10%">Visible</th>
        <th style="text-align:center; height:20px" width="12.5%">Format</th>
        <th style="text-align:center; height:20px" width="12.5%">Colour</th>

	</tr>
	
<%
        objRS.Open "SELECT * FROM tblColumnNames WHERE BudgetID = " & Session("BudgetID") & " AND Level1ID = " & Session("Level1ID") & " AND InputSheetID = " & Session("InputSheetID") & "",objCon
        'Response.Write "SELECT * FROM tblColumnNames WHERE BudgetID = " & Session("BudgetID") & " AND Level1ID = " & Session("Level1ID") & " AND InputSheetID = " & Session("InputSheetID") & ""
			Do until objRS.eof
				Response.Write "<TR><TD><img src=""../images/cross.png"" alt=""Delete GL Code"" onclick=""DeleteData(" & objRS("ColumnID") & ");""/></a></TD><TD Style=""Text-Align:Center;""><A Target=""_self"" HREF=""ColumnNames.asp?InputSheetID=" & objRS("InputSheetID") & "&ColumnID=" & objRS("ColumnID") & "&Level1ID=" & objRS("Level1ID") & """><B>&nbsp;" & objRS("ColumnID") & "</B></A></TD><TD>&nbsp;" & objRS("ColumnName") & "</B></TD><TD>&nbsp;" & objRS("ColumnWidth") & "</B></TD><TD style=""text-align:center"">" & objRS("IsTotal") & "</TD><TD style=""text-align:center"">" & objRS("Visible") & "</TD><TD style=""text-align:center"">" & objRS("Format") & "</TD><TD style=""text-align:center"">" & objRS("Colour") & "</TD></TR>"
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
       
		objRS.Open "SELECT * FROM tblColumnNames WHERE BudgetID = " & Session("BudgetID") & " AND Level1ID = " & Session("Level1ID") & " AND InputSheetID = " & Session("InputSheetID") & "",objCon	
        
            If objRS.EOF Then
               
                    If Session("Level1ID") <> 0 Then
                        objCon.Execute "spInsertDefaultColumnNames " & Session("BudgetID") & "," & Session("Level1ID") & "," & Session("InputSheetID") & ""
                    End If

            End If
            
        objRS.Close						


		objRS.Open "SELECT * FROM tblColumnNames WHERE BudgetID = " & Session("BudgetID") & " AND Level1ID = " & Session("Level1ID") & " AND InputSheetID = " & Session("InputSheetID") & " AND ColumnID = " & Session("ColumnID") & "",objCon							
		
            If Not objRS.EOF Then

                 intColumnID = objRS("ColumnID")
                 intInputSheetID = objRS("InputSheetID")
                 strColumnName = objRS("ColumnName")
                 intColumnWidth = objRS("ColumnWidth")
                 strIsTotal = objRS("IsTotal")
                 strVisible = objRS("Visible")
                 strColour = objRS("Colour")
                 strFormat = objRS("Format")
		  									
		    End if

		objRS.Close	
End Sub

Sub SaveDetails()	

Dim Vote

    objRS.Open "SELECT BusinessAreaCode FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon

        If Not objRS.EOF Then
            Vote = objRS("BusinessAreaCode")
        End If

    objRS.Close	


		
		 With objCmd
                .CommandType = 4
                .CommandText = "spColumnNameSave"
                
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Level1ID", adInteger, adParamInput) 
                .Parameters.Append objCmd.CreateParameter("ColumnID", adInteger, adParamInput) 
                .Parameters.Append objCmd.CreateParameter("InputSheetID", adInteger, adParamInput) 
                .Parameters.Append objCmd.CreateParameter("ColumnName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("ColumnWidth", adInteger, adParamInput) 
                .Parameters.Append objCmd.CreateParameter("WidthType", adVarChar, adParamInput, 50)                  
                .Parameters.Append objCmd.CreateParameter("IsTotal", adVarChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("Visible", adVarChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("Format", adVarChar, adParamInput, 4)
                .Parameters.Append objCmd.CreateParameter("Colour", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)  
              
				.Parameters("BudgetID") = Session("BudgetID")
                .Parameters("Level1ID") = Request.Form("Level1ID")	
                .Parameters("ColumnID") = Request.Form("ColumnID")
                .Parameters("InputSheetID") = Request.Form("InputSheetID")	
			    .Parameters("ColumnName") = Request.Form("ColumnName")
                .Parameters("ColumnWidth") = Request.Form("ColumnWidth")
                .Parameters("WidthType") = Request.Form("WidthType")
                .Parameters("IsTotal") = Request.Form("IsTotal")
                .Parameters("Visible") = Request.Form("Visible")  
                .Parameters("Format") = Request.Form("Format")  
                .Parameters("Colour") = Request.Form("Colour") 
                .Parameters("UpdatedBy") = Session("UserID")
                                        
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
            strMessage = "<B>RECORD SAVED.</B>"								
     		
					
End Sub	

Sub DeleteRecord(ColumnID,Level1ID,InputSheetID)
    
  
        objCon.Execute "DELETE FROM tblColumnNames WHERE BudgetID = " & Session("BudgetID") & " AND ColumnID = " & ColumnID & " AND Level1ID = " & Level1ID & " AND InputSheetID = " & InputSheetID & ""   
        strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
        strMessage = "<B>RECORD DELETED.</B>"
       
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

Public Sub CopyColumns() 

	objCon.Execute "spColumnNamesCopy " & Session("BudgetID") & "," & Session("InputSheetID") & "," & Session("Level1ID") & "," & Session("UserID") & ""	
	strMessage = "Column Names have been copied."  

End Sub

Set objRS = Nothing
Set objCon = Nothing


%>
