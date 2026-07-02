<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
If IsEmpty(Session("ColumnID")) Then Session("ColumnID") = 0
If IsEmpty(Session("InputSheetID")) Then Session("InputSheetID") = 0
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
Dim strFormulaName
Dim intSortOrder
Dim strMessageColour

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
    
    If Not IsEmpty(Request.QueryString("FormulaName")) Then	
     	strFormulaName = Request.QueryString("FormulaName")
    Else strFormulaName = "0"
	End If		

   	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Execute copy	
	If Request.QueryString("Action") = "Copy" Then
		CopyFormulas()
	End If

    If Request.QueryString("Action") = "Delete" Then
		DeleteRecord strFormulaName, Session("Level1ID"), Session("InputSheetID")
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

    if ((frm.Level1ID.value == 0)) {
        varAlert += "Please select an Account Class. \n \n";
        document.getElementById('Level1ID').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('Level1ID').style.backgroundColor = "ffffff"; 
   
    if(frm.FormulaName.value == "0")
	{
       varAlert += "Formula Name Cannot Be Blank. \n \n";
       document.getElementById('FormulaName').style.backgroundColor="ff8080";
       varSubmit = false;
    }
    else document.getElementById('FormulaName').style.backgroundColor = "ffffff";  
		
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
	self.location="InputFormulas.asp?Level1ID=" + frm.Level1ID.value
}

function ChangeIS() {
    self.location = "InputFormulas.asp?InputSheetID=" + frm.InputSheetID.value
}

function DeleteData(GEXPID) {
    
    if (window.confirm('Would you like to DELETE the selected record?') == true) {

        self.location = "InputFormulas.asp?Action=Delete&FormulaName=" + GEXPID;
    }

}

function CopyFormulas()
{
	
    if(window.confirm('Would you like to apply Default Formulas?')==true){
	self.location="InputFormulas.asp?Action=Copy"
	}
     
}

//-->
</script>
</head>
<body>
<h3>Input Sheet Formula Assignment Screen</h3>
<form action="InputFormulas.asp?Action=Save" method="POST" id="frm" name="frm">

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
        <th style="text-align:left; height:20px; width:20%;">&nbsp;Formulas</th>
		<td style="text-align:left; height:20px; width:30%;">
		    <select Style="Width:100%" tabindex="20" id="FormulaName" name="FormulaName" ><OPTION Value="0">Please select...</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblFormulas WHERE BudgetID = " & Session("BudgetID") & " AND FormulaTypeID = 2",objCon
    		
		    Do until objRS.EOF
			    If cstr(objRS("FormulaName")) = cstr(strFormulaName) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("FormulaName") & """" & strSelected & ">" & objRS("FormulaName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>	<th style="text-align:left; height:20px; width:20%;">&nbsp;Sort Order</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left" style="width:50%" id="SortOrder" name="SortOrder" TABINDEX="4" value="<%=intSortOrder%>"></td>
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
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="CopyFormulas()";><img src="../images/table_add.png" alt="" /> Apply Default</button></td>  
		<td class='locked' Width="100px"><button type="button" tabindex="19" onclick="DeleteData('<%=strFormulaName%>')";><img src="../images/cross.png" alt="" /> Delete </button></td>
        <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr>
</form>
<H3>List of Formulas Assigned</FONT></H3>
<table WIDTH="50%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
        <th style="text-align:center; height:20px" width="10%">Delete</th>	
		<th style="text-align:center; height:20px" width="80%">Formula Name</th>
        <th style="text-align:center; height:20px" width="10%">Sort Order</th>
	</tr>
	
<%
        objRS.Open "SELECT * FROM tblInputFormulas WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Level1ID = " & Session("Level1ID") & " AND ReportID = " & Session("InputSheetID") & "",objCon
            
            'Response.Write "SELECT * FROM tblColumnNames WHERE BudgetID = " & Session("BudgetID") & " AND Level1ID = " & Session("Level1ID") & " AND InputSheetID = " & Session("InputSheetID") & ""
		    Do until objRS.eof
			    Response.Write "<TR><TD><img src=""../images/cross.png"" alt=""Delete GL Code"" onclick=""DeleteData(" & objRS("FormulaName") & ");""/></a></TD><TD Style=""Text-Align:Center;""><A Target=""_self"" HREF=""InputFormulas.asp?InputSheetID=" & objRS("ReportID") & "&FormulaName=" & objRS("FormulaName") & "&Level1ID=" & objRS("Level1ID") & """><B>&nbsp;" & objRS("FormulaName") & "</B></A></TD><TD>" & objRS("SortOrder") & "</TD></TR>"
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
       
		objRS.Open "SELECT * FROM tblInputFormulas WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Level1ID = " & Session("Level1ID") & " AND ReportID = " & Session("InputSheetID") & " AND FormulaName = '" & strFormulaName & "'",objCon							
		
            If Not objRS.EOF Then

                 strFormulaName = objRS("FormulaName")
                 intSortOrder = objRS("SortOrder")

            Else

                intSortOrder = 2
       		  									
		    End if

		objRS.Close	

End Sub

Sub SaveDetails()	

Dim Vote
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spInputFormulaSave"
                
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput) 
                .Parameters.Append objCmd.CreateParameter("ReportID", adInteger, adParamInput) 
                .Parameters.Append objCmd.CreateParameter("Level1ID", adInteger, adParamInput) 
                .Parameters.Append objCmd.CreateParameter("FormulaName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("SortOrder", adInteger, adParamInput)                
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)  
              
				.Parameters("BudgetID") = Session("BudgetID")
                .Parameters("VersionID") = Session("VersionID")
                .Parameters("ReportID") = Request.Form("InputSheetID")	
                .Parameters("Level1ID") = Request.Form("Level1ID")
                .Parameters("FormulaName") = Request.Form("FormulaName")	
			    .Parameters("SortOrder") = Request.Form("SortOrder")
                .Parameters("UpdatedBy") = Session("UserID")
                                        
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
            strMessage = "<B>RECORD SAVED.</B>"								
     		
					
End Sub	

Sub DeleteRecord(FormulaName,Level1ID,InputSheetID)
    
  
        objCon.Execute "DELETE tblInputFormulas WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND FormulaName = '" & FormulaName & "' AND Level1ID = " & Level1ID & " AND ReportID = " & InputSheetID & ""   
       
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

Public Sub CopyFormulas() 

	objCon.Execute "spInputFormulasCopy " & Session("BudgetID") & "," & Session("InputSheetID") & "," & Session("Level1ID") & "," & Session("UserID") & ""
	
	strMessage = "<B>DEFAULT FORMULAS HAVE BEEN APPLIED.</V>" 
	strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
    strMessageColour = "Black" 

End Sub


Set objRS = Nothing
Set objCon = Nothing


%>
