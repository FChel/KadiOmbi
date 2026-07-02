<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
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

Dim lngProjectID
Dim strProjectName
Dim strProjectDesc
Dim strProjectNameL2
Dim strProjectDescL2
Dim strProjectNotes
Dim strParent
Dim strActive

'Declare and set default arrays

Dim arrActive(2)
	
	arrActive(1) = "Y"
	arrActive(2) = "N"
		
	'3. Capture Querystring variables	
	If Not IsEmpty(Request.QueryString("ProjectID")) Then		
		Session("ProjectID") = Request.QueryString("ProjectID")
		lngProjectID = Request.QueryString("ProjectID")					
	End If	
    
    '3. Capture Querystring variables	
	If Not IsEmpty(Request.QueryString("BusinessAreaID")) Then		
		Session("BusinessAreaID") = Request.QueryString("BusinessAreaID")
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

<script LANGUAGE="javascript">
<!--
function SaveData()
{	    
    var varSubmit = true						
    var varAlert="";	

	if(isWhitespace(frm.ProjectID.value) || frm.ProjectID.value=="0")
	{
       varAlert += "Project ID Cannot Be Blank. \n \n";
       document.getElementById('ProjectID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('ProjectID').style.backgroundColor="ffffff";
    
    
    if(isWhitespace(frm.ProjectName.value))
	{
       varAlert += "Project Name Cannot Be Blank. \n \n";
       document.getElementById('ProjectName').style.backgroundColor="ff8080";
       varSubmit = false;
    }
    else document.getElementById('ProjectName').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.ProjectNameL2.value)) {
        varAlert += "Project Name L2 Cannot Be Blank. \n \n";
        document.getElementById('ProjectNameL2').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('ProjectNameL2').style.backgroundColor = "ffffff";

		
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}


function ProjectIDSearch()
{	
	self.location="Projects.asp?ProjectID=" + frm.ProjectID.value
}

function BAIDSearch() {
    self.location = "Projects.asp?BusinessAreaID=" + frm.BusinessAreaID.value 
}
//-->
</script>
</head>
<body>
<h3>Project Administration Screen</h3>
<form action="Projects.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Vote</th>
		<td style="text-align:left; height:20px; width:30%;">
		    <select Style="Width:100%" tabindex="20" id="BusinessAreaID" name="BusinessAreaID" onchange="BAIDSearch()"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y' AND BusinessAreaID <> 1000",objCon
    		
		    Do until objRS.EOF
			    If clng(objRS("BusinessAreaID")) = clng(Session("BusinessAreaID")) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("BusinessAreaID") & """" & strSelected & ">" & objRS("BusinessAreaCode") & " - " & objRS("BusinessAreaName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
        <td colspan="2"></td>
	</tr>
    <tr><td colspan="4">&nbsp;</td></tr>	
	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Project ID</th>
		<td style="text-align:left; height:20px; width:30%; background-color:FFFFFF;">&nbsp;<input style="text-align:left; width:98%; background-color:FFFFFF;" id="ProjectID" name="ProjectID" maxlength="9" TABINDEX="1" onblur="ProjectIDSearch()" value="<%=lngProjectID%>"></td>		
    <td colspan="2"></td>
	</tr>
	<tr>
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Project Name Eng</th>
		<td style="text-align:left; height:20px; width:30%; background-color:FFFFFF;">&nbsp;<input style="text-align:left; width:98%; background-color:FFFFFF;" id="ProjectName" name="ProjectName" maxlength="50" TABINDEX="2" value="<%=strProjectName%>"></td>
	
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Project Name Swa</th>
		<td style="text-align:left; height:20px; width:30%; background-color:FFFFFF;">&nbsp;<input style="text-align:left; width:98%; background-color:FFFFFF;" id="ProjectNameL2" name="ProjectNameL2" maxlength="150" TABINDEX="3" value="<%=strProjectNameL2%>"></td>
       
   </tr>
	<tr><th style="text-align:left; height:20px; width:20%; ">&nbsp;Active</th><td style="text-align:left; height:20px; width:30%; background-color:FFFFFF;">
	 <select Style="Width:40%; background-color:FFFFFF;" tabindex="5" id="Active" name="Active">
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
        <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='Projects.asp?ProjectID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr>
</form>
<H3>List of Projects for Vote : <FONT Color="red"><%Response.Write Session("Vote") %></FONT></H3>
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
		<th style="text-align:center; height:20px" width="20%">Project ID</th>	
	    <th style="text-align:center; height:20px" width="25%">Project Name Eng</th>
	    <th style="text-align:center; height:20px" width="25%">Project Name Swa</th>		
	 	<th style="text-align:center; height:20px" width="10%">Active</th>
	</tr>
	
<%
        objRS.Open "SELECT * FROM tblProjects WHERE BudgetID = " & Session("BudgetID") & " AND Vote = '" & strVote & "' Order By ProjectID ASC",objCon
 
		Do until objRS.eof
			Response.Write "<TR><TD><A Target=""_self"" HREF=""Projects.asp?ProjectID=" & objRS("ProjectID") & """><B>&nbsp;" & objRS("ProjectID") & "</B></A></TD><TD>&nbsp;" & objRS("ProjectName") & "</B></TD><TD>&nbsp;" & objRS("ProjectNameL2") & "</B></TD><TD style=""text-align:center"">" & objRS("Active") & "</TD></TR>"
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
		
		objRS.Open "SELECT * FROM tblProjects WHERE BudgetID = " & Session("BudgetID") & " AND ProjectID = '" & Session("ProjectID") & "' AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon							
		If Not objRS.EOF Then
		    lngProjectID = objRS("ProjectID")
            strProjectName = objRS("ProjectName")
         
            strProjectNameL2 = objRS("ProjectNameL2")
       
            strProjectNotes = objRS("ProjectNotes")
            strActive = objRS("Active")										
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
                .CommandText = "spProjectSave"
                  
                .Parameters.Append objCmd.CreateParameter("ProjectID", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("ProjectName", adVarChar, adParamInput, 150)
                .Parameters.Append objCmd.CreateParameter("ProjectNameL2", adVarChar, adParamInput, 150)
                .Parameters.Append objCmd.CreateParameter("ProjectNotes", adLongVarChar, adParamInput, -1)
                .Parameters.Append objCmd.CreateParameter("Active", adVarChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)   
                .Parameters.Append objCmd.CreateParameter("Vote", adVarChar, adParamInput, 3)                                                                 
                
                .Parameters("ProjectID") = Request.Form("ProjectID")
				.Parameters("BudgetID") = Session("BudgetID")	
			    .Parameters("ProjectName") = Request.Form("ProjectName")
                .Parameters("ProjectNameL2") = Request.Form("ProjectNameL2")
                .Parameters("ProjectNotes") = ""
                .Parameters("Active") = Request.Form("Active")               
                .Parameters("UpdatedBy") = Session("UserID")
                .Parameters("Vote") = Vote
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessage = "<B>RECORD SAVED.</B>"
            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"								
     		Session("ProjectID") =  Request.Form("ProjectID")
					
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
