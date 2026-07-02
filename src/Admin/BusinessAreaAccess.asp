<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")

    Session("CurrentPage") = "Admin/CostCentreStatus.asp"
 
'Description:	Student Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			August 2004

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon
Dim strSort
Dim strOrder
Dim strBusinessAreaID
Dim strUserID

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngBusinessAreaID
Dim lngUserID
		
'2. Capture Querystring variables	
If Not IsEmpty(Request.QueryString("BusinessAreaID")) Then		
    Session("BusinessAreaID") = Request.QueryString("BusinessAreaID")
    lngBusinessAreaID = clng(Request.QueryString("BusinessAreaID"))					
End If		

If Not IsEmpty(Request.QueryString("UserID")) Then		
    Session("UserID1") = Request.QueryString("UserID")
    lngUserID = clng(Request.QueryString("UserID"))
    If lngUserID <> 0  Then strUserID = " AND UserID = " & lngUserID
End If		

    If Not IsEmpty(Request.QueryString("Sort")) Then
	   strSort = Request.QueryString("sort")
    Else
	   strSort = "BusinessAreaName"
    End If

    If Not IsEmpty(Request.QueryString("Ordered")) Then
	If Request.QueryString("Ordered") = "asc" Then
		strOrder = "desc"
	Else
	   strOrder = "asc"
	End If
    Else
	   strOrder = "asc"
    End If

'Execute save 	
If Request.QueryString("Action") = "Save" Then
    SaveDetails()
    lngUserID = Session("UserID1") 
    If lngUserID <> 0  Then strUserID = " AND UserID = " & lngUserID
End If

If Request.QueryString("Action") = "Delete" Then
    DeleteRecord()
    lngUserID = Session("UserID1") 
    If lngUserID <> 0  Then strUserID = " AND UserID = " & lngUserID
    strMessage = "<B>ACCESS REMOVED.</B>"
    strMessageIcon = "<img src=""../images/saveticksmall.jpg"" />"
End If

If Request.QueryString("Action") = "SaveAll" Then
    SaveAllRecord(Request.QueryString("UserID"))
    lngUserID = Session("UserID1") 
    If lngUserID <> 0  Then strUserID = " AND UserID = " & lngUserID
    strMessage = "<B>ACCESS TO ALL VOTES APPLIED.</B>"
    strMessageIcon = "<img src=""../images/saveticksmall.jpg"" />"
End If

'Execute save 	
If Request.QueryString("Action") = "Delete" Then
    deleteRecord()
End If

'Load page details
if IsEmpty(Request.QueryString("Show")) then
    LoadDetails()
End if	
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

    if(frm.UserID.value=="0")
	{
       varAlert += "User ID Cannot Be Blank. \n \n";
       document.getElementById('UserID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('UserID').style.backgroundColor="ffffff";
    
	if(frm.BusinessAreaID.value=="0")
	{
       varAlert += "Business Area Cannot Be Blank. \n \n";
       document.getElementById('BusinessAreaID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BusinessAreaID').style.backgroundColor="ffffff";
    
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}


function UserIDSearch()
{	        
	self.location="BusinessAreaAccess.asp?Show=True&BusinessAreaID=" + frm.BusinessAreaID.value + "&UserID=" + frm.UserID.value	
}

function deleteRecord()
{
     if(frm.UserID.value=="0" || frm.BusinessAreaID.value=="0")
	 {       
	    alert("Please select a record to delete");
     }   
     else
     {
        self.location="BusinessAreaAccess.asp?Action=Delete&BusinessAreaID=" + frm.BusinessAreaID.value + "&UserID=" + frm.UserID.value
     }
}
function SaveData2(){
	var varSubmit = true
	if(document.frm.UserID.value==0){
		alert("A User must be selected!");
		varSubmit = false;
	}
	if(varSubmit == true){
	if ( confirm("Would you like to give " + document.frm.UserID.options[document.frm.UserID.selectedIndex].text + " access to ALL VOTES for the selecetd Budget?"))
		self.location="BusinessAreaAccess.asp?Action=SaveAll&UserID=" + document.frm.UserID.value;
		//frm.submit();
	}else{
		//alert("Access NOT Updated!");
	}

}
//-->
</script>
</head>
<body>
<h3>Vote Access Administration Screen
<%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version : <FONT color=Red>" & Session("VersionName") & "</FONT>"
%></h3>

<form action="BusinessAreaAccess.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=50% ALIGN=Left BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
        <th style="text-align:left; height:20px; width:40%;">&nbsp;User</th>
		<td style="text-align:left; height:20px; width:60%;">
		    <select Style="Width:98%;height:20px" tabindex="20" id="UserID" name="UserID" onchange="UserIDSearch()"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblUsers WHERE Active = 'Y' Order By LName",objCon
    		
		    Do until objRS.EOF
			    If objRS("UserID") = lngUserID Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("UserID") & """" & strSelected & ">" & objRS("UserID") & " - " & objRS("FName") & " " & objRS("LName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
		
		
	</tr>
	
	<tr>
		<th style="text-align:left">&nbsp;Vote</th>
		<td>
		    <select Style="Width:98%;height:20px" tabindex="20" id="BusinessAreaID" name="BusinessAreaID"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon
    		
		    Do until objRS.EOF
			    If objRS("BusinessAreaID") = lngBusinessAreaID Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("BusinessAreaID") & """" & strSelected & ">" & objRS("BusinessAreaCode") & " -  " & objRS("BusinessAreaName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
	</tr>
  <tr>
	<td style="height:20px" colspan="4" align="left">&nbsp;</td>
	</tr>		
</table>
<br/>
<br/>
<br/>
<br/>

<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='BusinessAreaAccess.asp?BusinessAreaID=0&UserID=0'"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
        <td class='locked' Width="100px"><button type="button" tabindex="13" onclick="javascript:SaveData2();"><img src="../images/table_save.png" alt="" /> Save All</button></td>
		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="400px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>


<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
<%
	'Dynamically build the menu items depending on the sort selection 
	Response.write "<th>Edit</th><th>REMOVE</th>" & _
	    	"<th><B><A Target=""_self"" HREF=""BusinessAreaAccess.asp?Sort=LName&Ordered=" & strOrder & "&BusinessAreaID=" & lngBusinessAreaID & "&UserID="& lngUserID & """>User"
		If strSort = "LName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	Response.write "</A></B></th>" & _
	    	"<th><B><A Target=""_self"" HREF=""BusinessAreaAccess.asp?Sort=BusinessAreaID&Ordered=" & strOrder & "&BusinessAreaID=" & lngBusinessAreaID & "&UserID="& lngUserID & """>Vote"
		If strSort = "BusinessAreaID" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""BusinessAreaAccess.asp?Sort=BusinessAreaName&Ordered=" & strOrder & "&BusinessAreaID=" & lngBusinessAreaID & "&UserID="& lngUserID & """>Vote Name"
		If strSort = "BusinessAreaName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>"
		
	response.write"<th>Updated By</th>" & _
		"<th>Date Updated</th></tr>"

	    '<th align="Center">Edit</th>
	'	<th align="Center">User</th>
	'	<th align="Center">Business Area</th>	
	'	<th align="Center">Updated By</th>	
	'	<th align="Center">Date Updated</th>				
	 '   			
	'</tr>
	

    Dim sql 
       
    'If IsEmpty(Request.QueryString("Show")) then
        sql =  "SELECT * FROM qryBusinessAreaAccess WHERE BudgetID = " & Session("BudgetID") & strUserID & " Order By " & strSort & " " & strOrder
    'Else
	'If lngUserID <> 0 Then
        '	sql =  "SELECT * FROM qryBusinessAreaAccess WHERE BudgetID = " & Session("BudgetID") & " AND UserID=" & lngUserID
	'Else
	'	sql =  "SELECT * FROM qryBusinessAreaAccess WHERE BudgetID = " & Session("BudgetID")
	'End If
    'End if  
      
    objRS.Open sql,objCon
		Do until objRS.eof
			Response.Write "<TR><TD><A Target='_self' HREF='BusinessAreaAccess.asp?BusinessAreaID=" & objRS("BusinessAreaID") & "&UserID="& objRS("UserID") & "'><IMG SRC=""../images/edit.jpg""></TD><TD Style=""Text-align:Center""><A HREF=""BusinessAreaAccess.asp?Action=Delete&BusinessAreaID=" & objRS("BusinessAreaID") & "&UserID=" & objRS("UserID") & """>&nbsp;REMOVE ACCESS</A></TD>" & _
				"<TD><B>&nbsp;" & objRS("FName") & " " & objRS("LName") & "</TD><TD>&nbsp;" & objRS("BusinessAreaCode") & "</B></TD><TD>&nbsp;" & objRS("BusinessAreaName") & "</B></TD><TD style=""text-align:center"">" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
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
		
		objRS.Open "SELECT * FROM tblBusinessAreaAccess WHERE BusinessAreaID = " & clng(Session("BusinessAreaID")) & " AND UserID=" & clng(Session("UserID1")),objCon							
		If Not objRS.EOF Then
		    lngBusinessAreaID = objRS("BusinessAreaID")
            lngUserID = objRS("UserID")          						
		Else		                
		    'Do nothing                         
		End if

		objRS.Close	
End Sub

sub DeleteRecord()
    objCon.Execute("DELETE from tblBusinessAreaAccess WHERE BusinessAreaID="&lngBusinessAreaID & " AND UserID=" & lngUserID & " AND BudgetID = " & Session("BudgetID"))    
    'lngBusinessAreaID=0
    'lngUserID=0
End Sub

Sub SaveDetails()	
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spBusinessAreaAccessSave"
                  
                .Parameters.Append objCmd.CreateParameter("BusinessAreaID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("UserID", adInteger, adParamInput)                                
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                                                                    
              
				.Parameters("BusinessAreaID") = Request.Form("BusinessAreaID")
				.Parameters("BudgetID") = Session("BudgetID")		
                .Parameters("UserID") = Request.Form("UserID")                      
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessage = "<B>RECORD SAVED.</B>"
            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"								
     		Session("BusinessAreaID") =  Request.Form("BusinessAreaID")
     		Session("UserID1") = Request.Form("UserID")
		lngUserID = Request.Form("UserID")					
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

Sub SaveAllRecord(strUserIDSave)

  	With objCmd
                .CommandType = 4
                .CommandText = "spBusinessAreaAccessSaveAll"
                  
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("UserID", adInteger, adParamInput)                                
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                                                                    
	
		.Parameters("BudgetID") = Session("BudgetID")		
                .Parameters("UserID") = strUserIDSave                      
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessage = "<B>Record Saved.</B>"
            strMessageIcon = "<img src=""../images/saveticksmall.jpg"" />"									
     		Session("UserID1") = Request.Form("UserID")		
End Sub	


Set objRS = Nothing
Set objCon = Nothing


%>
