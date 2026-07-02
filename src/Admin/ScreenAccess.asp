<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
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
Dim strScreenID
Dim strUserTypeID

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngScreenID
Dim lngUserTypeID
		
'2. Capture Querystring variables	
If Not IsEmpty(Request.QueryString("ScreenID")) Then		
    Session("ScreenID") = Request.QueryString("ScreenID")
    lngScreenID = clng(Request.QueryString("ScreenID"))					
End If		

If Not IsEmpty(Request.QueryString("UserTypeID")) Then		
    Session("UserTypeID1") = Request.QueryString("UserTypeID")
    lngUserTypeID = clng(Request.QueryString("UserTypeID"))
    If lngUserTypeID <> 0  Then strUserTypeID = " AND UserTypeID = " & lngUserTypeID
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
    lngUserTypeID = Session("UserTypeID1") 
    If lngUserTypeID <> 0  Then strUserTypeID = " AND UserTypeID = " & lngUserTypeID
End If

If Request.QueryString("Action") = "Delete" Then
    DeleteRecord()
    lngUserTypeID = Session("UserTypeID1") 
    If lngUserTypeID <> 0  Then strUserTypeID = " AND UserTypeID = " & lngUserTypeID
End If

If Request.QueryString("Action") = "SaveAll" Then
    SaveAllRecord(Request.QueryString("UserTypeID"))
    lngUserTypeID = Session("UserTypeID1") 
    If lngUserTypeID <> 0  Then strUserTypeID = " AND UserTypeID = " & lngUserTypeID

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

    if(frm.UserTypeID.value=="0")
	{
       varAlert += "User ID Cannot Be Blank. \n \n";
       document.getElementById('UserTypeID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('UserTypeID').style.backgroundColor="ffffff";
    
	if(frm.ScreenID.value=="0")
	{
       varAlert += "Business Area Cannot Be Blank. \n \n";
       document.getElementById('ScreenID').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('ScreenID').style.backgroundColor="ffffff";
    
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}


function UserTypeIDSearch()
{	        
	self.location="ScreenAccess.asp?Show=True&ScreenID=" + frm.ScreenID.value + "&UserTypeID=" + frm.UserTypeID.value	
}

function deleteRecord()
{
     if(frm.UserTypeID.value=="0" || frm.ScreenID.value=="0")
	 {       
	    alert("Please select a record to delete");
     }   
     else
     {
        self.location="ScreenAccess.asp?Action=Delete&ScreenID=" + frm.ScreenID.value + "&UserTypeID=" + frm.UserTypeID.value
     }
}
function SaveData2(){
	var varSubmit = true
	if(document.frm.UserTypeID.value==0){
		alert("A User must be selected!");
		varSubmit = false;
	}
	if(varSubmit == true){
	if ( confirm("Would you like to give " + document.frm.UserTypeID.options[document.frm.UserTypeID.selectedIndex].text + " access to ALL Business Areas for the selecetd Budget?"))
		self.location="ScreenAccess.asp?Action=SaveAll&UserTypeID=" + document.frm.UserTypeID.value;
		//frm.submit();
	}else{
		//alert("Access NOT Updated!");
	}

}
//-->
</script>
</head>
<body>
<h3>User Role Screen Access Administration Screen</h3>


<form action="ScreenAccess.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=50% ALIGN=Left BORDER=1 CELLSPACING=1 CELLPADDING=1>
    <tr>
        <th style="text-align:left; width:20%;">&nbsp;User Role</th>
		<td style="text-align:left; width:30%;">
		    <select Style="Width:40%;height:20px" tabindex="20" id="UserTypeID" name="UserTypeID" onchange="UserTypeIDSearch()"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblUserTypes WHERE Active = 'Y' Order By UserTypeID",objCon
    		
		    Do until objRS.EOF
			    If objRS("UserTypeID") = lngUserTypeID Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("UserTypeID") & """" & strSelected & ">" & objRS("UserTypeID") & " - " & objRS("UserTypeName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
		
		
	</tr>
	
	<tr>
		<th style="text-align:left; width:20%;">&nbsp;Screen</th>
		<td style="text-align:left; width:30%;">
		    <select Style="Width:40%;height:20px" tabindex="20" id="ScreenID" name="ScreenID"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblScreens WHERE Secure = 'Y' Order By ScreenName",objCon
    		
		    Do until objRS.EOF
			    If objRS("ScreenID") = lngScreenID Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("ScreenID") & """" & strSelected & ">" & objRS("ScreenName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
	</tr>
<tr><td style="text-align:left;height:20px" colspan="4">&nbsp;</td></tr>		
</table>
<br>
<br />
<br />
<br />
<br />


<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='ScreenAccess.asp?ScreenID=0&UserTypeID=0'"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon%></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage%></TD>
	</tr>
</table>


<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
<%
	'Dynamically build the menu items depending on the sort selection 
	Response.write "<th>Edit</th><th>REMOVE</th>" & _
	    	"<th><B><A Target=""_self"" HREF=""ScreenAccess.asp?Sort=LName&Ordered=" & strOrder & "&ScreenID=" & lngScreenID & "&UserTypeID="& lngUserTypeID & """>Screen"
		If strSort = "LName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	Response.write "</A></B></th>" & _
	    	"<th><B><A Target=""_self"" HREF=""ScreenAccess.asp?Sort=ScreenID&Ordered=" & strOrder & "&ScreenID=" & lngScreenID & "&UserTypeID="& lngUserTypeID & """>User Role ID"
		If strSort = "ScreenID" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""ScreenAccess.asp?Sort=BusinessAreaName&Ordered=" & strOrder & "&ScreenID=" & lngScreenID & "&UserTypeID="& lngUserTypeID & """>User Role Name"
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
     
        If IsNull(Session("UserTypeID1")) or  Session("UserTypeID1") = "" Then
            Session("UserTypeID1") = 0
        End If
    
        sql =  "SELECT * FROM qryScreenAccess WHERE UserTypeID = " & Session("UserTypeID1") & " Order By ScreenName" 
       
    objRS.Open sql,objCon
		Do until objRS.eof
			Response.Write "<TR><TD><A Target='_self' HREF='ScreenAccess.asp?ScreenID=" & objRS("ScreenID") & "&UserTypeID="& objRS("UserTypeID") & "'><IMG SRC=""../images/edit.jpg""></TD><TD Style=""Text-align:Center""><A HREF=""ScreenAccess.asp?Action=Delete&ScreenID=" & objRS("ScreenID") & "&UserTypeID=" & objRS("UserTypeID") & """>&nbsp;REMOVE ACCESS</A></TD>" & _
				"<TD><B>&nbsp;" & objRS("ScreenName") & "</TD><TD>&nbsp;" & objRS("UserTypeID") & "</B></TD><TD>&nbsp;" & objRS("UserTypeName") & "</B></TD><TD style=""text-align:center"">" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
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
		
		objRS.Open "SELECT * FROM tblScreenAccess WHERE ScreenID = " & clng(Session("ScreenID")) & " AND UserTypeID=" & clng(Session("UserTypeID1")),objCon							
		If Not objRS.EOF Then
		    lngScreenID = objRS("ScreenID")
            lngUserTypeID = objRS("UserTypeID")          						
		Else		                
		    'Do nothing                         
		End if

		objRS.Close	
End Sub

sub DeleteRecord()
    objCon.Execute "DELETE from tblScreenAccess WHERE ScreenID=" & lngScreenID & " AND UserTypeID=" & lngUserTypeID & ""
    strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
    strMessage = "<B>RECORD DELETED.</B>"
End Sub

Sub SaveDetails()	
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spScreenAccessSave"
                  
                .Parameters.Append objCmd.CreateParameter("ScreenID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("UserTypeID", adInteger, adParamInput)                                
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                                                                    
              
				.Parameters("ScreenID") = Request.Form("ScreenID")
			    .Parameters("UserTypeID") = Request.Form("UserTypeID")                      
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
            strMessage = "<B>RECORD SAVED.</B>"										
     		Session("ScreenID") =  Request.Form("ScreenID")
     		Session("UserTypeID1") = Request.Form("UserTypeID")
		    lngUserTypeID = Request.Form("UserTypeID")
    					
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
