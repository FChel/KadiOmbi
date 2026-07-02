<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")
 
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

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngUserID
Dim strUserLogon
Dim strFName
Dim strLName
Dim strEmailAddress
Dim intUserTypeID
Dim strComments
Dim strActive
Dim strSQL

'2. Declare search variables

Dim strSearchType
Dim strFNameSearch
Dim strLNameSearch
Dim strUserLogonSearch
Dim strUserTypeSearch
dim strSort
Dim strOrder
Dim strBusinessAreaID
Dim strUserID
Dim intLanguage
Dim strEmployeeID 

'Declare and set default arrays

Dim arrYesNo(2)
	
	arrYesNo(1) = "N"
	arrYesNo(2) = "Y"
	
Dim arrLanguage(2)
	
	arrLanguage(1) = "ENGLISH"
	arrLanguage(2) = "SWAHALI"
	
	'3. Capture Querystring variables
	
	If Not IsEmpty(Request.QueryString("UserID")) Then
		
		Session("UserIDAdmin") = Request.QueryString("UserID")
		lngUserID = clng(Request.QueryString("UserID"))	
				
	End If
	
	If Not IsEmpty(Request.QueryString("UserLogon")) Then
		
		Session("UserLogon") = Request.QueryString("UserLogon")
		strUserLogon = Request.QueryString("UserLogon")	
				
	End If
	
	'4 Set up Search functions if required
	
	If Not IsEmpty(Request.QueryString("LNameSearch")) Then
		strLNameSearch = Request.QueryString("LNameSearch")
		strSearchType = "User Last Name search = " & "<FONT Color=""Navy"">" & strLNameSearch & "</FONT>"
		strLNameSearch = Request.QueryString("LNameSearch")	& "%"
	End If	
	
	If Not IsEmpty(Request.QueryString("FNameSearch")) Then
		strFNameSearch = Request.QueryString("FNameSearch")
		strSearchType = "User First Name search = " & "<FONT Color=""Navy"">" & strFNameSearch & "</FONT>"
		strFNameSearch = Request.QueryString("FNameSearch") & "%"
	End If	
	
	If Not IsEmpty(Request.QueryString("UserLogonSearch")) Then
		strUserLogonSearch = Request.QueryString("UserLogonSearch")
		strSearchType = "User logon search = " & "<FONT Color=""Navy"">" & strUserLogonSearch & "</FONT>"
		strUserLogonSearch = Request.QueryString("UserLogonSearch") & "%"
	End If	
	
	If Not IsEmpty(Request.QueryString("UserTypeSearch")) Then
		strUserTypeSearch = Request.QueryString("UserTypeSearch")
		strSearchType = "User Type search = " & "<FONT Color=""Navy"">" & strUserTypeSearch & "</FONT>"
		strUserTypeSearch = Request.QueryString("UserTypeSearch") & "%"
	End If
	
	If Not IsEmpty(Request.QueryString("EmployeeIDSearch")) Then
		strUserTypeSearch = Request.QueryString("EmployeeIDSearch")
		strSearchType = "User Type search = " & "<FONT Color=""Navy"">" & strUserTypeSearch & "</FONT>"
		strUserTypeSearch = Request.QueryString("EmployeeIDSearch") & "%"
	End If
	
	If Not IsEmpty(Request.QueryString("Sort")) Then
	   strSort = Request.QueryString("sort")
    Else
	   strSort = "LName"
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
	End If

	'Load page details
	If Not IsEmpty(Request.QueryString("UserLogon")) Then
	    LoadUserLogonDetails()
	Else
	    LoadDetails()
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

<script LANGUAGE="javascript">
<!--
function SaveData(){
	
var varSubmit = true
			
	if(isWhitespace(frm.LName.value)==true){
		warnInvalid(frm.LName,"You must enter a Last Name!") 
		return false;
	}
		
	if(isWhitespace(frm.FName.value)==true){
		warnInvalid(frm.FName,"You must enter a First Name!") 
		return false;
	}
	
	if (frm.Active.value == 0){
		alert("You must select either Active Y or N!")
		return false;
	}
	
	if(isWhitespace(frm.UserLogon.value)==true){
		warnInvalid(frm.UserLogon,"You must enter a User Logon!") 
		return false;
	}
	
	if (frm.UserTypeID.value == 0){
		alert("You must select a User Type ID!")
		return false;
	}
	
					
	if(varSubmit == true){
	frm.submit();
		
	}
}

function DeleteData(){
	if(window.confirm('Confirm delete')==true){
	self.location="User.asp?Action=Delete"
	}
}

function UserIDSearch(){
	if(frm.LNameEdit.checked){}else
	{self.location="User.asp?Action=UserIDSearch&UserIDSearch=" + frm.UserID.value}
}

function UserIDSearch()
{	
	self.location="User.asp?UserID=" + frm.UserID.value
}

function UserLogonSearch()
{	
	self.location="User.asp?UserLogon=" + frm.UserLogon.value
}
//-->
</script>
</head>
<body>
<main class="main py-3">
      <div class="container">
<h3>User Administration Screen</h3>
<form action="User.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
		<th height="20px" align="left">&nbsp;User ID</th>
		<td>&nbsp;<input style="text-align:left;width:90%" READONLY id="UserID" name="UserID" maxlength="50" TABINDEX="1" value="<%=lngUserID%>"></td>
		<th align="left">&nbsp;User Logon</th>
		<td>&nbsp;<input style="text-align:left;width:90%" id="UserLogon" name="UserLogon" maxlength="50" TABINDEX="2" value="<%=strUserLogon%>" ><img SRC="../images/searchtopic.gif" onClick="self.location='User.asp?Action=UserLogonSearch&UserLogonSearch='+(frm.UserLogon.value);"></td>
	</tr>
	<tr>
		<th height="20px" align="left">&nbsp;First Name</th>
		<td>&nbsp;<input style="text-align:left;width:90%"id="FName" name="FName" maxlength="50" TABINDEX="3" value="<%=strFName%>"><img SRC="../images/searchtopic.gif" onClick="self.location='User.asp?Action=FNameSearch&amp;FNameSearch='+(frm.FName.value);"></td>
		<th align="left">&nbsp;Last Name</th>
		<td>&nbsp;<input style="text-align:left;width:90%" id="LName" name="LName" maxlength="50" TABINDEX="4" value="<%=strLName%>"><img SRC="../images/searchtopic.gif" onClick="self.location='User.asp?Action=LNameSearch&amp;LNameSearch='+(frm.LName.value);"></td>
	</tr>
	<tr>
		<th height="20px" align="left">&nbsp;Email</th>
		<td colspan="1">&nbsp;<input style="text-align:left;width:90%" id="EmailAddress" name="EmailAddress" maxlength="50" TABINDEX="5" value="<%=strEmailAddress%>"></td>
		<th align="left">&nbsp;Employee ID</th>
		<td>&nbsp;<input style="text-align:left;width:90%" id="EmployeeID" name="EmployeeID" maxlength="20" TABINDEX="6" value="<%=strEmployeeID%>"><img SRC="../images/searchtopic.gif" onClick="self.location='User.asp?Action=EmployeeIDSearch&amp;EmployeeIDSearch='+(frm.EmployeeID.value);"></td>
	</tr>
    <tr>
		<th height="20px" align="left">&nbsp;User Type</th>
		<td><select Style="Width:40%" tabindex="7" id="UserTypeID" name="UserTypeID"><OPTION Value=0>Please Select..</OPTION>
	<%	
		objRS.Open "SELECT * FROM tblUserTypes WHERE Active = 'Y'",objCon
		
		Do until objRS.EOF
			If objRS("UserTypeID") = intUserTypeID Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("UserTypeID") & """" & strSelected & ">" & objRS("UserTypeName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		
	%></select> <img SRC="../images/searchtopic.gif" onClick="self.location='User.asp?Action=UserTypeSearch&amp;UserTypeSearch='+(frm.UserTypeID.value);"></td>
	
	<th align="left">&nbsp;Active</th><td><select Style="Width:40%" tabindex="8" id="Active" name="Active">
	<%
		For x = 1 to 2
			If arrYesNo(x) = cstr(strActive) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End If
				Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		Next
	%>
		</select></td>
	</tr>
	<tr>
		<td colspan="4" Align="left">&nbsp;</td>
	</tr>
	<tr>
		<th colspan="4" Align="left" Height="20px">&nbsp;Comments</th>
	</tr>
	<tr>
	    <td colspan="4"><TEXTAREA rows=4 cols=190 id=Comments name=Comments tabindex="8"><%=strComments%>
	
</TEXTAREA></td>
	</tr>
</table>
<br>
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<tr>
  <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
  <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
  <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='User.asp?UserID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp;</button></td>
  <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
  <TD class='locked' Width="600px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
</tr>
</table>
<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr><th height="20px" colspan="8" align="Left">&nbsp;Search Type&nbsp;:&nbsp;<font Color="Red"><%=strSearchType%></font></th></tr>
<tr><td colspan="8">&nbsp;</td></tr>

<%
	'Dynamically build the menu items depending on the sort selection 
	Response.write "<th><B><A Target=""_self"" HREF=""User.asp?Sort=EmployeeID&Ordered=" & strOrder & """>Employee ID"
		If strSort = "EmployeeID" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
		
	Response.write "<th><B><A Target=""_self"" HREF=""User.asp?Sort=LName&Ordered=" & strOrder & """>Last Name"
		If strSort = "LName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	Response.write "</A></B></th>" & _
	    	"<th><B><A Target=""_self"" HREF=""User.asp?Sort=FName&Ordered=" & strOrder & """>First Name"
		If strSort = "FName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""User.asp?Sort=UserLogon&Ordered=" & strOrder & """>User Logon"
		If strSort = "UserLogon" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
	
	response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""User.asp?Sort=UserTypeName&Ordered=" & strOrder & """>User Type"
		If strSort = "UserTypeName" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"

	response.write"</A></B></th>" & _
		"<th><B><A Target=""_self"" HREF=""User.asp?Sort=Active&Ordered=" & strOrder & """>Active"
		If strSort = "Active" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
		
	response.write"</A></B></th>"
		
	response.write"<th>Updated By</th>" & _
		"<th>Date Updated</th></tr>"	

If Request.QueryString("Action") = "LNameSearch" Then

		strLNameSearch =  Replace(strLNameSearch,"'","''")
strSQL = "SELECT Top 200 * FROM qryUsers WHERE LName Like '" & strLNameSearch & "' Order By LName ASC"

ElseIF Request.QueryString("Action") = "FNameSearch" Then

	strFNameSearch =  Replace(strFNameSearch,"'","''")
strsql = "SELECT Top 200 * FROM qryUsers WHERE FName Like '" & strFNameSearch & "' Order By FName ASC"

Elseif Request.QueryString("Action") = "UserLogonSearch" then

	strUserLogonSearch =  Replace(strUserLogonSearch,"'","''")
strsql = "SELECT Top 200 * FROM qryUsers WHERE UserLogon Like '" & strUserLogonSearch & "' Order By UserLogon ASC"

Elseif Request.QueryString("Action") = "UserTypeSearch" then

	strUserTypeSearch =  Replace(strUserTypeSearch,"'","''")
strsql = "SELECT Top 200 * FROM qryUsers WHERE UserTypeID Like '" & strUserTypeSearch & "' Order By LName ASC"

Elseif Request.QueryString("Action") = "EmployeeIDSearch" then

	strUserTypeSearch =  Replace(strUserTypeSearch,"'","''")
strsql = "SELECT Top 200 * FROM qryUsers WHERE EmployeeID Like '" & strUserTypeSearch & "' Order By EmployeeID ASC"

Else
strSQL = "SELECT top 200 * FROM qryUsers Order By " & strSort & " " & strOrder

End If

    objRS.Open strSQL,objCon
		Do until objRS.eof
			Response.Write "<TR><TD><A Target=""_self"" HREF=""User.asp?UserID=" & objRS("UserID") & """><B>&nbsp;" & objRS("EmployeeID") & "</B></A></TD><TD><A Target=""_self"" HREF=""User.asp?UserID=" & objRS("UserID") & """><B>&nbsp;" & objRS("LName") & "</B></A></TD><TD style=""text-align:center"">" & objRS("FName") & "</B></TD><TD style=""text-align:center"">" & objRS("UserLogon") & "</TD><TD style=""text-align:center"">" & objRS("UserTypeName") & "</TD><TD style=""text-align:center"">" & objRS("Active") & "</TD><TD style=""text-align:center"">" & objRS("UpdatedByName") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
			objRS.movenext
		Loop
			
	objRS.Close


%>
</table>

</div>
</main>
	  
</body>
<!-- #Include file=../CC/CAPSFooter.asp -->
</html>

<% 

Sub LoadDetails()

'Description:	Loads Caller's details into page if applicable.
		
		objRS.Open "SELECT * FROM tblUsers WHERE UserID = " & clng(Session("UserIDAdmin")) & "",objCon
				
			If Not objRS.EOF Then
				
				strFName = objRS("FName")
				strLName = objRS("LName")
				strUserLogon = objRS("UserLogon")
				strEmailAddress = objRS("EmailAddress")
				intUserTypeID = objRS("UserTypeID")
				strComments = objRS("Comments")
				strActive = objRS("Active")	
				intLanguage = objRS("Language")			
				strEmployeeID = objRS("EmployeeID")		
			Else
				
				strFName = ""
				strLName = ""
				strUserLogon = ""
				strEmailAddress = ""
				intUserTypeID = 0
				strComments = ""
				strActive = "Y"
				intLanguage = 1
				strEmployeeID = ""
				
			End If

		objRS.Close
	

End Sub

Sub LoadUserLogonDetails()

'Description:	Loads Caller's details into page if applicable.
		
		objRS.Open "SELECT * FROM tblUsers WHERE UserLogon = '" & Session("UserLogon") & "'",objCon
							
			If Not objRS.EOF Then
				
				lngUserID = objRS("UserID")
				strFName = objRS("FName")
				strLName = objRS("LName")
				strUserLogon = objRS("UserLogon")
				strEmailAddress = objRS("EmailAddress")
				intUserTypeID = objRS("UserTypeID")
				strComments = objRS("Comments")
				strActive = objRS("Active")				
				strEmployeeID  = objRS("EmployeeID")	
			Else
				
				strFName = ""
				strLName = ""
				strUserLogon = ""
				strEmailAddress = ""
				intUserTypeID = 0
				strComments = ""
				strActive = "Y"
				strEmployeeID = ""
				
			End If

		objRS.Close
	

End Sub

Sub SaveDetails()

		If Not IsEmpty(Request.Form("LName")) Then
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spUserSave"
                
                .Parameters.Append objCmd.CreateParameter("UserID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("UserLogon", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("FName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("LName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("EmailAddress", adVarChar, adParamInput, 150)
                .Parameters.Append objCmd.CreateParameter("UserTypeID", adInteger)
                .Parameters.Append objCmd.CreateParameter("Language", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Comments", adLongVarChar, adParamInput, -1)
                .Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
                .Parameters.Append objCmd.CreateParameter("EmployeeID", adVarChar, adParamInput, 20)
				
                .Parameters.Append objCmd.CreateParameter("UserIDOutput", adInteger, adParamOutput)
              
				.Parameters("UserID") = Session("UserIDAdmin")	
				.Parameters("UserLogon") = Request.Form("UserLogon")			
                .Parameters("FName") = Request.Form("FName")
                .Parameters("LName") = Request.Form("LName")             
                .Parameters("EmailAddress") = Request.Form("EmailAddress")
                .Parameters("UserTypeID") = Request.Form("UserTypeID")
                .Parameters("Language") = 1
                .Parameters("Comments") = Request.Form("Comments")
                .Parameters("Active") = Request.Form("Active")
                .Parameters("UpdatedBy") = Session("UserID")
                .Parameters("EmployeeID") = Request.Form("EmployeeID")
				
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
			'Return the result of the Save Function.
     		Session("UserIDAdmin") = objCmd.Parameters.Item("UserIDOutput")
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
