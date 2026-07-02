<%@ Language=VBScript %>
<% Option Explicit

Response.Expires = -1500

Dim objCon
Dim objRS
Dim strSelected

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
	
	objCon.Open Session("DBConnection")	

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="BERTStyle.css">
</HEAD>
<BODY>
<FORM action="Header3.asp" method=POST id=frm name=frm>

<TABLE WIDTH=100% BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<TH Style="Width:20%;height:25px">Version</TH>
		<TH Width=30%><select id="Version" name="Version" Style="Width:80%" onchange="top.location='ScrollingFrameset.asp?HeaderMenu=Header3.asp&VersionID=' + frm.Version.value">
<%	

    objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & "",objCon
	
	Do until objRS.EOF
		If objRS("VersionID") = clng(Session("VersionID")) Then
			strSelected = " SELECTED "
			Session("VersionName") = objRS("VersionName")
			Session("VersionTypeID") = objRS("VersionTypeID")
			Session("ColumnLock") = objRS("ColumnLock")
			Session("BaseBudgetVersionID") = objRS("BaseBudgetVersionID") 
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("VersionID") & """" & strSelected & ">" & objRS("VersionName") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close

	
%>
</select>
</TH>
<TH Width=50% colspan="2"></TH>

</TH>
</TR>
</TABLE>
<HR>
</FORM>
</BODY >
</HTML>