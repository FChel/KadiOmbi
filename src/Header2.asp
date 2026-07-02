<%@ Language=VBScript %>
<% Option Explicit

Response.Expires = -1500

Dim objCon
Dim objRS
Dim strSelected

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
	
	objCon.Open Session("DBConnection")	

    Session("Header") = "ScrollingFrameset.asp"

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="BERTStyle.css">
</HEAD>
<BODY>
<FORM action="Header2.asp" method=POST id=frm name=frm>

<TABLE WIDTH=100% BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<TR>
		<TH Width=20% Height="25px"><%=Session("BAName")%></TH>
		<TH Width=30%><select id="BusinessArea" name="BusinessArea" Style="Width:80%" onchange="parent.location='<%=Session("Header")%>?BusinessAreaID=' + frm.BusinessArea.value">
<%	

	objRS.Open "SELECT * FROM qryBusinessAreaAccess WHERE BudgetID = " & Session("BudgetID") & " AND UserID = " & Session("UserID") & "",objCon
		
	Do until objRS.EOF
		If objRS("BusinessAreaID") = clng(Session("BusinessAreaID")) Then
			strSelected = " SELECTED "
			Session("Vote") = objRS("BusinessAreaCode")
			Session("CCCeilingsOn") = objRS("CostCentreCeilingsOn")
            Session("ACCeilingsOn") = objRS("AccountClassCeilingsOn")
			Session("ProjectCeilingsOn") = objRS("ProjectCeilingsOn")
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("BusinessAreaID") & """" & strSelected & ">" & objRS("BusinessAreaCode") & " - " & objRS("BusinessAreaName") & "</OPTION>"
		objRS.Movenext
	Loop
	
    objRS.Close
	
%>
</select>
</TH>
	<TH Width=20%>Version</TH>
		<TH Width=30%><select id="Version" name="Version" Style="Width:80%" onchange="parent.location='<%=Session("Header")%>?VersionID=' + frm.Version.value">
<%	

    objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & "",objCon
	
	Do until objRS.EOF
		If objRS("VersionID") = clng(Session("VersionID")) Then
			strSelected = " SELECTED "
			Session("VersionName") = objRS("VersionName")
			Session("ColumnLock") = objRS("ColumnLock")
			Session("VersionTypeID") = objRS("VersionTypeID")
			Session("BaseBudgetVersionID") = objRS("BaseBudgetVersionID") 
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("VersionID") & """" & strSelected & ">" & objRS("VersionName") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close
	
	'Get Default Program and Sub Program
	
	objRS.Open "SELECT DefaultCostCentre FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon
	
	    If Not objRS.EOF Then
	        Session("CostCentreID") = objRS(0)
	    Else
	        Session("CostCentreID") = 0
	    End If
	    
	objRS.Close	

	
%>
</select>
</TH>

</TR>
</TABLE>
<HR>
</FORM>
</BODY >
</HTML>