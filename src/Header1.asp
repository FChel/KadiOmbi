<%@ Language=VBScript %>
<% Option Explicit

Response.Expires = -1500

Dim objCon
Dim objRS
Dim strSelected

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

	
	objCon.Open Session("DBConnection")
	
Dim arrStatus(7)

arrStatus(0) = "images/delete.png"
arrStatus(1) = "images/open.png"
arrStatus(2) = "images/ready.gif"
arrStatus(3) = "images/cross.png"
arrStatus(4) = "images/tick.png"
arrStatus(5) = "images/Closed.png"
arrStatus(6) = "images/bug.png"
arrStatus(7) = "images/wrench.png"


If IsNull(Session("BusinessAreaID")) Then Session("BusinessAreaID") = 1002

%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<LINK rel="stylesheet" type="text/css" href="BERTStyle.css">
</HEAD>
<BODY>
<FORM action="Header1.asp" method=POST id=frm name=frm>

<TABLE WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<TR>
	<TH Style="Width:10%"><%=Session("BAName")%></TH>
		<TH Style="Width:23%"><select id="BusinessArea" name="BusinessArea" Style="Width:90%" onchange="parent.location='<%=Session("Header")%>?BusinessAreaID=' + frm.BusinessArea.value">
		 <option value="0">Please Select..</option>
<%	

	objRS.Open "SELECT * FROM qryBusinessAreaAccess WHERE BudgetID = " & Session("BudgetID") & " AND UserID = '" & Session("UserID") & "'",objCon
		
	Do until objRS.EOF
		If clng(objRS("BusinessAreaID")) = clng(Session("BusinessAreaID")) Then
		   	strSelected = " SELECTED "
		   	Session("BusinessAreaCode") = objRS("BusinessAreaCode")
            Session("CCCeilingsOn") = objRS("CostCentreCeilingsOn")
            Session("ACCeilingsOn") = objRS("AccountClassCeilingsOn")
			Session("ProjectCeilingsOn") = objRS("ProjectCeilingsOn")
		Else
			strSelected = ""
            'Session("CCCeilingsOn") = "N"
            'Session("ACCeilingsOn") = "N"
		End if
			Response.Write "<option Value=""" & objRS("BusinessAreaID") & """" & strSelected & ">" & objRS("BusinessAreaCode") & " - " & objRS("BusinessAreaName") & "</OPTION>"
		objRS.Movenext
	Loop
	
    objRS.Close
    
    'objRS.Open "SELECT BMCeiling FROM tblBACeilingLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Level1ID = " & Session("BusinessAreaID") & "",objCon
    
        'If Not objRS.EOF Then
            'Session("BACeiling") = objRS("BMCeiling")
        'Else
            'Session("BACeiling") = 0
        'End If
        
    'objRS.Close
	
%>
</select>
</TH>
		<TH Style="Width:10%;Height:25px"><%=Session("CCName")%></TH>
		<TH Style="Width:23.5%"><select id="CostCentre" name="CostCentre" Style="Width:95%" onchange="parent.location='<%=Session("Header")%>?CostCentreID=' + frm.CostCentre.value">
		 <option value="0">Please Select..</option>
<%	

	'Set objRS = objSettings.CostCentresByBusinessArea_List (clng(Session("BudgetID")),clng(Session("BusinessAreaID")))
	objRS.Open "SELECT * FROM qryCostCentresByBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND CostObjectTypeID = 2 AND VersionID = " & Session("VersionID") & "",objCon
	
	Do until objRS.EOF
		If objRS("CostCentreID") = clng(Session("CostCentreID")) Then
		    strSelected = " SELECTED "
		    Session("CostCentreType") = objRS("CostCentreType")
		    Session("CCStatusID") = objRS("StatusID")
            'Session("Segment4") = objRS("ProgramCode")
            Session("InputSheetID") = objRS("InputSheetID")
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("CostCentreID") & """" & strSelected & ">" & objRS("ProgramCode") & " - " & objRS("CostCentreName") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close

%></select>
		<th><IMG SRC="<%=arrStatus(Session("CCStatusID"))%>"></th>	
		
	

		<TH Style="Width:10%">Version</TH>
		<TH Style="Width:23.5%"><select id="Version" name="Version" Style="Width:95%" onchange="parent.location='BERTFrameset.asp?VersionID=' + frm.Version.value">
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
</TR>
</TABLE>

<HR>
</FORM>
</BODY >
</HTML>
