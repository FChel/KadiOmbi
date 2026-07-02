<!-- #Include file=../ADOVBS.inc -->

<%	'Option Explicit
	Response.Expires = -1500
	
	Session("CurrentPage") = "Admin/FinStatementConfig.asp"

Dim objAdmin
Dim objRS
Dim objRS2
Dim objSettings
Dim objCon
Dim objCmd

Dim	intVersionID
Dim dteDateUpdated
Dim strUpdatedBy
Dim strSelected
Dim strMessage
Dim strMessageIcon
Dim intLevel1ID
Dim strLevel1Name
Dim strLevel1NameL2
Dim intReportGroup1ID
Dim intReportGroup2ID
Dim intDrillDownID
Dim strMathSign
Dim intSortOrder
Dim strGLType
Dim intDetailedGLCode
Dim intGLCode
Dim intBudgetClassID

Dim intLevel2ID
Dim intSortOrder2
Dim strTransactionID
Dim strPaymentRatio
Dim intTotalGrouping1
Dim intTotalGrouping2
Dim intTotalGrouping3
Dim intTotalGrouping4
Dim intTotalGrouping5
Dim intMathSignage
Dim intRLLevel2ID
Dim intRLLevel1ID
Dim strTransactionType

Dim intEdit
Dim strBackColour
Dim strBackColour2
Dim intLevelID
Dim arrType(2,5)
Dim x
Dim arrMathSign(2)
Dim intLevel1IDLast
Dim arrTransactionID(6)
Dim arrReport(99)
Dim arrDrill(16)

intLevel1ID = 0
intLevelID = 0
intLevel2ID = 0

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")

objCon.Open Session("DBConnection")

If Session("ReportID") = "" Or IsNull(Session("ReportID")) Then
	Session("ReportID") = 1
End If

If Not IsEmpty(Request.QueryString("Level1ID")) Then
	intLevel1ID = Request.QueryString("Level1ID")
End If

If Not IsEmpty(Request.QueryString("Level2ID")) Then
	intLevel2ID = Request.QueryString("Level2ID")
	if intLevel2ID = "" then intLevel2ID = 0
	intGLCode = Request.QueryString("Level2ID")
	
End If

If Not IsEmpty(Request.QueryString("VersionID")) Then
	Session("VersionID") = Request.QueryString("VersionID")
End If

If Not IsEmpty(Request.QueryString("TransactionID")) Then
	strTransactionID = Request.QueryString("TransactionID")
End If	

If Not IsEmpty(Request.QueryString("ReportID")) Then
	Session("ReportID") = Request.QueryString("ReportID")
End If

If Not IsEmpty(Request.QueryString("BudgetID")) Then
	Session("BudgetID") = Request.QueryString("BudgetID")
				
	'Session("BudgetName") = objSettings.BudgetName_Get(Session("BudgetID"))
	Session("VersionID") = 0
End If

If Not IsEmpty(Request.QueryString("Edit")) Then
	intEdit = Request.QueryString("Edit")
Else
	intEdit = 0
End If

Response.Write Request.QueryString("Edit")

'Set objAdmin = Server.CreateObject("BERTv5.Admin")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS2 = Server.CreateObject("ADODB.Recordset")
'Set objSettings = Server.CreateObject("BERTv5.Settings")

	'Save record.	
	If Request.QueryString("Action") = "Save" Then
		SaveRecord()
	End If
	
	'Delete record.	
	If Request.QueryString("Action") = "Delete" Then
		DeleteRecord Request.QueryString("LevelID"),Request.QueryString("Level2ID"),Request.QueryString("Level1ID"),Request.QueryString("TransactionID")
	End If
	
	'Run Auto Config.	
	If Request.QueryString("Action") = "AutoConfig" Then
		AutoConfig()
	End If

	'Load the details of the selected record to the text part of the screen, depending on what has been selected from the list.
	If intLevel1ID <> 0 AND intLevel2ID = 0 Then

		LoadLevel1()

		strBackColour = ""
		strBackColour2 = "style=background-color:#CCCCCC READONLY"
	End If

	If intLevel2ID <> 0 Then
		Response.Write "SELECT * FROM tblReportLayoutLevel2 WHERE ReportLayoutLevel2ID = " & strTransactionID & ""
		LoadLevel2()
		strBackColour = "Style=color:#CCCCCC style=background-color:#CCCCCC READONLY"
		strBackColour2 = ""
	End If

	'Build the GL Type Array
	arrType(1,1) = "A - (Asset)"
	arrType(2,1) = "A"
	arrType(1,2) = "C - (Capital)"
	arrType(2,2) = "C"
	arrType(1,3) = "E - (Expense)"
	arrType(2,3) = "E"
	arrType(1,4) = "L - (Liability)"
	arrType(2,4) = "L"
	arrType(1,5) = "R - (Revenue)"
	arrType(2,5) = "R"
	
	arrMathSign(1) = "+"
	arrMathSign(2) = "-"
	
	arrReport(1) = "1 : Operating Statement Report (OS)"
	arrReport(2) = "2 : Sub Budget Classifications Config (SBC)"
	arrReport(3) = "3 : TBD"
	arrReport(4) = "4 : TBD"
	arrReport(5) = "5 : Revenue Report Config (RR)"
	arrReport(21) = "21 : Revenue Report Config (RR)"

	'Set drill array strings
	arrDrill(2) = "2 - Basic Entry Multiple GL Schedule"
	arrDrill(3) = "3 - Single GL Multiple Entry Schedule"
	arrDrill(4) = "4 - Rate Based Entry Schedule"
	arrDrill(5) = "5 - Capital Expenditure Schedule"
	arrDrill(6) = "6 - Staffing Schedule"
	arrDrill(7) = "7 - Travel Schedule (Not Available)"
	arrDrill(8) = "8 - Unit Entry Schedule (NotAvailable)"
	arrDrill(9) = "9 - Internal Service Charge Schedule (Not Available)"
	arrDrill(10) = "10 - Client Statistics Schedule"
	arrDrill(11) = "11 - Cost Transferred Schedule"
	arrDrill(12) = "12 - Read Only Multiple GL Schedule"
	arrDrill(13) = "13 - Multiple GL Multiple Entry Schedule"
	arrDrill(14) = "14 - Unit Sales Entry Schedule"
	arrDrill(15) = "15 - Depreciation Schedule"
	arrDrill(16) = "16 - Multiple GL Multiple Entry Schedule - With Baseline"

	'Call the function to get the number of Level1ID's there are
	Level1IDLast()

%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
<script src="../ButtonRollOver.js">
</script>
<script src="../HelpText.js">
</script>
<script src="../formChek.js">
</script>

<script language=javascript>
//<!--
    function SaveData() {
        var varSubmit = true
        var varAlert = "";
        /*
        if ((isNonnegativeInteger(frm.Level1ID.value) == false) || (frm.Level1ID.value == 0)) {
            varAlert += "Please enter Level 1 ID. Level 1 ID must be a numeric value.  \n \n";
            document.getElementById('Level1ID').style.backgroundColor = "ff8080";
            varSubmit = false;
        }
        else document.getElementById('Level1ID').style.backgroundColor = "ffffff";    
      
        if (isWhitespace(frm.Level1Name.value) || frm.Level1Name.value == "0") {
            varAlert += "Level 1 name ENG cannot Be Blank. \n \n";
            document.getElementById('Level1Name').style.backgroundColor = "ff8080";
            varSubmit = false;
        }
        else document.getElementById('Level1Name').style.backgroundColor = "ffffff";


        if (isWhitespace(frm.Level1NameL2.value) || frm.Level1NameL2.value == "0") {
            varAlert += "Level 1 name ENG cannot Be Blank. \n \n";
            document.getElementById('Level1NameL2').style.backgroundColor = "ff8080";
            varSubmit = false;
        }
        else document.getElementById('Level1NameL2').style.backgroundColor = "ffffff";

        if ((isNonnegativeInteger(frm.SortOrder.value) == false) || (frm.SortOrder.value == 0)) {
            varAlert += "Please enter Sort Order. Sort Order must be a numeric value.  \n \n";
            document.getElementById('SortOrder').style.backgroundColor = "ff8080";
            varSubmit = false;
        }
        else document.getElementById('SortOrder').style.backgroundColor = "ffffff";

        if ((isNonnegativeInteger(frm.TotalGrouping1.value) == false) || (frm.TotalGrouping1.value == 0)) {
            frm.TotalGrouping1.value = 0
        }
      
        if ((isNonnegativeInteger(frm.TotalGrouping2.value) == false) || (frm.TotalGrouping2.value == 0)) {
            frm.TotalGrouping2.value = 0
        }

        if ((isNonnegativeInteger(frm.TotalGrouping3.value) == false) || (frm.TotalGrouping3.value == 0)) {
            frm.TotalGrouping3.value = 0
        }

        if ((isNonnegativeInteger(frm.TotalGrouping4.value) == false) || (frm.TotalGrouping4.value == 0)) {
            frm.TotalGrouping4.value = 0
        }

        if ((isNonnegativeInteger(frm.TotalGrouping5.value) == false) || (frm.TotalGrouping5.value == 0)) {
            frm.TotalGrouping5.value = 0
        }

        if ((isNonnegativeInteger(frm.DetailedGLCode.value) == false) || (frm.DetailedGLCode.value == 0)) {
            frm.DetailedGLCode.value = 0
        }
        */
        if (varSubmit == true) {
            frm.submit();
        }
        else {
            alert(varAlert);
        }
    }



function LevelText(){
	alert("LevelID = |" + frm.LevelID.value + "|");
}

function Level1Change(){
	//alert("LevelID = |" + frm.Level1ID2.value + "|");
	frm.Level1ID.value = frm.Level1ID2.value
}

function SaveData2(){
	var varSubmit = true
	if(document.frm.StatusID.value==0){
		alert("A Status must be selected!");
		varSubmit = false;
	}
	if(varSubmit == true){
	if ( confirm("Would you like to UPDATE all Business Areas to Status " + document.frm.StatusID.options[document.frm.StatusID.selectedIndex].text + " ?"))
		self.location="BusinessAreaStatus.asp?Action=SaveAll&StatusID=" + document.frm.StatusID.value;
	}else{
		//alert("Status NOT Updated!");
	}

}

function DeleteData(){
	if( confirm("Delete the selected record?") )
	{
		//self.location="FinStatementConfig.asp?Action=Delete&LevelID="+frm.LevelID.value+"&Level2ID="+frm.Level2ID.value+"&Level1ID="+frm.Level1ID.value+"&TransactionID="+frm.TransactionID.value;
		self.location="FinStatementConfig.asp?Action=Delete&LevelID="+frm.LevelID.value+"&Level2ID="+frm.RLLevel2ID.value+"&Level1ID="+frm.RLLevel1ID.value+"&TransactionID="+frm.TransactionID.value;
		frm.elements['msgbox'].value = 'Deleting...';}
}


function AutoCon() {
    if (confirm("Do you wish to run the Auto Configuration routine?")) {
       
        self.location = "FinStatementConfig.asp?Action=AutoConfig";
        frm.elements['msgbox'].value = 'Auto Config running...';
    }
}

MouseTip.AddTip("tip1","REPORT","<br>Select the REPORT you would like to view/edit details for");
MouseTip.AddTip("tip2","LEVEL 1 ID","<br>This is READ ONLY");
MouseTip.AddTip("tip3","LEVEL 1 NAME","<br>The Name which appears to Users in the report");
MouseTip.AddTip("tip4","CHANGE LEVEL 1 ID","<br>When a Level 2 item is selected change this (and click Save) to move it to a different Level 1 Item");
MouseTip.AddTip("tip5","SORT ORDER","<br>The ORDER in which the Level 1 Name will appear in the Report");
MouseTip.AddTip("tip6","REPORT GROUP 1 ID","<br>The NUMBER related to the Sub Total the Item will sum to.<br><br>This is displayed as the relevant name in the list below");
MouseTip.AddTip("tip7","REPORT GROUP 2 ID","<br>The NUMBER related to the Grand Total the Item will sum to.<br><br>(This is should be 1 for Net Profit Line 1 and 2 for Net Profit Line 2<br><br>- Balance Sheet items should be 2)");
MouseTip.AddTip("tip8","DRILL DOWN ID","<br>Level 2 screen which (related Level 2 GL) data can be viewed and edited<br><br>2 - Multiple GL Data entry<br><br>3 - Multiple GL Data Read Only<br><br>4 - FTE Salary Calculator/Staff Schedule<br><br>5 - Contractors Schedule<br><br>6 - Consultants Schedule<br><br>7 - Casual Salary Calculator/Staff Schedule<br><br>8 - Multiple GL Level 3 Data entry/Comments<br><br>9 -Travel Schedule");
MouseTip.AddTip("tip9","MATH SIGN","<br>Signage of item data when Summed for the Net Total");
MouseTip.AddTip("tip10","SORT ORDER","<br>The ORDER in which the Level 2 GL Name will appear in the Report, under the Level 1 Name");
MouseTip.AddTip("tip11","GL TYPE","<br>Type of the item (used in double-sided journals)");
MouseTip.AddTip("tip12","LEVEL 2 ID","<br>The GL Account Number of the Level 2 item (GL Account details are maintained on the GL Account Admin screen)");
MouseTip.AddTip("tip13","TRANSACTION ID","<br>Where transactions for the GL Account ID are posted.<br><br>The type of GL Account");
MouseTip.AddTip("tip14","TO ADD NEW LEVEL 1:","<br>Click on the CLEAR button.<br><br>Enter details in the left-hand side of the screen.<br><br>Click on the SAVE button");
MouseTip.AddTip("tip15","TO ADD NEW LEVEL 2:","<br>SELECT a parent Level 1 Item from the List below.<br><br>SELECT an existing Level 2 Item from the List below.<br><br>Change details in the right-hand side of the screen<br><br>Click on the SAVE button");
MouseTip.AddTip("tip16","SAVE","<br>Click to SAVE Level 1 OR Level 2 details");
MouseTip.AddTip("tip17","CLEAR","<br>Click to ADD a NEW Level 1 Item");
MouseTip.AddTip("tip18","DELETE","<br>Click to DELETE the currently selected Level 1 OR Level 2 item");

//-->
</script>
</head>
<body>
<h3>Sub Budget Class and Financial Statement Administration Screen
<%
	Response.write "for the currently selected Budget : <FONT color=Red>" &  Session("BudgetName") & "</FONT> and Version : <FONT color=Red>" & Session("VersionName") & "</FONT></H3>"
%>
<form action="FinStatementConfig.asp?Action=Save" method="POST" id="frm" name="frm">

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<TR>
	<th Height="20px" Width=10% onMouseOver="ShowMouseTip(this,event,'tip1')" onMouseOut="HideMouseTip()" >Report</th><th Width=20% onMouseOver="ShowMouseTip(this,event,'tip1')" onMouseOut="HideMouseTip()"><SELECT style="width:90%" id=ReportID name=ReportID onchange="self.location='FinStatementConfig.asp?ReportID=' + frm.ReportID.value">
	<%

    'objRS.Open "SELECT PLReportID,BusinessAreaCode FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & "",objCon
	objRS.Open "SELECT * FROM tblFinancialStatements WHERE BudgetID = " & Session("BudgetID") & "",objCon

	Do until objRS.EOF
		If objRS("FinancialStatementID") = clng(Session("ReportID")) Then
			strSelected = " SELECTED "
		
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("FinancialStatementID") & """" & strSelected & ">" & objRS("FinancialStatementName") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close

		'Also include all Business Areas as each one has its own PandL report.
		'objRS.Open "SELECT [BusinessAreaID], [BusinessAreaName] FROM tblBusinessArea WHERE BudgetID =" & Session("BudgetID"),objCon
		'Do until objRS.EOF
			'If objRS("BusinessAreaID") = clng(Session("ReportID")) Then
				'strSelected = " SELECTED "
			'Else
				'strSelected = ""
			'End if
				'Response.Write "<option Value=""" & objRS("BusinessAreaID") & """" & strSelected & ">" & objRS("BusinessAreaID") & " - " & objRS("BusinessAreaName") & "</OPTION>"
			'objRS.Movenext
		'Loop
	
		'objRS.Close

	%>
		</SELECT></th><th Width=10%></th>
		<TH Width=10%>Budget</TH>
		<TH Style=Width:20%><select id="BudgetID" name="BudgetID" onchange="self.location='FinStatementConfig.asp?BudgetID=' + frm.BudgetID.value">
		<%	
		Set objRS = Server.CreateObject("ADODB.Recordset")
		objRS.Open "SELECT * FROM tblBudget",objCon

	Do until objRS.EOF
		If objRS("BudgetID") = clng(Session("BudgetID")) Then
			strSelected = " SELECTED "
			Session("FinancialYearID") = objRS("FinancialYearID")
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("BudgetID") & """" & strSelected & ">" & objRS("BudgetName") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close	
			
%>
</select>
</TH>
<TH Width=10%>Version</TH>
<TH Width=20%><select id="Version" name="Version" onchange="self.location='FinStatementConfig.asp?VersionID=' + frm.Version.value">
<%	

    objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & "",objCon
	
	Do until objRS.EOF
		If objRS("VersionID") = clng(Session("VersionID")) Then
			strSelected = " SELECTED "
			Session("VersionName") = objRS("VersionName")
			Session("ColumnLock") = objRS("ColumnLock")
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
	</TR></table>
	
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
		<td colspan=4 Height="20px"></td>
	</tr>
	<tr><th Height="20px" colspan=2 onMouseOver="ShowMouseTip(this,event,'tip14')" onMouseOut="HideMouseTip()" Align="Center">&nbsp;Level 1</th><th colspan=2 onMouseOver="ShowMouseTip(this,event,'tip15')" onMouseOut="HideMouseTip()">Level 2</th></tr>
	<tr>
		<td colspan=4 Height="20px">&nbsp;</td>
	</tr>
	<tr>
		<th Width=15% onMouseOver="ShowMouseTip(this,event,'tip2')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Level 1 ID</th><td Width=35%>&nbsp;<INPUT <%=strBackColour%> type="text" style="width:20%;text-align:center" id="Level1ID" name="Level1ID" maxlength="3" value="<%=intLevel1ID%>"></td>
		<th Width=15% onMouseOver="ShowMouseTip(this,event,'tip12')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Level 2 ID</th><td Width=35%>&nbsp;<SELECT <%=strBackColour2%> id="Level2ID" name="Level2ID">
		<option value="0">Please select...</option>
<%
    objRS.Open "SELECT * FROM tblGLCodes WHERE BudgetID = " & Session("BudgetID") & "",objCon
	
	Do until objRS.EOF
		If objRS("GLCode") = clng(intLevel2ID) Then
			strSelected = " SELECTED "
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("GLCode") & """" & strSelected & ">" & objRS("SegmentValue") & " - " & objRS("GLCodeName") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close
	
%>
</select>
	</td>
	</tr>
	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip3')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Level 1 Name Eng</th><td Width=35%>&nbsp;<INPUT <%=strBackColour%> Style="Width:95%;text-align:left" type="text" id="Level1Name" name="Level1Name" maxlength="50" value="<%=strLevel1Name%>" ></td>
		<th Width=15% onMouseOver="ShowMouseTip(this,event,'tip10')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Sort Order</th><td Width=35%>&nbsp;<INPUT <%=strBackColour2%> Style=Width:"10%" style="text-align:center" type="text" id="SortOrder2" name="SortOrder2" value="<%=intSortOrder2%>"></td>	
	</tr>
    <tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip3')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Level 1 Name Swa</th><td Width=35%>&nbsp;<INPUT <%=strBackColour%> Style="Width:95%;text-align:left" type="text" id="Level1NameL2" name="Level1NameL2" maxlength="200" value="<%=strLevel1NameL2%>" ></td>
		<td colspan="2"></td>	
	</tr>
	
	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip8')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Drill Down Screen</th><td Width=35%>&nbsp;<SELECT id=DrillDownID name=DrillDownID Style=Width:"95%" <%=strBackColour%> >
		<option value="0">N/A</option>
	<%
		
	%> </SELECT></td>
		<th Width=15% onMouseOver="ShowMouseTip(this,event,'tip13')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Payment Ratio</th><td Width=35%>
        &nbsp;<SELECT <%=strBackColour2%> id="PaymentRatio" name="PaymentRatio">
        <option value="0">Please select...</option>
<%
    objRS.Open "SELECT * FROM tblPaymentRatio WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & "",objCon
	
	Do until objRS.EOF
		If objRS("PaymentRatio") = strPaymentRatio Then
			strSelected = " SELECTED "
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("PaymentRatio") & """" & strSelected & ">" & objRS("PaymentRatio") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close
	
%>
</select>
	</td>
	</tr>
	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip9')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Math Sign</th><td Width=35%>&nbsp;<SELECT id=MathSign name=MathSign Style=Width:"10%" <%=strBackColour%> >
	<%
		For x = 1 to 2
			If arrMathSign(x) = cstr(strMathSign) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End If
				Response.Write "<option Value=""" & arrMathSign(x) & """" & strSelected & ">" & arrMathSign(x) & "</OPTION>"
		Next
	%>
		</SELECT></td>
		<th Width=15% onMouseOver="ShowMouseTip(this,event,'tip4')" onMouseOut="HideMouseTip()"><font color=gray Align="Left">Change Level 1 ID</font></th><td Width=35%>&nbsp;<SELECT id=Level1ID2 name=Level1ID2 <%=strBackColour2%> onChange="Level1Change();">
		<option value="0">Please select...</option>
	<%
	objRS.Open "SELECT [Level1ID] FROM tblReportLayoutLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportID = " & Session("ReportID") & " ORDER BY [Level1ID]",objCon
	
	Do until objRS.EOF
		If objRS("Level1ID") = clng(intLevel1ID) Then
			strSelected = " SELECTED "
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("Level1ID") & """" & strSelected & ">" & objRS("Level1ID") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close
	
	%> </SELECT></td>	
	</tr>

	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip5')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Math Denominator</th><td Width=35%>&nbsp;<SELECT id=MathSignage name=MathSignage Style=Width:"10%" <%=strBackColour%> ><!--<INPUT <%=strBackColour%> Style=Width:"100%" type="text" id="MathSignage" name="MathSignage" value="<%=intMathSignage%>" ></td>-->
		
		<%
		For x = -1 to 1
			If x = clng(intMathSignage) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End If
				If x <> 0 Then Response.Write "<option Value=""" & x & """" & strSelected & ">" & x & "</OPTION>"
		Next
		%>
		</SELECT></td>
			<td Colspan="2" rowspan="9"></td>
	</tr>
	<tr><th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip13')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Transaction Type</th><td Width=35%>&nbsp;<SELECT <%=strBackColour%> id="TransactionID" name="TransactionID" >
    <option value="0">Please select...</option>
<%	

    objRS.Open "SELECT * FROM qryTransactionTypes",objCon
	
	Do until objRS.EOF
		If objRS("TransactionType") = cstr(strTransactionType) Then
			strSelected = " SELECTED "
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("TransactionType") & """" & strSelected & ">" & objRS("TransactionType") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close
	
%>
</select>
	</td>

	</tr>
	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip5')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Sort Order</th><td Width=35%>&nbsp;<INPUT <%=strBackColour%> style="text-align:center" Style=Width:"10%" type="text" id="SortOrder" name="SortOrder" value="<%=intSortOrder%>" ></td>
			
	</tr>
	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip5')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Total Grouping 1</th><td Width=35%>&nbsp;<INPUT <%=strBackColour%> style="text-align:center" Style=Width:"10%" type="text" id="TotalGrouping1" name="TotalGrouping1"  value="<%=intTotalGrouping1%>" ></td>
			
	</tr>
	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip5')" onMouseOut="HideMouseTip()" align="Left">&nbsp;Total Grouping 2</th><td Width=35%>&nbsp;<INPUT <%=strBackColour%> style="text-align:center" Style=Width:"10%" type="text" id="TotalGrouping2" name="TotalGrouping2" value="<%=intTotalGrouping2%>" ></td>
			
	</tr>
	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip5')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Total Grouping 3</th><td Width=35%>&nbsp;<INPUT <%=strBackColour%> style="text-align:center" Style=Width:"10%" type="text" id="TotalGrouping3" name="TotalGrouping3" value="<%=intTotalGrouping3%>" ></td>
		
	</tr>
	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip5')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Total Grouping 4</th><td Width=35%>&nbsp;<INPUT <%=strBackColour%> style="text-align:center" Style=Width:"10%" type="text" id="TotalGrouping4" name="TotalGrouping4" value="<%=intTotalGrouping4%>" ></td>
	
	</tr>
	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip5')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Total Grouping 5</th><td Width=35%>&nbsp;<INPUT <%=strBackColour%> style="text-align:center" Style=Width:"10%" type="text" id="TotalGrouping5" name="TotalGrouping5" value="<%=intTotalGrouping5%>" ></td>
		
	</tr>
	<tr>
		<th Height="20px" Width=15% onMouseOver="ShowMouseTip(this,event,'tip5')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Detailed GL Code</th><td Width=35%>&nbsp;<INPUT <%=strBackColour%> style="text-align:center" Style=Width:"10%" type="text" id="DetailedGLCode" name="DetailedGLCode" value="<%=intDetailedGLCode%>" ></td>
		
	</tr>
	
    <tr><th Width=15% onMouseOver="ShowMouseTip(this,event,'tip13')" onMouseOut="HideMouseTip()" Align="Left">&nbsp;Budget Class ID</th><td Width=35%>
        &nbsp;<SELECT <%=strBackColour2%> id="BudgetClassID" name="BudgetClassID">
        <option value="0">Please select...</option>
<%
    objRS.Open "SELECT * FROM tblBudgetClasses WHERE BudgetID = " & Session("BudgetID") & "",objCon
	
	Do until objRS.EOF
		If objRS("BudgetClassID") = cint(intBudgetClassID) Then
			strSelected = " SELECTED "
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS("BudgetClassID") & """" & strSelected & ">" & objRS("BudgetClassID") & " : " & objRS("BudgetClassName") & "</OPTION>"
		objRS.Movenext
	Loop
	
	objRS.Close
	
%>
</select><TD colspan="2"></TD></tr>
    <tr>
		<td Height="20px" colspan=4 height=20><INPUT Type="hidden" id=LevelID Name=LevelID value=<%=intLevelID%> ><INPUT Type="hidden" id=RLLevel2ID Name=RLLevel2ID value=<%=intRLLevel2ID%> ><INPUT Type="hidden" id=RLLevel1ID Name=RLLevel1ID value=<%=intRLLevel1ID%> ></td>
	</tr>
	</table>
<br>
<hr>

<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='FinStatementConfig.asp?Level1ID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
        <td class='locked' Width="100px"><button type="button" tabindex="11" onclick="DeleteData()";><img src="../images/cross.png" alt="" /> Delete </button></td>
        <td class='locked' Width="100px"><button type="button" tabindex="10" onclick="self.location='FinStatementConfig.asp?Level1ID=<%=intLevel1ID%>&Level2ID=1'"><img src="../images/table_add.png" alt="" /> + Level 2</button></td>
		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="400px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>


</form>
<hr>
<DIV ID="overDiv" name="overDiv" STYLE="position: absolute; visibility: hide; z-index: 5; left: 12; top: 361; width: 78; height: 19"></DIV>


<%ListRecords()%>
</table>
</body>
</html>
<%

Public Sub ListRecords()
'Function to dynamically build the list depending on the Level 1 selected, if any.
Dim objPL
Dim strReport
Dim strReportGroup1
Dim objRS2
Dim strLevel1Link
Dim strLevel1Link2
Dim strExtraColumns
Dim strBorder
Dim strSelectColor

'Set objPL = Server.CreateObject("BERTv5.ProfitLoss")
'Set objRS = objAdmin.ReportSchemaLevel1_List(Session("BudgetID"), Session("VersionID"),Session("ReportID"))'1)

objRS.Open "SELECT * FROM tblReportLayoutLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportID = " & Session("ReportID") & " Order By SortOrder",objCon
'Response.Write "SELECT * FROM tblReportLayoutLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportID = " & Session("ReportID") & " Order By SortOrder"    
If intLevel1ID <> 0 And intEdit <> 0 Then

	strExtraColumns = "<TD></TD><TD></TD><TD></TD>"
Else
	strExtraColumns = ""
End If

strBorder = "style=border:none"
strBorder = "Style=border-color:#FFFFFF"

'First write the table headers to the page depending on whether a level 1 Id has been selected or not.
Response.Write "<table WIDTH=100% BORDER=1 CELLSPACING=1 CELLPADDING=1><tr><th></TH><th>Report</th>" & _
		"<th>Level 1 ID</th><th>Level 1 Name</th><th></th>" & strExtraColumns & "<th>Transaction Type</th><th>Total Grouping 1</th>" & _
		"<th>Drill Down ID</th><th>Math Sign</th><th>Math Denominator</th><th>Sort Order</th><th>Detailed GL</th></tr>"

Do Until objRS.EOF
	'Determine the Report Name from the ReportID.
	If objRS("ReportID") = 1 Then
		strReport = "OS"
	End If
		
	If objRS("ReportID") = 2 Then
		strReport = "SBC"
	End If
	
	If objRS("ReportID") = 3 Then
	    strReport = "NA"
	End If

    If objRS("ReportID") = 4 Then
	    strReport = "NA"
	End If

    If objRS("ReportID") = 5 Then
	    strReport = "RR"
	End If
	
	strReportGroup1 = "No"'objPL.ReportGroupingName_Get(objRS("ReportID"),1,objRS("ReportGroup1ID"),clng(Session("BudgetID")))

	'Set the link for the Level1 unless it is currently selected.
	If clng(intLevel1ID) = clng(objRS("Level1ID")) And intEdit <> 0 Then
		strLevel1Link = "<A HREF=""FinStatementConfig.asp?Level1ID=0"">" & objRS("Level1Name") & "</A>"
		strLevel1Link2 = "<A HREF=""FinStatementConfig.asp?Level1ID=0""><img src=""../images/Minus.gif""></A>"
		strSelectColor = " style=background-color:#FFFFCC "
	Else
		strLevel1Link = "<A HREF=""FinStatementConfig.asp?Edit=1&Level1ID=" & objRS("Level1ID") & """>&nbsp;" & objRS("Level1Name") & "</A>"
		strLevel1Link2 = "<A HREF=""FinStatementConfig.asp?Edit=1&Level1ID=" & objRS("Level1ID") & """><img src=""../images/Plus.gif""></A>"
		strSelectColor = ""
	End If
	
	Response.Write "<TR><TD Style=""Text-Align:Center"" " & strSelectColor & "><A HREF=""FinStatementConfig.asp?Edit=" & intEdit & "&ReportID=" & objRS("ReportID") & "&Level1ID=" & objRS("Level1ID") & """><IMG SRC=""../images/edit.jpg""></A></TD>" & _
		"<TD Style=""Text-Align:Center"" " & strSelectColor & "><B>&nbsp;" & strReport & "</B></TD><TD Style=""Text-Align:Center"" " & strSelectColor & ">&nbsp;" & objRS("Level1ID") & "</TD>" & _
		"<TD Style=""Text-Align:Left"" " & strSelectColor & ">" & strLevel1Link & "</TD><TD Style=""Text-Align:Center"" " & strSelectColor & ">" & strLevel1Link2 & "</TD>" & strExtraColumns & "<TD Style=""Text-Align:Center"" " & strSelectColor & ">&nbsp;" & objRS("TransactionType") & "</TD><TD Style=""Text-Align:Center"" " & strSelectColor & ">&nbsp;" & objRS("TotalGrouping1") & "</TD><TD Style=""Text-Align:Center"" " & strSelectColor & ">&nbsp;" & objRS("DrillDown") & "</TD>" & _
		"<TD Style=""Text-Align:Center"" " & strSelectColor & ">&nbsp;" & objRS("MathSign") & "</TD><TD Style=""Text-Align:Center"" " & strSelectColor & ">&nbsp;" & objRS("MathSignage") & "</TD><TD Style=""Text-Align:Center"" " & strSelectColor & ">&nbsp;" & objRS("SortOrder") & "</TD><TD Style=""Text-Align:Center"" " & strSelectColor & ">&nbsp;" & objRS("DetailedGLCode") & "</TD></TR>"
	
	'Response.write " <style> TD { BORDER-RIGHT: 0pt; BORDER-TOP: 0pt; FONT-SIZE: 8pt; BORDER-LEFT: 0pt; BORDER-BOTTOM: 0pt; " & _
	'			"FONT-FAMILY: Tahoma; BACKGROUND-COLOR: white; TEXT-ALIGN: center </style> "
	
	'Write the ReportSchemaLevel2 to the screen list if one has been selected.
	If clng(intLevel1ID) = clng(objRS("Level1ID")) AND intEdit <> 0 Then
		'First write the column headers.
		Response.Write "<tr><th colspan=15></th></td>" & _
			"</tr><tr><td " & strBorder & " ></td><td></td>" & _
			"<td></td><td></td><td></td><th>Level 2 GL</th><th>GL Name</th><th>Payment Ratio</th><td></td><td></td>" & _
			"<td></td><td></td><td></td><td></td>" & _
			"<td></td></tr>"
			
		'Then open the recordset and write the records to the list.
		Set objRS2 = Server.CreateObject("ADODB.Recordset")
		'Set objRS2 = objAdmin.ReportSchemaLevel2_List(Session("BudgetID"), Session("VersionID"),Session("ReportID"))',1)
		
		objRS2.Open "SELECT * FROM qryReportLayoutLevel2List WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportID = " & Session("ReportID") & "",objCon
		
			Do Until objRS2.EOF
				If clng(intLevel1ID) = clng(objRS2("Level1ID"))  Then
				
					strLevel1Link2 = "<A HREF=""FinStatementConfig.asp?Level1ID=0&Level2ID=" & objRS2("GLCode") & """><img src=""../images/Minus.gif""></A>"
					
					If clng(intLevel2ID) = clng(objRS2("GLCode")) Then
						strSelectColor = " style=background-color:#FFFFCC "
					Else
						strSelectColor = ""
					End If
					
					'Write the Level2 values to the list.
					Response.Write "<TR><TD></TD>" & _
						"<TD " & strBorder & " ></TD><TD></TD>" & _
						"<TD Style=""Text-Align:Right"" " & strSelectColor & "><A HREF=""FinStatementConfig.asp?ReportID=" & objRS2("ReportID") & "&Level2ID=" & objRS2("GLCode") & "&Level1ID=" & objRS2("Level1ID") & "&Edit=" & intEdit & "&TransactionID=" & objRS2("ReportLayoutLevel2ID") & """>&nbsp;Edit&nbsp;</A></TD><TD Style=""Text-Align:Center"" " & strSelectColor & ">" & strLevel1Link2 & "</TD><TD" & strSelectColor & " Style=""Text-Align:Center"">&nbsp;" & objRS2("SegmentValue") & "</TD><TD Style=""Text-Align:Left"" " & strSelectColor & "><A HREF=""FinStatementConfig.asp?ReportID=" & objRS2("ReportID") & "&Level2ID=" & objRS2("GLCode") & "&Level1ID=" & objRS2("Level1ID") & "&Edit=" & intEdit & "&TransactionID=" & objRS2("ReportLayoutLevel2ID") & """>&nbsp;" & objRS2("GLCodeName") & "</A></TD><TD Style=""Text-Align:Center"" " & strSelectColor & " >&nbsp;" & objRS2("PaymentRatio") & "</TD><TD></TD><TD></TD><TD></TD>" & _
						"<TD></TD><TD></TD><TD Style=""Text-Align:Center"" " & strSelectColor & ">" & objRS2("SortOrder") & "</TD><TD></TD></TR>"
				End If
			
			objRS2.MoveNext
			
			Loop
		
		Response.Write "<tr><th colspan=15></th>"
		
		objRS2.Close
		
		Set objRS2 = Nothing
	
	End If
	
	objRS.MoveNext
Loop
	
Set objRS = Nothing

End Sub

Public Sub SaveRecord()


 Set ObjCmd = Server.CreateObject("ADODB.Command")
 
    If clng(Request.Form("LevelID")) <> 2 Then

  	    With objCmd
            .CommandType = 4
            .CommandText = "spReportLayoutLevel1Save"
         
            
            .Parameters.Append objCmd.CreateParameter("ReportLayoutLevel1ID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("ReportID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("Level1ID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("Level1Name", adVarChar, adParamInput, 200)
            .Parameters.Append objCmd.CreateParameter("Level1NameL2", adVarChar, adParamInput, 200)
            .Parameters.Append objCmd.CreateParameter("TotalGrouping1", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("TotalGrouping2", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("TotalGrouping3", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("TotalGrouping4", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("TotalGrouping5", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("DrillDown", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("MathSign", adVarChar, adParamInput, 1)
            .Parameters.Append objCmd.CreateParameter("MathSignage", adInteger, adParamInput)                
            .Parameters.Append objCmd.CreateParameter("TransactionType", adVarChar, adParamInput, 4)
            .Parameters.Append objCmd.CreateParameter("DetailedGLCode", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("BudgetClassID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("SortOrder", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
            

	    'If Request.Form("RLLevel1ID") = "" then
	    If not isnumeric(Request.Form("RLLevel1ID")) then
		.Parameters("ReportLayoutLevel1ID") = 0
	    Else
		                                                                                
	    	.Parameters("ReportLayoutLevel1ID") = Request.Form("RLLevel1ID")
	    End if                                                              
	    '.Parameters("ReportLayoutLevel1ID") = Request.Form("RLLevel1ID")
	    .Parameters("BudgetID") = Session("BudgetID")	
            .Parameters("VersionID") = Session("VersionID")
            .Parameters("ReportID") = Session("ReportID")
            
	    'If Request.Form("Level1ID") = "" then
	    If Request.Form("Level1ID") = 0 then

		'Call the procedure to get the last Level1 ID
		Call Level1IDLast
		.Parameters("Level1ID") = intLevel1IDLast + 1
	    Else
		                                                                                
	    	.Parameters("Level1ID") = Request.Form("Level1ID")
	    End if  
            '.Parameters("Level1ID") = Request.Form("Level1ID")
            .Parameters("Level1Name") = Request.Form("Level1Name")
            .Parameters("Level1NameL2") = Request.Form("Level1NameL2")
    
    
            If IsNumeric(Request.Form("TotalGrouping1")) = False Then
                .Parameters("TotalGrouping1") = 0    
            Else
                .Parameters("TotalGrouping1") = Request.Form("TotalGrouping1")
            End If

            If IsNumeric(Request.Form("TotalGrouping2")) = False Then
                .Parameters("TotalGrouping2") = 0    
            Else
                .Parameters("TotalGrouping2") = Request.Form("TotalGrouping2")
            End If
                      
            If IsNumeric(Request.Form("TotalGrouping3")) = False Then
                .Parameters("TotalGrouping3") = 0    
            Else
                .Parameters("TotalGrouping3") = Request.Form("TotalGrouping3")
            End If
    
            If IsNumeric(Request.Form("TotalGrouping4")) = False Then
                .Parameters("TotalGrouping4") = 0    
            Else
                .Parameters("TotalGrouping4") = Request.Form("TotalGrouping4")
            End If
    
            If IsNumeric(Request.Form("TotalGrouping5")) = False Then
                .Parameters("TotalGrouping5") = 0    
            Else
                .Parameters("TotalGrouping5") = Request.Form("TotalGrouping5")
            End If

           
            .Parameters("DrillDown") = 0
           

            .Parameters("MathSign") = Request.Form("MathSign")
            .Parameters("MathSignage") = Request.Form("MathSignage")
            .Parameters("TransactionType") = Request.Form("TransactionID")

            If IsNumeric(Request.Form("DetailedGLCode")) = False Then
                .Parameters("DetailedGLCode") = 0    
            Else
                .Parameters("DetailedGLCode") = Request.Form("DetailedGLCode")
            End If

            .Parameters("BudgetClassID") = Request.Form("BudgetClassID")

            If IsNumeric(Request.Form("SortOrder")) = False Then
                .Parameters("SortOrder") = 0    
            Else
                .Parameters("SortOrder") = Request.Form("SortOrder")
            End If

            .Parameters("UpdatedBy") = Session("UserID")
                       
            .ActiveConnection = objCon
            
        End With
        
        intLevel1ID = Request.Form("Level1ID")
                
    Else
	
	With objCmd
            .CommandType = 4
            .CommandText = "spReportLayoutLevel2Save"         
            
            .Parameters.Append objCmd.CreateParameter("ReportLayoutLevel2ID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("ReportID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("Level1ID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("GLCode", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("PaymentRatio", adVarChar, adParamInput, 50)               
            .Parameters.Append objCmd.CreateParameter("SortOrder", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)            
                                                                                                     '
	        .Parameters("ReportLayoutLevel2ID") = Request.Form("RLLevel2ID")
	        .Parameters("BudgetID") = Session("BudgetID")	
            .Parameters("VersionID") = Session("VersionID")
            .Parameters("ReportID") = Session("ReportID")                
            .Parameters("Level1ID") = Request.Form("Level1ID")
            .Parameters("GLCode") = Request.Form("Level2ID")                
            .Parameters("PaymentRatio") = Request.Form("PaymentRatio")                
            .Parameters("SortOrder") = Request.Form("SortOrder2")
            .Parameters("UpdatedBy") = Session("UserID")
                       
            .ActiveConnection = objCon
                
        End With
		
		intLevel2ID = Request.Form("Level2ID")
		intLevel1ID = Request.Form("Level1ID")
		
		intEdit = 1
	
	End If
        
        
        'response.write "Save=" & Request.Form("RLLevel1ID") & "," & intLevel1IDLast + 1 & "," & Request.Form("Level1Name") & "," & Request.Form("TotalGrouping1") & "," & Request.Form("TotalGrouping2") & "," & Request.Form("TotalGrouping3") & "," & Request.Form("TotalGrouping4") & "," & Request.Form("TotalGrouping5") & _
	'		"," & Request.Form("DrillDownID") & "," & Request.Form("MathSign") & "," & Request.Form("MathSignage") & "," & Request.Form("TransactionID") & "," & Request.Form("DetailedGLCode") & "," & Request.Form("SortOrder")
    objCmd.Execute   
    
     strMessage = "<B>RECORD IS SAVED.</B>"
     strMessageIcon = "<img src=""../images/saveticksmall.jpg"" />"
'Dim objAdmin

'	Set objAdmin = Server.CreateObject("BERTv5.Admin")
	
'		If clng(Request.Form("LevelID")) <> 2 Then
			
'			strMessage = objAdmin.ReportSchemaLevel1_Save(Session("BudgetID"),Session("ReportID"),Session("VersionID"),Request.Form("Level1ID"),Request.Form("Level1Name"),Request.Form("ReportGroup1ID"),Request.Form("ReportGroup2ID"),Request.Form("DrillDownID"),Request.Form("MathSign"),Request.Form("SortOrder"),Request.Form("GLType"),Session("UserID"))
			
'			intLevel1ID = Request.Form("Level1ID")
'		Else
	
'			strMessage = objAdmin.ReportSchemaLevel2_Save(Session("BudgetID"),Session("ReportID"),Session("VersionID"),Request.Form("Level2ID"),Request.Form("Level1ID"),Request.Form("SortOrder2"),Request.Form("TransactionID"),Session("UserID"))
			
'			intLevel2ID = Request.Form("Level2ID")
'			intLevel1ID = Request.Form("Level1ID")
		
'			intEdit = 1
			
'		End If
	
'	Set objAdmin = Nothing
	
End Sub

Public Sub LoadLevel1()

'Get Level 1 details and set variables.
'Set objRS = objAdmin.ReportSchemaLevel1_Get(Session("BudgetID"),Session("ReportID"),Session("VersionID"),intLevel1ID)'1,Session("VersionID"),intLevel1ID)
objRS.Open "SELECT * FROM tblReportLayoutLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportID = " & Session("ReportID") & " AND Level1ID = " & intLevel1ID & "",objCon

	If not objRS.EOF Then
	
	    intRLLevel1ID = objRS("ReportLayoutLevel1ID")
		intLevel1ID = objRS("Level1ID")
		strLevel1Name = objRS("Level1Name")
        strLevel1NameL2 = objRS("Level1NameL2")
		intReportGroup1ID = 0'objRS("ReportGroup1ID")
		intReportGroup2ID = 0'objRS("ReportGroup2ID")
		intDrillDownID = objRS("DrillDown")
		intTotalGrouping1 = objRS("TotalGrouping1")
		intTotalGrouping2 = objRS("TotalGrouping2")
		intTotalGrouping3 = objRS("TotalGrouping3")
		intTotalGrouping4 = objRS("TotalGrouping4")
		intTotalGrouping5 = objRS("TotalGrouping5")
		strMathSign = objRS("MathSign")
		intMathSignage = objRS("MathSignage")
		intSortOrder = objRS("SortOrder")
		strGLType = "N"'objRS("GLType")
		strTransactionType = objRS("TransactionType")
		intDetailedGLCode = 1224'objRS("DetailedGLCode")
		dteDateUpdated = objRS("DateUpdated")
		strUpdatedBy = objRS("UpdatedBy")
		If isNull(objRS("BudgetClassID")) or objRS("BudgetClassID") = "" Then
			intBudgetClassID = 0
		Else
	        	intBudgetClassID = objRS("BudgetClassID")
		End If
 		'Set the value of the onscreen level ID
		intLevelID = 1
		
	Else
	
	    intRLLevel1ID = 0
		'intLevel1ID = 0
		strLevel1Name = ""
        strLevel1NameL2 = ""
		intReportGroup1ID = 0
		intReportGroup2ID = 0
		intDrillDownID = 0
		intTotalGrouping1 = 0
		intTotalGrouping2 = 0
		intTotalGrouping3 = 0
		intTotalGrouping4 = 0
		intTotalGrouping5 = o
		strMathSign = ""
		intMathSignage = 0
		intSortOrder = 0
		strGLType = ""
		dteDateUpdated = ""
		strUpdatedBy = ""
		intDetailedGLCode = 0
		strTransactionType = ""
        intBudgetClassID = 0
		
		'Set the value of the onscreen level ID
		intLevelID = 0
		
	End If
	
	objRS.close

'Set objRS = Nothing

End Sub

Public Sub LoadLevel2()

'Get Level 2 details and set variables.
'Set objRS = objAdmin.ReportSchemaLevel2_Get(Session("BudgetID"),Session("ReportID"),Session("VersionID"),intLevel2ID,intLevel1ID, strTransactionID)'1,Session("VersionID"),intLevel2ID,intLevel1ID, strTransactionID)
If isnull(strTransactionID) or strTransactionID = 0 Then
strTransactionID = 0
    objRS2.Open "SELECT * FROM tblReportLayoutLevel2 WHERE ReportLayoutLevel2ID = " & strTransactionID & "",objCon
	Response.Write "SELECT * FROM tblReportLayoutLevel2 WHERE ReportLayoutLevel2ID = " & strTransactionID & ""
    'objRS.Open "SELECT * FROM tblReportLayoutLevel2 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Level1ID = " & intLevel1ID & "",objCon
else
    objRS2.Open "SELECT * FROM tblReportLayoutLevel2 WHERE ReportLayoutLevel2ID = " & strTransactionID & "",objCon
	Response.Write "SELECT * FROM tblReportLayoutLevel2 WHERE ReportLayoutLevel2ID = " & strTransactionID & ""
    'objRS.Open "SELECT * FROM tblReportLayoutLevel2 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Level1ID = " & intLevel1ID & "",objCon
end if
	If not objRS2.EOF Then
	
	    intRLLevel2ID = objRS2("ReportLayoutLevel2ID")
		intLevel2ID = objRS2("GLCode")
		intSortOrder2 = objRS2("SortOrder")
		intGLCode = objRS2("GLCode")
		strPaymentRatio = objRS2("PaymentRatio")
		dteDateUpdated = objRS2("DateUpdated")
		strUpdatedBy = objRS2("UpdatedBy")
		
		'Set the value of the onscreen level ID
		intLevelID = 2

	Else
	
	    intRLLevel2ID = 0
		intLevel2ID = 0
		'intLevel1ID = 0
		intSortOrder2 = 0
		intGLCode = 0
		strPaymentRatio = ""
		dteDateUpdated = ""
		strUpdatedBy = ""
		
		'Set the value of the onscreen level ID
		intLevelID = 0
	
	End If
	
	'Set the value of the onscreen level ID
	intLevelID = 2

Set objRS2 = Nothing

End Sub

Public Sub DeleteRecord(intLevelID,lngLevel2ID2,lngLevel1ID2,strTransactionID2)

	
		If intLevelID = 1 Then

		'objCon.Execute("DELETE from tblBusinessAreaAccess WHERE BusinessAreaID="&lngBusinessAreaID & " AND UserID=" & lngUserID & " AND BudgetID = " & Session("BudgetID"))
               objCon.Execute("DELETE FROM tblReportLayoutLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportLayoutLevel1ID = " & intLevel1ID & "")
               objCon.Execute("DELETE FROM tblReportLayoutLevel2 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Level1ID = " & intLevel1ID & "")
                strMessage = "DELETED ReportLayoutLevel1ID = " & intLevel1ID 
			    strMessageIcon = "<img src=""../images/saveticksmall.jpg"" />"
			intLevel1ID = 0
			intLevel2ID = 0
		Else
		    objCon.Execute("DELETE FROM tblReportLayoutLevel2 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportLayoutLevel2ID = " & intLevel2ID & "")
		        strMessage = "DELETED ReportLayoutLevel2ID = " & intLevel2ID
			    strMessageIcon = "<img src=""../images/saveticksmall.jpg"" />"
			lngLevel1ID2 = 0
			intLevel1ID = lngLevel1ID2
			strTransactionID = 0
			intEdit = 1
		End If

End Sub

Public Sub AutoConfig()

 Set ObjCmd = Server.CreateObject("ADODB.Command")

     With objCmd
              
            .CommandType = 4
            .CommandText = "spConfigureFinancialStatements"             

            .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
            .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)                                                                                                    '
	   
	        .Parameters("BudgetID") = Session("BudgetID")	
            .Parameters("VersionID") = Session("VersionID")
            .Parameters("UpdatedBy") = Session("UserID")
                       
            .ActiveConnection = objCon
                
        End With
        
         objCmd.Execute   
   
End Sub


Public Sub Level1IDLast()
'Get the last record number of Level1ID for the list.
Dim objRS3

Set objRS3 = Server.CreateObject("ADODB.Recordset")

'Set objRS = objAdmin.ReportSchemaLevel1_ListByLevel(Session("BudgetID"), Session("VersionID"),Session("ReportID"))'1)
objRS3.Open "SELECT [Level1ID] FROM tblReportLayoutLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportID = " & Session("ReportID") & " ORDER BY [Level1ID] DESC",objCon
	
	If Not objRS3.EOF Then
		'objRS3.MoveLast
		intLevel1IDLast = objRS3("Level1ID")
	End If

objRS3.Close

Set objRS3 = Nothing

End Sub

%>