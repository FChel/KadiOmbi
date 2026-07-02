<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	Positions Admin Screen
'Author:		Michael Giacomin
'Date:			April 2016

'Declare default variables

If IsEmpty(Session("EstablishmentID")) Then Session("EstablishmentID") = 0

Dim objCon
Dim objCmd
Dim objRS
Dim objRS1
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon
Dim strMessageColour
Dim arrYesNo(2)

arrYesNo(1) = "Y"
arrYesNo(2) = "N"

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim intPositionID
Dim intCostCentreID
Dim strEstablishmentName
Dim strEstablishmentNo
Dim strEstablishmentDesc
Dim intStartPeriod
Dim intEndPeriod
Dim dblRate
Dim strActive
Dim strAttribute1
Dim dteDateCeased
Dim strCesationReason
Dim strNewPosition
Dim strTransactionType

Dim arrPeriods(12)

	arrPeriods(1) = "Jul"
	arrPeriods(2) = "Aug"
    arrPeriods(3) = "Sep"
    arrPeriods(4) = "Oct"
    arrPeriods(5) = "Nov"
    arrPeriods(6) = "Dec"
    arrPeriods(7) = "Jan"
    arrPeriods(8) = "Feb"
    arrPeriods(9) = "Mar"
    arrPeriods(10) = "Apr"
    arrPeriods(11) = "May"
    arrPeriods(12) = "Jun"
	
'3. Capture Querystring variables
If Session("PositionID") = "" Then Session("PositionID") = 0

    If Not IsEmpty(Request.QueryString("PositionID")) Then
	   intPositionID = Request.QueryString("PositionID")
    End If

    If Not IsEmpty(Request.QueryString("CostCentreID")) Then
	   intCostcentreID = Request.QueryString("CostCentreID")
    End If

    If Not IsEmpty(Request.QueryString("EstablishmentID")) Then
	   Session("EstablishmentID") = Request.QueryString("EstablishmentID")
    End If
		
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Execute Recalculate procedure 	
	If Request.QueryString("Action") = "recalculate" Then
		RecalcStaff()
	End If
	
	'Execute Delete 	
	If Request.QueryString("Action") = "Delete" Then
		DeleteRecord Request.QueryString("EstablishmentName"),Request.QueryString("CostCentreID")
	End If
	
	'Load page details
	LoadDetails()
		
%>

<html>
<head>
<title></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
<script type="text/javascript" src="../formChek.js"></script>
<script src="../calender.js">
	</script>
<script type="text/javascript" language="javascript">
<!--
   function SaveData()
    {
        var varSubmit = true
        var varAlert =""  
        	    
	    if(isWhitespace(frm.EstablishmentName.value))
        {            
		    varAlert += "Please enter the Establishment Name. \n \n";
		    document.getElementById('EstablishmentName').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }
		else document.getElementById('EstablishmentName').style.backgroundColor = "ffffff";

		if (frm.PositionID.value == 0) {
		    varAlert += "Please select a Position. \n \n";
		    document.getElementById('PositionID').style.backgroundColor = "ff8080";
		    varSubmit = false;
		}
		else document.getElementById('PositionID').style.backgroundColor = "ffffff";

	    if(frm.CostCentreID.value == 0 )
	    {
		    varAlert += "Please select a Cost Centre. \n \n";
		    document.getElementById('CostCentreID').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }
		else document.getElementById('CostCentreID').style.backgroundColor = "ffffff";

		if (frm.StartPeriod.value == 0) {
		    varAlert += "Please select a Start Period. \n \n";
		    document.getElementById('StartPeriod').style.backgroundColor = "ff8080";
		    varSubmit = false;
		}
		else document.getElementById('StartPeriod').style.backgroundColor = "ffffff";

		if (frm.EndPeriod.value == 0) {
		    varAlert += "Please select a End Period. \n \n";
		    document.getElementById('EndPeriod').style.backgroundColor = "ff8080";
		    varSubmit = false;
		}
		else document.getElementById('EndPeriod').style.backgroundColor = "ffffff";

		
         	
	  if(varSubmit == true)
	  {
	        frm.submit();
	  }
	  else
	  {
	    window.alert ("" + varAlert);	    
	  }
  }
  
 
function StaffRecalc(varBud, varVer)
 {
 	if(window.confirm('Would you like to Recalculate ALL existing Staff values? \n \n For the currently selected Budget ' + varBud + ' and Version ' + varVer + '.')==true){
 	
 		document.getElementById('Progress').style.display = "inline";
        self.location='Establishments,asp?Action=recalculate';
	}
}
function DeleteData(){
	if( confirm("Delete " + frm.EstablishmentName.value + "?") )
	{
		self.location="Establishments.asp?Action=Delete&EstablishmentName="+frm.EstablishmentName.value+"&CostCentreID="+frm.CostCentreID.value;
		frm.elements['msgbox'].value = 'Deleting...';}
	}
//-->
</script>
</head>
<body>
<form action="Establishments.asp?Action=Save" method="POST" id="frm" name="frm">
<h3>Positions</h3>
<table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
	<tr>
		<th Style="Width:20%; text-align:left;">&nbsp;Position No</th>
	    <td Style="Width:30%; text-align:left;">&nbsp;<input style="text-align:left;width:90%" id="EstablishmentNo" name="EstablishmentNo" maxlength="50" tabindex="1" value="<%=strEstablishmentNo%>"></td>
		<th Style="Width:20%; text-align:left;">&nbsp;Cesation Reason</th><td Style="Width:30%;">
	     <select Style="Width:90%" tabindex="7" id="CesationReason" name="CesationReason" ><OPTION Value="N/A">Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblCesationReasons WHERE Active = 'Y'",objCon
    		
		    Do until objRS.EOF
			    If objRS("CesationReason") = strCesationReason Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CesationReason") & """" & strSelected & ">" & objRS("CesationReason") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select></td>
	</tr>
	<tr>
		<th Style="text-align:left;">&nbsp;Position Name</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="EstablishmentName" name="EstablishmentName" maxlength="50" tabindex="2" value="<%=strEstablishmentName%>"></td>
		<th style="text-align:left; height:20px;">&nbsp;Cesation Date</th><td>
        <input style="text-align:left" style="width:50%" readonly id="CesationDate" name="CesationDate" maxlength="50" TABINDEX="7" value="<%=dteDateCeased%>">&nbsp; &nbsp;<a href="javascript:CesationDate.popup();"><img src="../images/cal.gif" width="16" height="16" border="0" alt="Click Here to pick up the date"></a>
		&nbsp;<a href="javascript:clearField('DateApproved');"><img src="../Images/rubber.gif" border="0" alt="Click here to clear the date field"></a></td>
	</tr>
    <tr>
		<th Style="text-align:left;">&nbsp;Position Type</th><td>
	     <select Style="Width:90%" tabindex="2" id="PositionID" name="PositionID" ><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblPosition WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Active = 'Y'",objCon
    		
		    Do until objRS.EOF
			    If clng(objRS("PositionID")) = clng(intPositionID) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("PositionID") & """" & strSelected & ">" & objRS("PositionName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select></td>
			<th Style="Width:20%; text-align:left;">&nbsp;New Position</th><td Style="Width:30%;">
	     <select Style="Width:90%" tabindex="7" id="NewPosition" name="NewPosition" >
	       <%
				For x = 1 to 2
					If arrYesNo(x) = cstr(strNewPosition) Then
						strSelected = " SELECTED "
					Else
						strSelected = ""
					End If
					Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
				Next
                        %></select></td>
	</tr>
	<tr>
		<th Style="text-align:left;">&nbsp;Cost Centre</th><td>
	     <select Style="Width:90%" tabindex="3" id="CostCentreID" name="CostCentreID" ><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM qryCostCentresByBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND CostObjectTypeID = 2 AND Active = 'Y'",objCon
    		
		    Do until objRS.EOF
			    If clng(objRS("CostCentreID")) = clng(Session("CostCentreID")) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CostCentreID") & """" & strSelected & ">" & objRS("ProgramCode") & " - " & objRS("CostCentreName") &"</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select></td>
		<td colspan="2">&nbsp;</td>
	</tr>
		
	<tr>
		<th Style="text-align:left;">&nbsp;Start Period</th>		
		<td><select Style="Width:40%" tabindex="4" id="StartPeriod" name="StartPeriod"><option Value=0>Please Select....</option>
		<%
		For x = 1 to 12
			If intStartPeriod = x Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & x & """" & strSelected & ">" & arrPeriods(x) & "</OPTION>"
		Next
		%>
		</select> </td>
		<td colspan="2">&nbsp;</td>
	</tr>
    <tr>
		<th Style="text-align:left;">&nbsp;End Period</th>		
		<td><select Style="Width:40%" tabindex="5" id="EndPeriod" name="EndPeriod"><option Value=0>Please Select....</option>
		<%
		For x = 1 to 12
			If intEndPeriod = x Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & x & """" & strSelected & ">" & arrPeriods(x) & "</OPTION>"
		Next
		%>
		</select> </td>
		<td colspan="2">&nbsp;</td>
	</tr>
	    <tr>
		<th Style="text-align:left;">&nbsp;Gratuity</th>		
		<td><select Style="Width:40%" tabindex="6" id="Attribute1" name="Attribute1"><option Value=0>Please Select....</option>
	 <%
		For x = 1 to 2
			If arrYesNo(x) = cstr(strAttribute1) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End If
				Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		Next
                        %>
		</select> </td>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="4" align="left">&nbsp;</td>
	</tr>
	
</table>

<script LANGUAGE="javascript">
   
    var CesationDate;
    CesationDate = new calendar(document.forms(0).elements['CesationDate']);
    
</script>
<br/>
<div class="buttons">
<TABLE Width="1500px" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<TR>

    <td Width="100px"><button type="button" onclick="self.location='../ProfitLoss/SalaryRates.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
    <td Width="100px"><button type="button" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td> 
    <td Width="150px"><button type="button" onclick="self.location='Establishments.asp?EstablishmentID=0'" )""><img src="../images/add.png" alt="" /> Clear/Add New </button></td>
    <td Width="100px"><button type="button" onclick="DeleteData()";><img src="../images/delete.png" alt="" /> Delete </button></td>
    <!--<td Width="150px"><button type="button" onClick="StaffRecalc('<%=Session("BudgetName")%>','<%=Session("VersionName")%>');" )"" Title="Click to recalculate ALL staff data with Super Rates below (if changed)"><img src="../images/calculator.png" alt="" /> Recalculate </button></td>-->
    <TD Width="200px"><span id="Progress" style="display:none"><img src="../images/progress.gif">  &nbsp;&nbsp;&nbsp; <b>Recalculating Staff...</b></span></TD>
    <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
    <TD class='locked' align="left" Width="800x" style="BORDER-RIGHT:0px"><INPUT style="Align:Left; font-weight:Bold; width:100%; text-align:left; color:<%=strMessageColour%>;" type="text" id="msgbox" name="msgbox" value="<%=strMessage%>"></TD>
</TR>
</TABLE>
</div>

<hr>
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr><td colspan="10">&nbsp;<span style="color:gray; font-weight:bold;">Note:</span><span style="color:gray;"> Active Positions listed here will appear in the Salary Rates screen for each Cost Centre.</span></td></tr>
	<tr>
        <th Style="Width:7.5%;">Position No</th>
		<th Style="Width:15%;">Position Name</th>
		<th Style="Width:12.5%;">Position Type   </th>
		<th Style="Width:5%;">PF No</th>
		<th Style="Width:15%;">Employee Name</th>		
        <th Style="Width:17.5%;">Cost Centre </th>
        <th Style="Width:5%;">Start Period </th>
        <th Style="Width:5%;">End Period </th>
		<th Style="Width:10%;">Reason </th>
		<th Style="Width:7.5%;">Ceased </th>
	</tr>
<%
    objRS.Open "SELECT * FROM qryEstablishments WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " Order By EstablishmentName Asc",objCon

	Do until objRS.eof			
   	   Response.Write "<TR><TD>" & objRS("PositionNo") & "</TD><TD><A Target=""_self"" HREF=""Establishments.asp?EstablishmentID=" & objRS("EstablishmentID") & """>&nbsp;" & objRS("EstablishmentName") & "</TD><TD style=""text-align:left"">&nbsp;" & objRS("PositionName") & "</TD><TD>" & objRS("EmployeeID") & "</TD><TD>" & objRS("StaffClassificationDesc") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("CostCentreName") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("StartPeriod") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("EndPeriod") & "</TD><TD>" & objRS("CesationReason") & "</TD><TD>" & objRS("CesationDate") & "</TD></TR>"
           '' objRS1.Open "SELECT * FROM  qryEstablishmentFunding WHERE BudgetID = " & clng (Session("BudgetID")) & " AND VersionID = " & Session("VersionID") & " AND EstablishmentID = " & objRS("EstablishmentID") & "",objCon
           '' Do Until objRS1.EOF
                 '' Response.Write "<TR><TD Colspan=""2""></TD><TH><A Target=""_self"" HREF=""EstablishmentFunding.asp?EstablishmentID=" & objRS("EstablishmentID") & """>&nbsp;Funding Source:</A></TH><TD><A Target=""_self"" HREF=""EstablishmentFunding.asp?EstablishmentID=" & objRS("EstablishmentID") & "&EstablishmentFundingID=" & objRS1("EstablishmentFundingID") & """>&nbsp;" & objRS1("FundCode") & " - " & objRS1("FundName") & "</TD><TD style=""text-align:right"">&nbsp;" & objRS1("Weighting") * 100 & "%</TD><Th Colspan=""5""></Th></TR>"
               '' objRS1.Movenext
           '' Loop
           '' objRS1.Close
       objRS.movenext
	Loop
		
	objRS.Close
	
%>

</table>
</body>

</html>

<% 

Sub LoadDetails()

       'Description:	Loads Position details into page if applicable.
		objRS.Open "SELECT * FROM tblEstablishments WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND EstablishmentID = " & Session("EstablishmentID") & "",objCon

			If Not objRS.EOF Then
                intPositionID = objRS("PositionID")
                intCostcentreID = objRS("CostCentreID")
				strEstablishmentNo = objRS("PositionNo")
				strEstablishmentName = objRS("EstablishmentName")
				strEstablishmentDesc = objRS("EstablishmentDesc")
                intEndPeriod = objRS("EndPeriod")
                intStartPeriod = objRS("StartPeriod")
				strAttribute1 = objRS("Attribute1")
				strCesationReason = objRS("CesationReason")
				dteDateCeased = objRS("CesationDate")
				strTransactionType = objRS("TransactionType")
				If strTransactionType = "NSTA" Then
					strNewPosition = "Y"
				Else
					strNewPosition = "N"
				End If
			
			
    		Else

			  	strEstablishmentNo = ""
				strEstablishmentName = ""
				strEstablishmentDesc = ""
                intPositionID = 0
                intCostcentreID = 0
				strAttribute1 = "Y"
				strCesationReason = ""
				dteDateCeased = ""
				strTransactionType = ""
				strNewPosition = "Y"

           End If

		objRS.Close
	
End Sub

Sub SaveDetails()

		 With objCmd
                .CommandType = 4
                .CommandText = "spEstablishmentSave"
                
                .Parameters.Append objCmd.CreateParameter("EstablishmentID", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("PositionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("CostCentreID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("EstablishmentName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("StartPeriod", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("EndPeriod", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("EstablishmentDesc", adLongVarChar, adParamInput,-1)
				.Parameters.Append objCmd.CreateParameter("TransactionType", adVarChar, adParamInput, 50)
				.Parameters.Append objCmd.CreateParameter("Attribute1", adVarChar, adParamInput, 50)
				.Parameters.Append objCmd.CreateParameter("Attribute2", adVarChar, adParamInput, 50)
				.Parameters.Append objCmd.CreateParameter("Attribute3", adVarChar, adParamInput, 50)
				.Parameters.Append objCmd.CreateParameter("PositionNo", adVarChar, adParamInput, 50)
				.Parameters.Append objCmd.CreateParameter("CesationReason", adVarChar, adParamInput, 50)
				.Parameters.Append objCmd.CreateParameter("CesationDate", adDate, adParamInput)
			    .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
          	
                .Parameters("EstablishmentID") = clng(Session("EstablishmentID"))
				.Parameters("BudgetID") = clng(Session("BudgetID"))
   				.Parameters("VersionID") = clng(Session("VersionID"))
   				.Parameters("PositionID") = clng(Request.Form("PositionID"))
                .Parameters("CostCentreID") = clng(Request.Form("CostCentreID"))
				.Parameters("EstablishmentName") = Request.Form("EstablishmentName")	
                .Parameters("StartPeriod") = Request.Form("StartPeriod")
                .Parameters("EndPeriod") = Request.Form("EndPeriod")
				.Parameters("EstablishmentDesc") = ""'Request.Form("EstablishmentDesc")	
				.Parameters("TransactionType") = Request.Form("NewPosition")
				.Parameters("Attribute1") = Request.Form("Attribute1") 
				.Parameters("Attribute2") = "" 
				.Parameters("Attribute3") = ""  	  
				.Parameters("PositionNo") = Request.Form("EstablishmentNo")
				.Parameters("CesationReason") = Request.Form("CesationReason")
				If IsNull(Request.Form("CesationDate")) Then
					.Parameters("CesationDate") = Request.Form("CesationDate")
				Else
					.Parameters("CesationDate") = Null
				End If
                .Parameters("UpdatedBy") = Session("UserID")

                'Response.Write clng(Session("EstablishmentID")) & " A "
				'Response.Write clng(Session("BudgetID")) & " B "
   				'Response.Write clng(Session("VersionID")) & " C "
   				'Response.Write clng(Request.Form("PositionID")) & " D "
              '  Response.Write clng(Request.Form("CostCentreID")) & " E "
			'	Response.Write Request.Form("EstablishmentName") & " F "	 
             '   Response.Write clng(Request.Form("StartPeriod")) & " G "
             '   Response.Write clng(Request.Form("EndPeriod")) & " H "
			'	Response.Write Request.Form("EstablishmentDesc") & " I "	          
             '   Response.Write Session("UserID")
               '' Response.Write Request.Form("NewPosition")    
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
	   'Return the result of the Save Function.
            strMessage = "RECORD SAVED."
     	    strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
            strMessageColour = "Black"
	
End Sub	


Sub RecalcStaff()


		 With objCmd
                .CommandType = 4
                .CommandText = "spStaffDataSalaryUpdate"
                
				.Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
          	
				.Parameters("BudgetID") = clng(Session("BudgetID"))
   				.Parameters("VersionID") = clng(Session("VersionID"))
				
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
			'Return the result of the Save Function.
     	    strMessage = "Staff Values have been recalculated!"
	
End Sub	

Public Sub DeleteRecord(strPositionName, intCCID)
'Procedure to delete Salary Classification records if no values exist against it.

Dim lngEmployeeID

	'First get the EmployeeID of the related StaffingClassification Record
    objRS.Open "SELECT [EmployeeID] FROM tblStaffingClassifications With(NoLock) WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND EstablishmentID = " & Session("EstablishmentID") & "" ,objCon,adOpenStatic,adLockReadOnly
                	
        If objRS.EOF Then
        
            lngEmployeeID = 0
        Else
            lngEmployeeID = objRS("EmployeeID")
   
        End If
                	
    objRS.Close

	'Delete tblBudgetData records'
	objCon.Execute "DELETE tblBudgetData WHERE TransactionType In ('ESTA','NSTA') AND BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND SatelliteRecordID In (SELECT StaffDataID FROM tblStaffData WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND EmployeeID In (SELECT StaffingClassificationID FROM tblStaffingClassifications WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND EstablishmentID = " & Session("EstablishmentID") & ")) "

	
	'Delete the Salary Data Record
	objCon.Execute("DELETE FROM tblStaffData WHERE [BudgetID] = " & Session("BudgetID") & " AND EmployeeID = " & lngEmployeeID & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & intCCID & "")

	'Delete the Salary Classifications Record
	objCon.Execute("DELETE FROM tblStaffingClassifications WHERE [BudgetID] = " & Session("BudgetID") & " AND EstablishmentID = " & Session("EstablishmentID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & intCCID & "")

	'Delete the Establishment record
	objCon.Execute("DELETE FROM tblEstablishments WHERE [BudgetID] = " & Session("BudgetID") & " AND EstablishmentID = " & Session("EstablishmentID") & " AND VersionID = " & Session("VersionID") & "")
    	
    Response.Write "DELETE tblBudgetData WHERE TransactionType In ('ESTA','NSTA') AND BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND SatelliteRecordID In (SELECT StaffDataID FROM tblStaffData WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND EmployeeID In (SELECT StaffingClassificationID FROM tblStaffingClassifications WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND EstablishmentID = " & Session("EstablishmentID") & ")) "
End Sub

Set objRS = Nothing
Set objCon = Nothing


%>
