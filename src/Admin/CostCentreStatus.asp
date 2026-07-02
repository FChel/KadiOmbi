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
Dim objRS1
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon
Dim strColour

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

If Validate_Access(Session("UserTypeID"),Session("CurrentPage")) = "N" Then
     Response.Redirect "../AccessDenied.asp"
End If	

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngBudgetID
Dim lngCostCentreID
Dim lngStatusID
Dim strBudgetName
Dim intFinancialYearID
Dim intDefaultVersionID
Dim strBalanceSheet
Dim lngCashFlowGLCode
Dim lngBadDebtGLCode
Dim lngPrePaymentGLCode
Dim lngInvestmentGLCode
Dim lngLoanGLCode
Dim lngAPGLCode
Dim lngARGLCode
Dim intVariancePercentage
Dim strComments
Dim strActive
Dim lngInUseBy

Dim dblRevenue
Dim dblExpenditure
Dim dblNetTotal

Dim dblApprovedRev
Dim dblApprovedExp
Dim dblApprovedNet

Dim dblTotalRev
Dim dblTotalExp
Dim dblTotalNet

Dim dblOpexRev
Dim dblOpexExp
Dim dblOpexNet

Dim dblCapxRev
Dim dblCapxExp
Dim dblCapxNet

'Declare and set default arrays

Dim arrYesNo(2)
	
	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"
	
	'3. Capture Querystring variables
	
	If Not IsEmpty(Request.QueryString("CostCentreID")) Then
		
		lngCostCentreID = Request.QueryString("CostCentreID")
						
	End If
	
	If Not IsEmpty(Request.QueryString("StatusID")) Then
		
		lngStatusID = Request.QueryString("StatusID")
       
    End If

    If Not IsEmpty(Request.QueryString("InUseBy")) Then
		
		lngInUseBy = Request.QueryString("InUseBy")
       
    End If
	
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
		
	End If
    
Dim arrStatus(7)
Dim arrStatusName(7)

arrStatusName(0) = "ALL"
arrStatusName(1) = "OPEN"
arrStatusName(2) = "COMPLETED"
arrStatusName(3) = "REJECTED"
arrStatusName(4) = "APPROVED"
arrStatusName(5) = "CLOSED"
arrStatusName(6) = "ERROR"
arrStatusName(7) = "IN USE"

arrStatus(0) = "<IMG SRC='../images/delete.png'"
arrStatus(1) = "<IMG SRC='../images/open.png'"
arrStatus(2) = "<IMG SRC='../images/ready.gif'"
arrStatus(3) = "<IMG SRC='../images/cross.png'" 
arrStatus(4) = "<IMG SRC='../images/tick.png'"	
arrStatus(5) = "<IMG SRC='../images/Closed.png'"
arrStatus(6) = "<IMG SRC='../images/bug.png'"   
arrStatus(7) = "<IMG SRC='../images/wrench.png'"
    			
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
	<script src="../formChek.js">
	</script>
	<script type="text/javascript" language="javascript">

	    function SaveData() {
	        var varSubmit = true
	        var varAlert = ""

	        if ((isNonnegativeInteger(frm.CostCentreID.value) == false) || (frm.CostCentreID.value == 0)) {
	            varAlert += "Please select Program or Sub Program. \n \n";
	            document.getElementById('CostCentreID').style.backgroundColor = "ff8080";
	            varSubmit = false;
	        }
	        else document.getElementById('CostCentreID').style.backgroundColor = "ffffff";


	        if (frm.StatusID.value == 0) {
	            varAlert += "Please select a status. \n \n";
	            document.getElementById('StatusID').style.backgroundColor = "ff8080";
	            varSubmit = false;
	        }
	        else document.getElementById('StatusID').style.backgroundColor = "ffffff";

	        if (varSubmit == true) {
	            document.getElementById('Progress').style.display = "inline"; 
	            frm.submit();
	        }
	        else {
	            window.alert("" + varAlert);
	        }

	    }  

	 

</script>
	<script src="../ButtonRollOver.js">
	</script>
     <style> 
    <!--
div#tbl-container {
width: 100%;
height: 60%;
overflow: auto;
}

table {
table-layout: fixed;
border-collapse: inherit;
}

div#tbl-container table th {

}

thead th, thead th.locked	{

position:relative;
cursor: default; 
border-right: 1px solid silver;

}
	
thead th {
top: expression(document.getElementById("tbl-container").scrollTop-2); /*IE5+ only*/
z-index: 20;


}

thead th.locked {z-index: 30;}

td.locked,  th.locked{
left: expression(document.getElementById("tbl-container").scrollLeft); /*IE5+ only*/
position: relative;
z-index: 10;
border-right: 1px solid silver;


}
td.locked_left, th.locked_left {
    left            : expression(document.getElementById('tbl-container').scrollLeft);
    z-index         : 1;
    border-right: 1px solid silver;
	border-left: 1px solid silver;
	position: relative;
	align: center;

  }

    -->
   </style>
</head>

<body>
<div id='tbl-container'>
<form action="CostCentreStatus.asp?Action=Save" method="POST" id="frm" name="frm">
<h3><%=Session("CCName")%>Status : <FONT Color="Red"><%= arrStatusName(Session("CCStatusID")) %></FONT></h3>
<table Style="width:100%" border="1" cellspacing="1" cellpadding="1">

	
	<th align="left">&nbsp;<%=Session("CCName") %></th>
		<td style="background-color:FFFFFF;">
		    <select Style="Width:100%; background-color:FFFFFF;" tabindex="20" id="CostCentreID" name="CostCentreID"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM qryCostCentresByBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " Order By ProgramCode",objCon
    		
		    Do until objRS.EOF
			    If objRS("CostCentreID") = clng(lngCostCentreID) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("CostCentreID") & """" & strSelected & ">" & objRS("ProgramCode") & " - " & objRS("CostCentreName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
	    <td colspan="2" rowspan="6">&nbsp;<iframe id="Iframe1" name="framecontent" src="../Reports/BAStatusPieSmall.asp" Width="100%" frameborder="0" height="170px"></iframe></td>
	
	<tr>
	
		<th align="Left">&nbsp;Status</th>		
		<td style="background-color:FFFFFF;"><select Style="Width:100%;  background-color:FFFFFF;" tabindex="6" id="StatusID" name="StatusID"><option Value="0">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM qryStatus WHERE UserTypeID = " & Session("UserTypeID") & "",objCon
		
		Do until objRS.EOF
			If objRS("StatusID") = clng(lngStatusID) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("StatusID") & """" & strSelected & ">" & objRS("StatusName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>
		
	</tr>
	<tr>
	
		<th align="Left">&nbsp;In Use By</th>		
		<td style="background-color:FFFFFF;"><select Style="Width:100%;  background-color:FFFFFF;" tabindex="7" id="InUseBy" name="InUseBy"><option Value="0">No One</option>
		<%
		objRS.Open "SELECT * FROM qryCostCentreAccess WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID =  " & Session("BusinessAreaID") & " AND CostCentreID = " & clng(lngCostCentreID) & "",objCon
		
		Do until objRS.EOF
			If objRS("UserID") = clng(lngInUseBy) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("UserID") & """" & strSelected & ">" & objRS("FName") & " "  & objRS("LName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>
	
	</tr>
	
    <tr>
		<td style="height:20px" colspan="2" rowspan="4" align="left">&nbsp;</td>
	</tr>
	<tr>
		
	</tr>
	<tr>
		
	</tr>
	
</table>
<br />
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="10%" style="border-right:0px"><button type="button" tabindex="8" onclick="parent.location='../BudgetFrameset.asp'"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="10%" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="10%" style="border-right:0px"><button type="button" tabindex="17" onclick="window.open('CostCentreStatusExcel.asp')"><img src="../images/page_excel.png" alt="" /> Excel </button></td>
		<td Width="10%"><span id="Progress" style="display:none"><img src=../Images/progress.gif />  &nbsp;&nbsp;&nbsp; <b><FONT Face="Arial"></FONT></b></span></td>
		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="600px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr>
<table Style="width:100%" border="1" cellspacing="1" cellpadding="1">
    <tr>
        <td colspan="6">&nbsp;</td>
    </tr>
	<tr>
		<th Style="Height:20px; Width:15%"><%=Session("CCName") %></th>
		<th Width="20%"><%=Session("CCName") %> Name</th>
		<th Width="15%">Type</th>
		<th Width="15%">Status</th>
        <th Width="20%">In Use By</th>
        <th Width="15%">Last Edited</th>
	</tr>
	
<%

If Session("CCStatusID") = 0 Then

    objRS.Open "SELECT * FROM qryCostCentreApprovals WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " Order By ParentCostCentreID,CostCentreID ASC",objCon
	'Response.Write "SELECT * FROM qryCostCentreApprovals WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " Order By ParentCostCentreID,CostCentreID ASC"
Else
     
    objRS.Open "SELECT * FROM qryCostCentreApprovals WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND StatusID = " & Session("CCStatusID") & " Order By ParentCostCentreID,CostCentreID ASC",objCon
	'Response.Write "SELECT * FROM qryCostCentreApprovals WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND StatusID = " & Session("CCStatusID") & " Order By ParentCostCentreID,CostCentreID ASC"

End If
		Do until objRS.EOF		
			Response.Write "<TR><TD style=""text-align:center""><B><A Target=""_self"" HREF=""CostCentreStatus.asp?InUseBy=" & objRS("InUseBy") & "&CostCentreID=" & objRS("CostCentreID") & "&StatusID= " & objRS("StatusID") & """>" & objRS("ProgramCode") & "</A></B></TD><TD style=""text-align:left"">&nbsp;" & objRS("CostCentreName") & "</TD><TD style=""text-align:left"">&nbsp;" & objRS("CostObjectTypeName") & "</TD><TD style=""text-align:center""><B>" & objRS("StatusName") & "</B>&nbsp;&nbsp;&nbsp;" & arrStatus(objRS("StatusID")) & "</TD><TD style=""text-align:center""><B>" & objRS("InUseByName") & "</B></TD><TD style=""text-align:center""><B>" & objRS("LastEdited") & "</B></TD></TR>"
			objRS.movenext
		Loop
			
	objRS.Close
	

%>
</table>
</body>

</html>

<%

Sub SaveDetails()

Dim strCCStatus

      If Session("StatusID") = 1 Then
      
		    With objCmd
		    
                .CommandType = 4
                .CommandText = "spCostCentreStatusSave"
                .Parameters.Append objCmd.CreateParameter("CostCentreID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BusinessAreaID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("StatusID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("InUseBy", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
                .Parameters.Append objCmd.CreateParameter("Response", adVarChar, adParamOutput, 100)

     			.Parameters("CostCentreID") = clng(Request.Form("CostCentreID"))		
				.Parameters("BudgetID") = clng(Session("BudgetID"))			
                .Parameters("VersionID") = clng(Session("VersionID"))
                .Parameters("BusinessAreaID") = clng(Session("BusinessAreaID"))
                .Parameters("StatusID") = clng(Request.Form("StatusID"))       
                .Parameters("InUseBy") = clng(Request.Form("InUseBy"))       
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                       
            ' You cannot Approve Program if all Sub Programs not equal to completed.
            'strCCStatus = ValidateCCStatus(Request.Form("CostCentreID"),Request.Form("StatusID"))
            
                'If strCCStatus = "Y" Then
                    objCmd.Execute
                    'objCon.Execute "spUpdateCostCentreStatusTable " & Session("BudgetID") & "," & Session("VersionID") & "," & Session("UserID") & ""
                                     
			        'Return the result of the Save Function.
     		       strMessage = objCmd.Parameters.Item("Response")
                   If strMessage = "OK" Then
                        strMessage = "<B>Budget Status Record Saved.</B>"
                        strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
                   Else
                        strMessage = "<FONT Color=""Red""><B>" & strMessage & "</B></FONT>"
                        strMessageIcon = "<img src=""../images/warning.gif"" />"
                   End If 
     		    'Else
     		        'strMessage = objCmd.Parameters.Item("Response")
     		    'End If
     		
             Else
                strMessage = "<FONT Color=""Red""><B>BUDGET IS CLOSED, CHANGES CANNOT BE MADE.</B></FONT>"
                strMessageIcon = "<img src=""../images/warning.gif"" />"
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

Public Function Validate_Access(UserTypeID,Screen)

    If Session("UserTypeID") = 99 Then
        
        Validate_Access = "Y"
        
    Else
        
        objRS.Open "SELECT ScreenID FROM qryScreenAccess WHERE UserTypeID = " & UserTypeID & " AND PageName = '" & Screen & "'",objCon

            If objRS.EOF Then
                Validate_Access = "N" 
            Else
                Validate_Access = "Y"
            End If
    
        objRS.Close
    
    End If

End Function


Set objRS = Nothing
Set objCon = Nothing


%>
