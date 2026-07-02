<%@ Language=VBScript %>
<% Option Explicit
	
	Response.Expires = -1500

    If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
	
	Session("CurrentPage") = "Admin/AdminMenu.asp"

Dim objCon
Dim objRS

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

    objCon.Open Session("DBConnection")
    
If Validate_Access(Session("UserTypeID"),Session("CurrentPage")) = "N" Then
     'Response.Redirect "../AccessDenied.asp"
End If	

%>
<html>
<head>
<body scroll="auto">
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BertStyle.css">

<title>Isidore </title>
       <H3>Administration Menu</H3>
          
             <table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
               
                <tr>
	        	    <td Style="Width:50%;" colspan="3"><B>&nbsp;1&nbsp;General</B></td><td Style="Width:50%;" colspan="3" >&nbsp;<b>6&nbsp;User&nbsp;Role&nbsp;Configuration</b></td>
	            </tr>		
	            <tr>
	        	    <td  colspan="6">&nbsp;</td>
	            </tr>	
                <tr>
	        	    <td  width="5%">&nbsp;1.1</td><td  width="15%"><a Target="Body" HREF="Budget.asp">&nbsp;Budget</a></td><td  width="30%">&nbsp;Maintain Budget details.</td><td  width="5%">&nbsp;6.1</td><td  width="15%"><a Target="Body" HREF="ScreenAccess.asp?FinancialYearID=0">&nbsp;User Role Screen Access</a></td><td  width="30%">&nbsp;Insert and Edit Financial years details.</td>
	            </tr>
	            <tr>
		            <td >&nbsp;1.2</td><td ><a Target="Body" HREF="User.asp?UserID=0">&nbsp;User</a></td><td >&nbsp;User administration.</td><td >&nbsp;6.2</td><td ><a Targt="Body" HREF="User.asp?UserID=0">&nbsp;User</a></td><td >&nbsp;User administration.</td>
	            </tr>
            	<tr>
		            <td >&nbsp;1.3</td><td ><a Target="Body" HREF="Version.asp">&nbsp;Version</a></td><td >&nbsp;Version administration.</td><td colspan="3"></td>
	            </tr>
                 <tr>
		        <td  width="5%">&nbsp;1.4</td><td  width="15%"><a Target="Body" HREF="FinancialYears.asp?FinancialYearID=0">&nbsp;Financial Years</a></td><td  width="30%">&nbsp;Insert and Edit Financial years details.</td><td colspan="3"></td>		
	        </tr>
                <tr>
	        	    <td  colspan="6">&nbsp;</td>
	            </tr>	       
                <tr>
	        	    <td  colspan="3"><B>&nbsp;2&nbsp;Segment Administration</B></td><td  colspan="3"><B>&nbsp;7&nbsp;Input Sheet Configuration</B></td>
	            </tr>		
	            <tr>
	        	    <td  colspan="6">&nbsp;</td>
	            </tr>
	             <tr>
		            <td >&nbsp;2.1</td><td ><a Target="Body" HREF="GLCodes.asp">&nbsp;GFS Code</a></td><td >&nbsp;GFS Code administration.</td><td >&nbsp;7.1</td><td ><a Target="Body" HREF="ColumnNames.asp">&nbsp;Input Sheet Config</a></td><td >&nbsp;Input Sheet Configuration.</td>
	            </tr>
	            <tr>
		            <td >&nbsp;2.2</td><td ><a Target="Body" HREF="BusinessArea.asp">&nbsp;<%=Session("BAName")%></a></td><td >&nbsp;<%=Session("BAName")%> administration.</td><td >&nbsp;7.2</td><td ><a Target="Body" HREF="Formulas.asp">&nbsp;Formulas</a></td><td >&nbsp;Formula Configuration.</td>
	            </tr>
	            <tr>
		            <td >&nbsp;2.3</td><td ><a Target="Body" HREF="FinStatementConfig.asp">&nbsp;Financial Statement Config</a></td><td >&nbsp;Financial Statement Config.</td><td >&nbsp;7.3</td><td ><a Target="Body" HREF="CalculatedFields.asp">&nbsp;Calculated Fields</a></td><td >&nbsp;Calculated Field Configuration.</td>
	            </tr>	
	            <tr>
		            <td >&nbsp;2.4</td><td ><a Target="Body" HREF="CostCentre.asp?CostCentreID=0">&nbsp;<%=Session("CCName")%></a></td><td >&nbsp;<%=Session("CCName")%> administration.</td><td >&nbsp;7.4</td><td ><a Target="Body" HREF="CalculatedFieldValues.asp">&nbsp;Calculated Field Values</a></td><td >&nbsp;Calculated Field Value Administration.</td>
	            </tr>
                <tr>
		            <td >&nbsp;2.5</td><td ><a Target="Body" HREF="SegmentValues.asp">&nbsp;Segment Values</a></td><td >&nbsp;Segment Values administration.</td><td >&nbsp;7.5</td><td ><a Target="Body" HREF="InputFormulas.asp">&nbsp;Formula Assignment</a></td><td >&nbsp;Formula Assignment.</td>
	            </tr>
                <tr>
		            <td >&nbsp;2.6</td><td ><a Target="Body" HREF="Projects.asp">&nbsp;Projects</a></td><td>&nbsp;Project administration.</td><td colspan="3"></td>
	            </tr>
	            <tr>
		            <td >&nbsp;2.7</td><td ><a Target="Body" HREF="CCBARelationship.asp">&nbsp;<%=Session("BAName")%> / <%=Session("CCName")%>  Relationship</a></td><td >&nbsp;<%=Session("BAName")%> / <%=Session("CCName")%>  Relationship administration.</td><td >&nbsp;10.1</td><td ><a Target="Body" HREF="WarrantReleases.asp">&nbsp;Warrant Releases</a></td><td >&nbsp;Warrant Releases.</td>
	            </tr>
                <tr>
		            <td >&nbsp;2.8</td><td ><a Target="Body" HREF="Goals.asp">&nbsp;Goals</a></td><td >&nbsp;Goal administration.</td><td colspan="3"></td>
	            </tr>
                <tr>
		            <td >&nbsp;2.9</td><td ><a Target="Body" HREF="Clusters.asp">&nbsp;Clusters</a></td><td >&nbsp;Cluster administration.</td><td colspan="3"></td>
	            </tr>
	            <tr>
	        	    <td  colspan="6">&nbsp;</td>
	            </tr>
	            <tr>
	        	    <td  colspan="3"><B>&nbsp;3&nbsp;Access Control</B></td> <td  colspan="3"><B>&nbsp;8&nbsp;Ceiling Administration</B></td>
	            </tr>
	              <tr>
	        	    <td  colspan="6">&nbsp;</td>
	            </tr>	
	            <tr>
		            <td >&nbsp;3.1</td><td ><a Target="Body" HREF="BusinessAreaStatus.asp">&nbsp;<%=Session("BAName")%> Status</a></td><td >&nbsp;<%=Session("CCName")%> Ceilings.</td><td>&nbsp;8.1</td><td ><a Target="Body" HREF="CeilingsLevel2.asp">&nbsp;Ceiling Administration</a></td><td >&nbsp;Ceiling administration.</td>
	            </tr>
            	
	            <tr>
		            <td >&nbsp;3.2</td><td ><a Target="Body" HREF="BusinessAreaAccess.asp">&nbsp;<%=Session("BAName")%> Access</a></td><td >&nbsp;<%=Session("BAName")%> Access administration.</td><td>&nbsp;8.2</td><td ><a Target="Body" HREF="FundingPurposes.asp">&nbsp;Funding Purpose Administration</a></td><td >&nbsp;Funding Purpose administration.</td>
	            </tr>            	
	            <tr>
	        	    <td  colspan="6">&nbsp;</td>
	            </tr>	       
                <tr>
	        	    <td  colspan="3"><B>&nbsp;4&nbsp;Roll&nbsp;Over</B></td><td colspan="3"><B>&nbsp;9&nbsp;Staff Administration</B></td>
	            </tr>		
	            <tr>
	        	    <td  colspan="6">&nbsp;</td>
	            </tr>			
                <tr>
	        	    <td  width="5%">&nbsp;4.1</td><td  width="15%"><a Target="Body" HREF="BudgetRollOver.asp">&nbsp;Rollover</a></td><td  width="30%">&nbsp;Rollover administration.</td><td >&nbsp;9.1</td><td ><a Target="Body" HREF="SalaryClassifications.asp">&nbsp;Salary Classifications</a></td><td >&nbsp;Salary Classifications administration.</td>
	            </tr>				
	            <tr>
	        	    <td  width="5%">&nbsp;4.2</td><td  width="15%"><a Target="Body" HREF="BudgetDataBuildLog.asp">&nbsp;Recalculate Input Sheets</a></td><td  width="30%">&nbsp;Recalculate Input Sheets.</td><td >&nbsp;9.2</td><td ><a Target="Body" HREF="SuperFunds.asp">&nbsp;Super Funds</a></td><td >&nbsp;Super Funds administration.</td>
	            </tr>
                <tr>
	        	    <td  width="5%">&nbsp;4.3</td><td  width="15%"><a Target="Body" HREF="Indexation.asp">&nbsp;Indexation</a></td><td  width="30%">&nbsp;Indexation Administration.</td><td >&nbsp;9.2</td><td ><A HREF="../ScrollingFrameset.asp?CurrentPage=Admin/Establishments.asp&HeaderMenu=Header1.asp" Target="activeframe">&nbsp;Establishments</a></td><td >&nbsp;Establishments.</td>
	            </tr>	
				<tr>
	        	    <td  width="5%">&nbsp;4.4</td><td  width="15%"><a Target="Body" HREF="BudgetRolloverBA.asp">&nbsp;Business Area Rollover</a></td><td  width="30%">&nbsp;Indexation Administration.</td><td >&nbsp;9.2</td><td ><a Target="Body" HREF="SuperFunds.asp">&nbsp;Super Funds</a></td><td >&nbsp;Super Funds administration.</td>
	            </tr>	
                  <tr>
	        	    <td  colspan="3">&nbsp;</td><td >&nbsp;9.3</td><td ><a Target="Body" HREF="StaffingParametersDefault.asp">&nbsp;Staffing Parameters</a></td><td >&nbsp;Staffing Parameters administration.</td>
	            </tr>	
	              <tr>
	        	    <td  colspan="3"><B>&nbsp;5&nbsp;Epicor Interface</B></td><td >&nbsp;9.4</td><td ><a Target="Body" HREF="WorkDays.asp">&nbsp;Work Days</a></td><td >&nbsp;Work Days administration.</td>
	            </tr>		
	            <tr>
	        	    <td  colspan="3">&nbsp;</td><td >&nbsp;9.5</td><td ><a Target="Body" HREF="StaffingClassificationsDefault.asp">&nbsp;Salary Increment Update</a></td><td >&nbsp;Salary Increment Update administration.</td>
	            </tr>	
	            <tr>
		            <td >&nbsp;5.1</td><td ><a Target="Body" HREF="ActualsGLReconciliation.asp">&nbsp;Import Actuals from Epicor</a></td><td >&nbsp;Execute the Actuals Import procedure from Epicor.</td><td >&nbsp;9.6</td><td ><a Target="Body" HREF="Positions.asp">&nbsp;Position Types</a></td><td >&nbsp;Position Type administration.</td>	
	            </tr>
	            <tr>
		            <td >&nbsp;5.2</td><td ><a Target="Body" HREF="BudgetExport.asp">&nbsp;Load budget to Epicor.</a></td><td >&nbsp;Transfers budget from Isidore to Epicor via interface.</td><td >&nbsp;9.7</td><td ><a Target="Body" HREF="Establishments.asp">&nbsp;Position Pool</a></td><td >&nbsp;Position Pool administration.</td>
	            </tr> 
		        <tr>
	        	    <td  colspan="6">&nbsp;</td>
	            </tr> 
          </table>	 

	 


     	                            




</body>
</html>

<%
    
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