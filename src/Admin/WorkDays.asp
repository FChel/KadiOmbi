<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%
 
'Description:	Work Days Admin Screen
'Author:		MG
'Date:			March 2014

'Declare default variables
Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage

'Set ADO Database objects
Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection
objCon.Open Session("DBConnection")

Dim lngBM1
Dim lngBM2
Dim lngBM3
Dim lngBM4
Dim lngBM5
Dim lngBM6
Dim lngBM7
Dim lngBM8
Dim lngBM9
Dim lngBM10
Dim lngBM11
Dim lngBM12
Dim lngTotal
Dim lngOY1
Dim lngOY2
Dim lngOY3
Dim arrMonthName(12)

	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

If IsNull(Session("FirstMonth")) or Session("FirstMonth") = "" Then Session("FirstMonth") = "JUL"

    'Call the procedure to create the Month Names
    Call GetMonthNames()
    
	'Load page details
	LoadDetails()
		
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
	<script src="../formChek.js">
	</script>

<script LANGUAGE="javascript">
<!--
function SaveData()
{	    
    var varSubmit = true						
    var varAlert="";	

	if(isPositiveInteger(frm.BM1.value)==false)
	{
       varAlert += "Please enter a numeric value for BM1. \n \n";
       document.getElementById('BM1').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM1').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM2.value)==false)
	{
       varAlert += "Please enter a numeric value for BM2. \n \n";
       document.getElementById('BM2').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM2').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM3.value)==false)
	{
       varAlert += "Please enter a numeric value for BM3. \n \n";
       document.getElementById('BM3').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM3').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM4.value)==false)
	{
       varAlert += "Please enter a numeric value for BM4. \n \n";
       document.getElementById('BM4').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM4').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM5.value)==false)
	{
       varAlert += "Please enter a numeric value for BM5. \n \n";
       document.getElementById('BM5').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM5').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM6.value)==false)
	{
       varAlert += "Please enter a numeric value for BM6. \n \n";
       document.getElementById('BM6').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM6').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM7.value)==false)
	{
       varAlert += "Please enter a numeric value for BM7. \n \n";
       document.getElementById('BM7').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM7').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM8.value)==false)
	{
       varAlert += "Please enter a numeric value for BM8. \n \n";
       document.getElementById('BM8').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM8').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM9.value)==false)
	{
       varAlert += "Please enter a numeric value for BM9. \n \n";
       document.getElementById('BM9').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM9').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM10.value)==false)
	{
       varAlert += "Please enter a numeric value for BM10. \n \n";
       document.getElementById('BM10').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM10').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM11.value)==false)
	{
       varAlert += "Please enter a numeric value for BM11. \n \n";
       document.getElementById('BM11').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM11').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.BM12.value)==false)
	{
       varAlert += "Please enter a numeric value for BM12. \n \n";
       document.getElementById('BM12').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('BM12').style.backgroundColor="ffffff";        
    
    if(isPositiveInteger(frm.OY1.value)==false)
	{
       varAlert += "Please enter a numeric value for OY1. \n \n";
       document.getElementById('OY1').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('OY1').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.OY2.value)==false)
	{
       varAlert += "Please enter a numeric value for OY2. \n \n";
       document.getElementById('OY2').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('OY2').style.backgroundColor="ffffff";
    
    if(isPositiveInteger(frm.OY3.value)==false)
	{
       varAlert += "Please enter a numeric value for OY3. \n \n";
       document.getElementById('OY3').style.backgroundColor="ff8080";
       varSubmit = false;
    }   
    else document.getElementById('OY3').style.backgroundColor="ffffff";
    
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}

//-->
</script>
</head>
<body>
<form action="WorkDays.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>
	<tr>
	    <th colspan="2" style="width:50%">Work Days Administration</th>
	    <th colspan="2" style="width:50%">
	    <%
	    Response.write "for the currently selected Budget - <span style=""color:black;"">" &  Session("BudgetName") & "</span>"
	    %>
	    </th>
	</tr>
	<tr>
	    <td colspan="4">&nbsp;</td>
	</tr>
		
	<tr>
        <th width=20% align="left">&nbsp;<%=arrMonthName(1)%> - (BM1)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM1" name="BM1" maxlength="50" TABINDEX="1" value="<%=lngBM1%>"></td>
	    <th width=20% align="left">&nbsp;<%=arrMonthName(2)%> - (BM2)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM2" name="BM2" maxlength="50" TABINDEX="2" value="<%=lngBM2%>"></td>
	</tr>
	
	<tr>
        <th width=20% align="left">&nbsp;<%=arrMonthName(3)%> - (BM3)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM3" name="BM3" maxlength="50" TABINDEX="3" value="<%=lngBM3%>"></td>
	    <th width=20% align="left">&nbsp;<%=arrMonthName(4)%> - (BM4)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM4" name="BM4" maxlength="50" TABINDEX="4" value="<%=lngBM4%>"></td>
	</tr>
	
	<tr>
        <th width=20% align="left">&nbsp;<%=arrMonthName(5)%> - (BM5)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM5" name="BM5" maxlength="50" TABINDEX="5" value="<%=lngBM5%>"></td>
	    <th width=20% align="left">&nbsp;<%=arrMonthName(6)%> - (BM6)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM6" name="BM6" maxlength="50" TABINDEX="6" value="<%=lngBM6%>"></td>
	</tr>
	
	<tr>
        <th width=20% align="left">&nbsp;<%=arrMonthName(7)%> - (BM7)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM7" name="BM7" maxlength="50" TABINDEX="7" value="<%=lngBM7%>"></td>
	    <th width=20% align="left">&nbsp;<%=arrMonthName(8)%> - (BM8)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM8" name="BM8" maxlength="50" TABINDEX="8" value="<%=lngBM8%>"></td>
	</tr>
	
	<tr>
        <th width=20% align="left">&nbsp;<%=arrMonthName(9)%> - (BM9)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM9" name="BM9" maxlength="50" TABINDEX="9" value="<%=lngBM9%>"></td>
	    <th width=20% align="left">&nbsp;<%=arrMonthName(10)%> - (BM10)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM10" name="BM10" maxlength="50" TABINDEX="10" value="<%=lngBM10%>"></td>
	</tr>
	
	<tr>
        <th width=20% align="left">&nbsp;<%=arrMonthName(11)%> - (BM11)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM11" name="BM11" maxlength="50" TABINDEX="11" value="<%=lngBM11%>"></td>
	    <th width=20% align="left">&nbsp;<%=arrMonthName(12)%> - (BM12)</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="BM12" name="BM12" maxlength="50" TABINDEX="12" value="<%=lngBM12%>"></td>
	</tr>
	<tr>
        <th width=20% align="left">&nbsp;OY1</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="OY1" name="OY1" maxlength="50" TABINDEX="13" value="<%=lngOY1%>"></td>
	    <th width=20% align="left">&nbsp;OY2</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="OY2" name="OY2" maxlength="50" TABINDEX="14" value="<%=lngOY2%>"></td>
	</tr>
	<tr>
        <th width=20% align="left">&nbsp;OY3</th>
		<td width=30%>&nbsp;<input style="text-align:left; width:90%" id="OY3" name="OY3" maxlength="50" TABINDEX="15" value="<%=lngOY3%>"></td>
		<td width=30% colspan=2></td>
	</tr>
		
</table>
<br>
<div class="buttons">
<TABLE Width="1500px" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<TR>

    <td Width="100px"><button type="button" onclick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
    <td Width="100px"><button type="button" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td> 
    <td Width="150px"><button type="button" onclick="self.location='WorkDays.asp?WorkDaysID=0'" )""><img src="../images/add.png" alt="" /> Clear/Add New </button></td>
    <td Width="700px"><font Color="Red"><b><%=strMessage%></b></font></td>
</TR>
</TABLE>
</div>
<hr>
</form>
</body>

</html>

<% 

Sub LoadDetails()

'Description:	Loads Caller's details into page if applicable.
		
		objRS.Open "SELECT * FROM tblWorkDays WHERE BudgetID=" & clng(Session("BudgetID")) & "",objCon							
		If Not objRS.EOF Then		    
            lngBM1 = objRS("BM1")
            lngBM2= objRS("BM2")
            lngBM3= objRS("BM3")
            lngBM4= objRS("BM4")
            lngBM5= objRS("BM5")
            lngBM6= objRS("BM6")
            lngBM7= objRS("BM7")
            lngBM8= objRS("BM8")
            lngBM9= objRS("BM9")
            lngBM10= objRS("BM10")
            lngBM11= objRS("BM11")
            lngBM12= objRS("BM12")
            lngTotal= objRS("Total")	
            lngOY1 = objRS("OY1")
            lngOY2 = objRS("OY2")
            lngOY3 = objRS("OY3")	
		Else		                            
            lngBM1 = 0
            lngBM2 = 0
            lngBM3 = 0
            lngBM4 = 0
            lngBM5 = 0
            lngBM6 = 0
            lngBM7 = 0
            lngBM8 = 0
            lngBM9 = 0
            lngBM10 = 0
            lngBM11 = 0
            lngBM12 = 0
            lngTotal = 0
            lngOY1 = 0
            lngOY2 = 0
            lngOY3 = 0
		End if

		objRS.Close	
End Sub

Public Sub WriteList()
'Procedure to write the list of Work Days
    Response.Write "<table WIDTH=""50%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
	        "<tr><th align=""center"">Edit</th>" & _
	        "<th align=""center"">Budget</th>" & _				
	        "<th align=""center"">Total</th></tr>"
	    
    objRS.Open "SELECT * FROM qryWorkDays Where BudgetID=" & Session("BudgetID"),objCon
		Do until objRS.eof
			Response.Write "<TR><TD><A Target=""_self"" HREF=""WorkDays.asp"">Work Days</TD><TD style=""text-align:center""><B>" & objRS("BudgetID") & " - " & objRS("BudgetName") & "</TD><TD style=""text-align:center"">" & objRS("Total") & "</B></TD></TR>"
			objRS.movenext
		Loop
			
	objRS.Close
	
	Response.Write "</table>"
End sub

Sub SaveDetails()	
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spWorkDaysSave"
                                  
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM1", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM2", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM3", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM4", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM5", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM6", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM7", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM8", adInteger, adParamInput)                
                .Parameters.Append objCmd.CreateParameter("BM9", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM10", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM11", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BM12", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY1", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY2", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY3", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                                                                    
              
                .Parameters("BudgetID") = Session("BudgetID")		
				.Parameters("BM1") = Request.Form("BM1")
				.Parameters("BM2") = Request.Form("BM2")
				.Parameters("BM3") = Request.Form("BM3")
				.Parameters("BM4") = Request.Form("BM4")
				.Parameters("BM5") = Request.Form("BM5")
				.Parameters("BM6") = Request.Form("BM6")
				.Parameters("BM7") = Request.Form("BM7")
				.Parameters("BM8") = Request.Form("BM8")
				.Parameters("BM9") = Request.Form("BM9")
				.Parameters("BM10") = Request.Form("BM10")
				.Parameters("BM11") = Request.Form("BM11")
				.Parameters("BM12") = Request.Form("BM12")
				.Parameters("OY1") = Request.Form("OY1")
				.Parameters("OY2") = Request.Form("OY2")
				.Parameters("OY3") = Request.Form("OY3")
                .Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessage = "Work Days record saved !"									     		
					
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

Public Sub GetMonthNames()
'This is a procedure to get the order of Month names to be used as titles for Month Columns
Dim intFirstMonth

    'set the First Month name to an integer
    intFirstMonth = Month("21-" & Session("FirstMonth") & "-2012")
    'intFirstMonth = intFirstMonth -1
    arrMonthName(0) = intFirstMonth
    For x = 1 to 12
    
        arrMonthName(x) = Left(MonthName(intFirstMonth + x - 1),3)'intFirstMonth + x
  '      arrMonthName(x) = intFirstMonth'MonthName(intFirstMonth)
        
        'Once the count goes over 12 then go back to 1 to fill the remaining months
        If intFirstMonth + x - 1 > 11 Then 
            If intFirstMonth > 6 Then
                intFirstMonth = 2 - intFirstMonth
            Else
                intFirstMonth = (x - 1) * - 1
            End If
        End If
        
  '      intFirstMonth = x
        
    Next
    
End Sub

Set objRS = Nothing
Set objCon = Nothing


%>
