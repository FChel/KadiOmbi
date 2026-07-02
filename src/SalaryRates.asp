<%@ Language=VBScript %>
<%

Dim strClosePage
Dim objCon
Dim objRS
Dim objRS1
Dim strTransactionName
Dim lngRecordID
Dim strMessage1
Dim strMessage
Dim strSort
Dim strOrder, strOrder2
Dim strPositionNo
Dim intPositionID
Dim arrMonthName(12)
Dim arrHeadings(5)
Dim intFinYearPart1
Dim intFinYearPart2

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")	
	
	If Not IsEmpty(Request.QueryString("TransactionType")) Then
		Session("TransactionType") = Request.QueryString("TransactionType")
	End If	

	If Not IsEmpty(Request.QueryString("RecordID")) Then
		lngRecordID = Request.QueryString("RecordID")
	Else
    		lngRecordID = 0
	End If

	If Not IsEmpty(Request.QueryString("Sort")) Then
	   strSort = Request.QueryString("sort")
	Else
	   strSort = "StaffClassificationDesc"
	End If

	If Not IsEmpty(Request.QueryString("Ordered")) Then
		If Request.QueryString("Ordered") = "asc" Then
			strOrder = "desc"
			strOrder2 = "asc"
		Else
		 strOrder = "asc"
		 strOrder2 = "desc"
		End If
	Else
	   strOrder = "asc"
	   strOrder2 = "desc"
	End If

	If Session("TransactionType") = "DSAL" Then
	    strTransactionName = "Direct Salary Classifications"
	End If
	
	If Session("TransactionType") = "ISAL" Then
	    strTransactionName = "Indirect Salary Classifications"
	End If
	
    If Request.QueryString("Action") = "Save" Then
        If Session("StatusID") = 1 Then
            'Do not allow Read Only users to make changes
            If Session("UserTypeID") = 4 Then
                strMessage = "NOT SAVED. You are a READ ONLY User and cannot make any changes."
            Else
                Call Save_Classifications()
                strMessage = "Staffing Classification record has been saved!"
            End If
        Else
            strMessage = "Budget is closed, no changes can be made!"
        End If
    End If 
    
     If Request.QueryString("Action") = "SaveParameters" Then
        If Session("StatusID") = 1 Then
            'Do not allow Read Only users to make changes
            If Session("UserTypeID") = 4 Then
                strMessage = "NOT SAVED. You are a READ ONLY User and cannot make any changes."
            Else
                Call Save_Parameters()
	            strMessage1 = "Staffing Parameters have been saved!"
	        End If
        Else
            strMessage1 = "Budget is closed, no changes can be made!"
        End If
    End If 

   If Request.QueryString("Action") = "Delete" Then
   	If Session("StatusID") = 1 Then
		    'Do not allow Read Only users to make changes
            If Session("UserTypeID") = 4 Then
                strMessage = "NOT SAVED. You are a READ ONLY User and cannot make any changes."
            Else
        	    Call Delete_Data()
		'strMessage = "Staffing Classification record has been deleted."
		'strMessage = "Staffing Classification record has NOT been deleted."
		    End If
   	Else
        	strMessage = "Budget is closed, no changes can be made!"
   	End If
   End If  

   If Request.QueryString("Action") = "Transfer" Then
   	If Session("StatusID") = 1 Then
		    'Do not allow Read Only users to make changes
            If Session("UserTypeID") = 4 Then
                strMessage = "NOT TRANSFERRED. You are a READ ONLY User and cannot make any changes."
            Else
        	    Call Transfer(Request.QueryString("CostCentreIDTo"),Request.QueryString("EmployeeID"),Request.QueryString("StaffClassID"))
		'strMessage = "Staffing Classification record has been deleted."
		'strMessage = "Staffing Classification record has NOT been deleted."
	     End If
   	Else
        	strMessage = "Budget is closed, no changes can be made!"
   	End If
   End If  

  If Request.QueryString("Action") = "Copy" Then
   	If Session("StatusID") = 1 Then
		    'Do not allow Read Only users to make changes
            If Session("UserTypeID") = 4 Then
                strMessage = "NOT COPIED. You are a READ ONLY User and cannot make any changes."
            Else
        	    Call CopyData(Request.QueryString("StaffClassID"))
		
	     End If
   	Else
        	strMessage = "Budget is closed, no changes can be made!"
   	End If
   End If  

If isnull(Session("FirstMonth")) or Session("FirstMonth") = "" then Session("FirstMonth") = "JAN"

    'Set Headings
    For x = 0 to 5
	
	    intFinYearPart1 = cint(Session("FinancialYear")) + (x - 1)
	    intFinYearPart1 = Right(intFinYearPart1,2)
	    intFinYearPart2 = cint(Session("FinancialYear")) + x
	    intFinYearPart2 = Right(intFinYearPart2,2)

        'If the Financial Year starts in January then do not display the split/multiple years.
        If Session("FirstMonth") = "JAN" Then
            arrHeadings(x) = "20" & cstr(intFinYearPart1)
        Else
	        arrHeadings(x) = cstr(intFinYearPart1) & "/" & cstr(intFinYearPart2)
        End If
    Next

    'Call the procedure to create the Month Names
    Call GetMonthNames()
	
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
<script src="../formChek.js">
</script>
<script LANGUAGE="javascript">

function FlagSaveStatusClassifications(x){
	
	var Row
	var Col
	var Total
	var InputBox = x
	parts = InputBox.split(".");
	Row = parts[0];
	Col = parts[1];
		
	var varElem1 = parseInt(Row) + parseInt(Col);
	var varElem2 = parseInt(Row);
	var varElem3 = 100 + parseInt(Col);

	if(isFloat(document.getElementById(varElem2).value)==false){
		if(Col==0){
		}
		else{alert("Invalid Entry! Please do not include a comma in the number. ie 45,000 must be entered as 45000.");
		document.getElementById(varElem2).value=0;}
	}else{
		document.getElementById(varElem3).value = 'Y';
		}	
}

function FlagSaveStatus(x){
	
	var Row
	var Col
	var Total
	var InputBox = x
	parts = InputBox.split(".");
	Row = parts[0];
	Col = parts[1];
		
	var varElem1 = parseInt(Row) + parseInt(Col);
	var varElem2 = parseInt(Row);
	var varElem3 = 100 + parseInt(Col);

	
		document.getElementById(varElem3).value = 'Y';
			
}

function FlagSaveStatusParameters(x){
	
	var Row
	var Col
	var Total
	var InputBox = x
	parts = InputBox.split(".");
	Row = parts[0];
	Col = parts[1];	
	
	var varElem1 = parseInt(Row) + parseInt(Col);
	var varElem2 = parseInt(Row) + 200;
	var varElem3 = parseInt(Row) + 100;
	
	if(isFloat(document.getElementById(varElem2).value)==false){
		if(Col==0){
		}
		else{alert("Invalid Entry! Please do not include a comma in the number. ie 45,000 must be entered as 45000.");
		document.getElementById(varElem2).value=0;}
	}else{
		document.getElementById(varElem3).value = 'Y';
		}	
}

function SaveData(){

 var varSubmit = true						
 var varAlert="";

 if(isWhitespace(frm.StaffClassificationDesc.value))
        {            
		   
	    }
	    else
	    {
	      
	    if(isInteger(document.getElementById('PerformancePay').value)==false)
	        {
            varAlert += "Performance pay is invalid. \n \n";
            document.getElementById('PerformancePay').style.backgroundColor="ff8080";
            varSubmit = false;
            }   
            else 

	    {document.getElementById('PerformancePay').style.backgroundColor="ffffff";} 
 

	    if(document.getElementById('SuperFundID').value==0)
	        {
            varAlert += "Please select Super Fund. \n \n";
            document.getElementById('SuperFundID').style.backgroundColor="ff8080";
            varSubmit = false;
            }   
            else 

	    {document.getElementById('SuperFundID').style.backgroundColor="ffffff";}   


	    if(document.getElementById('StaffClassification').value==0)
	        {
            varAlert += "Please select Classification. \n \n";
            document.getElementById('StaffClassification').style.backgroundColor="ff8080";
            varSubmit = false;
            }   
            else 

	    {document.getElementById('StaffClassification').style.backgroundColor="ffffff";}  	            
	   	

	 }


	
              
	    				
	
	if(varSubmit == true)
	{
	   frm.submit();
	    frm.msgbox.value='Saving.......';		
	}
	else
	{
	    alert(varAlert);
	}
	
}

function SaveParameters(){
	frm1.submit();
	frm1.msgbox1.value='Saving.......';
}


function deleteRecord(RecordID)
{
   
        if(window.confirm('Would you like to DELETE the selected record?')==true){
	self.location="SalaryRates.asp?Action=Delete&RecordID=" + RecordID
	}
     
}
function TransferRecord(RecordID, EmployeeID, StaffClassID)
{
 varText = document.getElementById('CostCentre'+RecordID).options[document.getElementById('CostCentre'+RecordID).selectedIndex].text
 varValue = document.getElementById('CostCentre'+RecordID).options[document.getElementById('CostCentre'+RecordID).selectedIndex].value

	//alert(document.getElementById(RecordID).value);

   //alert(document.getElementById('CostCentre'+RecordID).options[document.getElementById('CostCentre'+RecordID).selectedIndex].text);
	
        if(window.confirm('Would you like to TRANSFER: ' + document.getElementById(RecordID).value + ' \n \n To: ' + varText + '?')==true){
	self.location="SalaryRates.asp?Action=Transfer&EmployeeID=" + EmployeeID + "&CostCentreIDTo=" + varValue + "&StaffClassID=" + StaffClassID
	}
     
}

function CopyData(RecordID, StaffClassID)
{
	
        if(window.confirm('Would you like to COPY: ' + document.getElementById(RecordID).value + '?')==true){
	self.location="SalaryRates.asp?Action=Copy&StaffClassID=" + StaffClassID
	}
     
}
</script>
</head>
<body Scroll="Auto">
<h3>Salary Rates</h3>
<form action="SalaryRates.asp?Action=Save&Sort=StaffClassificationDesc&Ordered=<%=strOrder2%>" method="POST" id="frm" name="frm">
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr><td style="border:0px none;" colspan="10"></td>
		<td style="border-top:1px solid black; border-left:1px solid black; border-right:1px solid black; background-color: #f2f2f2; text-align:center; color:gray; font-weight:bold;" colspan="15">Salary Monthly Increments</td>
		<td colspan="2"></td></tr>
	<tr><th Width="5%" Style="background-color:#ccccff;">Copy</th>
		<% Response.write "<th Width=""20%""><B><A Target=""_self"" HREF=""SalaryRates.asp?Sort=StaffClassificationDesc&Ordered=" & strOrder & """>Employee"
			If strSort = "StaffClassificationDesc" Then Response.write "<img height=20 =width=20 src=../images/" & strOrder & ".gif Align=absmiddle>"
		
		Response.write "</A></B></th>"
	%>
		<!--<th Width="10%">Employee</th>-->
		<th Width="5%">Position No.</th>
		<th Width="5%">Salary</th>
		<th Width="5%">Performance Pay</th>
		<th Width="10%">Class</th>
		<th Width="10%">Super</th>
		<th Width="10%">Position</th>
		<th Width="10%" Style="background-color:#ccccff;">Transfer To</th>
		<th Width="5%">Sort Order</th>
		

<%
    Response.Write "<th Width=""1%"">" & arrMonthName(1) & "</th><th Width=""1%"">" & arrMonthName(2) & "</th>" & _
                "<th Width=""1%"">" & arrMonthName(3) & "</th><th Width=""1%"">" & arrMonthName(4) & "</th>" & _
                "<th Width=""1%"">" & arrMonthName(5) & "</th><th Width=""1%"">" & arrMonthName(6) & "</th>" & _
                "<th Width=""1%"">" & arrMonthName(7) & "</th><th Width=""1%"">" & arrMonthName(8) & "</th>" & _
                "<th Width=""1%"">" & arrMonthName(9) & "</th><th Width=""1%"">" & arrMonthName(10) & "</th>" & _
                "<th Width=""1%"">" & arrMonthName(11) & "</th><th Width=""1%"">" & arrMonthName(12) & "</th>" & _
                "<TH Width=""1%"">" & arrHeadings(2) & "</TH><TH Width=""1%"">" & arrHeadings(3) & "</TH><TH Width=""1%"">" & arrHeadings(4) & "</TH>"
 %>
		<th Width="5%">Total Salary Cost</th>
		<th Width="1%"></th>
		
	</tr>
<%
	DisplayClassificationDetails()
%>	

<tr>
	<td><input id="StaffClassificationDesc" name="StaffClassificationDesc" type="text" style="Width:100%; Text-Align:Left"/></td>
	<td><input id="PositionNo" name="PositionNo" maxlength="50" type="text" style="Text-Align:left;"/></td>
    <td style="background-color:#FFFFFF;"><input id="Salary" READONLY name="Salary" type="text" style="background-color:#FFFFFF;"/></td>
    <td><input id="PerformancePay" name="PerformancePay" type="text" style="Text-Align:right;"/></td>
    <td><select id="StaffClassification" name="StaffClassification" style="Width:100%;"><OPTION Value="0">Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblSalaryClassifications WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Active = 'Y'",objCon
    		
		    Do until objRS.EOF
			  Response.Write "<option Value=""" & objRS("SalaryClassification") & """" & strSelected & ">" & objRS("SalaryClassification") &"</OPTION>"
			  objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
    <td><select id="SuperFundID" name="SuperFundID" style="Width:100%;"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblSuperFund WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Active = 'Y'",objCon
    		
		    Do until objRS.EOF
			    Response.Write "<option Value=""" & objRS("SuperFundID") & """" & strSelected & ">" & objRS("SuperFundName") &"</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td><th Colspan="21">&nbsp;</th></tr>
<tr><th colspan="27">&nbsp;</th></tr>
</table>
<br>
<hr>
<div class="buttons">
<table width="1000px" border="0" cellspacing="0" cellpadding="0">
<TR>

    <td Width="100px"><button type="button" onclick="self.location='../ProfitLoss/DataEntry6.asp'"><img src="../images/door.png" alt="" /> Close </button></td>     
    <td Width="100px" style="border-right:0px"><button type="button" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>    
    <td Width="100px"><button type="button" title="Click to Export this Statement to Microsoft Excel" onclick="window.open('SalaryRatesExcel.asp?Sort=StaffClassificationDesc&Ordered=<%=strOrder%>')"><img src="../images/page_excel.png" alt="" /> Excel </button></td>
    <td Width="100px"><button type="button" title="Click to Print this Statement" onclick="window.open('SalaryRatesPrint.asp?Sort=StaffClassificationDesc&Ordered=<%=strOrder%>')"><img src="../images/printer.png" alt="" /> Print </button></td>
    <TD align="left" Width="600x" style="BORDER-RIGHT:0px"><INPUT style="Align:Left; Font:Bold; Width:100%; Color:Red;" type="text" id="msgbox" name="msgbox" value="<%=strMessage%>"></TD>
</TR>
</table>
</div>
</form>
<form action="SalaryRates.asp?Action=SaveParameters" method="POST" id="frm1" name="frm1">
<br>
<table WIDTH="30%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
		<th Width="80%">Staffing Parameter</th>
		<th Width="20%">Rate</th>
	</tr>
<%
	DisplayParameterDetails()
%>	
</table>
<br>
<hr>
<div class="buttons">
<table width="800px" border="0" cellspacing="0" cellpadding="0">
<TR>
   
    <td Width="100px" style="border-right:0px"><button type="button" onclick="SaveParameters()";><img src="../images/tick.png" alt="" /> Save </button></td>    
    <TD align="left" Width="800x" style="BORDER-RIGHT:0px"><INPUT style="Align:Left; Font:Bold; Width:100%; Color:Red;" type="text" id="msgbox1" name="msgbox1" value="<%=strMessage%>"></TD>
</TR>
</table>
</div>
</form>
<br>
</body>
</html>
<%

Public Sub DisplayClassificationDetails()

Dim objRS
Dim x
Dim lngRow
Dim lngRow1
Dim lngRow2
Dim lngRow3
Dim lngRow4
Dim lngRow5
Dim lngRow6
Dim lngRow7
Dim lngRow8
Dim lngRow9
Dim lngRow10
Dim lngRow11
Dim lngRow12
Dim lngRow13
Dim lngRow14
Dim lngRow15
Dim lngRow16
Dim lngRow17
Dim lngRow18
Dim lngRow19
Dim lngRow20
Dim lngRow21
Dim lngRow22
Dim lngRow23
Dim lngRow24
Dim lngRow25
Dim dblSalary
Dim lngEmployeeID
Dim lngSuperFundID
Dim strStaffClassification
Dim dblTotalSalaryCost
Dim dblTSCTotal

Set objRS = Server.CreateObject("ADODB.Recordset")

     objRS.Open "SELECT * FROM qryStaffingClassifications WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " AND TransactionType = '" & Session("TransactionType") & "' AND Deleted = 'N' Order By " & strSort & " " & strOrder,objCon
	'objRS.Open "SELECT * FROM tblStaffingClassifications WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " AND TransactionType = '" & Session("TransactionType") & "' AND Deleted = 'N' Order By " & strSort & " " & strOrder,objCon
	
	
        x = 1
	    lngRow = 1
	    lngRow1 = 101
	    lngRow2 = 201	
	    lngRow3 = 301 
	    lngRow4 = 401
	    lngRow5 = 501
	    lngRow6 = 601
	    lngRow7 = 701
	    lngRow8 = 801
	    lngRow9 = 901
	    lngRow10 = 1001
	    lngRow11 = 1101
	    lngRow12 = 1201
	    lngRow13 = 1301
	    lngRow14 = 1401
	    lngRow15 = 1501
	    lngRow16 = 1601
	    lngRow17 = 1701
	    lngRow18 = 1801
	    lngRow19 = 1901
	    lngRow20 = 2001
	    lngRow21 = 2101	
	    lngRow22 = 2201
		lngRow23 = 2301
		lngRow24 = 2401
		lngRow25 = 2501

	Do until objRS.EOF	
		
		Response.Write "<TR><td><button type=""button"" onclick=""CopyData('" & lngRow & "','" & objRS("StaffingClassificationID") & "')"";><img src=""../images/table_add.png"" alt="""" /></button></td>" & _
			"<TD>&nbsp;<INPUT style=""width:95%; text-align:left;"" type=""text"" id=""" & lngRow & """ name=""" & lngRow & """ value=""" & objRS("StaffClassificationDesc") & """ onchange=""FlagSaveStatus('" & lngRow & "." & x & "');""><INPUT style=verticalalign:right type=""hidden"" style=width:50px readonly id=" & lngRow1 & " name=" & lngRow1 & " value=" & lngRow1 & "></Td>"
		
			    dblSalary = objRS("Salary")
			    lngEmployeeID = objRS("EmployeeID")
			    lngSuperFundID = objRS("SuperFundID")
			    strStaffClassification = objRS("StaffClassification")
			    strPositionNo = objRS("PositionNo")
				If IsNull(objRS("PositionID")) Then
					intPositionID = 0
				Else
			    		intPositionID = objRS("PositionID")
				End If

			    If IsNull(objRS("TotalSalaryCost")) Then
			    	dblTotalSalaryCost = 0
			    Else
					dblTotalSalaryCost = objRS("TotalSalaryCost")
			    End If		
				
					dblTSCTotal = dblTSCTotal + dblTotalSalaryCost
		
		'Write the new Position No cell
		Response.Write "<TD>&nbsp;<INPUT style=""width:95%; text-align:left;"" type=""text"" id="" & lngRow22 & "" maxlength=""50"" name=" & lngRow22 & " value=""" & strPositionNo & """ onchange=""FlagSaveStatus('" & lngRow22 & "." & x & "');""></Td>"
		
		        'Response.Write "<TD Align=""Right"" Style=""background-color:FFFFFF""><INPUT style=text-align:right style=width:100% type=""text"" Style=""background-color:FFFFFF"" id=" & lngRow2 & " name=" & lngRow2 & " value=""" & formatnumber(dblSalary,0,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow2 & "." & x & "');""><INPUT style=text-align:right style=width:0px type=""HIDDEN"" READONLY id=" & lngRow4 & " name=" & lngRow4 & " value=""" & objRS("EmployeeID") & """></TD><TD><INPUT style=text-align:right style=width:100% type=""text"" id=" & lngRow21 & " name=" & lngRow21 & " value=""" & objRS("PerformancePay") & """ onchange=""FlagSaveStatusClassifications('" & lngRow21 & "." & x & "');""></TD>"
		        Response.Write "<TD Align=""Right"" Style=""background-color:FFFFFF""><INPUT style=text-align:right style=width:100% type=""text"" Style=""background-color:FFFFFF"" id=" & lngRow2 & " name=" & lngRow2 & " value=""" & Round(dblSalary,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow2 & "." & x & "');""><INPUT style=text-align:right style=width:0px type=""HIDDEN"" READONLY id=" & lngRow4 & " name=" & lngRow4 & " value=""" & objRS("EmployeeID") & """></TD><TD><INPUT style=text-align:right style=width:100% type=""text"" id=" & lngRow21 & " name=" & lngRow21 & " value=""" & objRS("PerformancePay") & """ onchange=""FlagSaveStatusClassifications('" & lngRow21 & "." & x & "');""></TD>"
		        
			Response.Write "<td Align=""Right""><select style=""width:100%;"" id=" & lngRow3 & " name=" & lngRow3 & " onchange=""FlagSaveStatus('" & lngRow3 & "." & x & "');"">"
	        	
		            objRS1.Open "SELECT * FROM tblSalaryClassifications WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Active = 'Y'",objCon
    		
		                Do until objRS1.EOF
			                
			                If objRS1("SalaryClassification") = strStaffClassification Then
				                strSelected = " SELECTED "
			                Else
				                strSelected = ""
			                End if
				        
				                Response.Write "<option Value=""" & objRS1("Salaryclassification") & """" & strSelected & ">" & objRS1("SalaryClassification") & "</OPTION>"
			                
			                objRS1.Movenext
			                
		                Loop
    		
		            objRS1.Close
    		
	            	Response.Write "</select></td>"
			


			Response.Write "<td Align=""Right""><select style=""width:100%;"" id=" & lngRow20 & " name=" & lngRow20 & " onchange=""FlagSaveStatusClassifications('" & lngRow20 & "." & x & "');"">"
	        	
		            objRS1.Open "SELECT * FROM tblSuperFund WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Active = 'Y'",objCon
    		
		                Do until objRS1.EOF
			                
			                If objRS1("SuperFundID") = cint(lngSuperFundID) Then
				                strSelected = " SELECTED "
			                Else
				                strSelected = ""
			                End if
				        
				                Response.Write "<option Value=""" & objRS1("SuperFundID") & """" & strSelected & ">" & objRS1("SuperFundName") & "</OPTION>"
			                
			                objRS1.Movenext
			                
		                Loop
    		
		            objRS1.Close
    		
	            Response.Write "</select></td>"
			strSelected = ""
		Response.Write "<td Align=""Right""><select style=""width:100%;"" id=" & lngRow23 & " name=" & lngRow23 & " onchange=""FlagSaveStatusClassifications('" & lngRow23 & "." & x & "');""><option value=""0"">Please select...</option>"
	        	
		            objRS1.Open "SELECT * FROM tblPosition WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND Active = 'Y'",objCon
    		
		                Do until objRS1.EOF
			                
			                If objRS1("PositionID") = cint(intPositionID) Then
				                strSelected = " SELECTED "
			                Else
				                strSelected = ""
			                End if
				        
				                Response.Write "<option Value=""" & objRS1("PositionID") & """" & strSelected & ">" & objRS1("PositionName") & "</OPTION>"
			                
			                objRS1.Movenext
			                
		                Loop
    		
		            objRS1.Close
    		
	            Response.Write "</select></td>" & _
			"<td><select id=""CostCentre" & lngRow & """ name=""CostCentre" & lngRow & """ onchange=""TransferRecord('" & lngRow & "','" & lngRow4 & "','" & objRS("StaffingClassificationID") & "');"">"
		 	

	objRS1.Open "SELECT * FROM qryCostCentresByBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & " AND CostObjectTypeID = 2 AND VersionID = " & Session("VersionID") & "",objCon
	
	Do until objRS1.EOF
		If objRS1("CostCentreID") = clng(Session("CostCentreID")) Then
		    strSelected = " SELECTED "
		Else
			strSelected = ""
		End if
			Response.Write "<option Value=""" & objRS1("CostCentreID") & """" & strSelected & ">" & objRS1("ProgramCode") & " - " & objRS1("CostCentreName") & "</OPTION>"
		objRS1.Movenext
	Loop
	
	objRS1.Close



			Response.Write "</select></td><TD><INPUT style=""text-align:right; width:100%"" type=""text"" id=" & lngRow24 & " name=" & lngRow24 & " value=""" & objRS("SortOrder") & """ onchange=""FlagSaveStatusClassifications('" & lngRow24 & "." & x & "');""></TD>"
		        
			Response.Write "<TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow5 & """ name=""" & lngRow5 & """ value=""" & objRS("BM1Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow5 & "." & x & "');""></TD><TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow6 & """ name=""" & lngRow6 & """ value=""" & objRS("BM2Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow6 & "." & x & "');""></TD>" & _
			"<TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow7 & """ name=""" & lngRow7 & """ value=""" & objRS("BM3Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow7 & "." & x & "');""></TD><TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow8 & """ name=""" & lngRow8 & """ value=""" & objRS("BM4Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow8 & "." & x & "');""></TD>" & _
			"<TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow9 & """ name=""" & lngRow9 & """ value=""" & objRS("BM5Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow9 & "." & x & "');""></TD><TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow10 & """ name=""" & lngRow10 & """ value=""" & objRS("BM6Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow10 & "." & x & "');""></TD>" & _
			"<TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow11 & """ name=""" & lngRow11 & """ value=""" & objRS("BM7Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow11 & "." & x & "');""></TD><TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow12 & """ name=""" & lngRow12 & """ value=""" & objRS("BM8Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow12 & "." & x & "');""></TD>" & _
			"<TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow13 & """ name=""" & lngRow13 & """ value=""" & objRS("BM9Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow13 & "." & x & "');""></TD><TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow14 & """ name=""" & lngRow14 & """ value=""" & objRS("BM10Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow14 & "." & x & "');""></TD>" & _
			"<TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow15 & """ name=""" & lngRow15 & """ value=""" & objRS("BM11Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow15 & "." & x & "');""></TD><TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow16 & """ name=""" & lngRow16 & """ value=""" & objRS("BM12Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow16 & "." & x & "');""></TD>" & _
			"<TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow17 & """ name=""" & lngRow17 & """ value=""" & objRS("OY1Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow17 & "." & x & "');""></TD><TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow18 & """ name=""" & lngRow18 & """ value=""" & objRS("OY2Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow18 & "." & x & "');""></TD>" & _
			"<TD><INPUT style=""text-align:right; width:100%; background-color: #f2f2f2;"" READONLY id=""" & lngRow19 & """ name=""" & lngRow19 & """ value=""" & objRS("OY3Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow19 & "." & x & "');""></TD>"

		        'Response.Write "<TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow5 & " name=" & lngRow5 & " value=""" & objRS("BM1Index") & """ onchange=""FlagSaveStatusClassifications('" & lngRow5 & "." & x & "');""></TD><TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow6 & " name=" & lngRow6 & " value=""" & formatnumber(objRS("BM2Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow6 & "." & x & "');""></TD>" & _
			'"<TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow7 & " name=" & lngRow7 & " value=""" & formatnumber(objRS("BM3Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow7 & "." & x & "');""></TD><TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow8 & " name=" & lngRow8 & " value=""" & formatnumber(objRS("BM4Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow8 & "." & x & "');""></TD>" & _
			'"<TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow9 & " name=" & lngRow9 & " value=""" & formatnumber(objRS("BM5Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow9 & "." & x & "');""></TD><TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow10 & " name=" & lngRow10 & " value=""" & formatnumber(objRS("BM6Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow10 & "." & x & "');""></TD>" & _
			'"<TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow11 & " name=" & lngRow11 & " value=""" & formatnumber(objRS("BM7Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow11 & "." & x & "');""></TD><TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow12 & " name=" & lngRow12 & " value=""" & formatnumber(objRS("BM8Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow12 & "." & x & "');""></TD>" & _
			'"<TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow13 & " name=" & lngRow13 & " value=""" & formatnumber(objRS("BM9Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow13 & "." & x & "');""></TD><TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow14 & " name=" & lngRow14 & " value=""" & formatnumber(objRS("BM10Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow14 & "." & x & "');""></TD>" & _
			'"<TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow15 & " name=" & lngRow15 & " value=""" & formatnumber(objRS("BM11Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow15 & "." & x & "');""></TD><TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow16 & " name=" & lngRow16 & " value=""" & formatnumber(objRS("BM12Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow16 & "." & x & "');""></TD>" & _
			'"<TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow17 & " name=" & lngRow17 & " value=""" & formatnumber(objRS("OY1Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow17 & "." & x & "');""></TD><TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow18 & " name=" & lngRow18 & " value=""" & formatnumber(objRS("OY2Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow18 & "." & x & "');""></TD>" & _
			'"<TD><INPUT style=""text-align:right"" style=""width:0px"" type=""hidden"" id=" & lngRow19 & " name=" & lngRow19 & " value=""" & formatnumber(objRS("OY3Index"),2,0) & """ onchange=""FlagSaveStatusClassifications('" & lngRow19 & "." & x & "');""></TD></TD>"

Response.Write "<TD Style=""background-color:FFFFCC; text-align:right;"">" & formatnumber(dblTotalSalaryCost,0,0) & "</TD><TD Colspan=""2"" style=""text-align:right;""><INPUT type=""hidden"" readonly id=""" & lngRow25 & """ name=""" & lngRow25 & """ value=""" & objRS("StaffDataID") & """><IMG SRC=""../images/row_3_button_close.gif"" onclick=""deleteRecord(" & objRS("StaffingClassificationID") & ");""></TD>"					  

			    lngRow = lngRow + 1
			    lngRow1 = lngRow1 + 1
			    lngRow2 = lngRow2 + 1
			    lngRow3 = lngRow3 + 1
			    lngRow4 = lngRow4 + 1	
			    lngRow5 = lngRow5 + 1
			    lngRow6 = lngRow6 + 1
			    lngRow7 = lngRow7 + 1
			    lngRow8 = lngRow8 + 1
			    lngRow9 = lngRow9 + 1
			    lngRow10 = lngRow10 + 1
			    lngRow11 = lngRow11 + 1
			    lngRow12 = lngRow12 + 1
			    lngRow13 = lngRow13 + 1
			    lngRow14 = lngRow14 + 1
			    lngRow15 = lngRow15 + 1
			    lngRow16 = lngRow16 + 1
			    lngRow17 = lngRow17 + 1
			    lngRow18 = lngRow18 + 1    
			    lngRow19 = lngRow19 + 1    
			    lngRow20 = lngRow20 + 1
			    lngRow21 = lngRow21 + 1
			    lngRow22 = lngRow22 + 1
				lngRow23 = lngRow23 + 1
				lngRow24 = lngRow24 + 1
				lngRow25 = lngRow25 + 1
			    
			     x = x + 1
			     
			objRS.Movenext
			
	Loop
			Response.Write "</TR><TR><TH></TH><TH></TH><TH></TH><TH></TH><TH></TH><TH></TH><TH></TH><TH></TH><TH></TH><TH ""Style=Align:Right"" colspan=""16"">Total Cost</TH><TH ""Style=Text-Align:Right; Align:Right"" Colspan=""1"">" & formatnumber(dblTSCTotal,0,0) & "</TH><TH></TH>" & _
							"</TR><TR><TH Align=""Left"" Colspan=""28""><INPUT readonly type=""hidden"" Width=""5px;"" id=RowCount name=RowCount value=" & (lngRow-1) & ">&nbsp;Enter a new staff record in the fields below.</TH></TR>"
		
		objRS.Close			
End Sub

Public Sub DisplayParameterDetails()

Dim objRS
Dim x
Dim lngRow
Dim lngRow1
Dim lngRow2
Dim lngRow3
Dim lngDefault
Dim strPercentageSign	

Dim arrSubTotal(17)
Dim lngReportGroupingID

Set objRS = Server.CreateObject("ADODB.Recordset")

    objRS.Open "SELECT * FROM tblStaffingParameters WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " AND TransactionType = '" & Session("TransactionType") & "' AND Editable = 'Y'",objCon
		
        x = 1
	    lngRow = 601
	    lngRow1 = 701
	    lngRow2 = 801	 
	
	Do until objRS.EOF	
		
		Response.Write "<TR><TD>&nbsp;<INPUT style=width:100% style=""text-align:left"" readonly type=""text"" id=" & lngRow & " name=" & lngRow & " value=" & objRS("StaffingParameterID") & "><INPUT style=verticalalign:right style=width:5px readonly type=""hidden"" id=" & lngRow1 & " name=" & lngRow1 & "></Td>"
		
			    lngDefault = objRS("Rate")
			
		        Response.Write "<TD Align=""Right""><INPUT style=text-align:right style=width:80px type=""text"" id=" & lngRow2 & " name=" & lngRow2 & " value=""" & formatnumber(lngDefault,2,0) & """ onchange=""FlagSaveStatusParameters('" & lngRow & "." & x & "');""></TD>"
					  
			    lngRow = lngRow + 1
			    lngRow1 = lngRow1 + 1
			    lngRow2 = lngRow2 + 1
			   
			objRS.Movenext
			
	Loop

			Response.Write "</TR><TR><TH Colspan=4><INPUT readonly type=""hidden"" id=RowCount1 name=RowCount1 value=" & (lngRow - 401) & ">&nbsp;</TH></TR>"
		
		objRS.Close			
End Sub


Public Sub Save_Classifications()

Dim Row
Dim Col
Dim v
Dim w
Dim x
Dim y
Dim z
Dim a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p,q,r,s,t,aa
Dim strStaffClass
Dim intStaffDataID

    If Session("ModeID") = 2 or Session("ModeID") = 4 Then
    
        For u = 1 to cint(Request.Form("RowCount"))
    	   
		   v = 1 + u
		    w = 100 + u
    		y = 200 + u 
    		z = 300 + u 
    		x = 400 + u   
    		a = 500 + u 
    		b = 600 + u
    		c = 700 + u
    		d = 800 + u
    		e = 900 + u
    		f = 1000 + u
    		g = 1100 + u
    		h = 1200 + u
    		i = 1300 + u
    		j = 1400 + u
    		k = 1500 + u
    		l = 1600 + u
    		m = 1700 + u
    		n = 1800 + u
    		o = 1900 + u
    		p = 2000 + u
		    q = 2100 + u
		    r = 2200 + u
		s = 2300 + u
		t = 2400 + u
    		aa = 2500 + u

		        If Request.Form("" & w & "") = "Y" Then	
		        'Save the data in the row

				'Replace any apostrophe's in the text with triple apostrophe's
				strStaffClass = Request.Form("" & u & "")

				strStaffClass = cstr(replace(Request.Form("" & u & ""),"'","''"))
				'If the Salary Classification has not been saved in Data entry yet there will be no Staff DataID
				If isnull(Request.Form("" & aa & "")) or Request.Form("" & aa & "") = "" then
					intStaffDataID= 0
				Else
					intStaffDataID = Request.Form("" & aa & "")
				End if

		            'objCon.Execute "spStaffingClassificationsSave " & cint(Session("BudgetID")) & "," & cint(Session("VersionID")) & "," & clng(Session("CostCentreID")) & ",'" & Session("TransactionType") & "','" & Request.Form("" & u & "") & "'," & Request.Form("" & y & "") & ",'" & Request.Form("" & z & "") & "'," & Request.Form("" & x & "") & "," & Request.Form("" & a & "") & "," & Request.Form("" & b & "") & "," & Request.Form("" & c & "") & "," & Request.Form("" & d & "") & "," & Request.Form("" & e & "") & "," & Request.Form("" & f & "") & "," & Request.Form("" & g & "") & "," & Request.Form("" & h & "") & "," & Request.Form("" & i & "") & "," & Request.Form("" & j & "") & "," & Request.Form("" & k & "") & "," & Request.Form("" & l & "") & "," & Request.Form("" & m & "") & "," & Request.Form("" & n & "") & "," & Request.Form("" & o & "") & "," & Request.Form("" & p & "") & "," & Request.Form("" & q & "") & "," & Session("UserID") & ""
		         objCon.Execute "spStaffingClassificationsSave " & cint(Session("BudgetID")) & "," & cint(Session("VersionID")) & "," & clng(Session("CostCentreID")) & ",'" & Session("TransactionType") & "','" & strStaffClass & "'," & Request.Form("" & y & "") & ",'" & Request.Form("" & z & "") & "'," & Request.Form("" & x & "") & "," & Request.Form("" & t & "") & "," & Request.Form("" & a & "") & "," & Request.Form("" & b & "") & "," & Request.Form("" & c & "") & "," & Request.Form("" & d & "") & "," & Request.Form("" & e & "") & "," & Request.Form("" & f & "") & "," & Request.Form("" & g & "") & "," & Request.Form("" & h & "") & "," & Request.Form("" & i & "") & "," & Request.Form("" & j & "") & "," & Request.Form("" & k & "") & "," & Request.Form("" & l & "") & "," & Request.Form("" & m & "") & "," & Request.Form("" & n & "") & "," & Request.Form("" & o & "") & "," & Request.Form("" & p & "") & "," & Request.Form("" & q & "") & "," & Session("UserID") & ",'" & Request.Form("" & r & "") & "','" & Request.Form("" & s & "") & "','N','N'," & intStaffDataID & ""
		            '''*****line above is the live save line
			'Response.Write "spStaffingClassificationsSave " & cint(Session("BudgetID")) & "," & cint(Session("VersionID")) & "," & clng(Session("CostCentreID")) & ",'" & Session("TransactionType") & "','" & strStaffClass & "'," & Request.Form("" & y & "") & ",'" & Request.Form("" & z & "") & "'," & Request.Form("" & x & "") & "," & Request.Form("" & t & "") & "," & Request.Form("" & a & "") & "," & Request.Form("" & b & "") & "," & Request.Form("" & c & "") & "," & Request.Form("" & d & "") & "," & Request.Form("" & e & "") & "," & Request.Form("" & f & "") & "," & Request.Form("" & g & "") & "," & Request.Form("" & h & "") & "," & Request.Form("" & i & "") & "," & Request.Form("" & j & "") & "," & Request.Form("" & k & "") & "," & Request.Form("" & l & "") & "," & Request.Form("" & m & "") & "," & Request.Form("" & n & "") & "," & Request.Form("" & o & "") & "," & Request.Form("" & p & "") & "," & Request.Form("" & q & "") & "," & Session("UserID") & ",'" & Request.Form("" & r & "") & "','" & Request.Form("" & s & "") & "','N','N'," & intStaffDataID & ""

				'Response.Write "spStaffingClassificationsSave " & cint(Session("BudgetID")) & "," & cint(Session("VersionID")) & "," & clng(Session("CostCentreID")) & ",'" & Session("TransactionType") & "','" & strStaffClass & "'," & Request.Form("" & y & "") & ",'" & Request.Form("" & z & "") & "'," & Request.Form("" & x & "") & "," & Request.Form("" & t & "") & ",1,1,1,1,1,1,1,1,1,1,1,1,1,1,1," & Request.Form("" & p & "") & "," & Request.Form("" & q & "") & "," & Session("UserID") & ",'" & Request.Form("" & r & "") & "','" & Request.Form("" & s & "") & "','N','N'"
        		    strMessage = "Record sucessfully saved!"
        		    
        		    'Response.Write "here" & Request.Form("" & u & "")
        		           		    
        		End If
    		
	        Next    

		If Not IsNull(Request.Form("StaffClassificationDesc")) AND Request.Form("StaffClassificationDesc") <> "" or Request("SuperFundID") <> 0 AND Request.Form("StaffClassification") <> "0" Then

			'Replace any apostrophe's in the text with triple apostrophe's
			strStaffClass = Request.Form("StaffClassificationDesc")

			strStaffClass = replace(strStaffClass,"'","''")

			objCon.Execute "spStaffingClassificationInsert " & cint(Session("BudgetID")) & "," & cint(Session("VersionID")) & "," & clng(Session("CostCentreID")) & ",'" & Session("TransactionType") & "',0,'" & Request.Form("StaffClassification") & "','" & strStaffClass & "',0,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1," & Request.Form("SuperFundID") & "," & Request.Form("PerformancePay") & "," & Session("UserID") & ",'" & Request.Form("PositionNo") & "'"
			'response.Write "spStaffingClassificationInsert " & cint(Session("BudgetID")) & "," & cint(Session("VersionID")) & "," & clng(Session("CostCentreID")) & ",'" & Session("TransactionType") & "',0,'" & Request.Form("StaffClassification") & "','" & strStaffClass & "',0,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1," & Request.Form("SuperFundID") & "," & Request.Form("PerformancePay") & "," & Session("UserID") & ",'" & Request.Form("PositionNo") & "'"
			'Response.Write "spStaffingClassificationInsert " & cint(Session("BudgetID")) & "," & cint(Session("VersionID")) & "," & clng(Session("CostCentreID")) & ",'" & Session("TransactionType") & "',0,'" & Request.Form("StaffClassification") & "','" & strStaffClass   & "',0,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1," & Request.Form("SuperFundID") & "," & Request.Form("PerformancePay") & "," & Session("UserID") & ""
		End If

		strMessage = "Staffing Classification record has been saved."
    
    Else
    
        strMessage = "Data can only be saved in Budget Mode!"
    
    End If   	
    
End Sub

Public Sub Save_Parameters()

Dim Row
Dim Col
Dim w
Dim x
Dim y
Dim z


    If Session("ModeID") = 2 or Session("ModeID") = 4 Then
     
        For w = 1 to cint(Request.Form("RowCount1"))
    	   
		    x = 400 + w
		    y = 500 + w
    		z = 600 + w   
    		
		        If Request.Form("" & y & "") = "Y" Then	
		        'Save the data in the row
		            objCon.Execute "spStaffingParametersUpdate '" & Request.Form("" & x & "") & "'," & cint(Session("BudgetID")) & "," & cint(Session("VersionID")) & "," & clng(Session("CostCentreID")) & ",'" & Session("TransactionType") & "'," & Request.Form("" & z & "") & "," & Session("UserID") & ""
		            'Response.Write "spStaffingParametersUpdate '" & Request.Form("" & x & "") & "'," & cint(Session("BudgetID")) & "," & cint(Session("VersionID")) & "," & clng(Session("CostCentreID")) & ",'" & Session("TransactionType") & "'," & Request.Form("" & z & "") & "," & Session("UserID") & ""
        		    strMessage = "Record sucessfully executed!"
        		           		    
        		End If
    		
	        Next    
    
    Else
    
        strMessage = "Data can only be saved in Budget Mode!"
    
    End If   	
    
End Sub


Public Sub Delete_Data()

Dim objRS
Dim intTotal


    If Session("ModeID") = 2 or Session("ModeID") = 4 Then

	'Only Delete the record if there is no related record in the StaffData table with values.
	Set objRS = Server.CreateObject("ADODB.Recordset")

    	objRS.Open "SELECT * FROM qryStaffDataCostCentre WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " AND TransactionType = '" & Session("TransactionType") & "' AND StaffingClassificationID = " & lngRecordID & "",objCon	 
	
	IF NOt objRS.EOF Then	
		
		intTotal = objRS("CY") + objRS("BM1") + objRS("BM2")  + objRS("BM3") + objRS("BM4") + objRS("BM5") + objRS("BM6") + objRS("BM7") + objRS("BM8") + objRS("BM9") + objRS("BM10") + objRS("BM11") + objRS("BM12") + objRS("OY1") + objRS("OY2") + objRS("OY3")
		If intTotal = 0 or isnull(intTotal) Then	
	
			objCon.Execute "spStaffingClassificationsDelete " & lngRecordID & "," & Session("UserID") & ""	
			strMessage = objRS("StaffClassificationDesc") & " has been deleted!"
		Else
			strMessage = objRS("StaffClassificationDesc") & " has NOT been deleted! There are values against " & objRS("StaffClassificationDesc") & ".  Click the Close button and remove values on the Employees Schedule first!"
		End if
	Else
		
	End if

			
	objRS.Close	
	
        'objCon.Execute "spStaffingClassificationsDelete " & lngRecordID & "," & Session("UserID") & "" 
	'Response.Write "spStaffingClassificationsDelete " & lngRecordID & "," & Session("UserID") & "" 
        'strMessage = "Staffing Classification record has been deleted!"
    Else
        strMessage = "Data can only be saved in Budget Mode!"
    End If   

End Sub

Public Sub Transfer(strCostCentreTo, lngEmployeeID, lngStaffingClassificationID)
'Procedure to Transfer an Employee to another Cost Centre

    If Session("ModeID") = 2 or Session("ModeID") = 4 Then
	'response.write "spStaffingClassificationsTransfer " & lngStaffingClassificationID & "," & lngEmployeeID & "," & Session("CostCentreID") & "," & strCostCentreTo & "," & Session("UserID") & ""	
	objCon.Execute "spStaffingClassificationsTransfer " & lngStaffingClassificationID & "," & Session("CostCentreID") & "," & strCostCentreTo & "," & Session("UserID") & ""	
	strMessage = lngEmployeeID & " has been transferred!"
		
    Else
        strMessage = "Data can only be saved in Budget Mode!"
    End If   

End Sub


Public Sub CopyData(lngStaffingClassificationID)
'Procedure to Copy an existing Employee as a new record

    If Session("ModeID") = 2 or Session("ModeID") = 4 Then

	objCon.Execute "spStaffingClassificationsCopy " & lngStaffingClassificationID & "," & Session("UserID") & ""	
	strMessage = lngStaffingClassificationID & " has been Copied!"
		
    Else
        strMessage = "Data can only be saved in Budget Mode!"
    End If   

End Sub


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
%>