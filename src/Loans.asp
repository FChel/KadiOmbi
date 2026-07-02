    <%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<!-- #Include file=ADOVBS.inc -->

<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If IsEmpty(Session("BudgetID")) Then Response.Redirect("Timeout.asp")

Session("InputSheetID") = Session("PLReportID")




'Description:	Data Entry for General Expenses data
'Author:		MG
'Date:			Janaury 2014

	Response.Expires = -1500	

    Session("TransactionType") = "GEXP"

    If IsEmpty(Session("Level1ID")) Then Session("Level1ID") = 0 End If 
    If IsEmpty(Session("StrategicObjectiveID")) Then Session("StrategicObjectiveID") = 0 End If    
    If IsEmpty(Session("TargetID")) Then Session("TargetID") = 0 End If  
    If IsEmpty(Session("ActivityID")) Then Session("ActivityID") = 0 End If  

Dim objCon
Dim objRS
Dim objRS1
Dim objRS2
Dim objCmd

Dim strForeColour
Dim intMode
Dim dblTotal
Dim intFinYearPart1
Dim intFinYearPart2
Dim arrHeadings(5)
Dim strCostCentreName
Dim strVersionName
Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL

Dim intEmployeeID
Dim intCalculationMethodID
Dim intLocationID
Dim intCars
Dim intDays
Dim dblEmployeeContribution
Dim dteParkDate
Dim strDeclaration
Dim strNotes
Dim dteEndDate

'Set objIsidore = Server.CreateObject("ProcessFormulas.cProcessFormula")    

 'Set Headings
    For x = 0 to 4
	
	    intFinYearPart1 = cint(Session("FinancialYear")) + (x - 2)
	    intFinYearPart1 = Right(intFinYearPart1,2)
	    intFinYearPart2 = cint(Session("FinancialYear")) + x - 1
	    intFinYearPart2 = Right(intFinYearPart2,2)

	    arrHeadings(x) = cstr(intFinYearPart1) & "/" & cstr(intFinYearPart2)

    Next
    

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objRS1 = Server.CreateObject("ADODB.Recordset")
    Set objRS2 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")

    objCon.Open Session("DBConnection")	

    Session("CurrentPage") = "FBT/Loans.asp"


If Not IsEmpty(Request.QueryString("Level1ID")) Then
	Session("Level1ID") = Request.QueryString("Level1ID")
End If

If Not IsEmpty(Request.QueryString("EmployeeID")) Then
	Session("EmployeeID") = Request.QueryString("EmployeeID")
	Session("CarParkingID") = 0
End If


If Not IsEmpty(Request.QueryString("TransactionType")) Then
	Session("TransactionType") = Request.QueryString("TransactionType")
End If
 Session("InputSheetID") = 1
 

If Not IsEmpty(Request.QueryString("CarParkingID")) Then
	Session("CarParkingID") = Request.QueryString("CarParkingID")
End If

'Execute Action
If Request.QueryString("Action") = "Save" Then

   If Session("StatusID") = 1 Then
        'Do not allow Read Only users to make changes
        If Session("UserTypeID") = 4 Then
            strMessage = "NOT SAVED. You are a READ ONLY User and cannot make any changes."
        Else
            Call SaveCarParking()
        End If
   Else
        strMessage = "Budget is closed, no changes can be made!"
   End If
End If

If Request.QueryString("Action") = "Delete" Then
    If Session("StatusID") = 1 Then
        Call DeleteData(Request.QueryString("GeneralExpenseID"))
    Else
          Response.Write "&nbsp;&nbsp;<img src=""../images/warning.gif"" /><B><FONT Color=""Red"">&nbsp;&nbsp;WARNING - BUDGET IS NOT OPEN, CHANGES CANNOT BE MADE.</FONT></B>" 
          strMessage = "<FONT Color=""Red""><B>BUDGET IS NOT OPEN, CHANGES CANNOT BE MADE.</B></FONT>"
          strMessageIcon = "<img src=""../images/warning.gif"" />"
    End If
End If

    objRS.Open "Select CostCentreID,CostCentreName,ProgramCode from tblCostCentres where CostCentreID = " & clng(Session("CostCentreID")) & " AND BudgetID = " & clng(Session("BudgetID")),objCon,0,1

        If Not objRS.EOF Then
            strCostCentreName = objRS("CostCentreID") & " (" & objRS("CostCentreName") & ")" 
            Session("Segment2") = objRS("ProgramCode")
        Else
            strCostCentreName = ""
            Session("Segment2") = "0"
        End If

    objRS.Close

    objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & "",objCon,0,1

        If not objRS.EOF Then
            strVersionName = objRS("VersionName")
        End If

    objRS.Close

  
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "New" Then
			'Response.Write "CPID=" & Session("CarParkingID")
			Session("CarParkingID") = 0
		End If
	End If
	
	If isNull(Session("CarParkingID")) Or Session("CarParkingID") = "" Then 
		Session("CarParkingID") = 0
	End If
	
  Call LoadDetails()
  
%>

<html>
<head>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

 <!-- Bootstrap Core CSS -->
    <link href="css/bootstrap.min.css" rel="stylesheet">
	
<link rel="stylesheet" type="text/css" href="BertStyle.css">
<script src="formChek.js"></script>
<link rel="stylesheet" type="text/css" href="tcal.css" />
	<script type="text/javascript" src="tcal.js"></script> 
	<script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-3.3.1.min.js"></script>
<script LANGUAGE="javascript">

function FlagSaveStatus(x){
	
	var Row
	var Col
	var Total
	var InputBox = x
	parts = InputBox.split(".");
	Row = parts[0];
	Col = parts[1];
	
	var varElem1 = (parseInt(Row)) + parseInt(Col);
	var varElem2 = 100000 + parseInt(Row/100);
	var varElem3 = parseInt(Row) + 13;
	
		document.getElementById(varElem2).value = 'Y';
			
		GetTotal(x);
    
		var mynum = document.getElementById(varElem1).value;
		mynum = mynum.replace(/,/g, "");
		mynum = FormatNumber(mynum, 0, 0, 0, 1);
		document.getElementById(varElem1).value = mynum;
		
}

function Distribute(x){	
   
	var Row
	var Col
	var Total
	var InputBox = x
	var Amount
	parts = InputBox.split(".");
	Row = parts[0];
	Col = parts[1];
	
	var varElem3 = parseInt(Row) + 1;
	var Amount = (document.getElementById(varElem3).value);
	var str4 = Amount
	var Amount = str4.replace(/,/g, "");

	for (i=1;i<13;i++){
	varElem3 = parseInt(Row) + i;
	document.getElementById(varElem3).value = Math.round((Amount/12))
	}
	varElem3 = parseInt(Row) + 12;
	document.getElementById(varElem3).value =  (Amount - (Math.round((Amount/12))*11))
	
	FlagSaveStatus(x);
	GetTotal(x);
}

function SaveData(){
	var varSubmit = true
		frm.msgbox.value='Saving.......';
		frm.submit();
	//}
}


function GetTotal(x){
   
	var Row
	var Col
	var TotalR = 0
	var TotalC = 0
	var InputBox = x
	var RowCount = frm.RowCount.value
	parts = InputBox.split(".");
	Row = parts[0];
	Col = parts[1];
	
	
	var varElem1 
	var varElem2
	var varElem3
	
	for (i=1;i<13;i++){
	varElem1 = parseInt(Row) + i;		
		
	TotalR = TotalR + parseInt(document.getElementById(varElem1).value)
	}
		varElem3 =  parseInt(Row/100) + 200000;
        
		document.getElementById(varElem3).value = TotalR
		
	for (i=1;i<13;i++){Col = i;
		TotalC = 0
		for (j=1;j<RowCount;j++){varElem1 = (j*100) + Col;
		   
			TotalC = TotalC + parseInt(document.getElementById(varElem1).value)
			varElem3 =  parseInt(Col);
		
			//document.getElementById(varElem3).value = TotalC
		}
	}
}

function CloseScreen() {

    //if(top.Header1.Header2.document.form1.SaveStatus.value=='S')
    //{var x=window.confirm("Changes have been made, do you wish to save these changes?")
    //    if (x){
    //        SaveData();
    //        self.location='index.asp';
    //    }
    //    else
    //        self.location='index.asp';}
    //    else
    { self.location = '../home.asp'; }
}

setTimeout( 'ShowTimeoutWarning();', 1080000 );

function ShowTimeoutWarning () {     
    window.alert( "********** Warning! **********' \n \n 'You will be automatically logged out in 2 minutes unless you change screens, Close or Save!" ); 
}


function FormatNumber(num, decimalNum, bolLeadingZero, bolParens, bolCommas)
    /**********************************************************************
        IN:
            NUM - the number to format
            decimalNum - the number of decimal places to format the number to
            bolLeadingZero - true / false - display a leading zero for
                                            numbers between -1 and 1
            bolParens - true / false - use parenthesis around negative numbers
            bolCommas - put commas as number separators.
     
        RETVAL:
            The formatted number!
     **********************************************************************/ {
    if (isNaN(parseInt(num))) return "NaN";

    var tmpNum = num;
    var iSign = num < 0 ? -1 : 1;		// Get sign of number

    // Adjust number so only the specified number of numbers after
    // the decimal point are shown.
    tmpNum *= Math.pow(10, decimalNum);
    tmpNum = Math.round(Math.abs(tmpNum))
    tmpNum /= Math.pow(10, decimalNum);
    tmpNum *= iSign;					// Readjust for sign


    // Create a string object to do our formatting on
    var tmpNumStr = new String(tmpNum);

    // See if we need to strip out the leading zero or not.
    if (!bolLeadingZero && num < 1 && num > -1 && num != 0)
        if (num > 0)
            tmpNumStr = tmpNumStr.substring(1, tmpNumStr.length);
        else
            tmpNumStr = "-" + tmpNumStr.substring(2, tmpNumStr.length);

    // See if we need to put in the commas
    if (bolCommas && (num >= 1000 || num <= -1000)) {
        var iStart = tmpNumStr.indexOf(".");
        if (iStart < 0)
            iStart = tmpNumStr.length;

        iStart -= 3;
        while (iStart >= 1) {
            tmpNumStr = tmpNumStr.substring(0, iStart) + "," + tmpNumStr.substring(iStart, tmpNumStr.length)
            iStart -= 3;
        }
    }

    // See if we need to use parenthesis
    if (bolParens && num < 0)
        tmpNumStr = "(" + tmpNumStr.substring(1, tmpNumStr.length) + ")";

    return tmpNumStr;		// Return our formatted string!
}

function DeleteData(GEXPID) {
   
        if (window.confirm('Would you like to DELETE the selected record?') == true) {

            self.location = "Loans.asp?Action=Delete&GeneralExpenseID=" + GEXPID;
        }
        
}

    $(function(){           
        if (!Modernizr.inputtypes.date) {
            $('input[type=date]').datepicker({
                  dateFormat : 'yy-mm-dd'
                }
             );
        }
    });
	
	$(function(){
    $('#login').popover({
       
        placement: 'bottom',
        title: 'Popover Form',
        html:true,
        content:  $('#myForm').html()
    }).on('click', function(){
      // had to put it within the on click action so it grabs the correct info on submit
      $('.btn-primary').click(function(){
       $('#result').after("form submitted by " + $('#email').val())
        $.post('Stalls.asp',  {
            email: $('#email').val(),
            name: $('#name').val(),
            gender: $('#gender').val()
        }, function(r){
          $('#pops').popover('hide')
          $('#result').html('resonse from server could be here' )
        })
      })
  })
})

$(function(){
    $('.pops').popover({
       
        placement: 'bottom',
        title: 'Enter Details and Click Save',
        html:true,
        content:  $('#myForm').html()
    }).on('click', function(){
      // had to put it within the on click action so it grabs the correct info on submit
      $('.btn-primary').click(function(){
       $('#result').after("form submitted by " + $('#email').val())
        $.post('Stalls.asp?Action=Save',  {
            email: $('#email').val(),
            name: $('#name').val(),
            phone: $('#phone').val()
        }, function(r){
          $('#pops').popover('hide')
          $('#result').html('resonse from server could be here' )
        })
      })
  })
})

$(document).ready(function() {
    $("a").click(function(event) {
        //alert(event.target.id);
		//myForm.popForm.ID='dfdf'//event.target.id
		//document.getElementById('PopID').value = 0
		//$('#PopID').val=90
		//$('#PopID').val( "hello world" );
		$('input[name="PopID"]').val(event.target.id);
		$('input[name="FieldName"]').val(event.target.name);
    });
});
</script>


   

   
	
	
</head>
<body >

<hr Width="1760px">



               
                <!-- /.row -->

<div id="myForm" class="hide">
    <form action="Loans.asp?Action=Save" id="popForm" method="get">
        <div>
		<label>Please enter your DECLARATION details below.  This will appply the declaration to the expense you have entered.</label>
            <label for="name">First Name:</label>
            <input type="text" name="FirstName" id="FirstName" class="form-control input-md">
			<label for="name">Last Name:</label>
            <input type="text" name="LastName" id="LastName" class="form-control input-md">
			<label for="email">Email:</label>
            <input type="email" name="email" id="email" class="form-control input-md">
            <label for="phone">Phone:</label>
            <input type="text" name="phone" id="phone" class="form-control input-md">
			<label for="phone">Group:</label>
            
			<select class="form-control" id="Group">
			  <option>Please select your group...</option>
			  <option>Army</option>
			  <option>Navy</option>
			  <option>Air Force</option>
			  <option>CIOG</option>
			  <option>CFOG</option>
			  <option>DSTO</option>
			  <option>JOC</option>
			</select>
			<label for="ID"></label>
            <input type="hidden" name="PopID" id="PopID" class="form-control input-md">
			<label for="ID"></label>
            <input type="hidden" name="FieldName" id="FieldName" class="form-control input-md">
			<label for="ID"></label>
            <input type="hidden" name="Action" id="Action" class="form-control input-md" value="Save">
            <button type="submit" class="btn btn-primary" data-loading-text="Sending info.."><em class="icon-ok"></em> Save</button>
        </div>
    
</div>
<div id="result"></div>


<!-- Button trigger modal -->
<button type="button" class="btn btn-primary" data-toggle="modal" data-target="#exampleModalCenter">
  Launch demo modal
</button>

<!-- Modal -->
<div class="modal fade" id="exampleModalCenter" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">Loans Declaration Form</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>Please enter your DECLARATION details below.  This will appply the declaration to the expense you have entered.</label>
            <label for="name">First Name:</label>
            <input type="text" name="FirstName" id="FirstName" class="form-control input-md">
			<label for="name">Last Name:</label>
            <input type="text" name="LastName" id="LastName" class="form-control input-md">
			<label for="email">Email:</label>
            <input type="email" name="email" id="email" class="form-control input-md">
            <label for="phone">Phone:</label>
            <input type="text" name="phone" id="phone" class="form-control input-md">
			<label for="phone">Group:</label>
            
			<select class="form-control" id="Group">
			  <option>Please select your group...</option>
			  <option>Army</option>
			  <option>Navy</option>
			  <option>Air Force</option>
			  <option>CIOG</option>
			  <option>CFOG</option>
			  <option>DSTO</option>
			  <option>JOC</option>
			</select>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
        <button type="button" class="btn btn-primary">Save changes</button>
      </div>
    </div>
  </div>
</div>


<!--<div id='tbl-container'>-->
<form action="Loans.asp?Action=Save" method="POST" id="frm" name="frm">
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr>
    <th Style="height:20px">FBT - Loans</th>
</tr>
</table>
</br>
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">

<tr>
    <th Style="width:20%;">Employee</th><td style="width:30%; background-color:FFFFFF;height:20px;"><select style="width:100%; background-color:FFFFFF;" tabindex="2" id="EmployeeID" name="EmployeeID" onchange="self.location='Loans.asp?EmployeeID=' + frm.EmployeeID.value">
      
   <%	
    If Session("UType") = "Manager" Then
		Response.write "<option value=""0"">Show All</option>"
	End If
        objRS.Open "SELECT * FROM tblStaffingClassifications WHERE BudgetID = " & Session("BudgetID") & " AND CostCentreID = " & Session("CostCentreID") & " AND Deleted = 'N' Order By StaffClassificationDesc",objCon,0,1
   
    	Do until objRS.EOF
			If clng(objRS("StaffingClassificationID")) = clng(Session("EmployeeID")) Then
				strSelected = " SELECTED "
			
			Else
				strSelected = ""
			End if
				 
				If Session("UType") = "Employee" Then
					If clng(objRS("StaffingClassificationID")) = clng(Session("StaffingClassificationID")) Then
				Response.Write "<option Value=""" & objRS("StaffingClassificationID") & """" & strSelected & ">" & objRS("StaffingClassificationID") & " - " & objRS("StaffClassificationDesc") & "</OPTION>"
					End If
				
				Else
			    Response.Write "<option Value=""" & objRS("StaffingClassificationID") & """" & strSelected & ">" & objRS("StaffingClassificationID") & " - " & objRS("StaffClassificationDesc") & "</OPTION>"
				End If
			objRS.Movenext
		Loop
		
		objRS.Close
		
		
     %>
    
</select></td><td style="width:50%;"></td>
</tr>
</table>
<br />
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
<tr>

<th>Statutory Interest Rate (Benchmark)</th><td style="background-color:FFFFFF;">5.65%</td>
<td style="text-align:left; width:10%;"></td>
<th>Outstanding Loan Amount</th><td style="text-align:right; width=90%; background-color:FFFFFF;"><INPUT Name="Cars" ID="Cars" Value="<%=intCars%>" style="width=90%; background-color:FFFFFF;"></td>

</tr><tr>
<th>Notes</th><td style="background-color:FFFFFF; text-align:left; width=90%;"><INPUT Name="Notes" ID="Notes" Value="<%=strNotes%>" style="width:100%; background-color:FFFFFF; text-align:left;"></td>
<td></td>
    <th>Employee Contribution</th><td style="text-align:right; width=90%; background-color:FFFFFF;"><INPUT Name="EmployeeContribution" ID="EmployeeContribution" Value="<%=dblEmployeeContribution%>" style="width=90%; background-color:FFFFFF;"> </td>
</tr>
<tr Style="height:20px">
<td></td><td><INPUT Name="GLCodeID" ID="GLCodeID" Value="<%=Session("GLCodeID")%>"></td><td><INPUT Name="CarParkingID" ID="CarParkingID" Value="<%=Session("CarParkingID")%>" ></td>

<th>Days</th><td style="text-align:right; width=90%; background-color:FFFFFF;"><INPUT Name="Days" ID="Days" Value="<%=intDays%>" style="width=90%; background-color:FFFFFF;"/> </td>
</tr><tr Style="height:20px;">

<th style="background-color:#CEECF5; font-size:12;" colspan="2">Declaration: I confirm that I have incurred this expense as per ATO legislation</th><td><INPUT Type="Checkbox" Name="Declaration" ID="Declaration" Value="<%=strDeclaration%>" data-toggle="modal" data-target="#exampleModalCenter"><a class="pops" id="1" name="2" href="#">Add</a></td>
<th>Commencement Date</th><td style="text-align:right; width=90%; background-color:FFFFFF;"><INPUT class="tcal" Name="ParkDate" ID="ParkDate" Value="<%=dteParkDate%>" ></td>
</tr>
<tr Style="height:20px;">
<td></td><td></td><td></td>
<th>End Date</th><td style="text-align:right; width=90%; background-color:FFFFFF;"><INPUT class="tcal" Name="EndDate" ID="EndDate" Value="<%=dteEndDate%>" ></td>

</table>
</br>

<br />
<span style="font-size:12px; font-weight:bold;">
Conditions for FBT:</span><span style="font-size:12px;">
<br> You provide a loan fringe benefit if you give your employee a loan and charge no interest or a low rate of interest. A low rate of interest is one that is less than the benchmark interest rate.
<br /></span>
<br />
<br />

<TABLE Width="1200px" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<TR>

    <td class='locked' Width="100px" style="border-right:0px"><button type="button" onclick="CloseScreen()";><img src="images/door.png" alt="" /> Close </button></td>    
    <td class='locked' Width="100px" style="border-right:0px"><button type="button" onclick="SaveData()";><img src="images/tick.png" alt="" /> Save </button></td>   
	<td class='locked' Width="100px" style="border-right:0px"><button type="button" onclick="self.location='Loans.asp?Action=New'";><img src="images/add.png" alt="" /> New </button></td> 
	<td class='locked' Width="100px" style="border-right:0px"><button type="button" onclick="self.location='Loans.asp?Action=Filter&S6Filter=' + frm.S6Filter.checked + '&S7Filter=' + frm.S7Filter.checked + '&S8Filter=' + frm.S8Filter.checked + '&S9Filter=' + frm.S9Filter.checked + '&S10Filter=' + frm.S10Filter.checked + '&S11Filter=' + frm.S11Filter.checked + '&S12Filter=' + frm.S12Filter.checked + '&S13Filter=' + frm.S13Filter.checked";><img src="images/funnel.png" alt="" /> Filter </button></td> 
	<td class='locked' Width="100px" style="border-right:0px"><button type="button" ><img src="images/disk.png" alt="" /> Upload </button></td> 
    <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
    <TD class='locked' align="left" Width="800x" style="BORDER-RIGHT:0px"><INPUT style="Align:Left; font-weight:Bold; width:100%; text-align:left; color:<%=strMessageColour%>;" type="text" id="msgbox" name="msgbox" value="<%=strMessage%>"></TD>
    
</TR>
</TABLE>
<hr />
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">

	<% 
	Response.Write "<tr><TH Style=""Width:5%;"">Loan ID</TH><th Style=""Width:20%; Height:20px"">Cost Centre</th><Th Style=""Width:10%;"">Employee</TH>"& _
			"<Th Style=""Width:5%;"">Statutory Interest Rate</TH><Th Style=""Width:5%;"">Loan Amount</TH><Th Style=""Width:5%;"">Days</TH>" & _
			"<Th Style=""Width:5%;"">Employee Contribution</TH><Th Style=""Width:5%;"">Date</TH></tr>"
      
        
    
             %>   

<tbody>
<%
        
      DisplayTableDetails()
        
%>	
</tbody>
</table>





<!--</DIV>-->
</form>


    <!-- jQuery -->
    <script src="js/jquery.js"></script>

    <!-- Bootstrap Core JavaScript -->
    <script src="js/bootstrap.min.js"></script>

	
	
</body>
</html>
<%

Public Sub DisplayTableDetails()
Dim y
Dim dblEmpCont
Dim intDaysTotal
Dim intCarsTotal
Dim intEmpContTotal

If Session("EmployeeID") = 0 Then
	strSQL = "SELECT * FROM qryCarParking WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " AND BenefitType = 'Loan'"
Else
	strSQL = "SELECT * FROM qryCarParking WHERE EmployeeID = " & Session("EmployeeID") & " AND BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " AND BenefitType = 'Loan'"
End If

objRS.Open strSQL,objCon
    y = 1
    	
    Do until objRS.EOF 
		If isNull(objRS(9)) Then
			dblEmpCont = 0
		Else
			dblEmpCont = objRS(9)
		End If

		response.write "<TR><TD style=""text-align:center;""><a Target=""_self"" HREF=""Loans.asp?CarParkingID=" & objRS(0) & """>" & objRS(0) & "</a></TD>" & _
				"<TD><a Target=""_self"" HREF=""Loans.asp?CarParkingID=" & objRS(0) & """>" & objRS(3) & "</a></TD><TD><a Target=""_self"" HREF=""Loans.asp?CarParkingID=" & objRS(0) & """>" & objRS(4) & "</a></TD><TD style=""text-align:center;"">5.65%</TD>" & _
				"<TD style=""text-align:center;"">" & objRS(7) & "</TD><TD style=""text-align:center;"">" & objRS(8) & "</TD>" & _
				"<TD style=""text-align:center;"">" & objRS(9) & "</TD><TD style=""text-align:center;"">" & objRS(10) & "</TD></TR>"
				
				intDaysTotal = intDaysTotal + objRS("Days")
				intCarsTotal = intCarsTotal + objRS("Cars")
				intEmpContTotal = intEmpContTotal + objRS("EmployeeContribution")
				
		objRS.movenext
	Loop
	
	
	response.write "<TR><TH colspan=""4"">Total</TH>" & _
				"<TH style=""text-align:center;"">" & intCarsTotal & "</TH><TH style=""text-align:center;"">" & intDaysTotal & "</TH>" & _
				"<TH style=""text-align:center;"">" & intEmpContTotal & "</TH><TH style=""text-align:center;""></TH></TR>"
				
objRS.Close

End Sub


Sub LoadDetails()

       'Description:	Loads Position details into page if applicable.
		objRS.Open "SELECT * FROM tblCarParking WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND CostCentreID = " & Session("CostCentreID") & " AND CarParkingID = " & Session("CarParkingID") & " AND BenefitType = 'Loan'",objCon

			If Not objRS.EOF Then
                intEmployeeID = objRS("EmployeeID")
				Session("EmployeeID") = objRS("EmployeeID")
                intCalculationMethodID = objRS("CalculationMethodID")
				intLocationID = objRS("LocationID")
				intCars = objRS("Cars")
				intDays = objRS("Days")
                dblEmployeeContribution = objRS("EmployeeContribution")
                dteParkDate = objRS("ParkDate")
				strDeclaration = objRS("Declaration")	
				strNotes = objRS("Notes")	
				
    		Else
				Session("CarParkingID") = 0
			  	intEmployeeID = 0
				'Session("EmployeeID") = 0
                intCalculationMethodID = 0
				intLocationID = 0
				intCars = 0
				intDays =0
                dblEmployeeContribution = 0
                dteParkDate = ""
				strDeclaration = ""
				strNotes = ""
				
           End If

		objRS.Close
	
End Sub



Sub SaveCarParking()
Dim strDeclar

	If isNull(Request.Form("Declaration")) Or Request.Form("Declaration") = "" Then
		strDeclar = "checked"
	Else	
		strDeclar = Request.Form("Declaration") 
	End If

  	With objCmd

			.CommandType = 4
			.CommandText = "spCarParkingSave"

			.Parameters.Append objCmd.CreateParameter("CarParkingID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CostCentreID", adInteger, adParamInput)                
			.Parameters.Append objCmd.CreateParameter("EmployeeID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CalculationMethodID", adInteger, adParamInput) 
			.Parameters.Append objCmd.CreateParameter("LocationID",adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Cars", adDouble, adParamInput) 
			.Parameters.Append objCmd.CreateParameter("Days",adDouble, adParamInput)
			.Parameters.Append objCmd.CreateParameter("EmployeeContribution", adDouble, adParamInput)
			.Parameters.Append objCmd.CreateParameter("ParkDate", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Declaration", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("Notes", adVarChar, adParamInput, 200)
			.Parameters.Append objCmd.CreateParameter("BenefitType", adVarChar, adParamInput, 20)
			
			.Parameters("CarParkingID") = Request.Form("CarParkingID")
			.Parameters("BudgetID") = Session("BudgetID")	
			.Parameters("VersionID") = Session("VersionID")						
			.Parameters("CostCentreID") = Session("CostCentreID")
			.Parameters("EmployeeID") = Request.Form("EmployeeID")
			.Parameters("CalculationMethodID") = 0'Request.Form("CalculationMethodID") 
			.Parameters("LocationID") = 0'Request.Form("LocationID")             
			.Parameters("Cars") = Request.Form("Cars") 
			.Parameters("Days") = Request.Form("Days") 
			.Parameters("EmployeeContribution") = Request.Form("EmployeeContribution") 
			.Parameters("ParkDate") = Request.Form("ParkDate") 
			.Parameters("Declaration") = strDeclar'Request.Form("Declaration") 
			.Parameters("Notes") = Request.Form("Notes") 
			.Parameters("BenefitType") = "Loan"
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
		strMessageColour = "Black"
          

End Sub


Public Function Get_Variance_Colour(dblVariancePercentage,lngGLCode)

Dim strGLType

         Get_Variance_Colour = "Black"
         
         objRS1.Open "SELECT GLCodeType FROM tblGLCodes WHERE GLCode = " & lngGLCode & "",objCon
         
            If Not objRS.EOF Then
                strGLType = objRS1("GLCodeType")
            Else
                strGLType = "E"
            End If
         
         objRS1.Close
         
         If strGlType = "E" Then
        
             If dblVariancePercentage > 10 Then
    	         Get_Variance_Colour = "Orange"
		     End If
    		           	
		     If dblVariancePercentage > 20 Then
    	         Get_Variance_Colour = "Red"
		     End If
    		           	    
		     If dblVariancePercentage < -10 Then
    	         Get_Variance_Colour = "Green"
		     End If
    		           	    
		     If dblVariancePercentage < -20 Then
    	         Get_Variance_Colour = "Green"
		     End If	
		 
		 Else	
            
             If dblVariancePercentage > 10 Then
    	         Get_Variance_Colour = "Green"
		     End If
    		           	
		     If dblVariancePercentage > 20 Then
    	         Get_Variance_Colour = "Green"
		     End If
    		           	    
		     If dblVariancePercentage < -10 Then
    	         Get_Variance_Colour = "Orange"
		     End If
    		           	    
		     If dblVariancePercentage < -20 Then
    	         Get_Variance_Colour = "Red"
		     End If	
		     
		 End If

End Function


Public Function Validate_Access(UserTypeID,Screen)

    If Session("UserTypeID") = 99 Then
        
        Validate_Access = "Y"
        
    Else
        
        objRS.Open "SELECT ScreenID FROM qryScreenAccess WHERE UserTypeID = " & UserTypeID & " AND PageName = '" & Screen & "?TransactionType=" & Session("TransactionType") & "'",objCon

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
