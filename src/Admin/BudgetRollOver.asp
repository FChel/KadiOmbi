<%@ Language=VBScript%>
<!-- #Include file=../ADOVBS.inc -->
<%	
'Ensure that the page is not locally cached and therefore is a new page on every request.
Response.Expires = -1500
Server.ScriptTimeout = 6000

'Description:	Isidore ASP Template
'Author:		Isidore IT Pty Ltd
'Date:			18 August

'Instantiate Common Page Variables.
Dim objCon
Dim objCmd
Dim objCmd1
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objCmd1 = Server.CreateObject("ADODB.Command")
objCon.Open Session("DBConnection")

'Declare Database field variables.
Dim lngVersionID
Dim lngBudgetID
Dim strVersionName
Dim lngColumnLock
Dim arrMode(2)
Dim arrYesNo(2)
Dim strEstimates
Dim strRollOverMode

arrMode(1) = "Version"
arrMode(2) = "Budget"

arrYesNo(1) = "N"
arrYesNo(2) = "Y"

	'Setting the VersionID variable
	If Not IsEmpty(Request.QueryString("VersionID")) Then	
		lngVersionID = Request.QueryString("VersionID")					
	End If
	
	'Setting the BudgetID variable
	If Not IsEmpty(Request.QueryString("BudgetID")) Then	
		lngBudgetID = Request.QueryString("BudgetID")					
	End If

    If Not IsEmpty(Request.QueryString("Mode")) Then	
		strRollOverMode = Request.QueryString("Mode")					
	End If

	'Deleting the record
	If Request.QueryString("Action") = "Delete" Then	
		DeleteRecord()
		lngVersionID = 0	
	End If
		
	'Call Save function and save record to the database.	
	If Request.QueryString("Action") = "Save" Then				
		If Not isNull(lngVersionID) Then
			SaveRecord()
		Else			    						 	     
			strMessage = "Budget has not been rolled over!" 	    
	    End If					        
	End If
				
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
	<script src="../formChek.js">
	</script>
	<script src="../JavaDateCheck.js">
	</script>
	
	
<script LANGUAGE="javascript">
<!--
function SaveData(){
	
var varSubmit = true
var varAlert =""
var versionID;

    if (isWhitespace(frm.Mode.value) || frm.Mode.value == "0") {
        varAlert += "Rollover Mode Cannot Be Blank. \n \n";
        document.getElementById('Mode').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('Mode').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.Estimates.value) || frm.Estimates.value == "0") {
        varAlert += "Rollover Estimates Cannot Be Blank. \n \n";
        document.getElementById('Estimates').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('Estimates').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.BudgetOld.value) || frm.BudgetOld.value == "0") {
        varAlert += "Budget Being Rolled over Cannot Be Blank. \n \n";
        document.getElementById('BudgetOld').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('BudgetOld').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.BudgetNew.value) || frm.BudgetNew.value == "0") {
        varAlert += "Budget New Cannot Be Blank. \n \n";
        document.getElementById('BudgetNew').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('BudgetNew').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.VersionOld.value) || frm.VersionOld.value == "0") {
        varAlert += "Version Being Rolled over Cannot Be Blank. \n \n";
        document.getElementById('VersionOld').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('VersionOld').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.VersionNew.value) || frm.VersionNew.value == "0") {
        varAlert += "Version New Cannot Be Blank. \n \n";
        document.getElementById('VersionNew').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('VersionNew').style.backgroundColor = "ffffff";
    
       
	
	   	  			
	  if(varSubmit == true){
	    document.getElementById('Span1').style.display = "inline";
		frm.submit();
		}
	  
	  else{
		window.alert ("" + varAlert);
		}
}


-->
</script>
</head>
<body>

<form action="BudgetRollover.asp?Action=Save" method="POST" id="frm" name="frm">

<table class="custom_table" WIDTH="100%" CELLSPACING="1" CELLPADDING="1" >
	<tr>
		<th class="custom_table_th_4" style="width:20%" Height="25px">&nbsp;Budget Rollover Screen</th>
		<th Width="50%"  class="custom_table_th_5" align="left" style="width:80%"></th>
	</tr>
	<tr><td colspan="2">&nbsp</td></tr>
	<tr>
	<th style="height:25px; width:15%; text-align:left">&nbsp;Rollover Mode</th>
		<td width="15%">
		    <select Style="Width:15%" tabindex="1" id="Mode" name="Mode" onchange="self.location='BudgetRollOver.asp?Mode=' + frm.Mode.value"><OPTION Value=0>Please Select..</OPTION>
		    <%
		        For x = 1 to 2
			        If arrMode(x) = strRolloverMode Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrMode(x) & """" & strSelected & ">" & arrMode(x) & "</OPTION>"
		        Next
	        %>
	      </select>
	    </td>
     </tr>
     <tr>
	<th style="height:25px; width:15%; text-align:left">&nbsp;Rollover Estimates</th>
		<td width="15%">
		    <select Style="Width:15%" tabindex="1" id="Estimates" name="Estimates" ><OPTION Value="0">Please Select..</OPTION>
		    <%
		        For x = 1 to 2
			        If arrYesNo(x) = strEstimates Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrYesNo(x) & """" & strSelected & ">" & arrYesNo(x) & "</OPTION>"
		        Next
	        %>
	      </select>
	    </td>
     </tr>
     <tr><td colspan="2">&nbsp;</td></tr>
	 <tr>
		<th class="custom_table_th_2" height="20px">Budget being Rolled Over</th>			
		<td Width="40%" class="custom_table_td_1"><select Style="Width:15%" tabindex="3" id="BudgetOld" name="BudgetOld"><option value="0">--</option>			
		
		<%
		 'Open a recordset and build list using recordset details.			
		objRS.Open "SELECT * FROM tblBudget",objCon
		
		x = 1

		Do until objRS.eof	
		
			If lngBudgetIDOld = clng(objRS("BudgetID"))Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & objRS("BudgetID") & """" & strSelected & ">" & objRS("BudgetName") & "</option>"

		objRS.movenext
			
		Loop
		
		objRS.Close
		
	%>
		
	</select></td>
    </tr>      
	
   <tr>
		<th class="custom_table_th_2" height="20px">Budget New</th>			
		<td Width="40%" class="custom_table_td_1"><select Style="Width:15%" tabindex="3" id="BudgetNew" name="BudgetNew"><option value="0">--</option>			
		
		<%
		 'Open a recordset and build list using recordset details.	
        If strRollOverMode = "Budget" Then		
		    objRS.Open "SELECT * FROM tblBudget WHERE BudgetID <> " & Session("BudgetID") & "",objCon
        Else
            objRS.Open "SELECT * FROM tblBudget",objCon
        End iF
		
		x = 1

		Do until objRS.eof	
		
			If lngBudgetIDNew = clng(objRS("BudgetID"))Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & objRS("BudgetID") & """" & strSelected & ">" & objRS("BudgetName") & "</option>"

		objRS.movenext
			
		Loop
		
		objRS.Close
		
	%>
		
	</select></td>
    </tr>		        
</table>
<br>
<table class="custom_table" WIDTH="100%" CELLSPACING="1" CELLPADDING="1" >
	<tr>
		<th class="custom_table_th_4" height="20px" style="width:20%">&nbsp;Version of Budget Rolling From</th>
		<th class="custom_table_th_5"  style="width:80%"></th>
	
		
	</tr>
	<tr><td colspan="2">&nbsp</td></tr>
	 <tr>
		<th class="custom_table_th_2" height="20px">Version being Rolled Over</th>			
		<td Width="40%" class="custom_table_td_1"><select Style="Width:15%" tabindex="3" id="VersionOld" name="VersionOld"><option value="0">--</option>			
		
		<%
		 'Open a recordset and build list using recordset details.			
		objRS.Open "SELECT * FROM tblVersion where BudgetID = " & Session("BudgetID") & "",objCon
		
		x = 1

		Do until objRS.eof	
		
			If lngVersionIDOld = clng(objRS("VersionID"))Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & objRS("VersionID") & """" & strSelected & ">" & objRS("VersionName") & "</option>"

		objRS.movenext
			
		Loop
		
		objRS.Close
		
	%>
		
	</select></td><td colspan="2"></td>

    </tr>      
	
   <tr>
		<th class="custom_table_th_2" height="20px">Version New</th>			
		<td Width="40%" class="custom_table_td_1"><select Style="Width:15%" tabindex="4" id="VersionNew" name="VersionNew"><option value="0">--</option>			
		
		<%
		 'Open a recordset and build list using recordset details.
        If strRollOverMode = "Version" Then
            objRS.Open "SELECT * FROM tblVersion where BudgetID = " & Session("BudgetID") & "",objCon 
        Else			
		    objRS.Open "SELECT * FROM tblVersion where BudgetID = " & Session("BudgetID") + 1 & "",objCon 
		End If
		x = 1

		Do until objRS.eof	
		
			If lngVersionIDNew = clng(objRS("VersionID"))Then
				strSelected = " SELECTED " 
			Else
				strSelected = ""	
			End If
			
			Response.Write "<option Value=""" & objRS("VersionID") & """" & strSelected & ">" & objRS("VersionName") & "</option>"

		objRS.movenext
			
		Loop
		
		objRS.Close
		
	%>
		
	</select></td><td colspan="2"></td>

	
    </tr>		        
</table>
<br>
<HR>
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/wrench.png" alt="" /> Rollover </button></td>  
		<TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=UCASE(strMessage) %></TD>
        <td Width="300px"><span id="Span1" style="display:none"><img src=../Images/progress.gif />  &nbsp;&nbsp;&nbsp; <b></FONT>PLEASE WAIT BUDGET IS BEING ROLLED OVER...</b></span></td>
</tr>
</table>
</form>


<BR>
</body>

</html>

<% 

Sub SaveRecord()

            'Rollover Budget
	       
	        With objCmd
                .CommandType = 4
                .CommandTimeout = 6000
                .CommandText = "spRollOver"                
                
                .Parameters.Append objCmd.CreateParameter("BudgetID1", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetID2", adInteger, adParamInput)
		        .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
		        .Parameters.Append objCmd.CreateParameter("VersionIDNew", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Mode", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("RolloverInput", adVarChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)               
               	   		   
    		    .Parameters("BudgetID1") = clng(Request.Form("BudgetOld"))
    		    .Parameters("BudgetID2") = clng(Request.Form("BudgetNew"))
    		    .Parameters("VersionID") = clng(Request.Form("VersionOld"))
    		    .Parameters("VersionIDNew") = clng(Request.Form("VersionNew"))
                .Parameters("Mode") = cstr(Request.Form("Mode"))
                .Parameters("RollOverInput") = cstr(Request.Form("Estimates"))
            	.Parameters("UpdatedBy") = Session("UserID") 

               ' Response.Write clng(Request.Form("BudgetOld"))
                'Response.Write clng(Request.Form("BudgetNew"))
                'Response.Write clng(Request.Form("VersionOld"))
                'Response.Write clng(Request.Form("VersionNew"))
                'Response.Write cstr(Request.Form("Mode"))
               ' Response.Write cstr(Request.Form("Estimates"))
            					 				 				 				 
		.ActiveConnection = objCon
				
		objCmd.Execute	
			    		    
           End With 
			
            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
            strMessage = "<B>BUDGET ROLLED OVER.</B>"
     		
     	
     		
End Sub	

Set objRS = Nothing
Set objCon = Nothing
%>



