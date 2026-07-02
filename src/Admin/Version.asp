<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
'Description:	Student Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			August 2004

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngVersionID 
Dim strVersionName
Dim lngColumnLock
Dim lngVersionTypeID
Dim strCurrency
Dim strOYIndexType
Dim dblOY1Index
Dim dblOY2Index
Dim dblOY3Index
Dim dblOY4Index
Dim dblOY5Index
Dim strTrackChanges
Dim intPrevBudgetID
Dim intPrevVersionID
Dim intPrevOrigVersionID
Dim intBaseBudgetVersionID
Dim intReportingMonth
Dim strCeilingsOn

'Declare and set default arrays

Dim arrYesNo(2)

	arrYesNo(1) = "N"
	arrYesNo(2) = "Y"
	
Dim arrMonth(12)

    arrMonth(1) = "Jul"
    arrMonth(2) = "Aug"
    arrMonth(3) = "Sep"
    arrMonth(4) = "Oct"
    arrMonth(5) = "Nov"
    arrMonth(6) = "Dec"
    arrMonth(7) = "Jan"
    arrMonth(8) = "Feb"
    arrMonth(9) = "Mar"
    arrMonth(10) = "Apr"
    arrMonth(11) = "May"
    arrMonth(12) = "Jun"    
	
'3. Capture Querystring variables

    If Not IsEmpty(Request.QueryString("VersionID")) Then
        Session("VersionID") = Request.QueryString("VersionID")
	End If
		
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

	'Load page details
	LoadDetails()
	
If Request.QueryString("Action") = "Index" Then
    Build_Index()
End If

If Request.QueryString("Action") = "BuildBB" Then
    Build_Base_Budget()
End If

If Request.QueryString("Action") = "Consolidate" Then
    Consolidate()
End If
		
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0"/>
<link rel="stylesheet" type="text/css" href="../BERTStyle.css"/>
<script type="text/javascript" src="../formChek.js"></script>
<script type="text/javascript" src="../ButtonRollOver.js"></script>
<script type="text/javascript" language="javascript">
<!--
   function SaveData()
    {
        var varSubmit = true
        var varAlert =""  
        
        if((isNonnegativeInteger(frm.VersionID.value)==false) || (frm.VersionID.value == 0))     
        {            
		    varAlert += "Please enter Version ID. Version ID must be a numeric value. \n \n";
		    document.getElementById('VersionID').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('VersionID').style.backgroundColor="ffffff";		    
	
	    if((isNonnegativeInteger(frm.BudgetID.value)==false) || (frm.BudgetID.value == 0))
        {            
		    varAlert += "Please enter Budget ID. Budget ID must be a numeric value. \n \n";
		    document.getElementById('BudgetID').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('BudgetID').style.backgroundColor="ffffff";
	    
	    if(isWhitespace(frm.VersionName.value))
        {            
		    varAlert += "Please enter Version Name. \n \n";
		    document.getElementById('VersionName').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('VersionName').style.backgroundColor="ffffff";
	    
	    if(frm.ColumnLock.value < 0 )
	    {            
		    varAlert += "Please enter Column Lock. Column Lock must not be less than 0.  \n \n";
		    document.getElementById('ColumnLock').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }			
	    else document.getElementById('ColumnLock').style.backgroundColor="ffffff";		    
	
	
	    if(frm.VersionTypeID.value == 0 )
	    {
		    varAlert += "Please select a version type. \n \n";
		    document.getElementById('VersionTypeID').style.backgroundColor="ff8080";
		    varSubmit = false;
	    }	
	    else document.getElementById('VersionTypeID').style.backgroundColor="ffffff";	
	    
	 // if(frm.Currency.value == 0 )
	//    {
	//	    varAlert += "Please select a Currency. \n \n";
	//	    document.getElementById('Currency').style.backgroundColor="ff8080";
	//	    varSubmit = false;
	//    }	
	//    else document.getElementById('Currency').style.backgroundColor="ffffff";		   		  
	   	
	  if(varSubmit == true)
	  {
	        frm.submit();
	  }
	  else
	  {
	    window.alert ("" + varAlert);	    
	  }
	  
    }

   function DeleteData()
    {
 	if(window.confirm('Confirm delete')==true){
	self.location="Version.asp?Action=Delete"
	}
}

function ReIndex(){
    
        document.getElementById('Progress').style.display = "inline";
        self.location='Version.asp?Action=ReIndex';
        
}

function BuildBase(){
    
        document.getElementById('Progress').style.display = "inline";
        self.location='Version.asp?Action=BuildBB';

    }

    function Consolidate() {

        document.getElementById('Progress').style.display = "inline";
        self.location = 'Version.asp?Action=Consolidate';

    }



//-->
</script>
</head>
<body>
<h3>Version Administration Screen</h3>
<form action="Version.asp?Action=Save" method="POST" id="frm" name="frm">

<table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
	<tr>
		<th height="20px" align="Left">&nbsp;Version ID</th>
		<td>&nbsp;<input style="text-align:left;width:90%" id="VersionID" name="VersionID" maxlength="2" tabindex="1" value="<%=Session("VersionID")%>" /></td>
		<th align="left">&nbsp;Budget ID</th>
		<td>&nbsp;<input style="text-align:left;width:90%" id="BudgetID" name="BudgetID" READONLY maxlength="2" tabindex="2" value="<%=Session("BudgetID")%>" /></td>
	</tr>
	<tr>
		<th height="20px" align="left">&nbsp;Version Name</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="VersionName" name="VersionName" maxlength="50" tabindex="3" value="<%=strVersionName%>" /></td>
		<th align="left">&nbsp;Column Lock</th>
	    <td>&nbsp;<input style="text-align:left;width:90%" id="ColumnLock" name="ColumnLock" maxlength="2" tabindex="4" value="<%=lngColumnLock%>" /></td>
	</tr>    
    <tr>
		<th height="20px" align="left">&nbsp;Version Type</th>
		<td><select Style="Width:80%" tabindex="5" id="VersionTypeID" name="VersionTypeID"><option value="0">Please Select..</OPTION>
	<%	
		objRS.Open "SELECT * FROM tblVersionTypes WHERE Active = 'Y'",objCon
		
		Do until objRS.EOF
			If objRS("VersionTypeID") = lngVersionTypeID Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("VersionTypeID") & """" & strSelected & ">" & objRS("VersionTypeName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		
	%></select></td>
	
	
	<th height="20px" align="left">&nbsp;Currency</th>
		<td><select Style="Width:35%" tabindex="5" id="Currency" name="Currency"><option value="0">Please Select..</OPTION>
	<%	
		objRS.Open "SELECT DISTINCT(Currency) FROM tblCurrencies WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon
		
		Do until objRS.EOF
			If objRS("Currency") = strCurrency Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("Currency") & """" & strSelected & ">" & objRS("Currency") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		
	%></select></td>
	</tr>
	
	 <tr>
		<th height="20px" align="left">&nbsp;Base Budget Version</th>
		<td><select Style="Width:80%" tabindex="13" id="BaseBudgetVersionID" name="BaseBudgetVersionID"><option value="0">Please Select..</OPTION>
	<%	
		objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & "",objCon
		
		Do until objRS.EOF
			If cint(objRS("VersionID")) = cint(intBaseBudgetVersionID) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("VersionID") & """" & strSelected & ">" & objRS("VersionName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		
	%></select></td>
		<th align="left">&nbsp;Track Changes</th>
		<td>
	       <select Style="Width:40%" tabindex="14" id="TrackChanges" name="TrackChanges">
	        <%
		        For x = 1 to 2
			        If arrYesNo(x) = cstr(strTrackChanges) Then
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
	<tr>
		<th height="20px" align="left">&nbsp;Previous Budget</th>
		<td><select Style="Width:80%" tabindex="13" id="PrevBudgetID" name="PrevBudgetID"><option value="0">Please Select..</OPTION>
	<%	
		objRS.Open "SELECT * FROM tblBudget",objCon
		
		Do until objRS.EOF
			If objRS("BudgetID") = intPrevBudgetID Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("BudgetID") & """" & strSelected & ">" & objRS("BudgetName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		
	%></select></td>
	<th align="left">&nbsp;Previous Version</th>
		<td><select Style="Width:80%" tabindex="13" id="PrevVersionID" name="PrevVersionID"><option value="0">Please Select..</OPTION>
	<%	
		objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & intPrevBudgetID & "",objCon
		
		Do until objRS.EOF
			If objRS("VersionID") = cint(intPrevVersionID) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("VersionID") & """" & strSelected & ">" & objRS("VersionName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		
	%></select></td>
	</tr>
	
	<tr>
	    	<th height="20px" align="left">&nbsp;Previous Orig Version</th>
		<td><select Style="Width:80%" tabindex="14" id="PrevOrigVersionID" name="PrevOrigVersionID"><option value="0">Please Select..</OPTION>
	<%	
		objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & intPrevBudgetID & "",objCon
		
		Do until objRS.EOF
			If objRS("VersionID") = cint(intPrevOrigVersionID) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("VersionID") & """" & strSelected & ">" & objRS("VersionName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		
	%></select></td>
	
	  <th align="left">&nbsp;Ceilings On</th>
		<td>
	       <select Style="Width:40%" tabindex="15" id="CeilingsOn" name="CeilingsOn">
	        <%
		        For x = 1 to 2
			        If arrYesNo(x) = cstr(strCeilingsOn) Then
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
	<tr>
		<td colspan="4" align="left">&nbsp;</td>
	</tr>
	
</table>
<br/>

<table  WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td Width="100px" ><button type="button" tabindex="8" onclick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td Width="100px" ><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
	    <td Width="100px" ><button type="button" tabindex="10" onclick="self.location='Version.asp?VersionID=0'"><img src="../images/page_white_stack.png" alt="" /> Add New </button></td>
        <td Width="120px" ><button type="button" tabindex="13" onclick="javascript:ReIndex();"><img src="../images/wrench.png" alt="" /> Re-Build Indexes</button></td>
		<td Width="120px" ><button type="button" tabindex="14" onclick="javascript:BuildBase();"><img src="../images/wrench.png" alt="" /> Build Base Budget</button></td>
        <td Width="120px" ><button type="button" tabindex="15" onclick="javascript:Consolidate();"><img src="../images/table_multiple.png" alt="" /> Consolidate</button></td>
	</TR>
	<tr><td colspan="6">&nbsp;</td></TR>
    <tr>
    	  <TD Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon%></TD>
        <TD Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage%></TD><TD colspan="4"></TD>
    </tr>
	<tr>
    <TD colspan="6">&nbsp;<span id="Progress" style="display:none"><img src=../Images/progress.gif />  &nbsp;&nbsp;&nbsp; <b>Please wait...</b></span><span id="Span2" style="display:none"><img src=../Images/progress.gif />  &nbsp;&nbsp;&nbsp; <b>Building Base Budget...</b></span></td></tr>
	</tr>
</table>

<hr />
</form>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">

	<tr>		
	    <th height="20px">Version Name</th>
        <th>Version ID</th>
		<th>Column Lock</th>
		<th>Version Type</th>
		<th>HO Currency</th>
		<th>Base Version</th>
		<th>Updated By</th>
		<th>Date Updated</th>
	</tr>
    <tr><td colspan="10">&nbsp;</td></tr>
<%
    objRS.Open "SELECT VersionID, BudgetID, VersionName, ColumnLock, dbo.tblVersion.HOCurrency, dbo.tblVersion.OutYearIndexType, dbo.tblVersion.BaseBudgetVersionID, dbo.tblVersion.UpdatedBy, dbo.tblVersion.DateUpdated, VersionTypeName FROM  dbo.tblVersion INNER JOIN dbo.tblVersionTypes ON dbo.tblVersion.VersionTypeID = dbo.tblVersionTypes.VersionTypeID WHERE BudgetID = " & clng (Session("BudgetID"))

	Do until objRS.eof			
   	   Response.Write "<TR><TD><A Target=""_self"" HREF=""Version.asp?VersionID=" & objRS("VersionID") & """>" & objRS("VersionName") & "</TD><TD style=""text-align:center"">&nbsp;<b>" & objRS("VersionID") & "</B></TD><TD style=""text-align:center"">" & objRS("ColumnLock")+1 & "</TD><TD style=""text-align:center"">" & objRS("VersionTypeName") & "</TD><TD style=""text-align:center"">" & objRS("HOCurrency") & "</TD><TD style=""text-align:center"">" & objRS("BaseBudgetVersionID") & "</TD><TD style=""text-align:center"">" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD></TR>"
       objRS.movenext
	Loop
		
	objRS.Close
	
%>

</table>
</body>

</html>

<% 

Sub LoadDetails()

'Description:	Loads Caller's details into page if applicable.
		
		objRS.Open "SELECT * FROM tblVersion WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & clng(Session("VersionID")) & "",objCon
		
			If Not objRS.EOF Then
				
			  Session("VersionID") = objRS("VersionID")
			  Session("BudgetID") = objRS("BudgetID")
              strVersionName = objRS("VersionName")
              lngColumnLock = objRS("ColumnLock") + 1
              lngVersionTypeID = objRS("VersionTypeID")		
              strCurrency = objRS("HOCurrency")	
              strOYIndexTYpe = 	objRS("OutYearIndexType")
              dblOY1Index = objRS("OY1Index")
              dblOY2Index = objRS("OY2Index")
              dblOY3Index = objRS("OY3Index")
              dblOY4Index = objRS("OY4Index")
              dblOY5Index = objRS("OY5Index")
              intBaseBudgetVersionID = objRS("BaseBudgetVersionID")
              strTrackChanges = objRS("TrackChanges")
              intPrevBudgetID = objRS("PrevBudgetID")
              intPrevVersionID = objRS("PrevVersionID")
			  intPrevOrigVersionID = objRS("PrevVersionID")
              intReportingMonth = objRS("ReportingMonth")
              strCeilingsOn = objRS("CeilingsOn")
			
			Else
			
			  strVersionName = ""
              lngColumnLock = ""
              lngVersionTypeID = 0	
              strCurrency = ""
			
			  intPrevBudgetID = 0
			  intPrevVersionID = 0
			 ' strActive = "Y"	
			End If

		objRS.Close
	

End Sub

Sub SaveDetails()
Dim intColLock

	If IsNull(Request.Form("ColumnLock")) or Request.Form("ColumnLock") = "" Then
		intColLock = 0
	Else
		intColLock = Request.Form("ColumnLock") -1
	End If

       'response.Write(Request.Form("VersionID"))

		 With objCmd
                .CommandType = 4
                .CommandText = "spVersionSave"
                
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("ColumnLock", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionTypeID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("HOCurrency", adVarChar, adParamInput, 3)
                .Parameters.Append objCmd.CreateParameter("OutYearIndexType", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("OY1Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY2Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY3Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY4Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("OY5Index", adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BaseBudgetVersionID", adInteger)
                .Parameters.Append objCmd.CreateParameter("TrackChanges", adVarChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("PrevBudgetID", adInteger)
                .Parameters.Append objCmd.CreateParameter("PrevVersionID", adInteger)
                .Parameters.Append objCmd.CreateParameter("ReportingMonth", adInteger)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
                .Parameters.Append objCmd.CreateParameter("CeilingsOn", adVarChar, adParamInput, 1)
              
				.Parameters("VersionID") = Request.Form("VersionID")		
				.Parameters("BudgetID") = Session("BudgetID")			
                .Parameters("VersionName") = Request.Form("VersionName")
                .Parameters("ColumnLock") = intColLock 'Request.Form("ColumnLock") - 1            
                .Parameters("VersionTypeID") = Request.Form("VersionTypeID")
                .Parameters("HOCurrency") = "TZS" 'Request.Form("Currency")
                .Parameters("OutYearIndexType") = ""
                .Parameters("OY1Index") = 0
                .Parameters("OY2Index") = 0
                .Parameters("OY3Index") = 0
                .Parameters("OY4Index") = 0
                .Parameters("OY5Index") = 0
                .Parameters("BaseBudgetVersionID") = Request.Form("BaseBudgetVersionID")
                .Parameters("TrackChanges") = Request.Form("TrackChanges")
                .Parameters("PrevBudgetID") = Request.Form("PrevBudgetID")
                .Parameters("PrevVersionID") = Request.Form("PrevVersionID")
                .Parameters("ReportingMonth") = 0
                .Parameters("UpdatedBy") = Session("UserID")
                .Parameters("CeilingsOn") = Request.Form("CeilingsOn")
                          
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
			'Return the result of the Save Function.
     		'Session("VersionID") = objCmd.Parameters.Item("VersionIDOutput")
     		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
            strMessage = "<B>RECORD SAVED.</B>"
	
End Sub	

Public Sub Build_Index()

    objCon.Execute "spBuildIndexationVersion " & Session("BudgetID") & "," & Session("VersionID") & ",'GEXP'"
    strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
     strMessage = "<B>DATABASE HAS BEEN REINDEXED.</B>"

End Sub

Public Sub Build_Base_Budget()

    objCon.Execute "spBuildBaseBudget " & Session("BudgetID") & "," & Session("VersionID") & "," & Session("UserID") & ""
    
        strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
     strMessage = "<B>BASE BUDGET HAS BEEN BUILT.</B>"

End Sub

Public Sub Consolidate()

    objCon.Execute "spConsolidateGeneralExpenseTables " & Session("BudgetID") & "," & Session("VersionID") & ",1"

     strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
     strMessage = "<B>INPUT SHEET TABLES HAVE BEEN CONSOLIDATED.</B>"

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

Set objRS = Nothing
Set objCon = Nothing


%>
