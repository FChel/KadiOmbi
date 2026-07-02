<%@ Language=VBScript %>

<!-- #Include file=../ADOVBS.inc -->

<%

	Response.Expires = -1500

    If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
	
	Session("CurrentPage") = "Admin/BudgetExport.asp"	
	
Dim objCon
Dim objCmd
Dim objRS
Dim objRS1
Dim objRS2
Dim arrHeadings(5)
Dim intFinYearPart1
Dim intFinYearPart2
Dim strBackColour
Dim dblDisplay
Dim dblActual
Dim dblBudget
Dim dblBaseBudget
Dim dblVariance
Dim dblVariancePercentage
Dim strForeColour
Dim intMode
Dim dblTotal
Dim dblTotal1
Dim dblOriginal
Dim dblVarianceTotal
Dim dblVariancePercentageTotal
Dim dblOriginalTotal
Dim dblVarianceTotal1
Dim dblVarianceRunningTotal
Dim dblVariancePercentageTotal1
Dim x
Dim strGLType
Dim strTotalHeading
Dim strBudgetHeading
Dim dblOY1
Dim dblOY2
Dim dblOY3
Dim strBusinessAreaName
Public strMessage
Dim strMessageIcon
Dim strMessageIcon1

    If Session("ModeID") = 3 or Session("ModeID") = 1 Then
    	strTotalHeading = "YTD"
    	strBudgetHeading = "YTD Budget"
    Else
    	strTotalHeading = "Total"
	    strBudgetHeading = "Orig Budget"
    End If

    'Set Headings
    For x = 0 to 4
	
	    intFinYearPart1 = cint(Session("FinancialYear")) + (x - 2)
	    intFinYearPart1 = Right(intFinYearPart1,2)
	    intFinYearPart2 = cint(Session("FinancialYear")) + x - 1
	    intFinYearPart2 = Right(intFinYearPart2,2)

	    arrHeadings(x) = cstr(intFinYearPart1) & "/" & cstr(intFinYearPart2)

    Next

'Set database objects
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")
Set objRS2 = Server.CreateObject("ADODB.Recordset")


objCon.Open Session("DBConnection")
    
If Validate_Access(Session("UserTypeID"),Session("CurrentPage")) = "N" Then
     Response.Redirect "../AccessDenied.asp"
End If

If Not IsEmpty(Request.QueryString("BaseVersionID")) Then
    Session("BaseVersionID") = Request.QueryString("BaseVersionID")
End If    
    
If Request.QueryString("Action") = "Save" Then
    
    If Validate_Access(Session("UserTypeID"),"Budget Export") = "Y" Then
        	
        SaveRecord()
    Else
        Response.Write "&nbsp;&nbsp;<img src=""../images/warning.gif"" /><B><FONT Style=""font-family:Arial;Color:Red"">&nbsp;&nbsp;WARNING - YOU DO NOT HAVE AUTHORITY TO EXPORT BUDGETS.</FONT></B>"  
    End If
	 	 
End if   

If Request.QueryString("Action") = "Recalculate" Then
    Recalculate()
End If

'Get Indexation Status

objRS.Open "SELECT * FROM tblBusinessArea WHERE BudgetID =  " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon

    If Not objRS.EOF Then        
    
            strBusinessAreaName = objRS("BusinessAreaName") 
       
    End If
    
objRS.Close

%>

<html>
<head>
<style> 
    <!--
div#tbl-container {
width: 98%;
height: 90%;
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
    z-index         : 1
    border-right: 1px solid silver;
	border-left: 1px solid silver;
	position: relative;
	align: center;

  }

    -->
   </style>

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
<script src="../ButtonRollOver.js">
</script>
<script language="javascript">

    function SaveData()
    {
        document.getElementById('Progress').style.display = "inline";     
        frm.submit();
    }

    function FlagSaveStatus(x) {

        var Row
        var Col
        var Total
        var InputBox = x
        parts = InputBox.split(".");
        Row = parts[0];
        Col = parts[1];

        var varElem1 = (parseInt(Row)) + parseInt(Col);
        var varElem2 = 10000 + parseInt(Col);
        var varElem3 = parseInt(Row) + 13;

        if (isSignedFloat(document.getElementById(varElem1).value) == false) {
            if (Col == 0) {
            }
            else {
                alert("Invalid Entry! Please do not include a comma in the number. ie 45,000 must be entered as 45000.");
                document.getElementById(varElem1).value = 0;
            }
        } else {
            document.getElementById(varElem2).value = 'Y';
            //top.Header1.Header2.document.form1.text1.value='1';
        }

    }

    function FlagSaveStatusComments(x) {

        var Row
        var Col
        var Total
        var InputBox = x
        parts = InputBox.split(".");
        Row = parts[0];
        Col = parts[1];

        
        var varElem1 = (parseInt(Row)) + parseInt(Col);
        var varElem2 = 10000 + parseInt(Col);
        var varElem3 = parseInt(Row) + 13;
        var varElem4 = "bolApp1000";

        document.getElementById(varElem2).value = 'Y';
        if (varElem1 == 10001){
           if (document.getElementById(varElem4).checked==true) {
               alert('WARNING : By selecting Vote 000 all Votes will be exported to Epicor.');}
           }
        }
   
    
    function checkValue(obj)
    {
        alert(obj);                     
        
   }
</script>
</head>
<body>
<H3>Budget Export Status</H3>

<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	
		<th Height="20px" Width="10%" class='locked'>Vote</th>
		<th Width="15%">Date Exported</th>
		<th Width="5%">Exported</th>
        <th Width="5%">Export Status</th>
        <th Width="5%">Vote Status</th>	
        <th Width="10%">Record Count</th>	
        <th Width="10%">Recurrent</th>	
        <th Width="10%">Development</th>	
        <th Width="10%">Revenue</th>
        <th Width="15%">Imported On</th>	
        <th Width="5%">Has Error</th>				

<tbody>
<form name="frm" action="BudgetExport.asp?Action=Save" method="post">
<%
	DisplayTableDetails()
%>	

</table>

<hr Width="100%">

<TABLE Width="30%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<TR>
    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="15"  onclick="self.location='AdminMenu.asp'";><img src="../images/door.png" alt="" /> Close </button></td>    
    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="16" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
    <td Width="100px"><span id="Progress" style="display:none"><img src=../Images/progress.gif />  &nbsp;&nbsp;&nbsp; <b><FONT Face="Arial"></FONT></b></span></td>
    <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
    <TD class='locked' Width="800px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
    
</TR>
</TABLE>

</body>
</html>
<%
Public Sub DisplayTableDetails()

Dim x
Dim y

Dim intLength
Dim txtIdx
Dim txtCmt
Dim bolApp
Dim strChecked
Dim dblBMTotal
Dim dblBaseBudget
Dim dblVariance
Dim dblBudgetTotal
Dim dblBaseBudgetTotal
Dim dblExistingCeiling
Dim dblExistingCeilingTotal
Dim dblCeilingL2
Dim dblOY1CeilingTotal
Dim dblOY2CeilingTotal
Dim strColor
Dim dblOY1
Dim dblOY2
Dim dblBMCeiling
Dim strDisabled
Dim arrStatus(5)
Dim dblRecordCount
Dim dblRev
Dim dblRec
Dim dblDev

arrStatus(0) = "<IMG SRC='../images/delete.png'"
arrStatus(1) = "<IMG SRC='../images/open.png'"
arrStatus(2) = "<IMG SRC='../images/ready.gif'"
arrStatus(3) = "<IMG SRC='../images/cross.png'" 
arrStatus(4) = "<IMG SRC='../images/tick.png'"	
arrStatus(5) = "<IMG SRC='../images/Closed.png'"

lngRow = 1
lngRow100 = 10000
x = 1

'Open the master recordset

    objRS.Open "SELECT * FROM qryBudgetExportLog WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " Order by Vote",objCon

    y = 1  
	
    Do until objRS.EOF
    
        'Response.Write "Value : " & objRS("level1Name") & "<br>"
        txtIdx = "txtIdx" & objRS("BusinessAreaID")
        txtCmt = "txtCmt" & objRS("BusinessAreaID")
		bolApp = "bolApp" & objRS("BusinessAreaID")

        dblRecordCount = objRS("RecordsExported")
        dblRev = objRS("Revenue")
        dblRec = objRS("Recurrent")
        dblDev = objRS("Development")

        If IsNull(dblRecordCount) Then dblRecordCount = 0
        If IsNull(dblRev) Then dblRev = 0
        If IsNull(dblRec) Then dblRec = 0
        If IsNull(dblDev) Then dblDev = 0

        If objRS("StatusID") <> 4 Then
            strDisabled = "DISABLED"
        Else
            strDisabled = ""
        End If
   		
		'Insert Row Heading
	
        Response.Write "<TR><TD class='locked' class=""custom_table_th_" & y & """>&nbsp;" & objRS("Vote") & "</FONT></TD>"	
          
		If objRS("Status") = "on" Then
		    strChecked = "checked"
            strMessageIcon1 = "<img src=""../images/tick.png"" />"
            
		Else
		    strChecked = ""
            strMessageIcon1 = "<img src=""../images/cross.png"" />"
		End If
		
		strBackColour = "FFFFCC"
       	     
	     'Set the Total Row Total Variables 
	          
	        Response.Write "<TD Style=""Text-Align:Center""><input Style=""Width:100%; Text-Align:Right"" type='hidden' onkeyup='javascript:checkValue(this)' name='"& txtIdx &"' value='"& objRS("BusinessAreaID") &"' /><input Style=""Text-Align:Left; Width:100%"" type='text' name='"& txtCmt & "' value='"& objRS("DateExported") & "' onChange=""FlagSaveStatusComments('" & lngRow100 & "." & x & "');""/></TD><TD><input " & strDisabled & " Style=""Width:100%; Text-Align:Right"" type='checkbox' onkeyup='javascript:checkValue(this)' id='"& bolApp &"' name='"& bolApp &"' " & strChecked & " onChange=""FlagSaveStatusComments('" & lngRow100 & "." & x & "');"" /></td><td style=""text-align:center"">" & strMessageIcon1 & "</TD><TD Style=""Text-Align:Center"">&nbsp;&nbsp;&nbsp;" & arrStatus(objRS("StatusID")) & "</TD><TD Style=""Text-Align:Right"">" & formatnumber(dblRecordCount,0,0) & "</TD><TD Style=""Text-Align:Right"">" & formatnumber(dblRec,0,0) & "</TD><TD Style=""Text-Align:Right"">" & formatnumber(dblDev,0,0) & "</TD><TD Style=""Text-Align:Right"">" & formatnumber(dblRev,0,0) & "</TD><TD Style=""Text-Align:Center"">" & objRS("ImportedOn") & "</TD><TD Style=""Text-Align:Center"">" & objRS("HasError") & "</TD>"
	        txtBox = 10000 + lngRow	
	        Response.Write "<INPUT style=""verticalalign:right style=width:0px"" READONLY type=""hidden"" id=" & txtBox & " name=" & txtBox & " value=" & txtBox & "></TD>"
     	          
		Response.Write "</TR>"	
			x = x + 1	
			lngRow = lngRow + 1
		objRS.MoveNext
		
		'**** Insert the Total Rows ****

Loop


End Sub

'Function for saving Business Area Indexation record in the database
Sub SaveRecord() 
   
    Dim level1ID
    Dim indexation
    Dim comments
    Dim dblBudget
    Dim dblCeiling
    Dim approved
    Dim changed
    dim row
    Dim strBudClassCeiling
    
    row = 10001
    
    strMessage = "<b>BUDGET HAS BEEN EXPORTED.</b>" 
    strMessageIcon = "<img src=""../images/saveticksmall.jpg"" />"
      
    objRS.Open "SELECT * FROM qryBudgetExportLog WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " Order By Vote",objCon

        Do until objRS.EOF 
              
               level1ID = objRS("BusinessAreaID")
               indexation = Request.Form("txtIdx" & objRS("BusinessAreaID"))
               approved = Request.Form("bolApp" & objRS("BusinessAreaID"))
               changed = Request.Form("" & row & "")
               'Response.Write row
               If changed = "Y" Then                    
                    
                    if(indexation <>"") then
                
                        comments = Request.Form("txtCmt" & objRS("BusinessAreaID")) 
                    
                        If IsNull(indexation) Then indexation = 0
                        If IsNull(approved) Then approved = "off"
                        If IsEmpty(approved) Then approved = "off"
                        If approved = "" Then approved = "off"
                     
                    End If
                    
                        
	                    With objCmd
	                
                            .CommandType = 4
                            .CommandText = "spBudgetExportLogSave"
                   
                            .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                            .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                            .Parameters.Append objCmd.CreateParameter("BusinessAreaID", adInteger, adParamInput)
                            .Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput,3)   
                            .Parameters.Append objCmd.CreateParameter("ExportedBy", adChar, adParamInput,6)                                      
                           				 				 
    		                .Parameters("BudgetID") = Session("BudgetID")
				            .Parameters("VersionID") = Session("VersionID")
				            .Parameters("BusinessAreaID") = level1ID
				            .Parameters("Status") = approved
				            .Parameters("ExportedBy") = Session("UserID")
				            .ActiveConnection = objCon
            			
			                 objCmd.Execute
			                
			               'Response.Write " : "  & row
			               'Response.Write approved	& " - "
			               'Response.Write level1ID
                       
                       End With 
                       
                       Set ObjCmd = Server.CreateObject("ADODB.Command")
                           
                     
                 End If	     
            
     	objRS.MoveNext
     	row = row + 1
        Loop    
        objRS.Close   
     	
     	
End Sub	



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



%>