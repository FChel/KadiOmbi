<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->

<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
Session("CurrentPage") = "Admin/SegmentValues.asp"
If IsEmpty(Session("SegmentNo")) Then Session("SegmentNo") = 0
If IsEmpty(Session("SegmentID")) Then Session("SegmentID") = 0
 
'Description:	Segments Administration Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			March 2008

'Declare default variables

Dim objCon
Dim objCmd
Dim objRS
Dim strScreen
Dim strSelected
Dim x 
Dim strMessage
Dim strMessageIcon
Dim strSegmentNo
Dim intMax
Dim intMin

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngSegmentID
Dim strSegmentName
Dim strSegmentDesc
Dim strSegmentCode
Dim strSegmentNameL2
Dim strSegmentDescL2
Dim strSegmentNotes
Dim strParent
Dim strActive

'Declare and set default arrays

Dim arrActive(2)
	
	arrActive(1) = "Y"
	arrActive(2) = "N"
		
	'3. Capture Querystring variables	
	If Not IsEmpty(Request.QueryString("SegmentNo")) Then		
		Session("SegmentNo") = Request.QueryString("SegmentNo")
	End If	
    
    If Not IsEmpty(Request.QueryString("SegmentID")) Then		
		Session("SegmentID") = Request.QueryString("SegmentID")
	End If	

     If Not IsEmpty(Request.QueryString("SegmentCode")) Then		
		Session("SegmentCode") = Request.QueryString("SegmentCode")
	End If	


	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

    If Request.QueryString("Action") = "Delete" Then
        Call DeleteRecord(Request.QueryString("SegmentNo"),Request.QueryString("SegmentCode"))
    End If

	'Load page details
	LoadDetails()

		
%>

<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
	<script src="../formChek.js">
	</script>
	<script src="../ButtonRollOver.js">
	</script>

<script LANGUAGE="javascript">
<!--
function SaveData()
{	    
    var varSubmit = true						
    var varAlert="";	

	if(isWhitespace(frm.SegmentID.value) || frm.SegmentID.value=="0")
	{
	frm.SegmentID.value = 0
       //varAlert += "Segment ID Cannot Be Blank. \n \n";
       //document.getElementById('SegmentID').style.backgroundColor="ff8080";
       //varSubmit = false;
    }   
    else document.getElementById('SegmentID').style.backgroundColor="ffffff";
    
    
    if(isWhitespace(frm.SegmentName.value))
	{
       varAlert += "Segment Name Cannot Be Blank. \n \n";
       document.getElementById('SegmentName').style.backgroundColor="ff8080";
       varSubmit = false;
    }
    else document.getElementById('SegmentName').style.backgroundColor = "ffffff";

    if (isWhitespace(frm.SegmentNameL2.value)) {
        varAlert += "Segment Name L2 Cannot Be Blank. \n \n";
        document.getElementById('SegmentNameL2').style.backgroundColor = "ff8080";
        varSubmit = false;
    }
    else document.getElementById('SegmentNameL2').style.backgroundColor = "ffffff";

		
	if(varSubmit == true)
	{
	    frm.submit();		
	}
	else
	{
	    alert(varAlert);
	}
}


function SegmentIDSearch()
{	
	self.location="SegmentValues.asp?SegmentCode=" + frm.SegmentCode.value
}

function BAIDSearch() {
    self.location = "SegmentValues.asp?SegmentNo=" + frm.SegmentNo.value
}

function DeleteData() {
    if (frm.SegmentCode.value == '') {
        alert('Please select a record to DELETE!');
    } else {
        if (window.confirm('Would you like to DELETE the selected record?') == true) {

            self.location = "SegmentValues.asp?Action=Delete&SegmentNo=" + frm.SegmentNo.value + "&SegmentCode=" + frm.SegmentCode.value;
        }

    }
}

//-->
</script>
</head>
<body>
<h3>Segment ValuesAdministration Screen</h3>
<form action="SegmentValues.asp?Action=Save" method="POST" id="frm" name="frm">

<TABLE WIDTH=100% ALIGN=Center BORDER=1 CELLSPACING=1 CELLPADDING=1>

	<tr>
		<th style="text-align:left; height:20px; width:20%;">&nbsp;Segment</th>
		<td style="text-align:left; height:20px; width:30%;">
		    <select Style="Width:100%" tabindex="20" id="SegmentNo" name="SegmentNo" onchange="BAIDSearch()"><OPTION Value=0>Please Select..</OPTION>
	        <%	
		    objRS.Open "SELECT * FROM tblSegment",objCon
    		
		    Do until objRS.EOF
			    If clng(objRS("SegmentID")) = cint(Session("SegmentNo")) Then
				    strSelected = " SELECTED "
			    Else
				    strSelected = ""
			    End if
				    Response.Write "<option Value=""" & objRS("SegmentID") & """" & strSelected & ">" & objRS("SegmentCode") & " - " & objRS("SegmentName") & "</OPTION>"
			    objRS.Movenext
		    Loop
    		
		    objRS.Close
    		
	        %></select>
	    </td>
        <td colspan="2"></td>
	</tr>
    <tr><td colspan="4">&nbsp;</td></tr>	
	<tr>
		
		<input style="text-align:left" style="width:50%" type="hidden" id="SegmentID" name="SegmentID" maxlength="9" TABINDEX="1"  value="<%=lngSegmentID%>" />		
        <th style="text-align:left; height:20px; width:20%;">&nbsp;Segment Code</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left" style="width:50%" id="SegmentCode" name="SegmentCode" maxlength="9" TABINDEX="1" value="<%=Session("SegmentCode")%>" onblur="SegmentIDSearch()"></td>
		<td colspan="2"></td><!--<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left" style="width:50%" id="SegmentCode" name="SegmentCode" maxlength="9" TABINDEX="1" onblur="SegmentCodeSearch()" value="<%=strSegmentCode%>"></td>-->
	</tr>
	<tr>
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Segment Name Eng</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; width:98%" id="SegmentName" name="SegmentName" maxlength="50" TABINDEX="2" value="<%=strSegmentName%>"></td>
	
	    <th style="text-align:left; height:20px; width:20%;">&nbsp;Segment Name Swa</th>
		<td style="text-align:left; height:20px; width:30%;">&nbsp;<input style="text-align:left; width:98%" id="SegmentNameL2" name="SegmentNameL2" maxlength="200" TABINDEX="3" value="<%=strSegmentNameL2%>"></td>
       
   </tr>
	<tr><th style="text-align:left; height:20px; width:20%;">&nbsp;Active</th><td style="text-align:left; height:20px; width:30%;">
	 <select Style="Width:40%" tabindex="5" id="Active" name="Active">
	        <%
		        For x = 1 to 2
			        If arrActive(x) = cstr(strActive) Then
				        strSelected = " SELECTED "
			        Else
				        strSelected = ""
			        End If
				        Response.Write "<option Value=""" & arrActive(x) & """" & strSelected & ">" & arrActive(x) & "</OPTION>"
		        Next
	        %>
	       </select>
	    </td><td colspan="2"></td>
	</tr>		
    <tr>
	    <td style="text-align:left; height:20px" colspan="4">&nbsp;</td>
	</tr>
</table>
<br>
<table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="8" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="9" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="10" onClick="self.location='SegmentValues.asp?SegmentID=0';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
		<td class='locked' Width="100px"><button type="button" tabindex="19" onclick="DeleteData()";><img src="../images/cross.png" alt="" /> Delete </button></td>
        <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
<hr>
</form>
<H3>List of Values for Segment : <FONT Color="red"><%Response.Write Session("Segment") %></FONT></H3>
<table WIDTH="100%" BORDER="1" CELLSPACING="1" CELLPADDING="1">
	<tr>
		<th style="text-align:center; height:20px" width="20%">Segment Code</th>	
	    <th style="text-align:center; height:20px" width="30%">Segment Name Eng</th>
	    <th style="text-align:center; height:20px" width="30%">Segment Name Swa</th>		
	 	<th style="text-align:center; height:20px" width="10%">Active</th>
	</tr>
	
<%
        objRS.Open "SELECT * FROM tblSegmentValues WHERE BudgetID = " & Session("BudgetID") & " AND SegmentNo = '" & Session("SegmentNo") & "' AND BusinessAreaID = " & Session("BusinessAreaID") & " Order By SegmentCode ASC",objCon
       ' Response.Write "SELECT * FROM tblSegmentValues WHERE BudgetID = " & Session("BudgetID") & " AND SegmentNo = '" & Session("SegmentNo") & "' AND BusinessAreaID = " & Session("BusinessAreaID") & " Order By SegmentCode ASC"
		Do until objRS.eof
			Response.Write "<TR><TD><A Target=""_self"" HREF=""SegmentValues.asp?SegmentID=" & objRS("SegmentID") & "&SegmentCode=" & objRS("SegmentCode") & """><B>&nbsp;" & objRS("SegmentCode") & "</B></A></TD>" & _
				"<TD>&nbsp;" & objRS("SegmentName") & "</B></TD><TD>&nbsp;" & objRS("SegmentNameL2") & "</B></TD><TD style=""text-align:center"">" & objRS("Active") & "</TD></TR>"
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

        'Response.Write "SELECT * FROM tblSegmentValues WHERE BudgetID = " & Session("BudgetID") & " AND SegmentID = " & Session("SegmentID") & " AND SegmentNo = " & Session("SegmentNo") & ""
	If IsNull(Session("SegmentNo")) Or Session("SegmentNo") = "" Then Session("SegmentNo") = 0
	If IsNull(Session("SegmentIDo")) Or Session("SegmentID") = "" Then Session("SegmentID") = 0
		objRS.Open "SELECT * FROM tblSegmentValues WHERE BudgetID = " & Session("BudgetID") & " AND SegmentCode = '" & Session("SegmentCode") & "' AND SegmentNo = " & Session("SegmentNo") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon							
	
        If Not objRS.EOF Then
		    lngSegmentID = objRS("SegmentID")
            strSegmentName = objRS("SegmentName")
         
            strSegmentNameL2 = objRS("SegmentNameL2")
       
            strSegmentCode = objRS("SegmentCode")
            strActive = objRS("Active")										
		End if

		objRS.Close	
End Sub

Sub SaveDetails()	

Dim Vote

    objRS.Open "SELECT BusinessAreaCode FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaID = " & Session("BusinessAreaID") & "",objCon

        If Not objRS.EOF Then
            Vote = objRS("BusinessAreaCode")
        End If

    objRS.Close	

    objRS.Open "SELECT * FROM tblSegment WHERE SegmentID = " & Request.Form("SegmentNo") & "",objCon

        If Not objRS.EOF Then
            intMax = objRS("MaxLength")
            intMin = objRS("MinLength")
        End If

    objRS.Close	

    If Len(Request.Form("SegmentCode")) > intMax or Len(Request.Form("SegmentCode")) < intMin Then

        strMessage = "<FONT Color=""Red""><B>SEGMENT LENGTH IS INCORRECT.</B></FONT>"
        strMessageIcon = "<img src=""../images/warning.gif"" />"

	Else
    	
		 With objCmd
                .CommandType = 4
                .CommandText = "spSegmentValueSave"
                
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)    
                .Parameters.Append objCmd.CreateParameter("SegmentNo", adInteger, adParamInput) 
                .Parameters.Append objCmd.CreateParameter("SegmentID", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("SegmentCode", adVarChar, adParamInput, 50)  
                .Parameters.Append objCmd.CreateParameter("BusinessAreaID", adInteger, adParamInput)     
                .Parameters.Append objCmd.CreateParameter("SegmentName", adVarChar, adParamInput, 200)
                .Parameters.Append objCmd.CreateParameter("SegmentNameL2", adVarChar, adParamInput, 200)
                .Parameters.Append objCmd.CreateParameter("Active", adVarChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)  
              
				.Parameters("BudgetID") = Session("BudgetID")
                .Parameters("VersionID") = Session("VersionID")
                .Parameters("SegmentNo") = Request.Form("SegmentNo")	
                .Parameters("SegmentID") = Request.Form("SegmentID")
                .Parameters("SegmentCode") = Request.Form("SegmentCode")
                .Parameters("BusinessAreaID") = Session("BusinessAreaID")	
			    .Parameters("SegmentName") = Request.Form("SegmentName")
                .Parameters("SegmentNameL2") = Request.Form("SegmentNameL2")
                .Parameters("Active") = Request.Form("Active")               
                .Parameters("UpdatedBy") = Session("UserID")
                                        
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               			     				
     		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"	
            strMessage = "<B>RECORD SAVED.</B>"								
     		Session("SegmentID") =  Request.Form("SegmentID")

		End If
    			
End Sub	

Sub DeleteRecord(SegmentNo,SegmentCode)

Dim CCID

SELECT CASE SegmentNo

       Case 1

            objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND Segment1 = '" & SegmentCode & "'",objCon

       Case 2

            objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND Segment2 = '" & SegmentCode & "'",objCon

       Case 3

            objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND Segment3 = '" & SegmentCode & "'",objCon
       
       Case 4

            objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND Segment4 = '" & SegmentCode & "'",objCon

       Case 5

            objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND Segment5 = '" & SegmentCode & "'",objCon

       Case 6

            objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND Segment6 = '" & SegmentCode & "'",objCon

       Case 7

            objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND Segment7 = '" & SegmentCode & "'",objCon

       Case 8

            objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND Segment8 = '" & SegmentCode & "'",objCon

       Case 9

            objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND Segment9 = '" & SegmentCode & "'",objCon

       Case 10

            objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND Segment10 = '" & SegmentCode & "'",objCon

END SELECT 

    If objRS.EOF Then
        objCon.Execute "DELETE FROM tblSegmentValues WHERE BudgetID = " & Session("BudgetID") & " AND SegmentNo = " & SegmentNo & " AND SegmentCode = '" & SegmentCode & "' AND BusinessAreaID = " & Session("BusinessAreaID") & ""   
        
        Select Case SegmentNo

            Case 1

                objCon.Execute "DELETE tblGLCodes WHERE BudgetID = " & Session("BudgetID") & " AND GLCode = " & SegmentCode & ""                

            Case 4
            
                objCon.Execute "DELETE tblCostCentres WHERE BudgetID = " & Session("BudgetID") & " AND CostCentreID = " & SegmentCode & ""    
                CCID = Session("BusinessAreaID") & SegmentCode            
                objCon.Execute "DELETE tblCCBARelationship WHERE BudgetID = " & Session("BudgetID") & " AND CostCentreID = " & CCID & ""
      
            Case 6

                objCon.Execute "DELETE tblBid WHERE BudgetID = " & Session("BudgetID") & " AND SegmentCode = " & SegmentCode & " AND Left(CostCentreID,4) = " & Session("BusinessAreaID") & ""   

        End Select

        strMessage = "<B>RECORD DELETED.</B>"
        strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
              
    Else
        strMessage = "<FONT Color=""Red""><B>SEGMENT CANNOT BE DELETED BECAUSE ENTRIES EXIST FOR THIS SEGMENT.</B></FONT>"
        strMessageIcon = "<img src=""../images/warning.gif"" />"

    End If

objRS.Close

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
