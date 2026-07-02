<%@  language="VBScript" %>
<% Option Explicit %>
<!-- #Include file=../ADOVBS.inc -->
<%

Response.Expires = -1500

If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")
 
'Description:	GL Codes Screen
'Author:		Andrew Bull - Isidore Limited (www.isidore.com - 07969 589413)
'Date:			November 2007

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

If IsNull(Session("GLCodeID")) Then Session("GLCodeID") = -1 End If

'Open database connection

objCon.Open Session("DBConnection")

'1. Declare screen specific variables naming convention = variable type prefix + database field name

Dim lngGLCode
Dim lngGLCodeID
Dim strGLCodeName
Dim strGLCodeNameL2
Dim strGLCodeDesc
Dim strGLCodeType
Dim lngActualMapping
Dim strPrepayment
Dim strBalanceSheet
Dim lngContraGLCode
Dim lngBSLevel1ID
Dim lngCFLevel1ID
Dim strCEType
Dim strCEGroup
Dim strCESubGroup
Dim strTriDataCEType
Dim strTriDataCEGroup
Dim strTriDataCESubGroup
Dim strPASP
Dim strReportingCategory1
Dim strReportingCategory2
Dim strReportingCategory3
Dim strReportingCategory4
Dim strReportingCategory5
Dim strActive
Dim strReportItalic
Dim strSegmentValue
Dim lngPostingGLCode

'Declare and set default arrays

Dim arrYesNo(2)
	
	arrYesNo(1) = "Y"
	arrYesNo(2) = "N"
	
Dim arrGlCodeType(5)
Dim arrGlCodeTypeName(5)
	
	arrGLCodeType(1) = "R"
	arrGLCodeType(2) = "E"
	arrGLCodeType(3) = "A"
	arrGLCodeType(4) = "L"
	arrGLCodeType(5) = "C"	

	arrGLCodeTypeName(1) = "Revenue"
	arrGLCodeTypeName(2) = "Expense"
	arrGLCodeTypeName(3) = "Asset"
	arrGLCodeTypeName(4) = "Liability"
	arrGLCodeTypeName(5) = "Capital"
	
	'3. Capture Querystring variables
	
	If Not IsEmpty(Request.QueryString("GLCodeID")) AND IsNumeric(Request.QueryString("GLCodeID")) Then
		
		Session("GLCodeID") = Request.QueryString("GLCodeID")
		lngGLCode = clng(Request.QueryString("GLCodeID"))

	End If
	
	'Execute save 	
	If Request.QueryString("Action") = "Save" Then
		SaveDetails()
	End If

    If Request.QueryString("Action") = "Delete" Then
        Call DeleteRecord(Request.QueryString("GLCode"))
    End If

	'Load page details
	LoadDetails()
		
%>
<html>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <link rel="stylesheet" type="text/css" href="../BERTStyle.css">

    <script src="../formChek.js">
    </script>

    <script src="../ButtonRollOver.js">
    </script>

    <script language="javascript">
<!--
        function SaveData() {
            var varSubmit = true
            var varAlert = "";

            if (isWhitespace(frm.GLCodeName.value) || frm.GLCodeName.value == "0") {
                varAlert += "GLCode name cannot Be Blank. \n \n";
                document.getElementById('GLCodeName').style.backgroundColor = "ff8080";
                varSubmit = false;
            }
            else document.getElementById('GLCodeName').style.backgroundColor = "ffffff";


            if ((isNonnegativeInteger(frm.GLCode.value) == false) || (frm.GLCode.value == 0)) {
                varAlert += "Please enter GL Code. GL Code must be a numeric value.  \n \n";
                document.getElementById('GLCode').style.backgroundColor = "ff8080";
                varSubmit = false;
            }
            else document.getElementById('GLCode').style.backgroundColor = "ffffff";
         
            if (isWhitespace(frm.GLCodeType.value)) {
                varAlert += "GL Code Type cannot Be Blank. \n \n";
                document.getElementById('GLCodeType').style.backgroundColor = "ff8080";
                varSubmit = false;
            }
            else document.getElementById('GLCodeType').style.backgroundColor = "ffffff";

	        if ((isNonnegativeInteger(frm.SegmentValue.value) == false) || (frm.SegmentValue.value == 0)) {
                varAlert += "Please enter Segment Value. Segment Value must be a numeric value.  \n \n";
                document.getElementById('SegmentValue').style.backgroundColor = "ff8080";
                varSubmit = false;
            }
            else document.getElementById('SegmentValue').style.backgroundColor = "ffffff";

             if (isWhitespace(frm.CESubGroup.value) || frm.CESubGroup.value == "0") {
                varAlert += "CE Sub Group name cannot Be Blank. \n \n";
                document.getElementById('CESubGroup').style.backgroundColor = "ff8080";
                varSubmit = false;
            }
            else document.getElementById('CESubGroup').style.backgroundColor = "ffffff";


            if (varSubmit == true) {
                frm.submit();
            }
            else {
                alert(varAlert);
            }
        }

        function GLCodeIDSearch() {
           
            self.location = "GLCodes.asp?GLCodeID=" + frm.GLCode.value
        }

        function DeleteData() {
            if (frm.GLCode.value == 0) {
                alert('Please select a record to DELETE!');
            } else {
                if (window.confirm('Would you like to DELETE the selected record?') == true) {

                    self.location = "GLCodes.asp?Action=Delete&GLCode=" + frm.GLCode.value;
                }

            }
        }

//-->
    </script>

</head>
<body>
<h3>GFS Code Administration Screen</h3>
    <form action="GLCodes.asp?Action=Save" method="POST" id="frm" name="frm">
    <table width="100%" align="Center" border="1" cellspacing="1" cellpadding="1">
           <tr>
            <th style="text-align:left; height:20px; width:20%;">
                &nbsp;GFS Code ID
            </th>
            <td style="text-align:left; height:20px; width:30%;">
                &nbsp;<input style="text-align: left" style="width: 90%" id="GLCode" name="GLCode"
                    maxlength="7" tabindex="1" value="<%=lngGLCode%>" onblur="GLCodeIDSearch()">
            </td>
            <th style="text-align:left; height:20px; width:20%;">
                &nbsp;GL Code Name Eng
            </th>
            <td style="text-align:left; height:20px; width:30%;">
                &nbsp;<input style="text-align: left;width: 90%" id="GLCodeName" name="GLCodeName"
                    maxlength="50" tabindex="2" value="<%=strGLCodeName%>">
            </td>
        </tr>
        <tr>
              <th style="text-align:left;height:20px">
                &nbsp;GL Code Name Swa
            </th>
             <td>
                &nbsp;<input style="text-align: left;width: 90%" id="GLCodeNameL2" name="GLCodeNameL2"
                    maxlength="100" tabindex="2" value="<%=strGLCodeNameL2%>">
            </td>
             <th style="text-align:left;height:20px">
                &nbsp;Segment Value
            </th>
             <td>
                &nbsp;<input style="text-align: left;width: 90%" id="SegmentValue" name="SegmentValue"
                    maxlength="100" tabindex="3" value="<%=strSegmentValue%>">
            </td>
        </tr>
        <tr>
            <th style="text-align:left;height:20px">
                &nbsp;GL Code Type
            </th>
            <td>
                <select style="width: 40%" tabindex="4" id="GLCodeType" name="GLCodeType">
                    <%
		For x = 1 to 5
			If arrGLCodeType(x) = cstr(strGLCodeType) Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End If
				Response.Write "<option Value=""" & arrGLCodeType(x) & """" & strSelected & ">" & arrGLCodeTypeName(x) & "</OPTION>"
		Next
                    %>
                </select>
            </td> 
            
                   <th style="text-align:left;height:20px">
                    &nbsp;Active
                </th>
                <td>
                    <select style="width: 40%" tabindex="10" id="Active" name="Active">
                        <%
		For x = 1 to 2
			If arrYesNo(x) = cstr(strActive) Then
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
            <th style="text-align:left;height:20px">
                &nbsp;Posting GL Code
            </th>
          
		<td><select style="Width:95%;" tabindex="11" id="PostingGLCode" name="PostingGLCode"><option Value="0">Please Select....</option>
		<%
		objRS.Open "SELECT * FROM tblGLCodes WITH(NOLOCK) WHERE BudgetID = " & Session("BudgetID") & " AND Active = 'Y'",objCon,0,1
		
		Do until objRS.EOF
			If objRS("GLCode") = lngPostingGLCode Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("GLCode") & """" & strSelected & ">" & objRS("GLCode") & " : " & objRS("GLCodeName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>	
         <th style="text-align:left;height:20px">
                &nbsp;CE Sub Group
            </th>
          
		<td><select style="Width:95%;" tabindex="12" id="CESubGroup" name="CESubGroup"><option Value="0">Please Select....</option>
		<%
		objRS.Open "SELECT DISTINCT(SecondGroupingCalculatedFieldName) FROM tblReportLayoutLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND ReportID = 999 AND SecondGroupingCalculatedFieldName <> 'DEXP'",objCon,0,1
		
		Do until objRS.EOF
			If objRS("SecondGroupingCalculatedFieldName") = strCESubGroup Then
				strSelected = " SELECTED "
			Else
				strSelected = ""
			End if
				Response.Write "<option Value=""" & objRS("SecondGroupingCalculatedFieldName") & """" & strSelected & ">" & objRS("SecondGroupingCalculatedFieldName") & "</OPTION>"
			objRS.Movenext
		Loop
		
		objRS.Close
		%>
		</select> </td>
        <tr>  
            <td colspan="4" align="left">
                &nbsp;
            </td>
        </tr>
        <tr>
            <th colspan="4" style="text-align:left;height:20px">
                &nbsp;Description
            </th>
        </tr>
        <tr>
            <td colspan="4">
                <textarea rows="4" cols="190" id="GLCodeDesc" name="GLCodeDesc" tabindex="11"><%=strGLCodeDesc%>
	
</textarea>
            </td>
        </tr>
    </table>
    <br>
   <table Class="CallCentre" WIDTH="100%" BORDER="0" CELLSPACING="0" CELLPADDING="0">
	<tr>
	    <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="12" onClick="self.location='AdminMenu.asp';"><img src="../images/door.png" alt="" /> Close </button></td>    
        <td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="13" onclick="SaveData()";><img src="../images/tick.png" alt="" /> Save </button></td>  
        <td class='locked' Width="100px"><button type="button" tabindex="14" onClick="self.location='GLCodes.asp?GLCodeID=-1';"><img src="../images/page_white_stack.png" alt="" /> New&nbsp;&nbsp; </button></td>
        <td class='locked' Width="100px"><button type="button" tabindex="19" onclick="DeleteData()";><img src="../images/cross.png" alt="" /> Delete </button></td>
        <TD class='locked' Width="100px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessageIcon %></TD>
        <TD class='locked' Width="300px" style="BORDER-RIGHT:0px;vertical-align:middle;align-content:center"><%=strMessage %></TD>
	</tr>
</table>
    <hr>
    </form>
    <table width="100%" border="1" cellspacing="1" cellpadding="1">
       <tr>
            <th style="height:20px">
                GFS Code ID
            </th>
            <th>
                GFS Code Name Eng
            </th>
             <th>
                GFS Code Name Swa
            </th>
            <th>
                GL Code Type
            </th>
            <th>
                Active
            </th>
            <th>
                Updated By
            </th>
            <th>
                Date Updated
            </th>
        </tr>
        <%

 objRS.Open "SELECT * FROM tblGLCodes WHERE BudgetID = " & Session("BudgetID") & " Order By GLCode ASC",objCon
		Do until objRS.eof
			Response.Write "<TR><TD><A Target=""_self"" HREF=""GLCodes.asp?GLCodeID=" & objRS("GLCode") _
			& """>&nbsp;" & objRS("GLCode") & "</TD><TD>&nbsp;" _
			& objRS("GLCodeName") & "</B></TD><TD style=""text-align:left"">&nbsp;" & objRS("GLCodeNameL2") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("GLCodeType") & "</TD>" _
			& "</TD><TD style=""text-align:center"">&nbsp;" & objRS("Active") & "</TD><TD style=""text-align:center"">&nbsp;" & objRS("UpdatedBy") _
			& "</TD><TD style=""text-align:center"">&nbsp;" & objRS("DateUpdated") & "</TD></TR>"
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
		
		objRS.Open "SELECT * FROM tblGLCodes WHERE BudgetID = " & Session("BudgetID") & " AND GLCode = " & Session("GLCodeID") & "",objCon			
			
            If Not objRS.EOF Then
				lngGLCode = objRS("GLCode")
				strGLCodeName = objRS("GLCodeName")
                strGLCodeNameL2 = objRS("GLCodeNameL2")
				strGLCodeDesc = objRS("GLCodeDesc")
				strGLCodeType = objRS("GLCodeType")
				lngActualMapping = objRS("ActualGLCodeMapping")
				strPrepayment = objRS("Prepayment")
				strBalanceSheet = objRS("BalanceSheet")
				lngBSLevel1ID = objRS("BSLevel1ID")
				lngCFLevel1ID = objRS("CFLevel1ID")		
				strCEType = objRS("CEType")
				strCEGroup = objRS("CEGroup")
				strCESubGroup = objRS("CESubGroup")
				'strTriDataCEType = objRS("TriDataCEType")
				'strTriDataCEGroup = objRS("TriDataCEGroup")
				'strTriDataCESubGroup = objRS("TriDataCE")
				'strPASP = objRS("PASP")
				strReportingCategory1 = objRS("ReportingCategory1")
				strReportingCategory2 = objRS("ReportingCategory2")
				strReportingCategory3 = objRS("ReportingCategory3")
				strReportingCategory4 = objRS("ReportingCategory4")
				strReportingCategory5 = objRS("ReportingCategory5")										
				strActive = objRS("Active")	
                strSegmentValue = objRS("SegmentValue")			
				If IsNull(objRS("ReportItalic")) Then
				    strReportItalic = "N"
				Else
				    strReportItalic = objRS("ReportItalic")
				End If
                lngPostingGLCode = objRS("PostingGLCode")
				
			Else
				
				'lngGLCode = ""
				strGLCodeName = ""
                strGLCodeNameL2 = ""
				strGLCodeDesc = ""
				strGLCodeType = ""
				lngActualMapping = 0
				strPrepayment = "N"
				strBalanceSheet = "N"
				strActive = "Y"
				strReportItalic = "N"
                strSegmentValue = 0
				
			End If

		objRS.Close
	

End Sub

Sub SaveDetails()
Dim RepCat1,RepCat2,RepCat3,RepCat4, RepCat5

		If Not IsEmpty(Request.Form("GLCode")) Then
		
		 With objCmd
                .CommandType = 4
                .CommandText = "spGLCodeSave"
                
                .Parameters.Append objCmd.CreateParameter("GLCode", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("GLCodeName", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("GLCodeNameL2", adVarChar, adParamInput, 150)
                .Parameters.Append objCmd.CreateParameter("GLCodeDesc", adLongVarChar, adParamInput, -1)
                .Parameters.Append objCmd.CreateParameter("GLCodeType", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("ActualMapping", adInteger, adParamInput)   
                .Parameters.Append objCmd.CreateParameter("PostingGLCode", adInteger, adParamInput)   
                .Parameters.Append objCmd.CreateParameter("Prepayment", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("BalanceSheet", adChar, adParamInput, 1)
                .Parameters.Append objCmd.CreateParameter("BSLevel1ID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("CFLevel1ID", adInteger, adParamInput)
                
                .Parameters.Append objCmd.CreateParameter("CEType", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("CEGroup", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("CESubGroup", adVarChar, adParamInput, 50)
		        .Parameters.Append objCmd.CreateParameter("GLHierarchyLevel1", adVarChar, adParamInput, 50)
		        .Parameters.Append objCmd.CreateParameter("GLHierarchyLevel2", adVarChar, adParamInput, 50)
		        .Parameters.Append objCmd.CreateParameter("GLHierarchyLevel3", adVarChar, adParamInput, 50)
		        .Parameters.Append objCmd.CreateParameter("GLHierarchyLevel4", adVarChar, adParamInput, 50)
		        .Parameters.Append objCmd.CreateParameter("GLHierarchyLevel5", adVarChar, adParamInput, 50)

                .Parameters.Append objCmd.CreateParameter("TriDataCEType", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("TriDataCEGroup", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("TriDataCESubGroup", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("PASP", adVarChar, adParamInput, 1)
                
                .Parameters.Append objCmd.CreateParameter("ReportingCategory1", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("ReportingCategory2", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("ReportingCategory3", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("ReportingCategory4", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("ReportingCategory5", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("ReportItalic", adChar, adParamInput, 1)                 
                .Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput, 1)
                
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)                
                .Parameters.Append objCmd.CreateParameter("SegmentValue", adVarChar, adParamInput, 50)
              
				.Parameters("GLCode") = Request.Form("GLCode")	
				.Parameters("BudgetID") = Session("BudgetID")	
				.Parameters("GLCodeName") = Request.Form("GLCodeName")
                .Parameters("GLCodeNameL2") = Request.Form("GLCodeNameL2")				
                .Parameters("GLCodeDesc") = Request.Form("GLCodeDesc")               
                .Parameters("GLCodeType") = Request.Form("GLCodeType") 
                .Parameters("ActualMapping") = Request.Form("GLCode")
                .Parameters("PostingGLCode") = Request.Form("PostingGLCode")
                .Parameters("Prepayment") = "N"
                .Parameters("BalanceSheet") = "N"
                .Parameters("BSLevel1ID") = 2
                .Parameters("CFLevel1ID") = 3 
                .Parameters("CEType") = Request.Form("GLCodeType") 
                .Parameters("CEGroup") = ""
                .Parameters("CESubGroup") = Request.Form("CESubGroup")
		        .Parameters("GLHierarchyLevel1") = ""'Request.Form("GLHierarchyLevel1") 
		        .Parameters("GLHierarchyLevel2") = ""'Request.Form("GLHierarchyLevel2")
		        .Parameters("GLHierarchyLevel3") = ""'Request.Form("GLHierarchyLevel3")
		        .Parameters("GLHierarchyLevel4") = ""'Request.Form("GLHierarchyLevel4")
		        .Parameters("GLHierarchyLevel5") = ""'Request.Form("GLHierarchyLevel5")

                .Parameters("TriDataCEType") = ""
                .Parameters("TriDataCEGroup") = ""
                .Parameters("TriDataCESubGroup") = ""
                .Parameters("PASP") = "Y"
               

		If Request.Form("ReportingCategory1") = "0" Then
			RepCat1 = ""
		Else
			RepCat1 = ""'Request.Form("ReportingCategory1")
		End If
                
		If Request.Form("ReportingCategory2") = "0" Then
			RepCat2 = ""
		Else
			RepCat2 = ""'Request.Form("ReportingCategory2")
		End If

		If Request.Form("ReportingCategory3") = "0" Then
			RepCat3 = ""
		Else
			RepCat3 = ""'Request.Form("ReportingCategory3")
		End If

		If Request.Form("ReportingCategory4") = "0" Then
			RepCat4 = ""
		Else
			RepCat4 = ""'Request.Form("ReportingCategory4")
		End If
		
		If Request.Form("ReportingCategory5") = "0" Then
			RepCat5 = ""
		Else
			RepCat5 = ""'Request.Form("ReportingCategory5")
		End If

		        .Parameters("ReportingCategory1") = RepCat1 
                .Parameters("ReportingCategory2") = RepCat2 
                .Parameters("ReportingCategory3") = RepCat3 
                .Parameters("ReportingCategory4") = RepCat4 
                .Parameters("ReportingCategory5") = RepCat5 

	
                .Parameters("ReportItalic") = "N"
                .Parameters("Active") = Request.Form("Active")
                .Parameters("UpdatedBy") = Session("UserID")
                .Parameters("SegmentValue") = Request.Form("SegmentValue")
                            
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute          
               
			'Return the result of the Save Function.
     		Session("GLCode") = Request.Form("GLCodeID")
     		strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
            strMessage = "<B>RECORD SAVED.</B>"
			
						
	End If
				
	

End Sub	

Sub DeleteRecord(GLCode)
    
objRS.Open "SELECT * FROM tblBudgetData WHERE BudgetID = " & Session("BudgetID") & " AND GLCode = " & GLCode & "",objCon

    If objRS.EOF Then
        objCon.Execute "DELETE FROM tblGLCodes WHERE BudgetID = " & Session("BudgetID") & " AND GLCode = " & GLCode & "" 
        objCon.Execute "DELETE FROM tblReportLayoutLevel2 WHERE BudgetID = " & Session("BudgetID") & " AND GLCode = " & GLCode & ""     
        strMessage = "<B>RECORD DELETED.</B>"
        strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
        
    Else
        strMessage = "<FONT Color=""Red""><B>GLCODE CANNOT BE DELETED BECAUSE ENTRIES EXIST FOR THIS GLCODE.</B></FONT>"
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
