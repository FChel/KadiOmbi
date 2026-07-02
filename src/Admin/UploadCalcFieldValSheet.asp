<%@ Language=VBScript %>
<!-- #include file="upload.asp" -->
<!-- #Include file=../ADOVBS.inc -->
<%
    If IsEmpty(Session("BudgetID")) Then Response.Redirect("../Timeout.asp")

    Session("CurrentPage") = "Appropriation/UploadInputSheet.asp"

%>
 
<html>
<head>
<link rel="stylesheet" type="text/css" href="../BERTStyle.css">
<script language=javascript>
function upload()
{   
    if(document.getElementById('FILE1').value=="")
    {
        alert("Please select the file to upload");
    }
    else
    {   
        if(getFileExt(document.getElementById('FILE1').value)==".xls")
        {
            document.getElementById('Progress').style.display = "inline";
            frm.submit();
        } 
        else
        {
            alert("Please enter a valid Excel(.xls) file");
        }
    }
}

 function getFileExt(filename)
 {
     var s
     s = filename.charAt(filename.length-4) + filename.charAt(filename.length-3) + filename.charAt(filename.length-2)+ filename.charAt(filename.length-1);
     return s;
 }

 function DeleteBudgetData() {

     var message
     var uploadid = prompt('Enter the Upload ID of the batch you wish to delete.');

     message = 'Would you like to DELETE all the existing data for upload batch upload ' + uploadid + '?'

     if (window.confirm(message) == true) {
         
         document.getElementById('Progress').style.display = "inline";
         self.location = "UploadInputSheet.asp?Action=Delete&UploadID=" + uploadid
     }
     else {

     }
 }

 function CloseScreen() {        

    self.location = 'FundsSource.asp';
                  
}

</script>
</head>
<body>
<H3>Upload Resource Envelope Funding Transactions</H3>
<form action="UploadInputSheet.asp?Action=Save" method="POST" enctype="multipart/form-data" id="frm" name="frm">
<table BORDER="1" CELLSPACING="1" CELLPADDING="1" width="100%">
<tr>
    <th width=100% align=left height="25px">Funding Sources Upload Screen</th>
</tr>
</table>

<br />
<table BORDER="1" CELLSPACING="1" CELLPADDING="1" width="100%">
<tr>
    <th align=left width="180" height="25px">Select a file to upload</th>
    <td><INPUT TYPE=FILE SIZE=50 NAME="FILE1" id="FILE1" ></td>
</tr>
<tr>
    <td></td>
    <td Height="25px"><button type="button" name=tabindex="15" onclick="javascript:upload();"><img src="../images/database_save.png" alt="" /> Load File</button></td>
</tr>
<tr><td Height="25px">&nbsp;</td><td><span id="Progress" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Processing...</b></span></td></tr>
</table>
<br>
<a href="Funding_Upload_Template.xls" target="_top">Click here to download Funding Sources Upload Template.....</a>
<br>
<BR>
<hr>
<table Width="400px" BORDER="0" CELLSPACING="0" CELLPADDING="0">
<tr>
	<td class='locked' Width="100px" style="border-right:0px"><button type="button" tabindex="15" onclick="CloseScreen('<%=Session("TransactionType")%>')";><img src="../images/door.png" alt="" /> Close </button></td> 
	<td class='locked' Width="200px"><button type="button" tabindex="11" onclick="DeleteBudgetData()";><img src="../images/wrench.png" alt="" /> Delete Existing Data</button></td>    


</tr>
</table>
<hr>
<table BORDER="1" CELLSPACING="1" CELLPADDING="1" width="100%">
    <tr>
        <th Height="20px" Style="text-align:left" colspan="5">&nbsp;Existing Uploads</th>
    </tr>
    <tr><td colspan="5">&nbsp;</td></tr>
    <tr>
        <th Height="20px">Upload ID</th><th>Uploaded By</th><th>Date Uploaded</th><th>Record Count</th><th>Upload Total</th>
    </tr>


<%

'On Error Resume Next 
Dim objCon
Dim objExcelCon
Dim objCmd
Dim objRS
Dim objRS1
Dim objRS2
Dim strUploadPath
dim errors
dim lineNo
dim strCalcField
dim strProgramCode
dim strDescription
dim dblBM1
dim dblBM2
dim dblBM3
dim dblBM4
dim dblBM5
dim dblBM6
dim dblBM7
dim dblBM8
dim dblBM9
dim dblBM10
dim dblBM11
dim dblBM12
dim dblOY1
dim dblOY2
dim dblOY3
dim dblOY4
dim dblOY5
dim updatedBy
dim strUploadID



public intBusinessAreaID

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
objCon.Open Session("DBConnection")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")
Set objRS2 = Server.CreateObject("ADODB.Recordset")

If Validate_Access(Session("UserTypeID"),Session("CurrentPage")) = "N" Then
     Response.Redirect "../AccessDenied.asp"
End If
  
 If Request.QueryString("Action") = "Delete" Then
    If Session("StatusID") = 1 Then
        Call DeleteData(Request.QueryString("UploadID"))
 Else
         Response.Write "&nbsp;&nbsp;<img src=""../images/warning.gif"" /><B><FONT Color=""Red""><H3>WARNING - BUDGET IS NOT OPEN, CHANGES CANNOT BE MADE.</H3></FONT></B><BR>" 
        strMessage = "BUDGET IS NOT OPEN, CHANGES CANNOT BE MADE."
    End If
 End If

    objRS.Open "SELECT * FROM qryFundAppropriationBatchUploadLog WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND UpdatedBy = " & Session("UserID") & "",objCon

        Do Until objRS.EOF

                Response.Write "<TR><TD>" & objRS("UploadID") & "</TD><TD Style=""text-align:Center"">" & objRS("FName") & " " & objRS("LName") & "</TD><TD Style=""text-align:Center"">" & objRS("DateUpdated") & "</TD><TD Style=""text-align:Right"">" & objRS("RecordCount") & "</TD><TD Style=""text-align:Right"">" & formatnumber(objRS("Total"),0,0) & "</TD></TR>"
                dblTotal = dblTotal + objRS("Total")
            objRS.Movenext

        Loop

    objRS.Close
    If IsNull(dblTotal) Then dblTotal = 0
    Response.Write "<TR><TH Align=""Right"" Colspan=""4"">Total</TH><TD Style=""text-align:Right""><B>" & formatnumber(dblTotal,0,0) & "</B></TD></TR>"

errors = "" 
lineNo = 1    
updatedBy = Session("UserID")

Dim Uploader, File, filePath

Set Uploader = New FileUploader

' This starts the upload process
Uploader.Upload()

if request.QueryString("Action")="Save" AND Session("StatusID") = 1 Then
  
' Check if any file is sucessfully uploaded
    If Uploader.Files.Count = 0 Then
	    Response.Write "File(s) not uploaded."
    Else
        ' Loop through the uploaded file
	    For Each File In Uploader.Files.Items					  		    		    
		    'set the upload path
            strUploadPath = Server.MapPath(GetFilePath()) & "\Attachments"
			File.SaveToDisk strUploadPath			
		    filePath = Server.MapPath(GetFilePath()) & "\Attachments\" & File.FileName
   
	    Next
	    
	    'After uploading, Read excel file
	    
	    Set objExcelCon = Server.CreateObject("ADODB.connection")     
        objExcelCon.Open "DBQ=" & filePath & "; DRIVER={Microsoft Excel Driver (*.xls)};" 
        'objExcelCon.Open "DBQ=" & filePath & "; Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=path to xls/xlsx/xlsm/xlsb file"
          
        'sFileConnectionString = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq="&sFilePath&";DefaultDir="&sDataDir&";"
        
        'Write the SQL Query 
        objRS.open "SELECT * FROM UploadZone", objExcelCon  
	           
	    ReadExcel 
	    
	    'Check for errors
        if(errors<>"") then
            'Print the errors and return
             Response.Write "<img src=""../images/attention.jpg"" /><FONT Color=""Red""><H3>File not uploaded. Please correct the followings errors and try again.</H3></FONT><BR> "& errors &"<br>"         
                        
        else  
            'If the no errors found in the ReadExcel method, then start uploading the records in the database       	    	    
	        UploadExcel

              strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
              strMessage = "&nbsp;&nbsp;<B>UPLOAD SUCCESSFUL.</B>"
	          Response.Write "<BR>" & strMessageIcon & "   " & strMessage & "<BR><BR>"
	       
	    End if
	    	    	                
    'Close the recordset/connection 
    objRS.Close 
    objExcelCon.Close 
	    
    End If
    
Else


        

End if

'Function for validating the excel file values
Sub ReadExcel    

    'First loop to check all the values
    Do until objRS.EOF 
        'Read each recors present in the Excel file and check for the validation 

        strVote = objRS("Vote") 'Validate against Votes
            objRS1.Open "SELECT BusinessAreaCode,PrimaryBusinessArea,BusinessAreaID FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaCode = '" & strVote & "'",objCon
                If objRS1.EOF Then                
                    errors = errors & "Error in row no. " & lineNo + 1 & ": " & strVote & " - Vote is not valid.<br>"
                    strBudClassChk = "N" 
                    intBusinessAreaID = 0
                Else
                    strBudClassChk = objRS1(1)
                    intBusinessAreaID = objRS1(2)
                End IF              
            objRS1.Close

         strVote = objRS("Vote") 'Validate against Votes
            'Response.Write "SELECT Approved FROM tblBACeilingLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = 1000 AND Level1ID = " & intBusinessAreaID & "" 
            objRS1.Open "SELECT Approved FROM tblBACeilingLevel1 WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND BusinessAreaID = 1000 AND Level1ID = " & intBusinessAreaID & " AND Approved = 'on'",objCon
                If Not objRS1.EOF Then                
                    errors = errors & "Error in row no. " & lineNo + 1 & ": Expenditue Ceiling for Vote has been approved no changes can be made to Resource Envelope.<br>"
                 
                Else
                 
                End IF              
            objRS1.Close

        strCalcField = objRS("Calculated Field") 'Validate against Calculated Field
          
            objRS1.Open "SELECT tblCalculatedFields WHERE BudgetID = " & Session("BudgetID") & " AND CalculatedField = '" & strCalcField & "'",objCon
            
                If objRS1.EOF Then                
                    errors = errors & "Error in row no. " & lineNo + 1 & ": " & strCalcField & " - Calculated Field code is not valid or is not valid for the Vote.<br>"
             
                End IF              
            objRS1.Close
            
        strDescription = objRS("Description") 'Should be integer and Not Null values  
        if(strDescription="" OR isNull(strDescription)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Description must NOT be empty <br>"
        End if
    
        dblBM1 = objRS("Period1") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM1)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 1 must be a numeric and NOT Null value <br>" 
        End if
    
        dblBM2 = objRS("Period2") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM2)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 2 must be a numeric and NOT Null value <br>" 
        End if

        dblBM3 = objRS("Period3") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM3)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 3 must be a numeric and NOT Null value <br>" 
        End if

        dblBM4 = objRS("Period4") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM4)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 4 must be a numeric and NOT Null value <br>" 
        End if

        dblBM5 = objRS("Period5") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM5)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 5 must be a numeric and NOT Null value <br>" 
        End if

        dblBM6 = objRS("Period6") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM6)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 6 must be a numeric and NOT Null value <br>" 
        End if

        dblBM7 = objRS("Period7") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM7)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 7 must be a numeric and NOT Null value <br>" 
        End if

        dblBM8 = objRS("Period8") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM8)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 8 must be a numeric and NOT Null value <br>" 
        End if

        dblBM9 = objRS("Period9") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM9)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 9 must be a numeric and NOT Null value <br>" 
        End if

        dblBM10 = objRS("Period10") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM10)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 10 must be a numeric and NOT Null value <br>" 
        End if

        dblBM11 = objRS("Period11") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM11)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 11 must be a numeric and NOT Null value <br>" 
        End if

        dblBM12 = objRS("Period12") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblBM12)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Period 12 must be a numeric and NOT Null value <br>" 
        End if

        dblOY1 = objRS("Year + 1") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblOY1)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Year + 1 must be a numeric and NOT Null value <br>" 
        End if     
        
        dblOY1 = objRS("Year + 2") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblOY2)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Year + 2 must be a numeric and NOT Null value <br>" 
        End if 

        dblOY2 = objRS("Year + 3") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblOY3)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Year + 3 must be a numeric and NOT Null value <br>" 
        End if 

        dblOY1 = objRS("Year + 4") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblOY4)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Year + 4 must be a numeric and NOT Null value <br>" 
        End if 

        dblOY1 = objRS("Year + 5") 'Should be integer and Not Null values       
        if NOT (IsNumeric(dblOY5)) then
            errors = errors & "Error in row no. " & lineNo + 1 & ": Year + 5 must be a numeric and NOT Null value <br>" 
        End if 
                          
            lineNo = lineNo + 1            
            
        objRS.movenext 
    
        If IsNull(objRS("Vote")) Then       
          
            exit Sub

        End If
    
    Loop        
    
End Sub	

'Function for validating the excel file values
Sub UploadExcel
              
    objRS.MoveFirst()
    
    Do until objRS.EOF 
        
        strVote = objRS(0)
        strProgramCode = objRS(1)
        strCalcField = objRS(2)   
        dblBM1 = objRS(3)
        dblBM2 = objRS(4)
        dblBM3 = objRS(5)
        dblBM4 = objRS(6)
        dblBM5 = objRS(7)
        dblBM6 = objRS(8)
        dblBM7 = objRS(9)
        dblBM8 = objRS(10)
        dblBM9 = objRS(11)
        dblBM10 = objRS(12)
        dblBM11 = objRS(13)
        dblBM12 = objRS(14)
        dblOY1 = objRS(15)
        dblOY2 = objRS(16)
        dblOY3 = objRS(17)
        dblOY4 = objRS(18)
        dblOY5 = objRS(19)
        strDescription = objRS(20)
               
        'Save the record
        SaveRecord strVote,strProgramCode,strCalcField,dblBM1,dblBM1,dblBM2,dblBM3,dblBM4,dblBM5,dblBM6,dblBM7,dblBM8,dblBM9,dblBM10,dblBM11,dblBM12,dblOY1,dblOY2,dblOY3,dblOY4,dblOY5,strUploadID
     
     objRS.movenext 
    
    If IsNull(objRS("Vote")) Then               
       
       exit Sub
       
    End If
        
    Loop 
    
End Sub	

Sub SaveRecord(Vote,ProgramCode,CalculatedField,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,Y1,Y2,Y3,Y4,Y5)

 
    Dim BusinessAreaID
    Dim CostCentreID
    Dim strMessage
    Dim strMessageIcon    

    'Convert Vote to BusinessAreaID

    objRS1.Open "SELECT BusinessAreaID FROM tblBusinessArea WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaCode = '" & Vote & "'",objCon
        If Not objRS1.EOF Then                
            BusinessAreaID = objRS1(0)
        Else
            BusinessAreaID = 0
        End IF              
    objRS1.Close

     objRS1.Open "SELECT CostCentreID FROM tblCostCentres WHERE BudgetID = " & Session("BudgetID") & " AND BusinessAreaCode = '" & Vote & "' AND ProgramCode = '" & ProgramCode & "'",objCon
        If Not objRS1.EOF Then                
            CostCentreID = objRS1(0)
        Else
            CotCentreID = 0
        End IF              
    objRS1.Close

   Set ObjCmd = Server.CreateObject("ADODB.Command")
  	   With objCmd
                .CommandType = 4
                .CommandText = "spFundAppropriationSave"
                
                .Parameters.Append objCmd.CreateParameter("FundAppropriationID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("BudgetID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("VersionID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("YearID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("FundID",adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("Amount1",adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Amount2",adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("Amount3",adDouble, adParamInput)
                .Parameters.Append objCmd.CreateParameter("AdviceDate", adDate)     
                .Parameters.Append objCmd.CreateParameter("CostObjectID", adInteger, adParamInput)                                                                            
                .Parameters.Append objCmd.CreateParameter("Comments", adLongVarChar, adParamInput,-1) 
                .Parameters.Append objCmd.CreateParameter("Approved", adChar, adParamInput,1) 
                .Parameters.Append objCmd.CreateParameter("ReferenceID", adVarChar, adParamInput, 50)   
                .Parameters.Append objCmd.CreateParameter("Attachment", adVarChar, adParamInput, 500)   
                .Parameters.Append objCmd.CreateParameter("UpdatedBy", adChar, adParamInput,6) 
                .Parameters.Append objCmd.CreateParameter("UploadID",  adVarChar, adParamInput, 50)  
                .Parameters.Append objCmd.CreateParameter("CostCentreID",  adVarChar, adParamInput, 50) 
                .Parameters.Append objCmd.CreateParameter("FundingPurposeID", adInteger, adParamInput) 
                .Parameters.Append objCmd.CreateParameter("ProjectID", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("FundingPurposeL2ID", adInteger, adParamInput)
                .Parameters.Append objCmd.CreateParameter("FundTypeID", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("FinancingID", adVarChar, adParamInput, 50)
                .Parameters.Append objCmd.CreateParameter("SatelliteRecordID", adInteger, adParamInput)  
                .Parameters.Append objCmd.CreateParameter("Response",  adVarChar, adParamOutput, 100)                          
                .Parameters.Append objCmd.CreateParameter("FundAppropriationIDOutput", adInteger, adParamOutput)
                .Parameters.Append objCmd.CreateParameter("A", adDouble, adParamOutput)
                .Parameters.Append objCmd.CreateParameter("B", adDouble, adParamOutput)
                .Parameters.Append objCmd.CreateParameter("C", adDouble, adParamOutput)
                .Parameters.Append objCmd.CreateParameter("D", adDouble, adParamOutput)
               				 				 
    		    .Parameters("FundAppropriationID") = 0
                .Parameters("BudgetID") = Session("BudgetID")
				.Parameters("VersionID") = Session("VersionID")
				.Parameters("YearID") =  1
				.Parameters("FundID") = FundCode				
                .Parameters("Amount1") = clng(Amount)
                .Parameters("Amount2") = clng(Year1)
                .Parameters("Amount3") = clng(Year2)
			    .Parameters("AdviceDate") = Null
                .Parameters("CostObjectID") = cint(BusinessAreaID)														
				.Parameters("Comments") = Description		
				.Parameters("Approved") = Approved
				.Parameters("ReferenceID") = Null
			    .Parameters("Attachment") = Null																									
				.Parameters("UpdatedBy") = Session("UserID")
                .Parameters("UploadID") = UploadID
                .Parameters("CostCentreID") = 0
                .Parameters("FundingPurposeID") = cint(strFundingPurposeID)
                .Parameters("ProjectID") = Project
                .Parameters("FundingPurposeL2ID") = cint(strFundingPurposeL2ID)
                .Parameters("FundTypeID") = FundType
                .Parameters("FinancingID") = Financing
                .Parameters("SatelliteRecordID") = 0
				
				.ActiveConnection = objCon
				
			    objCmd.Execute	

                strMessage = objCmd.Parameters.Item("Response")   
                lngFundAppropriationID = objCmd.Parameters.Item("FundAppropriationIDOutput")
                strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
                If strMessage <> "OK" Then
                     Response.Write strMessageIcon & "   " & strMessage
                End If
    End With                            			     				     		     		
                        
End Sub

'Get virtual Pathname and set to where spreadsheets are created and opened from.
Function GetFilePath()
  Dim lsPath, arPath

  ' Obtain the virtual file path. The SCRIPT_NAME
  lsPath = Request.ServerVariables("SCRIPT_NAME")
    
  ' Split the path along the /s.
  arPath = Split(lsPath, "/")

  ' Set the last item in the array to blank string
  ' (The last item actually is the file name)
  arPath(UBound(arPath,1)) = ""

  ' Join the items in the array.
  GetFilePath = Join(arPath, "/")
End Function

If Err.Number <> 0 then
    Response.Write "<font face=arial size=1><b>Error occured while trying to process the request<b><br>Please contact system administrator</font><br>"
    Error.Clear
End If

Public Sub DeleteData(UploadId)

Dim strCostCentreName
Dim strCeilingStatus

          If Session("StatusID") = 1 Then
         
                 
                    objRS.Open "SELECT DISTINCT(UpdatedBy) FROM tblFundAppropriation WHERE BudgetID = " & Session("BudgetID") & " AND VersionID = " & Session("VersionID") & " AND UploadID = '" & UploadID & "' AND UpdatedBy = " & Session("UserID") & "",objCon
                
                        If Not objRS.EOF Then
           
                           objCon.Execute "spPurgeFundAppropriationData " & Session("BudgetID") & "," & Session("VersionID") & ",'" & UploadID & "'," & Session("UserID") & ""
		                    'Response.Write "spPurgeSubProgramBudgetData " & Session("BudgetID") & "," & Session("VersionID") & "," & Session("SubProgrammeID") & ",'" & Session("TransactionType") & "'," & Session("UserID") & ""
                            strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" />"
                            Response.Write "<BR>" & strMessageIcon & "<b>All data for batch upload <I>" & UploadId & "</I> has been be DELETED!</b><br><BR>"
                     
                        
                        Else
                            Response.Write "<img src=""../images/warning.gif"" />&nbsp;&nbsp;<font color=""red""><b>Invalid Upload ID. Please ensure Upload ID entered is valid.</font><br><BR>"
                        End If

                    objRS.Close    
            
                          
        Else
                 Response.Write "&nbsp;&nbsp;<img src=""../images/warning.gif"" />&nbsp;&nbsp;<B><FONT Color=""Red"">WARNING - BUDGET IS NOT OPEN, CHANGES CANNOT BE MADE.</FONT></B><BR><BR>" 
                  strMessage = "BUDGET IS NOT OPEN, CHANGES CANNOT BE MADE."
        End If  



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

Set objRS = Nothing
Set objRS1 = Nothing
Set objRS2 = Nothing
Set objCon = Nothing


 %>
    </table>
</form>
 </body>
</html>

