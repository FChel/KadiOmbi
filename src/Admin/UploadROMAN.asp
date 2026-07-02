<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #include file="upload.asp" -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #include file="../CC/CAPSFunctions.asp" -->
<%
'Description:	Genera Expenses Upload Administration screen
'Author:		MG
'Date:			April 2013

    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")
	
'Instantiate Common Page Variables.
Dim objCon
Dim objCmd
Dim objRS

Dim strDeleteCheck
Dim dteBatchDate

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

If request.QueryString("Action")="Save" Then

	Call StartLoad()
End If

If Not IsEmpty(Request.QueryString("Reload")) Then

	Call StartLoad()
End If

'If the local load has been clicked then call the procedure to load the network file rather than uploading it
If request.QueryString("Action")="SaveFileLocal" Then

	Call StartLoadLocal()
	
End If

If request.QueryString("Action")="SaveFileDB" Then

	Call StartLoadDB()
End If

'''''---New May 2025 ---Load the CMS data directly from the ProMaster database
'''via a Linked SQL server, so all functionality is in the Stored Procedure
If request.QueryString("Action")="SaveFileLinked" Then
	Call StartLoadLinked()
End If



If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
End If

'Response.Write "B=" & 	dteBatchDate
'dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
'Response.Write " 2B=" & 	dteBatchDate

'	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
'	Response.Write " 3B=" & 	dteBatchDateFormat
 %>
		
<script language=javascript>

function upload()
{   
    if(document.getElementById('FILE1').value=="")
    {
        alert("Please select the file to upload (click the Browse button above)");
    }
    else
    {   
        if(getFileExt(document.getElementById('FILE1').value)==".csv" || getFileExt(document.getElementById('FILE1').value)==".txt")
        {
            if(window.confirm('This will overwrite any existing CDMC file data! \n \n Continue?')==true)
                {
                document.getElementById('Progress').style.display = "inline";
                frm.submit();
            }
        } 
        else
        {
            alert("Please enter a valid CSV(.csv) file");
        }
    }
}

function UploadLinkedCMS()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="UploadROMAN.asp?Action=SaveFileLinked"
}

function UploadLocal()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="UploadROMAN.asp?Action=SaveFileDB"
}

function UploadLocalG()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="UploadROMAN.asp?Action=SaveFileLocal&ActionType=Service"
}

 function getFileExt(filename)
 {
     var s
     s = filename.charAt(filename.length-4) + filename.charAt(filename.length-3) + filename.charAt(filename.length-2)+ filename.charAt(filename.length-1);
     return s;
}

function ShowButtons() {
	document.getElementById('ShowButt').style.display = "inline";
}

jQuery(document).ready(function($) {
    $(".clickable-row").click(function() {
        window.location = $(this).data("href");
    });
});

function GetDateSelect() {
  //var xhttp = new XMLHttpRequest();
  //xhttp.onreadystatechange = function() {
  alert(this.responseText);
  //  if (this.readyState == 4 && this.status == 200) {
  //   document.getElementById("EmployeeIDST").innerHTML = this.responseText;
  //  }
  //};
  //xhttp.open("GET", "../../CC/AJAX/GetEmployees.asp?EmpID=" + frm.EmpIDS.value + "&FName=" + frm.FNamms.value + "&LName=" + frm.LNamms.value + "", true);
  //xhttp.send();
}

function DatePickChange() {
	self.location="UploadROMAN.asp?FileDate=" + document.getElementById("CSDate").value;
}
</script>
<script src="../js/jquery.js"></script>

<body>
<main class="main py-3">
      <div class="container">
<form action="UploadROMAN.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	<div class="row" id="basic-table">
  <div class="col-3">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../images/defence_logo_dark.png" height="40px" width="110px" title="Diners"> Chart of Accounts File Load</h4>
        </div>
      <div class="card-content">
        <div class="card-body">
		

<div class="col-lg-12 col-md-12">

</div>
<script>
	//Script to display the file selected in the File Load Input field (it doesn't by default)
	$('#FILE1').on('change',function(){
	//alert('asas');
		//get the file name
		var fileName = $(this).val();
		
		fileName = fileName.replace("C:\\fakepath\\", "");
		
		//replace the "Choose a file" label
		$(this).next('.custom-file-label').html(fileName);
	})
</script>
<div class="form-body">

<div class="row col-12" style="margin:5px; padding:0px;">

	<button type="button" class="btn btn-primary btn-sm" onclick="UploadLinkedCMS();" Title="Click to Load Cost Centres for ERP from ProMaster"><i class="fa fa-upload"></i> Load ERP from CMS</button>&nbsp;
	
</div>
<br><br>

<div class="row col-12" style="margin:5px; padding:0px;">

	<button type="button" class="btn btn-outline-secondary btn-sm" onclick="ShowButtons();" Title="Click to Load from old buttons (Pre ERP)"><i class="fa fa-eye"></i> Show Old Load Buttons</button>
 
</div>
<div class="row col-12" style="margin:5px; padding:0px; display:none" id="ShowButt">

	<button type="button" class="btn btn-outline-secondary btn-sm" onclick="UploadLocal();" Title="Click to Load any existing file in the G Drive Imports Folder"><i class="fa fa-upload"></i> Load CMS</button>&nbsp;
	<button type="button" class="btn btn-outline-secondary btn-sm" onclick="UploadLocalG();" Title="Click to Load any existing file in the G Drive Imports Folder"><i class="fa fa-upload"></i> Load G</button>&nbsp;
 
</div>
</div>

<div class="col-lg-12 col-md-12">
<div class="py-3"> 
<span id="Progress" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>
<br>
</div>
</div>

<div class="col-lg-12 col-md-12">

</div>

		</div>
	  </div>
    </div>
   </div>
  

  <div class="col-7">
    <div class="card">
     
      <div class="card-content">
        <div class="card-body">
		
		<%
        
      DisplaySummary()
        
%>	
		</div>
	  </div>
    </div>
   </div>
   
   <div class="col-2">
    <div class="card">
     
      <div class="card-content">
        <div class="card-body">
		
		<%
        
      DisplayFileSummary()
        
%>	
		</div>
	  </div>
    </div>
	
	 <div class="card">
     
      <div class="card-content">
        <div class="card-body">
		
		<%
      DisplayFileSummaryG()
				%>	
		</div>
	  </div>
    </div>
	
   </div>
   
   
  </div>
  
    <!-- Example DataTables Card-->
     <div class="row" id="basic-table">
  <div class="col-12">
    <div class="card">
     
      <div class="card-content">
        <div class="card-body">

          <!-- Table with outer spacing -->
          <div class="table-responsive">
         
				<%
        
      DisplayTableDetails()
        
%>	

          </div>
        </div> 
          </div>
        </div>
       </div>
      </div>
    </div>
</div>
</form>
 </div>
    </main>
</body>

<!-- #Include file=../CC/CAPSFooter.asp -->
</html>
<%
Sub DisplayTableDetails()

Dim strWhere

If Not IsEmpty(Request.QueryString("BatchNo")) Then
	If IsNull(Request.QueryString("BatchNo")) or Request.QueryString("BatchNo")= "" Then 
		
	Else
		strWhere = "WHERE FileID = " & Request.QueryString("BatchNo") & ""
	
		If Request.QueryString("BatchNo") = 0 Then strWhere = strWhere & " OR FileID IS NULL"
	End If
Else
	strWhere = ""
End If

objRS.Open "SELECT TOP 50 * FROM qryCAPSCostCentreSummary "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no COA data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Upload Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover table-compact mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing COA Data already in CAPS2</th></tr>" & _
                "<tr><td colspan=""16"" style=""text-align:left; color:red;font-size:20px""><B>WARNING! The data below will be deleted if you Upload a new COA file!</B></td></tr><tr>" & _
		        "<th>File ID</th>" & _
				"<th>Rec Type</th>" & _
				"<th>Cost Centre Number</th><th>Cost Centre Name</th>" & _	
		        "<th>Company Code</th>" & _
	 	        "<th>Updated By</th>" & _	
	 	        "<th>Date Updated</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR><TD>" & objRS("FileID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("RecType") & "</TD><TD style=""text-align:center"">" & objRS("CostCentreNumber") & "</TD><TD style=""text-align:center"">" & objRS("CostCentreName") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("CompanyCode") & "</TD><TD style=""text-align:center"">" & objRS("UpdatedBy") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD>" & _
			                    "</TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close

        Response.Write "</table>"
		
End Sub

Sub DisplaySummary()

Dim lngCards
Dim lngEmployees
Dim lngBatchNo
Dim lngDTC
Dim lngCMC
Dim lngOther
Dim strCardType
Dim lngTotalRecords
Dim lngBatchNo1

Dim dteBatchDateFormat
Dim srStatus1
Dim strDateUpdated
Dim x
Dim strDateLoadColour
Dim strDateTitle

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		objRS.Open "SELECT TOP 6 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE DateLoaded > '" & dteBatchDate & "' AND (FileType = 'ERPCostCentres' OR FileType = 'ROMANCostCentres') ORDER BY FileLoadID DESC",objCon
		'objRS.Open "SELECT TOP 20 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'ROMANCostCentres' ORDER BY FileSeqNum DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no COA Cost Centre data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-compact mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""5"" style=""text-align:left"">COA Cost Centres Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
		        "<tr><th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th><th>Deleted</th></tr>" 
				
		    End If
		    
		    Do until objRS.eof
				
				x = x + 1
				
				If DateDiff("d",objRS("DateLoaded"),Now()) = 0 Then
					strDateLoadColour = " Color:Green; font-weight:bold;"
					strDateTitle = " title=""COA File for today has been LOADED"" "
				Else
					strDateLoadColour = ""
					strDateTitle = ""
				End If
				
				Response.Write "<TR><TD><a href=""UploadROMAN.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center"">" & objRS("RecordCount") & "</TD><TD style=""text-align:center"">" & objRS("Status") & "</TD>" & _
							"<TD style=""text-align:center; " & strDateLoadColour & """ " & strDateTitle & ">" & objRS("DateLoaded") & "</TD><TD style=""text-align:center"">" & objRS("Deleted") & "</TD></TR>"
							
    			objRS.Movenext			
		    Loop
    			
			
								
	    objRS.Close

        Response.Write "</table>"
		
End Sub

Sub StartLoad()

'On Error Resume Next 
Dim objExcelCon
Dim strUploadPath
Dim strFileName

Dim errors
Dim lineNo

				
errors = "" 
lineNo = 1    
strUpdatedBy = Session("UserID")

Dim Uploader, File, filePath

Set Uploader = New FileUploader

' This starts the upload process
Uploader.Upload()

if request.QueryString("Action")="Save" Then

	'If Session("StatusID") <> 1 Then
   
    '    Response.write "<Font Color=red><B>Budget is closed, no changes can be made!</B></FONT>"
	'Else
	' Check if any file is sucessfully uploaded
	    If Uploader.Files.Count = 0 Then
		    Response.Write "No File(s) uploaded."
	    Else
			'Set the global (page) variable to the Check Delete value for use when processing data
			strDeleteCheck = Uploader.Form("chkDelete")
			
		    ' Loop through the uploaded file
		    For Each File In Uploader.Files.Items					  		    		    
			    'set the upload path
	            strUploadPath = Server.MapPath(GetFilePath()) & "\Attachments"
				File.SaveToDisk strUploadPath
			    filePath = Server.MapPath(GetFilePath()) & "\Attachments\" & File.FileName
				strFileName = File.FileName
		    Next
			
			'Call the relevant procedure depending on whether the file is .xls or .txt
			If Right(filePath,3) = "xls" Then
			
				'After uploading, Read excel file
				Set objExcelCon = Server.CreateObject("ADODB.connection")     
				'objExcelCon.Open "DBQ=" & filePath & "; DRIVER={Microsoft Excel Driver (*.xls)};" 
				'objExcelCon.Open "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & filePath & ";DefaultDir=c:\Apps\CAPS2\ASP2\Admin\CAPSAdmin\Attachments;" 
				objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=YES"";"
				
				'objExcelCon.Open strUploadPath & "\CAPSCSFromDiners.dsn"
				
				'Write the SQL Query 
				objRS.open "SELECT * FROM [CSFromDinersFile$]", objExcelCon  
		    
				ReadExcel 
		    
				
				'Check for errors
				if(errors<>"") then
					'Print the errors and return
					Response.Write "<font face=arial size=1><b>File not uploaded. Please correct the followings errors and try again</b> <br> "& errors &"</font><br>"         
				else  
					'If the no errors found in the ReadExcel method, then start uploading the records in the database       	    	    
					UploadExcel Uploader.Form("chkDelete")
					Response.Write "<b>COA File Successfully Uploaded!!<b><br>"
				End if
				
				
				
				'Close the recordset/connection 
				objRS.Close 
				objExcelCon.Close 
		    
			Else
				ReadText filePath,strFileName
			End If
			
	    'End If
	End IF
	
End if

End Sub

Sub StartLoadLocal()
'Procedure to load the local file from within the network, rather than loading the file to the server
Dim strFileName
Dim File, filePath

	If Request.QueryString("Action")="SaveFileLocal" Then

		filePath = "D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\Attachments\"
		strFileName ="ZCMS_COA_20200623.txt"

		'Call the procedure to read the file
		ReadText filePath,strFileName
			
	End if

End Sub

Sub StartLoadDB()
'Procedure to load the local file from within the network, rather than loading the file to the server
Dim strFileName
Dim File, filePath,strFileSeqNum, strFileDateTime, lngFileLoadID, intRecord, strFileType

	strFileSeqNum = GetSystemAdmin("ROMANCostCentres")
	Response.Write strFileSeqNum

	strFileDateTime = Day(Now()) & Month(Now()) & Year(Now())
	
	intRecord = 0
	lngFileLoadID = SaveFileLoadID ("ROMANCostCentres","","",0,0,0,0,0,0,0,intRecord,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
	'lngFileLoadID = SaveFileLoadID ("ROMANCostCentres","","",y,0,0,0,0,0,0,intRecord,"","","Imported",Session("UserID"),"N")
	
	response.Write lngFileLoadID

	If Request.QueryString("Action")="SaveFileDB" Then

		'Call the procedure to read the file
		  	With objCmd 
  	 
                .CommandType = 4
                .CommandText = "spCAPSLoadCMSCostCentres"

				.Parameters.Append objCmd.CreateParameter("FileID", adInteger)
				.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
				.Parameters.Append objCmd.CreateParameter("RecordCount", adInteger, adParamOutput)
		
				.Parameters("FileID") = lngFileLoadID
				.Parameters("UpdatedBy") = Session("UserID")
                           
               .ActiveConnection = objCon
                
          End With
                
            'objCmd.Execute        
            
            'Return the result of the Save Function.
     		'intRecord = objCmd.Parameters.Item("RecordCount")   
			intRecord = 1234
			
			Response.Write "XXX"
		
''			
			
		'response.write "fie=" & lngFileLoadID
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
		Call UpdateFileLoadSummary ("ROMANCostCentres",strFileSeqNum, strFileName,lngFileLoadID)
			
	End if

End Sub


Sub StartLoadLinked()
'Procedure to load the local file from within the network, rather than loading the file to the server
Dim strFileName
Dim File, filePath,strFileSeqNum, strFileDateTime, lngFileLoadID, intRecord, strFileType

	strFileSeqNum = GetSystemAdmin("ERPCostCentres")
	
	strFileDateTime = Day(Now()) & Month(Now()) & Year(Now())
	
	''Set the File Name as a fixed name as no file is used
	strFileName = "CMSLinkedServer_" & strFileDateTime
	
	intRecord = 0
	lngFileLoadID = SaveFileLoadID ("ERPCostCentres",strFileName,"",0,0,0,0,0,0,0,intRecord,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
	'lngFileLoadID = SaveFileLoadID ("ROMANCostCentres","","",y,0,0,0,0,0,0,intRecord,"","","Imported",Session("UserID"),"N")

	If Request.QueryString("Action")="SaveFileLinked" Then

		'Call the procedure to read the file
		  	With objCmd 
  	 
                .CommandType = 4
                .CommandText = "spCAPSProMasterCostCentreLinkedImport"

				.Parameters.Append objCmd.CreateParameter("FileID", adInteger)
				.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
				.Parameters.Append objCmd.CreateParameter("ProMasterCCIDOutput", adInteger, adParamOutput)
		
				.Parameters("FileID") = lngFileLoadID
				.Parameters("UpdatedBy") = Session("UserID")
                           
               .ActiveConnection = objCon
                
          End With
                
            objCmd.Execute        
            
            'Return the result of the Save Function.
     		intRecord = objCmd.Parameters.Item("ProMasterCCIDOutput")   
			'intRecord = 1234
			
		'''Notify the User that the Load has finished
		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-info alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>" & intRecord & " Cost Centre Records Imported..... strFileSeqNum=" & strFileSeqNum & ", strFileName=" & strFileName & ", lngFileLoadID=" & lngFileLoadID & "</span></div></div></div>"
			
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
		Call UpdateFileLoadSummary ("ERPCostCentres",strFileSeqNum, strFileName,lngFileLoadID)
		
		
		'''Notify the User that the Load has finished
		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
			"<span aria-hidden=""true"">&times;</span></button>" & _
			"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
			"<span>ERP Cost Centre File Seq Num """ & strFileSeqNum & """ load COMPLETE!</span></div></div></div>"
			
	End if

End Sub



'Function for validating the excel file values
Sub ReadExcel 

Dim errors
Dim strFileSeqNum
Dim strEID

    'First loop to check all the values
    Do until objRS.EOF 
        'Read each recors present in the Excel file and check for the validation
 
                 
        strFileSeqNum = objRS("FileSeqNum") 'Should be integer and Not Null values       
        If IsNull(strFileSeqNum) Then
            errors = errors & "Error in line no. " & lineNo & ": FileSeqNum should NOT be Null value <br>" 
        End if
        
		strEID = objRS("EIDNo") 'Should be integer and Not Null values       
        If IsNull(strEID) Then
            errors = errors & "Error in line no. " & lineNo & ": Email Address should NOT be Null value <br>" 
        End if
       
	   strSurname = objRS("Surname") 'Should be integer and Not Null values       
        If IsNull(strSurname) Then
            errors = errors & "Error in line no. " & lineNo & ": Surname/LastName should NOT be Null value <br>" 
        End if
                          
            lineNo = lineNo + 1            
            
        objRS.movenext 
    
        If IsNull(objRS("EIDNo")) AND IsNull(objRS("EIDNo")) Then       
        
            Exit Sub

        End If
    
    Loop        
    
End Sub	


Sub ReadText(strFileNamePath,strFileName)

Const ForReading = 1
Dim strLine
Dim strCardType
Dim strRow
Dim x, y

Dim strCostCentreID
Dim strRecType
Dim strCostCentre
Dim strCostCentreName
Dim strCompanyCode

Dim strFooterCount
Dim lngFileLoadID
Dim objFSO
Dim objStartFolder
Dim objFolder
Dim strFileNameDefault
Dim filePath
Dim colFiles
Dim objFile
Dim objTextFile
Dim strFileDateTime
Dim strFileSeqNum
Dim lngFileID

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass

Set objNetwork = CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'''---Start the New Service Account Login section
If Request.QueryString("ActionType")="Service" Then

'Get the System Parameter for the start of the Training File Location
strServer = GetSystemAdmin("GDriveFilePath")

'Get the System Parameter for the Service Account UserName and Password
strUser = GetSystemAdmin("CAPSServiceAccountName")
strPass = GetSystemAdmin("CAPSServiceAccountPassword")

objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

	'objStartFolder = strServer
	
End If
'''---End the New Service Account Login section

	
	'''---Start the New Service Account Login section
	If Request.QueryString("ActionType")="Service" Then
		objStartFolder = strServer
		'objStartFolder = strServer & "ROMAN\"
	Else			
		'Get the ROMAN Data File Starting folder from the System Parameters
		objStartFolder = GetSystemAdmin("ServerFilePath")
		
		objStartFolder = objStartFolder & "\Admin\CAPSAdmin\Attachments\ROMAN\"
		
	End If
	
	'If there is no System Parameter for the ROMAN Data starting folder then set to the VBMRSN05 server
	'If IsNull(objStartFolder) Or objStartFolder = "" then
	
		'objStartFolder = "D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\ROMAN\"
	'	objStartFolder = objStartFolder & "\Admin\CAPSAdmin\Attachments\ROMAN\"
	'Else
	'	objStartFolder = objStartFolder & "\Admin\CAPSAdmin\Attachments\ROMAN\"
	'End If

	Set objFolder = objFSO.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files

	'Get the System Parameter for the start of the ROMAN fileName
	strFileNameDefault = GetSystemAdmin("ROMANFileName")
	
	If IsNull(strFileNameDefault) or strFileNameDefault = "" Then

		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
				"<span aria-hidden=""true"">&times;</span></button>" & _
				"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
				"<span>NOT LOADED! There is no System Parameter for COA File Names (""ROMANFileName""). See System Admin.</span></div></div></div>"
			Exit Sub
	End If
	
	For Each objFile in colFiles

		If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) Then
			strFileName = objFile.Name
			filePath = objStartFolder & "\" & strFileName
			
			'Response.Write objFile.Name & "</br>"
		End If
		
	Next

	If IsNull(filePath) or filePath = "" Then

		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
				"<span aria-hidden=""true"">&times;</span></button>" & _
				"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
				"<span>" & strFileNameDefault & " NOT LOADED! There is no COA File on the Server Folder to Load! Copy the 'ZCMS_COA__YYYYMMDD.txt' file to the location: " & objStartFolder & "</span></div></div></div>"
			Exit Sub
	End If


Set objTextFile = objFSO.OpenTextFile (filePath, ForReading)
'Set objTextFile = objFSO.OpenTextFile ("c:\mytextfile.txt", ForReading)

x = 0 
'Set the variable which normally gets the checkbox value to overwrite, to "on" so previous loads are always deleted before loading the current file
strDeleteCheck = "on"
 
	Do Until objTextFile.AtEndOfStream
		
		'Count the rows for use in line counts, summary and for getting header
		x = x + 1
					
		strLine = objTextFile.Readline
		
		'The first row of the CS file has a header with FileDateTime and  FileSequenceNumber
		If x = 1 Then

			'The fileDate and Number are only in the header row
			strFileDateTime = Mid(strLine,3,8)
			'strFileSeqNum = Mid(strLine,11,12)
		
			'strFileSeqNum = GetLastFileLoadID("ROMANCostCentres",strFileName)
			'strFileSeqNum = strFileSeqNum + 1
			strFileSeqNum = GetSystemAdmin("ROMANCostCentres")
			'strFileSeqNum = GetFileLoadID("ROMAN",strFileSeqNum,strFileName)
			'''Set the File Number to 1 as the file will always be replaced by what is being loaded.
			'''Also avoid the check and always delete
			
			'Check to see if the same FileSeqNum for the same FileType has already been loaded
		
			'lngFileID = GetLastFileLoadID("ROMANCostCentres",strFileName)
			lngFileID = GetSystemAdmin("ROMANCostCentres")
			
			'lngFileID = lngFileID + 1
			'Check to see if the same FileSeqNum for the same FileType has already been loaded
			'If strFileSeqNum = "" Then
				
			'Else
				'If the checkbox to overwrite is checked then load the data, otherwise do not load
				If strDeleteCheck = "on" Then
					'Delete any existing CS From Diners Records
					objCon.Execute "DELETE FROM tblCAPSCostCentre"
				Else
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! The COA File Seq Num """ & strFileSeqNum & """ has already been loaded! <a href=""UploadROMAN.asp?Reload=1&FileSeqNum=" & strFileSeqNum & """ Style=""font-weight:bold; color:white;""> Click here to Load anyway and overwrite the existing COA file...</a></span></div></div></div>"
					Exit Sub
				End If
			'End If
			
		Else
			
			'If the first character is a T then it is the final row
			If Mid(strLine, 1, 1) = "F" Then
				strFooterCount = Mid(strLine, 2, 6)
			Else
		  
				strCostCentreID = 0
				strRecType = Mid(strLine, 1, 2) 
				strCostCentre = Mid(strLine, 3, 24) 
				strCostCentreName = Mid(strLine, 27, 50) 
				strCompanyCode = Mid(strLine, 77, 25) 
				
				'response.write strCostCentreID & "," & strRecType & "," & strCostCentre & "," & strCostCentreName & "," & strCompanyCode
				
				If IsNull(strRecType) or strRecType = "" Then strRecType = "0"
				
				'Only load the rec types = "03" as the rest of the file (100K records) are not used in CAPS
				If strRecType = "03" Then
					
					y = y + 1
					
					SaveRecord strCostCentreID,cstr(strRecType),cstr(strCostCentre),cstr(strCostCentreName),strCompanyCode,lngFileID, y
					
					'response.write strCostCentreID & "," & strRecType & "," & strCostCentre & "," & strCostCentreName & "," & strCompanyCode
			  
					'outPut.WriteLine(id_no & "_" & strEID);
		
				End If
				
			End If
		End If
	Loop


	If x > 1 Then
		'the CS FRom Diners contains a header and footer row, so remove then from the count
		If x > 2 Then x = x -2
		
		'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
		lngFileLoadID = SaveFileLoadID ("ROMANCostCentres",strFileName,filePath,y,0,0,0,0,0,0,strFooterCount,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
		'response.write "fie=" & lngFileLoadID
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
		Call UpdateFileLoadSummary ("ROMANCostCentres",strFileSeqNum, strFileName,lngFileLoadID)
		'response.write "UpdateFileLoadSummary (""CSFRomDiners""," & strFileSeqNum & "," & lngFileLoadID & ")"
		
		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>COA Cost Centre File Seq Num """ & strFileSeqNum & """ load COMPLETE!</span></div></div></div>"
		
		'Move files to Loaded folder
		Set objTextFile = Nothing
		
		'''---If the file is loaded from the G Drive then do not move it after loading, otherwise from the server move it to the loaded folder
		If Request.QueryString("ActionType")="Service" Then
		
		Else
			objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\ROMAN\"
			strFileNamePath = objStartFolder & strFileName
			
			objStartFolder = objStartFolder & "Loaded\"
			
			'Delete any files with the same name before moving
			Call DeleteExistingFile(objStartFolder,strFileName)
			
			objFSO.MoveFile strFileNamePath,objStartFolder & strFileName	
			'objFSO.MoveFile strFileNamePath,objStartFolder & "Loaded\" & strFileName			
		End If			
	End If
	
'outPut.Close();

'''---Start the New Service Account Login section
If Request.QueryString("ActionType")="Service" Then

	objNetwork.RemoveNetworkDrive strServer, True, False

	Set objNetwork = Nothing
End If
	
Set objFolder = Nothing
Set colFiles = Nothing
Set objFSO = Nothing
'Set outPut = Nothing

End Sub

'Function for validating the excel file values
Sub UploadExcel(chkDelete)

Dim x

Dim strCSFromDinersID
Dim strFileDateTime
Dim strFileSeqNum
Dim strEIDNo
Dim strCardNo
Dim strCardUpdateInd
Dim strCardExpiryDate
Dim strCardStatus
Dim strTitle
Dim strSurname
Dim strGivenNames
Dim strNameOnCard
Dim strAddress1
Dim strAddress2
Dim strAddress3
Dim strSuburb
Dim strState
Dim strPostCode
Dim strHomePhone
Dim strWorkPhone
Dim strMobilePhone
Dim strEmail
Dim strReportGroup
Dim strCreditLimit
Dim strRelationship
Dim strCat2
Dim strAccountNumber
Dim strActivationFlag
Dim strPlasticID
Dim strCompanion
Dim strStatus
Dim strNotes
Dim intCardUpdated
Dim strAccountBlockCode1
Dim strAccountBlockCode2
Dim strCardLevelBlockCode
Dim strCardLevelCreditLimit
Dim strCashHoldFlag
Dim strCashAllowFlag
Dim strZeroes

    objRS.MoveFirst()
    
    Do until objRS.EOF 
        
		strCSFromDinersID = 0'objRS("CSFromDinersID") 
		strFileDateTime = objRS("FileDateTime") 
		strFileSeqNum = objRS("FileSeqNum") 
		strEIDNo = objRS("EIDNo") 
		strCardNo = objRS("CardNo") 
		strCardUpdateInd = objRS("CardUpdateInd") 
		strCardExpiryDate = objRS("CardExpiryDate") 
		strCardStatus = objRS("CardStatus") 
		strTitle = objRS("Title") 
		strSurname = objRS("Surname") 
		strGivenNames = objRS("GivenNames") 
		strNameOnCard = objRS("NameOnCard") 
		strAddress1 = objRS("Address1") 
		strAddress2 = objRS("Address2") 
		strAddress3 = objRS("Address3") 
		strSuburb = objRS("Suburb") 
		strState = objRS("State") 
		strPostCode = objRS("PostCode") 
		strHomePhone = objRS("HomePhone") 
		strWorkPhone = objRS("WorkPhone") 
		strMobilePhone = objRS("MobilePhone") 
		strEmail = objRS("Email") 
		strReportGroup = objRS("ReportGroup") 
		strCreditLimit = objRS("CreditLimit") 
		strRelationship = objRS("Relationship") 
		strCat2 = CheckString(objRS("Cat2"))
		strAccountNumber = CheckNumber(objRS("AccountNumber"))
		strActivationFlag = objRS("ActivationFlag") 
		strPlasticID = objRS("PlasticID") 
		strCompanion = objRS("Companion") 
		strStatus = objRS("Status") 
		strNotes = objRS("Notes") 
		intCardUpdated = objRS("CardUpdated") 
		strAccountBlockCode1 = CheckNumber(objRS("AccountBlockCode1"))
		strAccountBlockCode2 = CheckNumber(objRS("AccountBlockCode2"))
		strCardLevelBlockCode = objRS("CardLevelBlockCode") 
		strCardLevelCreditLimit = objRS("CardLevelCreditLimit") 
		strCashHoldFlag = objRS("CashHoldFlag") 
		strCashAllowFlag = objRS("CashAllowFlag") 
		strZeroes = objRS("Zeroes") 


        'Only save the record if the DeleteCheckbox is checked otherwise it will overwrite zero values and clear existing data
        'If Cstr(Trim(chkDelete)) = " " Then
		'	Reponse.Write "Not Loaded"
        'Else
            x = x + 1
            'Save the record
			
			'Response.Write  "exec spGeneralExpensesSave ="& strCSFromDinersID & "," & strFileDateTime & "," & strFileSeqNum & "," & x
			
            SaveRecord strCSFromDinersID,strFileDateTime,strFileSeqNum,strEIDNo,strCardNo,strCardUpdateInd,strCardExpiryDate,strCardStatus,strTitle,strSurname, _
						strGivenNames,strNameOnCard,strAddress1,strAddress2,strAddress3,strSuburb,strState,strPostCode,strHomePhone,strWorkPhone,strMobilePhone,  _
						strEmail,strReportGroup,strCreditLimit,strRelationship,strCat2,strAccountNumber,strActivationFlag,strPlasticID,strCompanion,strStatus,  _
						strNotes,intCardUpdated,strAccountBlockCode1,strAccountBlockCode2,strCardLevelBlockCode,strCardLevelCreditLimit,strCashHoldFlag,strCashAllowFlag,strZeroes, x
            
        
        'End IF
        
     objRS.movenext 
    
    If IsNull(objRS("EIDNo")) Then               
       
       exit Sub
       
    End If
        
    Loop 
    
End Sub	


Sub SaveRecord(strCostCentreID,strRecType,strCostCentre,strCostCentreName,strCompanyCode, lngFileID, x)

Dim intRecord

If IsNumeric(strCompanyCode) Then
Else
	strCompanyCode = 0
End If

  	With objCmd
  	
  	    'If the procedure has akready run then don't create the parameter objects again (more than once)
  	    If x = 1 then
                .CommandType = 4
                .CommandText = "spCAPSROMANSave"
                
				.Parameters.Append objCmd.CreateParameter("CostCentreID", adInteger)
				.Parameters.Append objCmd.CreateParameter("RecType", advarchar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("CostCentre", advarchar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("CostCentreName", advarchar, adParamInput,100)
				.Parameters.Append objCmd.CreateParameter("CompanyCode", advarchar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("FileID", adInteger)
				.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger)
				
				.Parameters.Append objCmd.CreateParameter("CostCentreIDOutput", adInteger, adParamOutput)				
            
        End If
                 
				.Parameters("CostCentreID") = strCostCentreID
				.Parameters("RecType") = strRecType
				.Parameters("CostCentre") = cstr(left(strCostCentre,10))
				.Parameters("CostCentreName") = strCostCentreName
				.Parameters("CompanyCode") = cstr(Left(strCompanyCode,10))
				.Parameters("FileID") = lngFileID
				.Parameters("UpdatedBy") = Session("UserID")
           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute        
            
            'Return the result of the Save Function.
     		intRecord = objCmd.Parameters.Item("CostCentreIDOutput")    
     		                  			     				     		     		
       'response.write  "exec spGeneralExpensesSave =0," & Session("BudgetID") & "," & Session("VersionID") & "," & CostCentreID & ",'GEXP'" & GLCode & "," & BM1 & "," & BM2 & "," & BM3 & "," & BM4 & "," & BM5 & "," & _
       '                     BM6 & "," & BM7 & "," & BM8 & "," & BM9 & "," & BM10 & "," & BM11 & "," & BM12 & "," & OY1 & "," & OY2 & "," & OY3 & ",'" & Comments & "','" & UpdatedBy & "'," & Session("ColumnLock")
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

Public Sub GetMonthNames()
'This is a procedure to get the order of Month names to be used as titles for Month Columns
Dim intFirstMonth

    'set the First Month name to an integer
    'intFirstMonth = Month("21-" & Session("FirstMonth") & "-2012")
	intFirstMonth = Month("21-1-2012")
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

Public Function CheckNumber(strNumber)

	If IsNull(strNumber) or strNumber = "" Then
		CheckNumber = 0
	Else
		CheckNumber = strNumber
	End If

End Function

Public Function CheckString(strString)

	If IsNull(strString) or strString = "" Then
		CheckString = ""
	Else
		CheckString = strString
	End If

End Function


Public Sub DisplayFileSummary()

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile

Set objFSO = CreateObject("Scripting.FileSystemObject")

	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\ROMAN\")
	objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\ROMAN\"

	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be loaded. Click 'Load Server' button to Load COA File.""></i></th></tr>"
	
	intCount = 0
	
	For Each objFile in colFiles

		intCount = intCount + 1
		
		If intCount < 6 Then
			If IsNull(objFile.Name) or objFile.Name = "" Then
				strFile = ""
			Else
				strFile = Left(objFile.Name,10)
			End If
			
			Response.Write "<TR><TD>" & strFile & "...xls</TD></TR>"
		End If
		
	Next
	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"
	 
Set objFSO = Nothing

End Sub


Public Sub DisplayFileSummaryG()

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile
Dim strFileSize

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass
Dim strFileNameDefault

Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get the System Parameter for the start of the Training File Location
strServer = GetSystemAdmin("GDriveFilePath")

'Get the System Parameter for the Service Account UserName and Password
strUser = GetSystemAdmin("CAPSServiceAccountName")
strPass = GetSystemAdmin("CAPSServiceAccountPassword")

'Get the System Parameter for the fileName
strFileNameDefault = GetSystemAdmin("ROMANFileName")
			
objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

	objStartFolder = strServer
	
	'objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Diners\DinersFrom"

	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">G Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the G Drive waiting to be loaded. Click 'Load G' button to Load the COA File from this location.""></i></th></tr>"
	
	intCount = 0
	
	For Each objFile in colFiles
		'only display the files with the same name as the System Parameter
		If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) Then
		
			intCount = intCount + 1
			
			If intCount < 6 Then
				If IsNull(objFile.Name) or objFile.Name = "" Then
					strFile = ""
					strFileSize = 0
				Else
					strFile = Left(objFile.Name,10) & ".." & Right(objFile.Name,4)
					strFileSize = Round(objFile.Size/1024000,2)
				End If
				
				Response.Write "<TR><TD Title=""" & objFile.Name & " -- Size: " & strFileSize & " MB"">" & strFile & "</TD></TR>"
			End If
		End If
		
		
	Next
	 If intCount > 1 Then
		Response.Write "<tr><th style=""text-align:left""><FONT Color=""RED"">WARNING MORE THAN 1 FILE EXISTS IN LOAD FOLDER: " & intCount & "</FONT>&nbsp;<i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""There should only be 1 file in the load folder.  This can cause issues with the loading sequence.""></i></th></tr></table>"
	 Else
		Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr>"
		'Response.Write "<tr><th style=""text-align:left"">Size: " & strFileSize & " MB</th></tr></table>"
		Response.Write "</table>"
	 End If

	 
objNetwork.RemoveNetworkDrive strServer, True, False
	 
Set objFSO = Nothing
Set objNetwork = Nothing

'Set objFSO = Nothing
'Set outPut = Nothing

End Sub



Public Sub DeleteExistingFile(objStartFolder,strFileName)
'Procedure to delete any existing file with the same name and location as passed in

Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile

Set objFSO = CreateObject("Scripting.FileSystemObject")

	Set objFolder = objFSO.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files
		
	For Each objFile in colFiles
		
		If IsNull(objFile.Name) or objFile.Name = "" Then
			strFile = ""
		Else
			strFile = objFile.Name
		End If
		
		If strFileName = strFile Then
			'Delete any files already n the Export folder for the current user
			objFSO.DeleteFile strFileName, True
				
		End If
	Next
	
End Sub


Set objRS = Nothing
Set objCon = Nothing

 %>


