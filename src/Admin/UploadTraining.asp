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
Dim errors

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

If Request.QueryString("Action")="Save" Then

	Call StartLoad()
End If

'If the Process button has been clicked next to a file, then call the Process procedure
If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "Process" Then
		Call ProcessFile(Request.QueryString("FileLoadID"))
	End If
End If

If Not IsEmpty(Request.QueryString("Reload")) Then

	Call StartLoad()
End If

If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
End If

'If the local load has been clicked then call the procedure to load the network file rather than uploading it
If Request.QueryString("Action")="SaveFileLocal" Then

	'New Check to see if the Load if from the Local OR G button
	'If the ActionType is Service then load the Training file from the G Drive location (System Parameter)
	If Request.QueryString("ActionType")="Service" Then

		strDeleteCheck = "true"'Request.QueryString("chkDelete")
		Call StartLoadGDrive()
		
		'*******---- NEW 6th December 2021 ---********
		'Update the CourseID in the Training table to make sure the leading zeroes exist.  Loading from CSV removes the leading zeroes
		objCon.Execute "UPDATE tblCAPSTraining SET CourseID=CAST('000'+CourseID as varchar(8)) FROM tblCAPSTraining WHERE LEN(CourseID)=5 "
		
		Response.Write "Load from G Drive complete. Please check that the number of Total Records loaded (in the file summary list below) is correct"
	'Else load from the local Server file location
	Else
		'Call StartLoadLocal()
		strDeleteCheck = "true"'Request.QueryString("chkDelete")
		Call StartLoadLocal()
		
		'*******---- NEW 6th December 2021 ---********
		'Update the CourseID in the Training table to make sure the leading zeroes exist.  Loading from CSV removes the leading zeroes
		objCon.Execute "UPDATE tblCAPSTraining SET CourseID=CAST('000'+CourseID as varchar(8)) FROM tblCAPSTraining WHERE LEN(CourseID)=5 "
		
	End If
	
End If

'Response.Write "B=" & 	dteBatchDate
'dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
'Response.Write " 2B=" & 	dteBatchDate

'	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
'	Response.Write " 3B=" & 	dteBatchDateFormat
 %>

<html>
<head>

	
	
<script language=javascript>
function upload()
{   
    if(document.getElementById('FILE1').value=="")
    {
        alert("Please select the file to upload (click the Browse button above)");
    }
    else
    {   
        if(getFileExt(document.getElementById('FILE1').value)==".xls" || getFileExt(document.getElementById('FILE1').value)==".txt")
        {
            if(window.confirm('This will overwrite any existing Training file data! \n \n Continue?')==true)
                {
                document.getElementById('Progress').style.display = "inline";
                frm.submit();
            }
        } 
        else
        {
            alert("Please enter a valid Excel(.xls) or TEXT (.txt) file");
        }
    }
}

function UploadLocal()
{
	document.getElementById('Progress').style.display = "inline";
	document.getElementById('LoadButton').disabled=true;
	self.location="UploadTraining.asp?Action=SaveFileLocal"
}

function UploadLocalG()
{
	document.getElementById('Progress').style.display = "inline";
	document.getElementById('LoadButton').disabled=true;
	document.getElementById('LoadButton2').disabled=true;
	self.location="UploadTraining.asp?Action=SaveFileLocal&ActionType=Service"
}

 function getFileExt(filename)
 {
     var s
     s = filename.charAt(filename.length-4) + filename.charAt(filename.length-3) + filename.charAt(filename.length-2)+ filename.charAt(filename.length-1);
     return s;
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
	self.location="UploadTraining.asp?FileDate=" + document.getElementById("CSDate").value;
}
</script>

<style>

    table.newd th, table.newd td{

        padding: 4px; 

    }

</style>
</head>
<body>

<!-- Modal -->
	<div class="loader" id="ModApp">
        <div class="wrap">
            <div class="spinner"></div>
            <span class="loading-message">Loading...</h6>
        </div>
    </div>
	
<main class="main py-3">
      <div class="container">
<form action="UploadTraining.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	 <div class="row" id="basic-table">
  <div class="col-4">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../CC/img/DFG_Logo.png" height="40px" width="100px" title="Training File from CAMPUS"> Training File Load</h4>
        </div>
      <div class="card-content">
        <div class="card-body">
		
<!--
<div class="col-lg-12 col-md-12">
<fieldset class="form-group">
	<label for="basicInputFile">Select a File to Load</label>
	<div class="custom-file">
		<input type="file" class="custom-file-input" NAME="FILE1" id="FILE1">
		<label class="custom-file-label" for="FILE1">Choose file</label>
	</div>
</fieldset>
</div>-->

<div class="form-body">
<div class="row col-12">
<!--<div class="col-auto mr-auto">
	<div class="form-group">
	   <div class="checkbox">
		<input type="checkbox" class="checkbox-input" id="chkDelete" name="chkDelete">
		<label for="chkDelete">Overwrite Existing Batch</label>
	  </div>
	</div>
  </div>-->
   <div class="col-auto text-right">
	<button type="button" name="LoadButton2" id="LoadButton2" class="btn btn-primary btn" onclick="UploadLocalG();" Title="Click to Load any existing file in the Training Folder on the G Drive (G Files to be loaded list to the right)-->"><i class="fa fa-upload"></i> Load G</button>
	&nbsp;&nbsp;
	<button type="button" name="LoadButton" id="LoadButton" class="btn btn-outline-secondary btn" onclick="UploadLocal();" Title="ONLY USE if Load G button does not work! Click to Load any existing file in the Training Folder on the Server (Files to be loaded list to the right)-->"><i class="fa fa-upload"></i> Load Local</button>
  </div>
  <!--<div class="col-auto">
	<button type="button" class="btn btn-primary btn-xs" onclick="upload();"><i class="fa fa-upload"></i> Upload</button>
  </div>-->
</div>
</div>
<p class="text-left">
<span id="Progress" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>
</br></br>
<font color="red" size="2"><b>NOTE: </Font><font color="black" size="2">When loading from Excel the Worksheet MUST be named 'Sheet1' (no spaces)
            <BR>* The file must be '.csv' only
			<BR>* The file must be name 'DTCTraining.csv' exactly
            <BR>* Do not change the first row headers</B></Font>
</p>
<!--<div class="col-lg-12 col-md-12">
<fieldset class="form-group">
    <button type="button" class="btn btn-outline-primary btn-xs" onclick="window.open('TrainingTemplateExcel.asp')"><i class="fa fa-file"></i> Training Template Excel </button>
</fieldset>
</div>-->

 <script>
	//Script to display the file selected in the File Load Input field (it doesn't by default)
	$('#FILE1').on('change',function(){
		//get the file name
		var fileName = $(this).val();
		//replace the "Choose a file" label
		$(this).next('.custom-file-label').html(fileName);
	})
</script>

		</div>
	  </div>
    </div>
   </div>
  

  <div class="col-6">
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


<!-- #Include file=../CC/CAPSFooter.asp -->
</body>
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

objRS.Open "SELECT TOP 50 * FROM qryCAPSTraining WITH(NOLOCK) "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no Training data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Upload Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing Training Data already in CAPS</th></tr>" & _
		        "<th Style=""width:20px;""Training ID</th>" & _
				"<th>Employee ID</th>" & _
				"<th>Course ID</th><th>Offering ID</th>" & _
	 	        "<th>First Name</th>" & _	
	 	        "<th>Last Name</th>" & _
		        "<th>Completion Date</th><th>Loaded</th>" & _
                "<th>File ID</th><th>Date Updated</th><th>Updated By Name</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR class='clickable-row' data-href='UploadTrainingDetail.asp?BatchNo=" & objRS("FileID") & "&EIDNo=" & objRS("EmployeeID") & "' style=""cursor: pointer;""><TD>" & objRS("TrainingID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("EmployeeID") & "</TD><TD style=""text-align:center"">" & objRS("CourseID") & "</TD><TD style=""text-align:center"">" & objRS("OfferingID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("FirstName") & "</TD><TD style=""text-align:center"">" & objRS("LastName") & "</TD><TD style=""text-align:center"">" & objRS("CompletionDate") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("Loaded") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("FileID") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD><TD style=""text-align:center"">" & objRS("UpdatedByName") & "</TD>" & _
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
Dim dteDateLoaded
Dim strAction
Dim strView
Dim strStatus

Dim strDateLoadColour
Dim strDateTitle
Dim strDateLoadedTitle
Dim lngRecordCount

Dim strRecordCount
Dim strRecordCountMessage
Dim strRecordCountColour


	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		objRS.Open "SELECT TOP 10 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE DateLoaded > '" & dteBatchDate & "' AND FileType = 'Training' ORDER BY DateLoaded DESC",objCon
		'objRS.Open "SELECT TOP 10 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'Training' ORDER BY DateLoaded DESC",objCon
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        'Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                '"<tr><th colspan=""11"" style=""text-align:left"">There is no Training data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
				Response.Write "<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th>No Training Data Loaded</th></tr>" 
				
		    Else
		    
		         Response.Write "<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th>" & _
				"<th>View</th></tr>"
				
				 'Response.Write "<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                '"<tr><th colspan=""11"" style=""text-align:left"">Training Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
		        '"<th Style=""width:20px;"">Batch No.</th>" & _
				'"<th>Total Records</th>" & _
				'"<th>Status</th>" & _
	 	        '"<th>Date Loaded</th>" & _
				'"<th>View</th></tr>"
				
		    End If
		    
		    Do until objRS.eof
				
				If IsNull(objRS("Status")) Then
					strStatus = ""
				Else
					strStatus = objRS("Status")
				End If
				
				If strStatus = "Imported" Then
					strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='UploadTraining.asp?Action=Process&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to Process the CS File and update changes to cards in CAPS from the CS File loaded " & objRS("DateLoaded") & """>Process</button>"
					'strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='UploadANZ.asp?Action=Process&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to Process the ANZ File and update changes to cards in CAPS from the ANZ File loaded " & objRS("DateLoaded") & """><i class=""fa fa-cogs""></i> Process</button>"
				Else
					strAction = ""
				End If
				
				'If IsNull(objRS("DateLoaded")) Then
				'	dteDateLoaded = ""
				'Else
				'	dteDateLoaded = FormatDateTime(objRS("DateLoaded"),vbShortDate)
				'End If
				
				If DateDiff("d",objRS("DateLoaded"),Now()) = 0 Then
					strDateLoadColour = " Color:Green; font-weight:bold;"
					strDateTitle = " - Training File for today has been LOADED "
				Else
					strDateLoadColour = ""
					strDateTitle = ""
				End If
				
				'Get the Date Loaded and Format for display with the full date as ther title
				If IsNull(objRS("DateLoaded")) Then
					dteDateLoaded = ""
				Else
					dteDateLoaded = FormatDateTime(objRS("DateLoaded"),vbShortDate)
					strDateLoadedTitle = "title=""" & objRS("DateLoaded") & " " & strDateTitle & " """
				End If

				'Get the Total records and format
				If IsNull(objRS("RecordCount")) or objRS("RecordCount") = "" Then
					strRecordCount = 0
				Else
					If IsNumeric(objRS("RecordCount")) Then
						strRecordCount = FormatNumber(objRS("RecordCount"),0)
						'check to make sure that the number of records loaded is about normal (above previous numbers) as loading an incomplete file will cause processing issues and cancel cards
						'If red text is displayed then the file should be checked and loaded again if it is not the full file count in CAPS (CAPS and CSV do not reconcile)
						If strRecordCount < 100000 Then
							strRecordCountMessage = "title=""The Number of records loaded looks to be below normal. Please SEE SYSTEM ADMINISTRATORS before continuing!"""
							strRecordCountColour = "color:red; font-weight:bold;"
						Else
							strRecordCountMessage = ""
							strRecordCountColour = ""
						End If
					Else
						strRecordCount = 0
					End If
				End If


				'Create the View button detail
				strView = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='TrainingTransactions.asp?FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to view details of the Training File loaded " & objRS("DateLoaded") & """>View</button>"
				
				
				Response.Write "<TR><TD>" & objRS("FileSeqNum") & "</TD>" & _
							"<TD style=""text-align:center; " & strRecordCountColour & """ " & strRecordCountMessage & ">" & strRecordCount & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("Status") & "</TD><TD style=""text-align:center; " & strDateLoadColour & """ " & strDateLoadedTitle & ">" & dteDateLoaded & "</TD><TD style=""text-align:center"">" & strView & "</TD></TR>"
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
			If Right(filePath,3) = "csv" Then
				
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
					Response.Write "<b>CS From Diners File Successfully Uploaded!!<b><br>"
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
Dim objFSO
Dim objStartFolder
Dim strFileNameDefault
Dim objFolder
Dim colFiles
Dim objFile
Dim lngFileID
Dim objExcelCon
Dim intRecordsLoaded
Dim strFileDateTime
Dim lngFileLoadID
Dim strFileType
Dim x

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass

Set objNetwork = CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")

'''---Start the New Service Account Login section
If Request.QueryString("ActionType")="Service" Then

'Get the System Parameter for the start of the Training File Location
strServer = GetSystemAdmin("TrainingFilePath")

'Get the System Parameter for the Service Account UserName and Password
strUser = GetSystemAdmin("CAPSServiceAccountName")
strPass = GetSystemAdmin("CAPSServiceAccountPassword")

objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

	'objStartFolder = strServer
	
End If
'''---End the New Service Account Login section


	If Request.QueryString("Action")="SaveFileLocal" Then	

		'filePath = "D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\Attachments\deptofdefence07052021_3.xls"
		'strFileName ="deptofdefence07052021_3.xls"
		
		'Set objFSO = CreateObject("Scripting.FileSystemObject")

			'''---Start the New Service Account Login section
			If Request.QueryString("ActionType")="Service" Then
				objStartFolder = strServer
			Else
				'Get the Training Data File Starting folder from the System Parameters
				objStartFolder = GetSystemAdmin("ServerFilePath")
							
				'If there is no System Parameter for the Training Data starting folder then set to the VBMRSN05 server
				If IsNull(objStartFolder) Or objStartFolder = "" Then			
					objStartFolder = objStartFolder & "\Admin\CAPSAdmin\Attachments\Training\"				
				Else
					objStartFolder = objStartFolder & "\Admin\CAPSAdmin\Attachments\Training\"
				End If
			
			End If
			
			Set objFolder = objFSO.GetFolder(objStartFolder)
			Set colFiles = objFolder.Files

			'Get the System Parameter for the start of the Training fileName
			strFileNameDefault = GetSystemAdmin("TrainingFileName")
			
			If IsNull(strFileNameDefault) or strFileNameDefault = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! There is no System Parameter for Training File Names (""TrainingFileName""). See System Admin.</span></div></div></div>"
					Exit Sub
			End If
			
			For Each objFile in colFiles

				If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) Then
					strFileName = objFile.Name
					filePath = objStartFolder & "" & strFileName
				
					'Response.Write objFile.Name & "</br>"
				End If
				
			Next
			
			If IsNull(strFileName) or strFileName = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>" & strFileNameDefault & " NOT LOADED! There is no Training File in the Server Folder to Load! Copy the 'Defence Travel Card Course Completions_DDMMMYYYY.xls' file to the location " & objStartFolder & "</span></div></div></div>"
					Exit Sub
			End If

		'<!--------- start the load procedure, which shares similar code as the upload procedure above
		'Could be re-written to reduce duplicated code --->
		
		'Check to see if the same FileSeqNum for the same FileType has already been loaded
			lngFileID = GetFileLoadID("Training","",strFileName)

			If lngFileID = "" Then
				'Get the next fileID Number for the ANZCardlist File
				lngFileID = GetNextTrainingFileID
			Else

				'If the checkbox to overwrite is checked then load the data, otherwise do not load
				If strDeleteCheck = "true" Then
					'Delete any existing CS From Diners Records
					objCon.Execute "DELETE FROM tblCAPSTraining WHERE [FileID] = " & lngFileID & ""
				Else
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! The Training File """ & strFileName & """ has already been loaded! <a href=""UploadTraining.asp?Reload=1&FileSeqNum=" & lngFileID & """ Style=""font-weight:bold; color:white;""> Check the 'Overwrite Existing Batch' box and load again to overwrite...</a></span></div></div></div>"
					Exit Sub
				End If
			End If
		
			'Call the relevant procedure depending on whether the file is .xlsx (or .xls -- not enabled)
			If Right(filePath,3) = "csv" Then
				
				'After uploading, Read excel file
				Set objExcelCon = Server.CreateObject("ADODB.connection")     
				'objExcelCon.Open "DBQ=" & filePath & "; DRIVER={Microsoft Excel Driver (*.xls)};" 
				objExcelCon.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & objStartFolder & ";Extensions=asc,csv,tab,txt;ColNameHeader=Yes;"
				
				'objExcelCon.Open "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & filePath & ";DefaultDir=c:\Apps\CAPS2\ASP2\Admin\CAPSAdmin\Attachments;" 
				
				'OLEDB Driver not on VBMRSN05 Server---*****
				'objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=YES"";"
				
				'objExcelCon.Open strUploadPath & "\CAPSCSFromDiners.dsn"
				
				'Delete all of the existing records from the Training Table before loading the new dataset
				objCon.Execute "TRUNCATE TABLE tblCAPSTraining"
				
				'Write the SQL Query 
				objRS.open "SELECT * FROM [" & strFileName & "]", objExcelCon
				
				'Call the procedure to check the Excel fields and data are correct
				Call ReadExcel
		   
				'Check for errors
				If(errors<>"") Then
					'Print the errors and return
					Response.Write "<font face=arial size=1><b>File not uploaded. Please correct the followings errors and load again</b> <br> "& errors &"</font><br>"         
				Else  
					strFileDateTime = Right(strFileName,14)
					strFileDateTime = Left(strFileName,9)
					
					'If the no errors found in the ReadExcel method, then start uploading the records in the database       	    	    
					'UploadExcel strDeleteCheck,lngFileID, strFileName,filePath,strFileDateTime
					intRecordsLoaded = UploadExcel(strDeleteCheck,lngFileID)
					'UploadExcel Uploader.Form("chkDelete"),lngFileID, strFileName,filePath,strFileDateTime
					
					'Response.Write "<b>ANZ Cardlist File Successfully Uploaded!!<b><br>"
				End if
				
				
				
				'Close the recordset/connection 
				objRS.Close 
				objExcelCon.Close 
		    
			Else
				'Training file is xls converted from XLSX for the VBMRSN05 server, so do not enable text load
				'ReadText(filePath)
			End If
			
			'<!--------- End of duplicated code --->
		
		If intRecordsLoaded > 1 Then

			'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
			lngFileLoadID = SaveFileLoadID ("Training",strFileName, filePath,intRecordsLoaded,0,0,0,0,0,0,0,strFileDateTime,lngFileID,"Imported",Session("UserID"),"N")
			
			'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			Call UpdateFileLoadSummary ("Training",lngFileID, strFileName, lngFileLoadID)
			'response.write "UpdateFileLoadSummary (""ANZCardlist""," & lngFileID & "," & strFileName & "," & lngFileLoadID & ")"
			
			'Move the file to the loaded folder (run through up to 10 numbers at the end of the file name if the same file name exists in the loaded folder)
			'First set the name of the Loaded folder
			strFileType = "Loaded"
			'response.write strFileName
			If objFSO.FileExists(objStartFolder & "\" & strFileType & "\" & strFileName) Then
				For x = 1 to 10
					If objFSO.FileExists(objStartFolder & "\" & strFileType & "\" & x & strFileName) Then
					Else
						'Move the file to the Loaded folder
						objFSO.MoveFile objStartFolder & "\" & strFileName,objStartFolder & "\" & strFileType & "\" & x & Left(strFileName,Len(strFileName)-4) & Year(now()) & PadDigits(Month(now()),2) & Day(now()) & Hour(now()) & Minute(now()) & Right(strFileName,4)
						
						x = 10
					End If
				Next
			Else
				'Move the file to the Loaded folder
				objFSO.MoveFile objStartFolder & "\" & strFileName,objStartFolder & "\" & strFileType & "\" & Left(strFileName,Len(strFileName)-4) & Year(now()) & PadDigits(Month(now()),2) & Day(now()) & Hour(now()) & Minute(now()) & Right(strFileName,4)
				'strFileName = objStartFolder & "\" & objFile.Name
			End If
	
	
			Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
							"<span aria-hidden=""true"">&times;</span></button>" & _
							"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
							"<span>Training File """ & strFileName & """ load COMPLETE! " & intRecordsLoaded & " records loaded.</span></div></div></div>"
						
		End If
	
	End if

End Sub


'Function for validating the excel file values
Sub ReadExcel 

'Dim errors
Dim strCourseID
Dim strEID
Dim strCompletionDate
Dim lineNo

    'First loop to check all the values
    Do until objRS.EOF 
        'Read each records present in the Excel file and check for the validation
    
        strCourseID = objRS("Course ID") 'Should be string and Not Null values       
        If IsNull(strCourseID) Then
            errors = errors & "Error in line no. " & lineNo & ": Course ID should NOT be Null value <br>" 
        End if
        
		strEID = objRS("PMKeyS/ODS ID") 'Should be a string and Not Null values       
        If IsNull(strEID) Then
            errors = errors & "Error in line no. " & lineNo & ": Employee ID should NOT be Null value <br>" 
        End if
       
	   strCompletionDate = objRS("Completion Date") 'Should be a date/string and Not Null values       
        If IsNull(strCompletionDate) Then
            errors = errors & "Error in line no. " & lineNo & ": Completion Date should NOT be Null value <br>" 
        End if
                          
            lineNo = lineNo + 1            
            
        objRS.movenext 
    
        If IsNull(objRS("PMKeyS/ODS ID")) AND IsNull(objRS("PMKeyS/ODS ID")) Then       
        
            Exit Sub

        End If
    
    Loop        
    
End Sub	



'Function for validating the excel file values
Function UploadExcel(chkDelete,lngFileID)

Dim x
Dim lngTrainingID
Dim strCourseID
Dim strOfferingID
Dim strCourseTitle
Dim strPMKeySID
Dim strFirstName
Dim strLastName
Dim strEmail
Dim strCompletionDate

    objRS.MoveFirst()
    
    Do until objRS.EOF 
        
		lngTrainingID = 0 
		strCourseID = objRS("Course ID") 
		strOfferingID = objRS("Offering ID") 
		strCourseTitle = objRS("Course Title") 
		strPMKeySID = cstr(objRS("PMKeyS/ODS ID")) 
		strFirstName = objRS("First Name") 
		strLastName = objRS("Last Name") 
		strEmail = objRS("Email") 
		strCompletionDate = objRS("Completion Date") 
		
        'Only save the record if the DeleteCheckbox is checked otherwise it will overwrite zero values and clear existing data
        'If Cstr(Trim(chkDelete)) = " " Then
		'	Reponse.Write "Not Loaded"
        'Else
            x = x + 1
            'Save the record
			
			'Response.Write  "exec spGeneralExpensesSave ="& strCSFromDinersID & "," & strFileDateTime & "," & strFileSeqNum & "," & x
			
            SaveRecord lngTrainingID,strCourseID,strOfferingID,strCourseTitle,strPMKeySID,strFirstName,strLastName,strEmail,strCompletionDate,lngFileID, x
            
        
        'End IF
        
		 objRS.movenext 
		
		'Set the Function to the number of records loaded to return
		UploadExcel = x
		
		If IsNull(objRS("PMKeyS/ODS ID")) Then               
		   
		   Exit Function
		   
		End If
        
    Loop 
    
End Function	


Sub SaveRecord(lngTrainingID,strCourseID,strOfferingID,strCourseTitle,strPMKeySID,strFirstName,strLastName,strEmail,strCompletionDate,lngFileID, x)

Dim intRecord

  	With objCmd
  	
  	    'If the procedure has already run then don't create the parameter objects again (more than once)
  	    If x = 1 then
                .CommandType = 4
                .CommandText = "spCAPSTrainingSave"
                
				.Parameters.Append objCmd.CreateParameter("TrainingID", adInteger)
				.Parameters.Append objCmd.CreateParameter("CourseID", adVarChar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("OfferingID", adVarChar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("CourseTitle", adVarChar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarChar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("FirstName", adVarChar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("LastName", adVarChar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("Email", adVarChar, adParamInput,100)
				.Parameters.Append objCmd.CreateParameter("CompletionDate", adDate, adParamInput)
				.Parameters.Append objCmd.CreateParameter("Loaded", adChar, adParamInput,1)
				.Parameters.Append objCmd.CreateParameter("FileID", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)  
				.Parameters.Append objCmd.CreateParameter("TrainingIDDOutput", adInteger, adParamOutput)				
            
        End If
                 
				.Parameters("TrainingID") = lngTrainingID
				.Parameters("CourseID") = strCourseID
				.Parameters("OfferingID") = strOfferingID
				.Parameters("CourseTitle") = strCourseTitle
				.Parameters("EmployeeID") = strPMKeySID
				.Parameters("FirstName") = strFirstName
				.Parameters("LastName") = strLastName
				.Parameters("Email") = strEmail
				.Parameters("CompletionDate") = strCompletionDate
				.Parameters("Loaded") = "Y"
				.Parameters("FileID") = lngFileID
				.Parameters("UpdatedBy") = Session("UserID")
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute        
            
            'Return the result of the Save Function.
     		intRecord = objCmd.Parameters.Item("TrainingIDDOutput")    
     		                  			     				     		     		
       'response.write  "exec spGeneralExpensesSave =0," & Session("BudgetID") & "," & Session("VersionID") & "," & CostCentreID & ",'GEXP'" & GLCode & "," & BM1 & "," & BM2 & "," & BM3 & "," & BM4 & "," & BM5 & "," & _
       '                     BM6 & "," & BM7 & "," & BM8 & "," & BM9 & "," & BM10 & "," & BM11 & "," & BM12 & "," & OY1 & "," & OY2 & "," & OY3 & ",'" & Comments & "','" & UpdatedBy & "'," & Session("ColumnLock")
End Sub

Public Function ProcessFile(strFileID)
'Function to Process an ANZ File which has been loaded into the database (update all changes from the ANZ file and add changes to the audit log)
Dim intRecord

	With objCmd

		.CommandType = 4
		.CommandText = "spCAPSDinersProcessCSFile"

		.Parameters.Append objCmd.CreateParameter("UserID", adInteger)
		.Parameters.Append objCmd.CreateParameter("FileID", advarchar,adParamInput,6)
		.Parameters.Append objCmd.CreateParameter("CAPSDinersProcessCSFileOutput", adInteger, adParamOutput)
		
		.Parameters("UserID") = Session("UserID")
		.Parameters("FileID") = strFileID
		
		.ActiveConnection = objCon
		 
	End With
   
	objCmd.Execute        
  
	'Return the result of the Save Function.
	intRecord = objCmd.Parameters.Item("CAPSDinersProcessCSFileOutput") 

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Records Processed in CS Upload File " & strFileID & ". Please notify System Administrators.</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert""> " & intRecord & " CS Upload File records Processed in file " & strFileID & "</div>"
	End If
	
	ProcessFile = intRecord
	
End Function

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


Public Function GetNextTrainingFileID()

	'Description:	Gets the next FILE ID number for ANZ Cardlist files.
	objRS.Open "SELECT TOP 1 [FileSeqNum] FROM tblCAPSFileLoad WHERE FileType = 'Training' AND [Deleted] = 'N' ORDER BY [FileLoadID] DESC ",objCon

		If Not objRS.EOF Then
			GetNextTrainingFileID = objRS("FileSeqNum") + 1
		Else
			GetNextTrainingFileID = 1
		End If

	objRS.Close
	
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

	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\Training\")
	objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Training\"

	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be loaded. Click 'Load Local' button to Load Training File.""></i></th></tr>"
	
	intCount = 0
	
	For Each objFile in colFiles

		intCount = intCount + 1
		
		If intCount < 6 Then
			If IsNull(objFile.Name) or objFile.Name = "" Then
				strFile = ""
			Else
				strFile = Left(objFile.Name,10)
			End If
			
			Response.Write "<TR><TD>" & strFile & "...csv</TD></TR>"
		End If
		
	Next
	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"
	 
Set objFSO = Nothing

End Sub


Public Sub DisplayFileSummaryG()

''''This procedure gets details of files on the G Drive using the CAP service account.
''''EXIT this temporarily until functionality is tested.
'Exit Sub 

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile
Dim strFileSize
Dim strFileAttributes

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass

Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get the System Parameter for the start of the Training File Location
strServer = GetSystemAdmin("GDriveFilePath")
'strServer = "\\groupdata.rus.car.drn.defence.mil.au\groupdata\CFO\CFO\CMS Admin\CAPS\Import Files\Training\"

'Get the System Parameter for the Service Account UserName and Password
strUser = GetSystemAdmin("CAPSServiceAccountName")
strPass = GetSystemAdmin("CAPSServiceAccountPassword")
'strUser = "DRN\svc_CAPS_VBMRSN05_Ad"
'strPass = "$dw2zt%2V9D2"

objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

	'Use the Starting folder from System Parameters and add the Training folder
	objStartFolder = strServer & "\Training"

	

	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"" Title=""Load From G Location: " & objStartFolder & """>G Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the G: Drive waiting to be loaded. Click 'Load G' button to Load Training File listed here.""></i></th></tr>"
	
	intCount = 0
	
	For Each objFile in colFiles

	'	intCount = intCount + 1
		
	'	If intCount < 6 Then
	'		If IsNull(objFile.Name) or objFile.Name = "" Then
	'			strFile = ""
	'		Else
	'			strFile = Left(objFile.Name,10)
	'		End If
	'		
	'		Response.Write "<TR><TD>" & strFile & "...csv</TD></TR>"
	'	End If
		
	'Next
	
	
	If Right(objFile.Name,3) = "csv" Then
		
		intCount = intCount + 1
		
			If intCount < 6 Then
				If IsNull(objFile.Name) or objFile.Name = "" Then
					strFile = ""
					strFileSize = 0
				Else
					strFile = Left(objFile.Name,10) & ".." & Right(objFile.Name,4)
					strFileSize = Round(objFile.Size/1024000,2)
					
					strFileAttributes =  "Created: " & objFile.DateCreated
					strFileAttributes = strFileAttributes & " Last Accessed: " & objFile.DateLastAccessed
					strFileAttributes = strFileAttributes & " Last Modified: " & objFile.DateLastModified  
	
				End If
				
				Response.Write "<TR><TD Title=""" & objFile.Name & " " & strFileAttributes & """>" & strFile & "</TD></TR>"
			End If
			
		End If
		
	Next
	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"

objNetwork.RemoveNetworkDrive strServer, True, False
	 
Set objFSO = Nothing
Set objNetwork = Nothing

End Sub


Sub StartLoadGDrive()
'Procedure to load the local file from within the network, rather than loading the file to the server
Dim strFileName
Dim File, filePath
Dim objFSO
Dim objStartFolder
Dim strFileNameDefault
Dim objFolder
Dim colFiles
Dim objFile
Dim lngFileID
Dim objExcelCon
Dim intRecordsLoaded
Dim strFileDateTime
Dim lngFileLoadID
Dim strFileType
Dim x

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


	If Request.QueryString("Action")="SaveFileLocal" Then	

		'filePath = "D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\Attachments\deptofdefence07052021_3.xls"
		'strFileName ="deptofdefence07052021_3.xls"
		
		'Set objFSO = CreateObject("Scripting.FileSystemObject")

			'''---Start the New Service Account Login section
			If Request.QueryString("ActionType")="Service" Then
				objStartFolder = strServer & "\Training\"
				'objStartFolder = strServer & "Training\"

			Else
				'Get the Training Data File Starting folder from the System Parameters
				objStartFolder = GetSystemAdmin("ServerFilePath")
							
				'If there is no System Parameter for the Training Data starting folder then set to the VBMRSN05 server
				If IsNull(objStartFolder) Or objStartFolder = "" Then			
					objStartFolder = objStartFolder & "\Admin\CAPSAdmin\Attachments\Training\"				
				Else
					objStartFolder = objStartFolder & "\Admin\CAPSAdmin\Attachments\Training\"
				End If
			
			End If
			
			Set objFolder = objFSO.GetFolder(objStartFolder)
			Set colFiles = objFolder.Files

			'Get the System Parameter for the start of the Training fileName
			strFileNameDefault = GetSystemAdmin("TrainingFileName")
			
			If IsNull(strFileNameDefault) or strFileNameDefault = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! There is no System Parameter for Training File Names (""TrainingFileName""). See System Admin.</span></div></div></div>"
					Exit Sub
			End If
			
			For Each objFile in colFiles

				'File must be exactly the System Parameter name with no extra after (dates etc..) -- line commented out accepts any start file name the same and system parameters not the whole file name
				'If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) Then
				If Left(objFile.Name,Len(objFile.Name)-4) = Trim(strFileNameDefault) Then
					strFileName = objFile.Name
					filePath = objStartFolder & "" & strFileName
				
					'Response.Write objFile.Name & "</br>"
				End If
				
			Next
			
			If IsNull(strFileName) or strFileName = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>" & strFileNameDefault & " NOT LOADED! There is no Training File in the G Drive Folder to Load! Copy the Training File (" & strFileNameDefault & ".csv) file to the location " & objStartFolder & "</span></div></div></div>"
					Exit Sub
			End If

		'<!--------- start the load procedure, which shares similar code as the upload procedure above
		'Could be re-written to reduce duplicated code --->
		
		'Check to see if the same FileSeqNum for the same FileType has already been loaded
			lngFileID = GetFileLoadID("Training","",strFileName)

			If lngFileID = "" Then
				'Get the next fileID Number for the ANZCardlist File
				lngFileID = GetNextTrainingFileID
				
			Else

				'If the checkbox to overwrite is checked then load the data, otherwise do not load
				If strDeleteCheck = "true" Then
					'Delete any existing CS From Diners Records
					objCon.Execute "DELETE FROM tblCAPSTraining WHERE [FileID] = " & lngFileID & ""
				Else
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! The Training File """ & strFileName & """ has already been loaded! <a href=""UploadTraining.asp?Reload=1&FileSeqNum=" & lngFileID & """ Style=""font-weight:bold; color:white;""> Check the 'Overwrite Existing Batch' box and load again to overwrite...</a></span></div></div></div>"
					Exit Sub
				End If
			End If
		
			'Call the relevant procedure depending on whether the file is .xlsx (or .xls -- not enabled)
			If Right(filePath,3) = "csv" Then
				
				'After uploading, Read excel file
				Set objExcelCon = Server.CreateObject("ADODB.connection")     
				'objExcelCon.Open "DBQ=" & filePath & "; DRIVER={Microsoft Excel Driver (*.xls)};" 
				''''VBMRSN05 Sever ---objExcelCon.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & objStartFolder & ";Extensions=asc,csv,tab,txt;ColNameHeader=Yes;"
				
				'objExcelCon.Open "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & filePath & ";DefaultDir=c:\Apps\CAPS2\ASP2\Admin\CAPSAdmin\Attachments;" 
				
				'OLEDB Driver not on VBMRSN05 Server---*****
				objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & objStartFolder & ";Extended Properties=""text;HDR=YES;FMT=Delimited;"";"
				'objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;"";"

				'objExcelCon.Open "Driver={Microsoft Access Text Driver (*.txt; *.csv)};Dbq=" & filePath & ";Extensions=asc,csv,tab,txt;ColNameHeader=Yes;"

				'objExcelCon.Open strUploadPath & "\CAPSCSFromDiners.dsn"
				
				'Delete all of the existing records from the Training Table before loading the new dataset
				objCon.Execute "TRUNCATE TABLE tblCAPSTraining"
				
				'Write the SQL Query 
				objRS.open "SELECT * FROM [" & strFileName & "]", objExcelCon
				
				'Call the procedure to check the Excel fields and data are correct
				Call ReadExcel
		   
				'Check for errors
				If(errors<>"") Then
					'Print the errors and return
					Response.Write "<font face=arial size=1><b>File not uploaded. Please correct the followings errors and load again</b> <br> "& errors &"</font><br>"         
				Else  
					strFileDateTime = Right(strFileName,14)
					strFileDateTime = Left(strFileName,9)
					
					'If the no errors found in the ReadExcel method, then start uploading the records in the database       	    	    
					'UploadExcel strDeleteCheck,lngFileID, strFileName,filePath,strFileDateTime
					intRecordsLoaded = UploadExcel(strDeleteCheck,lngFileID)
					'UploadExcel Uploader.Form("chkDelete"),lngFileID, strFileName,filePath,strFileDateTime
					
					'Response.Write "<b>ANZ Cardlist File Successfully Uploaded!!</b><br>"
				End if
				
				
				
				'Close the recordset/connection 
				objRS.Close 
				objExcelCon.Close 
		    
			Else
				'Training file is xls converted from XLSX for the VBMRSN05 server, so do not enable text load
				'ReadText(filePath)
			End If
			
			'<!--------- End of duplicated code --->
	
		If intRecordsLoaded > 1 Then

			'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
			lngFileLoadID = SaveFileLoadID ("Training",strFileName, filePath,intRecordsLoaded,0,0,0,0,0,0,0,strFileDateTime,lngFileID,"Imported",Session("UserID"),"N")
			
			'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			Call UpdateFileLoadSummary ("Training",lngFileID, strFileName, lngFileLoadID)
			'response.write "UpdateFileLoadSummary (""ANZCardlist""," & lngFileID & "," & strFileName & "," & lngFileLoadID & ")"
			
			
			'''---Start the New Service Account Login section
			'If Request.QueryString("ActionType")="Service" Then

				
			'Else
				'Move the file to the loaded folder (run through up to 10 numbers at the end of the file name if the same file name exists in the loaded folder)
				'First set the name of the Loaded folder
				strFileType = "Loaded"
				'response.write strFileName
				If objFSO.FileExists(objStartFolder & "\" & strFileType & "\" & strFileName) Then
					For x = 1 to 10
						If objFSO.FileExists(objStartFolder & "\" & strFileType & "\" & x & strFileName) Then
						Else
							'Move the file to the Loaded folder
							objFSO.MoveFile objStartFolder & "\" & strFileName,objStartFolder & "\" & strFileType & "\" & x & Left(strFileName,Len(strFileName)-4) & Year(now()) & PadDigits(Month(now()),2) & Day(now()) & Hour(now()) & Minute(now()) & Right(strFileName,4)
							
							x = 10
						End If
					Next
				Else
					'Move the file to the Loaded folder
					objFSO.MoveFile objStartFolder & "\" & strFileName,objStartFolder & "\" & strFileType & "\" & Left(strFileName,Len(strFileName)-4) & Year(now()) & PadDigits(Month(now()),2) & Day(now()) & Hour(now()) & Minute(now()) & Right(strFileName,4)
					'strFileName = objStartFolder & "\" & objFile.Name
				End If
				
			''--End of Service Type IF when moving files after load
			'End If
	
			Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
							"<span aria-hidden=""true"">&times;</span></button>" & _
							"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
							"<span>Training File """ & strFileName & """ load COMPLETE! " & intRecordsLoaded & " records loaded.</span></div></div></div>"
						
		End If
	
	End if

'''---Start the New Service Account Login section
If Request.QueryString("ActionType")="Service" Then

	objNetwork.RemoveNetworkDrive strServer, True, False

	Set objNetwork = Nothing
End If
	
Set objFSO = Nothing

End Sub

Set objRS = Nothing
Set objCon = Nothing

 %>


