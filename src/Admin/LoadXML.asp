<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #include file="../CC/CAPSFunctions.asp" -->
<!-- #Include file=../ADOVBS.inc -->
<% 'Option Explicit 

'Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

'Description:	AE602 Form XML Upload Administration screen
'Author:		Michael Giacomin
'Date:			September 2020

    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")
	'Response.Write "DPC Provider = " 
	'Response.Write Session("DPCProvider")
	
'Instantiate Common Page Variables.
Dim objCon
Dim objCmd
Dim objRS
Dim objCmd2
Dim objCmd3
Dim objCmd4
Dim objCmd5
Dim objCmd6
Dim objCmd7
Dim objCmd8
Dim objCmd9

Dim strDeleteCheck
Dim dteBatchDate

Dim strSaveString
Dim strSaveString2
Dim y
Dim intRecord
Dim intFiles
Dim intParseErrors

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objCmd2 = Server.CreateObject("ADODB.Command")
Set objCmd3 = Server.CreateObject("ADODB.Command")
Set objCmd4 = Server.CreateObject("ADODB.Command")
Set objCmd5 = Server.CreateObject("ADODB.Command")
Set objCmd6 = Server.CreateObject("ADODB.Command")
Set objCmd7 = Server.CreateObject("ADODB.Command")
Set objCmd8 = Server.CreateObject("ADODB.Command")
Set objCmd9 = Server.CreateObject("ADODB.Command")

objCon.Open Session("DBConnection")

If request.QueryString("Action")="SaveFileLocal" Then
	Call LoadXMLDocs()
End If

If request.QueryString("Action")="Save" Then

Response.Write "Not available yet"
	'Call LoadXMLDocs()
End If

If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
End If

If Request.QueryString("Action")="ProcessApplicationContacts" Then

	'Call the procedure to Format CDMDC Details for CAPS (Out Address details/fields)
	Call UpdateApplicationCDMC(0)
	
	'Call the procedure to Add Contact details to applications
	Call UpdateApplicationContacts()
	
	'Call the procedure to check all AE602 applications for errors
	Call UpdateApplicationErrors()
	
	Call UpdateEmailErrorTemplate()
	
	'Call the procedure to Auto Approve applications with no errors which are Awaiting Review
	Call AutoApproveApplications()
	
	'Call the procedure to Auto reduce Credit Limit applications

	'Removed by AB 300125 - Credit Limit Apps are now loaded into credit limit tables on Application release
	Call AutoApproveLimitApplications()
	
	'Call the procedure to Delete Applications that have passed 14 days from the warning date.
	'Call DeleteWarningApplications()
	''''***** 28th June 2021
	''''***** Above has been turned off as it is run in another procedure....
	
End If

If Request.QueryString("Action")="PostErrorEmails" Then

	'Call the procedure to post error emails
	Call PostErrorEmails()	
	
End If



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
        if(getFileExt(document.getElementById('FILE1').value)==".xml" || getFileExt(document.getElementById('FILE1').value)==".txt")
        {
            if(window.confirm('This will overwrite any existing CDMC file data! \n \n Continue?')==true)
                {
                document.getElementById('Progress').style.display = "inline";
                frm.submit();
            }
        } 
        else
        {
            alert("Please enter a valid XML(.xml) file");
        }
    }
}

function UploadLocal()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="LoadXML.asp?Action=SaveFileLocal"
}

function UploadLocalG()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="LoadXML.asp?Action=SaveFileLocal&ActionType=Service"
}

jQuery(document).ready(function($) {
    $(".clickable-row").click(function() {
        window.location = $(this).data("href");
    });
});

function RefreshGXML() {
  self.location="LoadXML.asp"
}

function myFunction() {
  self.location="LoadXML.asp"
}

function DatePickChange() {
	self.location="LoadXML.asp?FileDate=" + document.getElementById("CSDate").value;
}

</script>


</head>
<body>

<!-- Modal -->
<div class="modal fade" id="ModApp" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
   <div class="loader">
        <div class="wrap">
            <div class="spinner"></div>
				<span class="loading-message">Loading...</h6>
        </div>
    </div>
</div>

<!--Loading Wait Spinner-->
	<div class="modal fade bd-example-modal-lg modalWait" id="ModalWait" data-backdrop="static" data-keyboard="false" tabindex="-1">
    <div class="modal-dialog modal-sm">
        <div class="modal-content" style="width: 88px">
            <span style="color:black;" class="spinner-border spinner-border-lg"></span>
        </div>
    </div>
</div>

<main class="main py-3">
      <div class="container">	
<form action="LoadXML.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	 <div class="row" id="basic-table">
  <div class="col-4">
    <div class="card">
     <div class="card-header">
         <h4 class="card-title"><img src="../images/defence_logo_dark.png" height="40px" width="110px" title="Applications"> Load Applications</h4>
        </div>
      <div class="card-content">
        <div class="card-body">
		

<div class="col-lg-12 col-md-12">

</div>

<div class="form-body">
<div class="row col-12">

  <div class="col-auto text-center">
	
	<button type="button" class="btn btn-primary btn-sm" onclick="UploadLocalG();" Title="Click to Load any existing file in the G Drive Applications Folder"><i class="fa fa-upload"></i> Load G</button> <i class="fa help-tooltip fa-question-circle" data-toggle="tooltip" title="Click to Load All XML Applications on the G Drive ('G Files To Be Loaded')."></i>
	<button type="button" class="btn btn-outline-secondary btn-sm" onclick="UploadLocal();" Title="Click to Load any existing file in the CAPS Server Applications Folder"><i class="fa fa-upload"></i> Load Apps</button> <i class="fa help-tooltip fa-question-circle" data-toggle="tooltip" title="Click to Load All XML Applications on the CAPS Server ('Files To Be Loaded')."></i>
	</div>
  
</div>
</div>

<div class="col-lg-12 col-md-12">
<br>
<span id="Progress" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>
<br>

	<fieldset class="form-group">
		<!--<button type="button" class="btn btn-outline-secondary btn-sm" onclick="setTimeout(myFunction, 3000)" data-toggle="modal" data-target="#ModalWait">-->
		<!--<button type="button" class="btn btn-outline-secondary btn-sm" onclick="self.location='LoadXML.asp?Action=ProcessApplicationContacts22222'" data-toggle="modal" data-target="#ModApp">-->
		<button type="button" class="btn btn-outline-secondary btn-sm" onclick="self.location='LoadXML.asp?Action=ProcessApplicationContacts'" data-toggle="modal" data-target="#ModalWait">
			<i class="fa fa-file"></i> Process AE602 Applications </button> <i class="fa help-tooltip fa-question-circle" data-toggle="tooltip" title="Click to update AE602 applications with contact details from CDMC (where the application has no contact details AND where already updated application details are different to CDMC details) AND Check applications for ERRORS."></i>
	</fieldset>
	<span style="font-weight:bold;">Functions:</span><BR>
		1. Adds CDMC details to applications<br>
		2. Checks applications for errors (and resolves)<br>
		3. Formats contact details (Address, phone)<br>
	    4. Adds Error Code for Error Emails<br>
		5. Auto Approves Applications<br>
		6. Auto Reduce Temporary Limit Applications<br>
		<div>
			<fieldset class="form-group">
			<!--<button type="button" class="btn btn-outline-secondary btn-sm" onclick="setTimeout(myFunction, 3000)" data-toggle="modal" data-target="#ModalWait">-->
			<!--<button type="button" class="btn btn-outline-secondary btn-sm" onclick="self.location='LoadXML.asp?Action=ProcessApplicationContacts22222'" data-toggle="modal" data-target="#ModApp">-->
			<button type="button" class="btn btn-outline-secondary btn-sm" onclick="self.location='LoadXML.asp?Action=PostErrorEmails'" data-toggle="modal" data-target="#ModalWait">
				<i class="fa fa-envelope-open-text"></i> Generate Error Emails </button> <i class="fa help-tooltip fa-question-circle" data-toggle="tooltip" title="Click to Generate Error Emails for Applications."></i>
		</fieldset>
		</div>
</div>

<div class="col-lg-12 col-md-12">
<!--<fieldset class="form-group">
    <button type="button" class="btn btn-outline-secondary btn-sm" onclick="window.open('CAPSADMIN/TemplateExcel.asp?T=tblCAPSANZCardlist')"><i class="fa fa-file"></i> Blank AE602 Form </button>
</fieldset>-->
</div>

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
   
   
  </div>
  
      <!-- Example DataTables Card-->
     <div class="row" id="basic-table">
  <div class="col-12">
    <div class="card">     
      <div class="card-content">
        <div class="card-body">
          <!-- Table with outer spacing -->
          <div class="table-responsive">         
				<%DisplayTableDetails()%>
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
	
	<!-- jQuery -->
    <script src="../js/jquery.js"></script>
	
	<!-- Bootstrap Core JavaScript -->
    <script src="../js/bootstrap.min.js"></script>
	
	
</body>

<!-- #Include file=../CC/CAPSFooter.asp -->
</html>

<%

Sub DisplayTableDetails()

Dim strWhere
Dim strAppType
Dim strAppName
Dim strAppEID
Dim strCardType
Dim x
Dim strApplicationVersion
Dim strDPCProvider

'Get the current applicatio nversion accepted from the System Parameters
strApplicationVersion = GetSystemAdmin("ApplicationVersion")
strDPCProvider = GetSystemAdmin("DPCProvider")

If IsNull(strApplicationVersion) OR strApplicationVersion = "" Then strApplicationVersion="No Version"

If Not IsEmpty(Request.QueryString("BatchNo")) Then
	If IsNull(Request.QueryString("BatchNo")) or Request.QueryString("BatchNo")= "" Then 
		
	Else
		strWhere = "WHERE DateLoaded = '" & Request.QueryString("BatchNo") & "'"
	
		If Request.QueryString("BatchNo") = "" Then strWhere = strWhere & " OR DateLoaded IS NULL"
	End If
Else
	strWhere = ""
End If

objRS.Open "SELECT TOP 500 * FROM qryCAPSXMLApplications WITH(NOLOCK) "  & strWhere & " ORDER BY [XMLApplicationID] DESC",objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no Application data loaded</B></th></tr>" & _
		        "<tr><td colspan=""8"" style=""text-align:left"">Okay to Upload Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 newd table-compact"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""5"" style=""text-align:left"">Sample of Existing Applications Loaded already in CAPS</th><th colspan=""3"">Application Versions Accepted: <span style=""color:red; font-weight:bold;"">" & strApplicationVersion & "</span> <span style=""font-weight:normal;"">and above</span></th></tr>" & _
		        "<th>Application ID</th><th>Date Loaded</th><th>Application Type</th>" & _
				"<th>EmployeeID</th><th>Applicant Name</th><th>App Version</th>" & _
				"<th style=""text-align:center"">Card Type</th><th>Loaded By</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
			
				x = x + 1
				
				'Determine the name fields to be displayed based on the application type
				strAppName = objRS("fldAppGivenNames") & " " & objRS("fldAppFamilyName")
				strAppEID = objRS("fldAppEmployeeID")
				strCardType = objRS("fldCardType")
				
				'Determine the application type (Card) based on the Apply For Field
				If IsNull(objRS("grpIAmApplyingFor")) or objRS("grpIAmApplyingFor") = "" Then
					strAppType = ""
				Else
					Select Case objRS("grpIAmApplyingFor")
					
						CASE 1
							strAppType = objRS("grpIAmApplyingFor") & " - DTC"
							strCardType = "DTC"							
						CASE 2
							strAppType = objRS("grpIAmApplyingFor") & " - Lodge Only"
							strCardType = "DTC"		
						CASE 3
							strAppType = objRS("grpIAmApplyingFor") & " - DUAL (CiH + Lodge)"
							strCardType = "DTC"							
						CASE 4
							If strDPCProvider = "ANZ" Then
								strAppType = objRS("grpIAmApplyingFor") & " - ANZ DPC"
								strCardType = "DPC"
							Else
								strAppType = objRS("grpIAmApplyingFor") & " - DPC"
								strCardType = "DPC"
							End If
						CASE 5
							strAppType = objRS("grpIAmApplyingFor") & " - DTC Limit Change"
							strAppName = objRS("subLimitChange_fldCDGivenNames") & " " & objRS("subLimitChange_fldCDFamilyName")
							strAppEID = objRS("subLimitChange_fldCDEmployeeID")
							strCardType = "DTC"
						CASE 6
							strAppType = objRS("grpIAmApplyingFor") & " - DPC Limit Change"
							strAppName = objRS("subLimitChange_fldCDGivenNames") & " " & objRS("subLimitChange_fldCDFamilyName")
							strAppEID = objRS("subLimitChange_fldCDEmployeeID")
							strCardType = "DPC"
						CASE 7
							strAppType = objRS("grpIAmApplyingFor") & " - DTC CiH Only"
							strCardType = "DTC"							
						Case 8
							strAppType = objRS("grpIAmApplyingFor") & " - Lodge Limit Change"
							strAppName = objRS("subLimitChange_fldCDGivenNames") & " " & objRS("subLimitChange_fldCDFamilyName")
							strAppEID = objRS("subLimitChange_fldCDEmployeeID")
							strCardType = "DTC"	
						Case Else
							strAppType = objRS("grpIAmApplyingFor")
							
					END Select
					
				End If
				
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR class='clickable-row' data-href='DisplayDataset.asp?tbl=tblCAPSXMLApplication&W=WHERE XMLApplicationID=" & objRS("XMLApplicationID") & "' data-target='_blank'><TD>" & objRS("XMLApplicationID") & "</TD><TD>" & objRS("DateLoaded") & "</TD><TD>" & strAppType & "</TD>" & _
			                    "<TD>" & strAppEID & "</TD><TD>" & strAppName & "</TD><TD style=""text-align:center"">" & objRS("AppVersion") & "</TD><TD style=""text-align:center"">" & strCardType & "</TD><TD>" & objRS("LoadedByName") & "</TD>" & _
			                    "</TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close

        Response.Write "<tr><th colspan=""6"" style=""text-align:right;""><th>Total</th><th>" & x & "</th></tr></table>"
		
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
Dim strDPCProvider

Dim strDateShort

strDPCProvider = GetSystemAdmin("DPCProvider")


	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)
	
		objRS.Open "SELECT TOP 9 * FROM qryCAPSApplicationSummary WITH(NOLOCK) WHERE DateLoaded >= '" & dteBatchDate & "' ORDER BY [DateLoaded] DESC",objCon
			
		'objRS.Open "SELECT TOP 20 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'Applications' ORDER BY FileSeqNum DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no Application data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd table-compact"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">Application Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
		        "<tr><th style=""text-align:right"">Date Loaded</th>" & _
				"<th style=""text-align:center"">Total Apps</th>" & _
				"<th style=""text-align:center"">DTC</th><th style=""text-align:center"">CiH</th><th style=""text-align:center"">DUAL DTC</th><th style=""text-align:center"">DPC</th><th style=""text-align:center"">DTC Limit</th><th style=""text-align:center"">DPC Limit</th><th style=""text-align:center"">Lodge Limit Change</th>" & _
	 	        "</tr>" 
				
		    End If
		    
		    Do until objRS.eof
				
				x = x + 1
				
				If IsNull(objRS("DateLoaded")) or objRS("DateLoaded")= "" Then
					strDateShort = ""
				Else
					strDateShort = Day(objRS("DateLoaded")) & " " & Left(MonthName(Month(objRS("DateLoaded"))),3)
				End If
				
				Response.Write "<TR><TD Title=""" & objRS("DateLoaded") & """ style=""text-align:right;""><a href=""LoadXML.asp?BatchNo=" & objRS("DateLoaded") & """>" & strDateShort & "</A></B></TD>" & _
							"<TD style=""text-align:center; font-weight:bold;"">" & objRS("Applications") & "</TD><TD style=""text-align:center;"">" & objRS("DTC") & "</TD><TD style=""text-align:center;"">" & objRS("DTCCiH") & "</TD><TD style=""text-align:center;"">" & objRS("DUALDTC") & "</TD>"  &_
							"<TD style=""text-align:center;"">" & objRS("DPC") & "</TD><TD style=""text-align:center;"">" & objRS("DTCLimitChange") & "</TD><TD style=""text-align:center;"">" & objRS("DPCLimitChange") & "</TD><TD style=""text-align:center;"">" & objRS("LodgeLimitChange") & "</TD></TR>"
				
				
    			objRS.Movenext			
		    Loop
    			
			
								
	    objRS.Close

        Response.Write "<tr><td colspan=""9""><i>Current DPC Provider:</i> <span style=""font-weight:bold;"">NAB</span></td></tr></table>"
		
End Sub

Public Sub DisplayFileSummary()

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile
Dim strFileExtension

Set objFSO = CreateObject("Scripting.FileSystemObject")

	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\Applications\")
	objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Applications\"
	
	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be loaded. Click 'Load Server' button to Load All XML Applications.""></i></th></tr>"
			
	For Each objFile in colFiles

		intCount = intCount + 1
		
		If intCount < 6 Then
			If IsNull(objFile.Name) or objFile.Name = "" Then
				strFile = ""
			Else
				strFile = Left(objFile.Name,10)
			End If
			
			'Display the correct file extension
			strFileExtension = Right(objFile.Name,3)
			
			Response.Write "<TR><TD>" & strFile & "..." & strFileExtension & "</TD></TR>"
			
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
Dim strFileExtension

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass

Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get the System Parameter for the start of the Training File Location
strServer = GetSystemAdmin("ApplicationImportFilePath")
'strServer = "\\groupdata.rus.car.drn.defence.mil.au\groupdata\CFO\CFO\CMS Admin\CAPS\Import Files\Training\"

'Get the System Parameter for the Service Account UserName and Password
strUser = GetSystemAdmin("CAPSServiceAccountName")
strPass = GetSystemAdmin("CAPSServiceAccountPassword")

On Error Resume Next

objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

If Err.Number <>0 Then
	'Write an error to the top of the screen
	'Response.Write "<div class=""container"" style=""position:relative; z-index:100; top:40px; left:40px;""><div class=""alert alert-danger"" role=""alert"" style=""position: absolute; top:40px; left:40px; z-index:100;"">Error! Server path not found: " & strServer & "</div></div>"
	'Write a message in the G Drive Div area
	Response.Write "<div class=""alert alert-danger"" role=""alert"" style=""position: absolute; top:0px; left:0px; z-index:100;"">Error! G Drive path not found: " & strServer & "</div>"			

	Err.Clear
	On Error Goto 0
	Exit Sub
	
End If

On Error Goto 0

	objStartFolder = strServer
	
	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"" title=""Click to Refresh. Server: " & strServer & """>G Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be loaded. Click 'Load Server' button to Load All XML Applications.""></i> <button type=""button"" name=""RefreshG"" id=""RefreshG"" class=""btn btn-secondary btn-xs"" onclick=""RefreshGXML();""><i class=""fa fa-redo""></i></button></th></tr>"
			
	For Each objFile in colFiles

		strFileExtension = Right(objFile.Name,3)
		
		If strFileExtension = "xml" Then
			
			intCount = intCount + 1
			
			If intCount < 5 Then
				If IsNull(objFile.Name) or objFile.Name = "" Then
					strFile = ""
				Else
					strFile = Left(objFile.Name,5)
				End If
				
				'Display the correct file extension
				strFileExtension = Right(objFile.Name,3)
				
				Response.Write "<TR><TD title=""" & objFile.Name & """>" & strFile & "..." & strFileExtension & "</TD></TR>"
				
			End If
		
		Else
		
		End If
		
	Next
	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"

'''---Start the New Service Account Login section
If Request.QueryString("ActionType")="Service" Then

	objNetwork.RemoveNetworkDrive strServer, True, False

	Set objNetwork = Nothing
End If

Set objFSO = Nothing

End Sub



Public Sub LoadXMLDocs()

Dim x
Dim objFSO
Dim strFileType
Dim objStartFolder
Dim objFolder
Dim colFiles
Dim objFile
Dim strFileName
Dim objXML
Dim CellCount
Dim objRoot
Dim objLevel1
Dim nodeName1
Dim nodeName2
Dim nodeName3
Dim nodeName4
Dim nodeText1
Dim nodeText2
Dim nodeText3
Dim nodeText4
Dim child1
Dim child2
Dim child3
Dim child4
Dim strSaveString
Dim strSaveString2
Dim objLevel2
Dim objLevel3
Dim objLevel4
Dim fldCount

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass
Dim strMoveNotXML
Dim strFileNameNew

Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")


'''---Start the New Service Account Login section
If Request.QueryString("ActionType")="Service" Then

	'Get the System Parameter for the start of the Training File Location
	strServer = GetSystemAdmin("ApplicationImportFilePath")

	'Get the System Parameter for the Service Account UserName and Password
	strUser = GetSystemAdmin("CAPSServiceAccountName")
	strPass = GetSystemAdmin("CAPSServiceAccountPassword")

	objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

		objStartFolder = strServer
		
		'Set objFolder = objFSO.GetFolder(objStartFolder)

Else
	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\Applications\")
	objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Applications\"

End If

'response.write objStartFolder & " <--</br>"

Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

intFiles = 0

'Rename Files with too long file name
For Each objFile in colFiles

	strFileName = objStartFolder & "\" & objFile.Name
	strFileNameNew = objFile.Name
	strFileNameNew = Replace(strFileNameNew,"[SEC=OFFICIAL]","")
	strFileNameNew = Replace(strFileNameNew,"Attachment1","")
	strFileNameNew = Replace(strFileNameNew,"Sensitive ACCESS=Personal-Privacy","")	
	
	If objFSO.FileExists(objStartFolder & "\" & strFileNameNew) Then
	
		'Response.Write "<div class=""alert alert-danger"" role=""alert"">There was an error moving application:" & strFilename & " TO:" & objStartFolder & "\" & strFileNameNew & ". Application already exists.</div>"

	Else
		objFSO.MoveFile strFilename, objStartFolder & "\" & strFileNameNew
	End If
	
Next

For Each objFile in colFiles
    'Response.Write objStartFolder & "\" & objFile.Name & "</br>"
	
	strSaveString = ""
	strSaveString2 = ""
	x = 0
	fldCount = 1
	
	strFileName = objStartFolder & "\" & objFile.Name
	
	'Response.Write Server.MapPath(strFileName)

	'strFileName = Server.MapPath("Attachment2.xml")

	'response.write strFileName & "</br>"

	'Only attempt to load XML or XFDF files
	If Right(strFileName,3) = "xml" OR Right(strFileName,3) = "xdp" Then
	
		'Set the File Type Parameter to XML for the move process at the end of this procedure
		strMoveNotXML = "XML"
		
		Set objXML = Server.CreateObject("Microsoft.XMLDOM")

		objXML.Async = False
		objXML.SetProperty "ServerHTTPRequest", True

		objXML.ResolveExternals = True

		objXML.ValidateOnParse = True

		objXML.Load(strFileName)

		CellCount = 0

		If (objXML.parseError.errorCode = 0) Then

		  Set objRoot = objXML.documentElement

		  If IsObject(objRoot) = False Then

			 'Response.Write "There was an error parsing xml"

		  Else
			
			For Each objLevel1 in objRoot.ChildNodes
				nodeName1 = objLevel1.NodeName
				nodeText1 = objLevel1.Text
				child1 = objLevel1.childNodes.length
				
				If child1<2 Then
					
					strSaveString = "," & strSaveString & nodename1 & " = " & nodeText1

					x = x + 1
					strSaveString2 = "," & strSaveString2 & x & " = " & nodeText1
					
					y = y + 1
					'Add Field Count to Digital Signature fields
					If nodeName1 = "subDigitalSignature" Then
						nodeName1 = nodeName1 & fldCount
						fldCount = fldCount + 1
						'response.Write(nodeName1 & " = " & nodeText1 & " a<br>")
					End If
					'Response.Write(nodeName1 & " = " & nodeText1 & " 1st<br>")
					Call SaveXMLTemp (nodename1,nodeText1)
				Else
					'response.Write(nodeName1 & " b<br>")
					strSaveString = "," & strSaveString & nodename1
					
					'x = x + 1
					'strSaveString2 = "," & strSaveString2 & x
				ENd If
				
				'second level
				If child1>1 Then
					For Each objLevel2 in objLevel1.ChildNodes
						nodeName2 = objLevel2.NodeName
						nodeText2 = objLevel2.Text
						child2 = objLevel2.childNodes.length
						
						If child2<2 Then
							
							''''' OLD DELETE 'response.Write("&nbsp;&nbsp;&nbsp;" &  nodename2 & " = " & nodeText2 & "<br>")
							strSaveString = "," & strSaveString & nodeName1 & "_" & nodename2 & " = " & nodeText2
							
							x = x + 1
							strSaveString2 = "," & strSaveString2 & x & " = " & nodeText2
							
							y = y + 1
							'Add Field Count to Digital Signature fields
							If nodename1 = "subDigitalSignature" Then
								nodename1 = nodename1 & fldCount
								fldCount = fldCount + 1
								'response.Write(nodeName1 & "_" & nodename2 & " : " & nodeText2 & " b<br>")
							End If
							Call SaveXMLTemp (nodeName1 & "_" & nodename2,nodeText2)
							'Response.Write nodename2 & "ZZ<BR>"
							'Response.Write("&nbsp;&nbsp;&nbsp;" & nodeName1 & "_" & nodename2 & " = " & nodeText2 & " 2nd<br>")
						Else
							'response.Write("&nbsp;&nbsp;&nbsp;" & nodeName1 & "_" & nodename2 & " dXX<br>")
							''''' OLD DELETE 'response.Write("&nbsp;&nbsp;&nbsp;" &  nodename2 & "<br>")
							strSaveString = "," & strSaveString & nodeName1 & "_" & nodename2
						End If
						
						'third level
						If child2>1 Then
							For Each objLevel3 in objLevel2.ChildNodes
								nodeName3 = objLevel3.NodeName
								nodeText3 = objLevel3.Text
								child3 = objLevel3.childNodes.length
								
								If child3<2 Then
									
									''''' OLD DELETE 'response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & nodename3 & " = " & nodeText3 & "<br>")
									strSaveString = "," & strSaveString & nodeName1 & "_" & nodename2 & "_" & nodename3 & " = " & nodeText3
									
									x = x + 1
									strSaveString2 = "," & strSaveString2 & x & " = " & nodeText3
									
									y = y + 1
									'Add Field Count to Digital Signature fields
									If nodename2 = "subDigitalSignature" Then
										nodename2 = nodename2 & fldCount
										fldCount = fldCount + 1
										'response.Write(nodeName1 & "_" & nodename2 & "_" & nodename3 & " : " & nodeText3 & " c<br>")
									End If
									'Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & nodeName1 & "_" & nodename2 & "_" & nodename3 & " = " & nodeText3 & " 3rd<br>")
									Call SaveXMLTemp (nodeName1 & "_" & nodename2 & "_" & nodename3,nodeText3)
							
								Else
									'response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & nodeName1 & "_" & nodename2 & "_" & nodename3 & " f<br>")
									''''' OLD DELETE response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & nodename3 & "<br>")
									strSaveString = "," & strSaveString & nodeName1 & "_" & nodename2 & "_" & nodename3
								End If
								
								
								'fourth level
								If child3>1 Then
									For Each objLevel4 in objLevel3.ChildNodes
										nodeName4 = objLevel4.NodeName
										nodeText4 = objLevel4.Text
										child4 = objLevel4.childNodes.length
										
										If child4<2 Then
											'response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &  nodename4 & " = " & nodeText3 & " g<br>")
											strSaveString = "," & strSaveString & nodename4 & " = " & nodeText3
											
											x = x + 1
											strSaveString2 = "," & strSaveString2 & x & " = " & nodeText3
									
											y = y + 1
												'Add Field Count to Digital Signature fields
												'If nodename4 = "subDigitalSignature" Then
													'nodename4 = nodename4 & fldCount
													'fldCount = fldCount + 1
													'response.Write (nodename4 & " : " & nodeText3 & " d <BR>")
												'End If
											Call SaveXMLTemp (nodename4,nodeText3)
											'Response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &  nodename4 & " = " & nodeText3 & " 4th<br>")
										Else
											'response.Write("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &  nodename4 & " h<br>")
											strSaveString = "," & strSaveString & nodename4
										End If
										
										
									Next
								End If
										
														
								
							Next
						End If
								
					Next
				End If
				
			Next


		  End If

		Else

			intParseErrors = intParseErrors + 1
		   'Response.Write "There was an error parsing xml " & objXML.parseError.errorCode
		   Response.Write "<div class=""alert alert-danger"" role=""alert"">There was an error parsing xml " & objXML.parseError.errorCode & " in " & intParseErrors & " files. File:" & strFileName & "</div>"

		End If

	Else
		If Right(strFileName,3) = "pdf" Then
			strMoveNotXML = "PDF"
		Else
			Response.Write "<div class=""alert alert-danger"" role=""alert"">" & strFileName & " is not a valid XML/XFDF file</div>" '" <-- Not an XML File."
		End If
		
	End If

	'Response.Write strSaveString2
	
	'Only run the count, load and move if the file in an XML file
	If strMoveNotXML = "XML" Then
		intFiles = intFiles + 1

		'Check to see if the application is old and if so move it to a not loaded folder and mark it in the application table
		If GetApplicationType = 0 Then
			'If the load is from the G drive then set the correct Loaded/Extracted folder
			If Request.QueryString("ActionType")="Service" Then
				'If the File is a PDF then move it to the relevant folder
				If strMoveNotXML = "PDF" Then
					strFileType = "Extracted"
				Else
					strFileType = "Loaded"
				End If
			Else
				strFileType = "Loaded"
			End If
		Else
			'If the load is from the G drive then set the correct NotLoaded/Rejected folder
			If Request.QueryString("ActionType")="Service" Then
				'If the File is a PDF then move it to the relevant folder
				If strMoveNotXML = "PDF" Then
					strFileType = "Extracted"
				Else
					strFileType = "Rejected"
				End If
			Else
				strFileType = "NotLoaded"
			End If
		End If
		
		''Execute the procedure to save the XML loaded into the temporary table above into the transposed Application table tblCAPSXMLApplication
		Call SaveXMLApplication(strFileName,strFileType)
		
		If objFSO.FileExists(objStartFolder & "\" & strFileType & "\" & objFile.Name) Then
			For x = 1 to 10
				If objFSO.FileExists(objStartFolder & "\" & strFileType & "\" & x & objFile.Name) Then
				Else
					'Move the file to the Loaded folder
					objFSO.MoveFile strFileName,objStartFolder & "\" & strFileType & "\" & x & objFile.Name
					
					x = 10
				End If
			Next
		Else
			'Move the file to the Loaded folder
			'Response.Write strFileName & "<BR>"
			'Response.Write objStartFolder & "\" & strFileType & "\" & objFile.Name

			objFSO.MoveFile strFileName,objStartFolder & "\" & strFileType & "\" & objFile.Name
			'strFileName = objStartFolder & "\" & objFile.Name
		End If
		
	'End of the check to see if the file is an XML (to avoid count, load and move)
	End If
	
'End of the loop through all files in the folder
Next

'''---Start the New Service Account Login section
If Request.QueryString("ActionType")="Service" Then

	objNetwork.RemoveNetworkDrive strServer, True, False

	Set objNetwork = Nothing
End If
	
Set objFSO = Nothing
Set objXML = Nothing

	If intFiles = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Applications Loaded from folder " & objStartFolder & ".</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert""> " & intFiles & " Application files loaded from " & objStartFolder & "</div>"
	End If
	
End Sub


Sub SaveRecord(s1,s2,s3,s4,s5,s6,s7,s8,s9,s10,s11,s12,s13,s14,s15,s16,s17,s18,s19,s20,s21,s22,s23,s24,s25,s26,s27,s28,s29,s30,s31,s32,s33,s34,s35,s36,s37,s38,s39,s40,s41,s42,s43,s44,s45,s46,s47,s48,s49,s50, _ 
			s51,s52,s53,s54,s55,s56,s57,s58,s59,s60,s61,s62,s63,s64,s65,s66,s67,s68,s69,s70,s71,s72,s73,s74,s75,s76,s77,s78,s79,s80,s81,s82,s83,s84,s85,s86,s87,s88,s89,s90,s91,s92,s93,s94,s95,s96,s97,s98,s99,s100, _ 
			s101,s102,s103,s104,s105,s106,s107,s108,s109,s110,s111,s112,s113,s114,s115,s116,s117,strFileName)

'''''NOTE'''''
'Procedure to Load the XML Application in one go after getting all of the NODES (Name and Text) rather than each node one at a time.
'Not used as there were more issues with this than each node individually in a table (as individual records) then transposing to the XML table tblCAPSXMLApplication
'Left here just in case it is needed in the future as there are a lot of records (saves re-typing)!

Dim intRecord

If IsNumeric(strCompanyCode) Then
Else
	strCompanyCode = 0
End If

  	With objCmd
  	
  	    'If the procedure has already run then don't create the parameter objects again (more than once)
  	    If x = 1 then
                .CommandType = 4
                .CommandText = "spCAPSXMApplicationSave"
                
				.Parameters.Append objCmd.CreateParameter("XMLApplicationID ", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("dteGFOValidDate ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldAppEmployeeID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldAppFamilyName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldAppGivenNames ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldCaption ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldCardType ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldCCSCOnlyShowBtns ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldClassification_body_p ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldFormNumber ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldForwardIdent ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldGFOEmployeeID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldGFOFamilyName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldGFOGivenNames ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldIdent ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldIsSigned ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldSigningStatus ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("fldTriggerRelayoutFlag ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("grpIAmApplyingFor ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkAckConditions ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkActivateMasterCard ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkAppDeleted ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkCheckedAndConfirm ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkContactDetailsCorrect ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkCurrentActive ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkCustodyofCard ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkDinersNotAccepted ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkDoesNotReplace ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkOnDRN ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkRetainReceipts ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkTravellingOS ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkUnbrandedCard ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_chkUsesSamePIN ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_ddlAACompanyCode ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_ddlAppGender ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_ddlCivilianTitle ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_ddlCMCGroup ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_ddlCMSAccountHolder ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_dteAppAge ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_dteCMCValidDate ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_dteDTPCValidDate ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_dteDTPSupeValidDate ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldAACostCentre ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldAAFund ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldAAInternalOrder ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldAAWBS ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldAccHolderCMSID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldAccHolderEmployeeID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldAccHolderName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldAppGroup ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldCMCEmployeeID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldCMCFamilyName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldCMCGivenNames ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldDTPSupeEmployeeID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldDTPSupeFamilyName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldDTPSupeGivenNames ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_fldWorkEmailAddress ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_grpEmailDomain ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subCompanionMCPrintAttachSubmit_fldCompanionMCShowBtns ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDigitalSignature1_fldSignFlag ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDigitalSignature2_fldSignFlag ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDigitalSignature3_ddlCMAGroup ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDigitalSignature3_fldCMAEmployeeID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDigitalSignature3_fldCMAFamilyName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDigitalSignature3_fldCMAGivenNames ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDigitalSignature3_fldCMAWorkEmailAddress ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDigitalSignature3_fldSignFlag ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDPCAgreement_subPoint4 ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDPCAgreement_subPoint5 ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDPCAgreement_subPoint6 ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDPCAgreement_subPoint7 ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDPCSupePrintAttachSubmit_fldDTPCSupeShowBtns ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDTCAgreement_subPoint2 ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDTCAgreement_subPoint3 ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDTCAgreement_subPoint4 ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDTCAgreement_subPoint6 ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subApply_subDTPCApplicantPrintAttachSubmit_fldDTPCAppShowBtns ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subDigitalSignature6_fldCMAFamilyName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subDigitalSignature6_fldCMAGivenNames ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subDigitalSignature6_fldSAEmployeeID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subDigitalSignature6_fldSignFlag ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subGroupCFOPrintAttachSubmit_fldGroupCFOShowBtns ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_chkIncreaseDecreaseLimit ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_ddlCDGroup ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_ddlRaisePurchaseOrder ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_dteCSValidDate ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_dteDTCAppValidDate ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldCDEmployeeID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldCDFamilyName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldCDGivenNames ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldCDWorkEmailAddress ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldCSEmployeeID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldCSFamilyName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldCSGivenNames ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldDTCAppEmployeeID ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldDTCAppFamilyName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldDTCAppGivenNames ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldJustificationForChange ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_fldLastFourDigits ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_grpChangesPermanent ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_numCurrentLimit ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_numNewLimit ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subCardHolderPrintAttachSubmit_fldCardHolderShowBtns ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subCreditLimitJustification_ddlHigherLimitTransactions ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subCreditLimitJustification_ddlNatureofTransaction ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subCreditLimitJustification_ddlOperationorExercise ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subCreditLimitJustification_ddlProposedCreditLimit ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subDigitalSignature4_fldSignFlag ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subDigitalSignature5_dteValidDate ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subDigitalSignature5_fldCardHoldersName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subDigitalSignature5_fldSignFlag ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subDTCAppPrintAttachSubmit_fldDTCAppShowBtns ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subPeriodofChangeFromDateToDate_dteFrom ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subPeriodofChangeFromDateToDate_dteTo", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subRequestedTransactionAmountDates_numLimitAmount", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subRequestedTransactionAmountDates_subRequestedPeriodFromDateToDate_dteRequestedDateFrom", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("subLimitChange_subRequestedTransactionAmountDates_subRequestedPeriodFromDateToDate_dteRequestedDateTo", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("FileName ", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("Loaded ", adchar, adParamInput,1)
				.Parameters.Append objCmd.CreateParameter("LoadedBy ", adInteger, adParamInput,0)
				.Parameters.Append objCmd.CreateParameter("DateImported ", addate, adParamInput,0)
				.Parameters.Append objCmd.CreateParameter("ImportedBy ", adInteger, adParamInput,0)

				.Parameters.Append objCmd.CreateParameter("CostCentreIDOutput", adInteger, adParamOutput)				
            
        End If
                 
				.Parameters("XMLApplicationID") = s1
				.Parameters("dteGFOValidDate") = s2
				.Parameters("fldAppEmployeeID") = s3
				.Parameters("fldAppFamilyName") = s4
				.Parameters("fldAppGivenNames") = s5
				.Parameters("fldCaption") = s6
				.Parameters("fldCardType") = s7
				.Parameters("fldCCSCOnlyShowBtns") = s8
				.Parameters("fldClassification_body_p") = s9
				.Parameters("fldFormNumber") = s10
				.Parameters("fldForwardIdent") = s11
				.Parameters("fldGFOEmployeeID") = s12
				.Parameters("fldGFOFamilyName") = s13
				.Parameters("fldGFOGivenNames") =s14
				.Parameters("fldIdent") = s15
				.Parameters("fldIsSigned") = s16
				.Parameters("fldSigningStatus") = s17
				.Parameters("fldTriggerRelayoutFlag") = s18
				.Parameters("grpIAmApplyingFor") = s19
				.Parameters("subApply_chkAckConditions") = s20
				.Parameters("subApply_chkActivateMasterCard") = s21
				.Parameters("subApply_chkAppDeleted") = s22
				.Parameters("subApply_chkCheckedAndConfirm") = s23
				.Parameters("subApply_chkContactDetailsCorrect") = s24
				.Parameters("subApply_chkCurrentActive") = s25
				.Parameters("subApply_chkCustodyofCard") = s26
				.Parameters("subApply_chkDinersNotAccepted") = s27
				.Parameters("subApply_chkDoesNotReplace") = s28
				.Parameters("subApply_chkOnDRN") = s29
				.Parameters("subApply_chkRetainReceipts") = s30
				.Parameters("subApply_chkTravellingOS") = s31
				.Parameters("subApply_chkUnbrandedCard") = s32
				.Parameters("subApply_chkUsesSamePIN") = s33
				.Parameters("subApply_ddlAACompanyCode") = s34
				.Parameters("subApply_ddlAppGender") = s35
				.Parameters("subApply_ddlCivilianTitle") = s36
				.Parameters("subApply_ddlCMCGroup") = s37
				.Parameters("subApply_ddlCMSAccountHolder") = s38
				.Parameters("subApply_dteAppAge") = s39
				.Parameters("subApply_dteCMCValidDate") = s40
				.Parameters("subApply_dteDTPCValidDate") = s41
				.Parameters("subApply_dteDTPSupeValidDate") = s42
				.Parameters("subApply_fldAACostCentre") = s43
				.Parameters("subApply_fldAAFund") = s44
				.Parameters("subApply_fldAAInternalOrder") = s45
				.Parameters("subApply_fldAAWBS") = s46
				.Parameters("subApply_fldAccHolderCMSID") = s47
				.Parameters("subApply_fldAccHolderEmployeeID") = s48
				.Parameters("subApply_fldAccHolderName") = s49
				.Parameters("subApply_fldAppGroup") = s50
				.Parameters("subApply_fldCMCEmployeeID") = s51
				.Parameters("subApply_fldCMCFamilyName") = s52
				.Parameters("subApply_fldCMCGivenNames") = s53
				.Parameters("subApply_fldDTPSupeEmployeeID") = s54
				.Parameters("subApply_fldDTPSupeFamilyName") = s55
				.Parameters("subApply_fldDTPSupeGivenNames") = s56
				.Parameters("subApply_fldWorkEmailAddress") = s57
				.Parameters("subApply_grpEmailDomain") = s58
				.Parameters("subApply_subCompanionMCPrintAttachSubmit_fldCompanionMCShowBtns") = s59
				.Parameters("subApply_subDigitalSignature1_fldSignFlag") = s60
				.Parameters("subApply_subDigitalSignature2_fldSignFlag") = s61
				.Parameters("subApply_subDigitalSignature3_ddlCMAGroup") = s62
				.Parameters("subApply_subDigitalSignature3_fldCMAEmployeeID") = s63
				.Parameters("subApply_subDigitalSignature3_fldCMAFamilyName") = s64
				.Parameters("subApply_subDigitalSignature3_fldCMAGivenNames") = s65
				.Parameters("subApply_subDigitalSignature3_fldCMAWorkEmailAddress") = s66
				.Parameters("subApply_subDigitalSignature3_fldSignFlag") = s67
				.Parameters("subApply_subDPCAgreement_subPoint4") = s68
				.Parameters("subApply_subDPCAgreement_subPoint5") = s69
				.Parameters("subApply_subDPCAgreement_subPoint6") = s70
				.Parameters("subApply_subDPCAgreement_subPoint7") = s71
				.Parameters("subApply_subDPCSupePrintAttachSubmit_fldDTPCSupeShowBtns") = s72
				.Parameters("subApply_subDTCAgreement_subPoint2") = s73
				.Parameters("subApply_subDTCAgreement_subPoint3") = s74
				.Parameters("subApply_subDTCAgreement_subPoint4") = s75
				.Parameters("subApply_subDTCAgreement_subPoint6") = s76
				.Parameters("subApply_subDTPCApplicantPrintAttachSubmit_fldDTPCAppShowBtns") = s77
				.Parameters("subDigitalSignature6_fldCMAFamilyName") = s78
				.Parameters("subDigitalSignature6_fldCMAGivenNames") = s79
				.Parameters("subDigitalSignature6_fldSAEmployeeID") = s80
				.Parameters("subDigitalSignature6_fldSignFlag") = s81
				.Parameters("subGroupCFOPrintAttachSubmit_fldGroupCFOShowBtns") = s82
				.Parameters("subLimitChange_chkIncreaseDecreaseLimit") = s83
				.Parameters("subLimitChange_ddlCDGroup") = s84
				.Parameters("subLimitChange_ddlRaisePurchaseOrder") = s85
				.Parameters("subLimitChange_dteCSValidDate") = s86
				.Parameters("subLimitChange_dteDTCAppValidDate") = s87
				.Parameters("subLimitChange_fldCDEmployeeID") = s88
				.Parameters("subLimitChange_fldCDFamilyName") = s89
				.Parameters("subLimitChange_fldCDGivenNames") = s90
				.Parameters("subLimitChange_fldCDWorkEmailAddress") = s91
				.Parameters("subLimitChange_fldCSEmployeeID") = s92
				.Parameters("subLimitChange_fldCSFamilyName") = s93
				.Parameters("subLimitChange_fldCSGivenNames") = s94
				.Parameters("subLimitChange_fldDTCAppEmployeeID") = s95
				.Parameters("subLimitChange_fldDTCAppFamilyName") = s96
				.Parameters("subLimitChange_fldDTCAppGivenNames") = s97
				.Parameters("subLimitChange_fldJustificationForChange") = s98
				.Parameters("subLimitChange_fldLastFourDigits") = s99
				.Parameters("subLimitChange_grpChangesPermanent") = s100
				.Parameters("subLimitChange_numCurrentLimit") = s101
				.Parameters("subLimitChange_numNewLimit") = s102
				.Parameters("subLimitChange_subCardHolderPrintAttachSubmit_fldCardHolderShowBtns") = s103
				.Parameters("subLimitChange_subCreditLimitJustification_ddlHigherLimitTransactions") = s104
				.Parameters("subLimitChange_subCreditLimitJustification_ddlNatureofTransaction") = s105
				.Parameters("subLimitChange_subCreditLimitJustification_ddlOperationorExercise") = s106
				.Parameters("subLimitChange_subCreditLimitJustification_ddlProposedCreditLimit") = s107
				.Parameters("subLimitChange_subDigitalSignature4_fldSignFlag") = s108
				.Parameters("subLimitChange_subDigitalSignature5_dteValidDate") = s109
				.Parameters("subLimitChange_subDigitalSignature5_fldCardHoldersName") = s110
				.Parameters("subLimitChange_subDigitalSignature5_fldSignFlag") = s111
				.Parameters("subLimitChange_subDTCAppPrintAttachSubmit_fldDTCAppShowBtns") = s112
				.Parameters("subLimitChange_subPeriodofChangeFromDateToDate_dteFrom") = s113
				.Parameters("subLimitChange_subPeriodofChangeFromDateToDate_dteTo") =s114
				.Parameters("subLimitChange_subRequestedTransactionAmountDates_numLimitAmount") = s115
				.Parameters("subLimitChange_subRequestedTransactionAmountDates_subRequestedPeriodFromDateToDate_dteRequestedDateF") = s116
				.Parameters("subLimitChange_subRequestedTransactionAmountDates_subRequestedPeriodFromDateToDate_dteRequestedDateT") = s117
				.Parameters("FileName") = strFileName
				.Parameters("Loaded") = "Y"
				.Parameters("LoadedBy") = Session("UserID")
				.Parameters("DateImported") = NULL
				.Parameters("ImportedBy") = ""
           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute        
            
            'Return the result of the Save Function.
     		intRecord = objCmd.Parameters.Item("CostCentreIDOutput")    
     		                  			     				     		     		
       'response.write  "exec spGeneralExpensesSave =0," & Session("BudgetID") & "," & Session("VersionID") & "," & CostCentreID & ",'GEXP'" & GLCode & "," & BM1 & "," & BM2 & "," & BM3 & "," & BM4 & "," & BM5 & "," & _
       '                     BM6 & "," & BM7 & "," & BM8 & "," & BM9 & "," & BM10 & "," & BM11 & "," & BM12 & "," & OY1 & "," & OY2 & "," & OY3 & ",'" & Comments & "','" & UpdatedBy & "'," & Session("ColumnLock")
End Sub



Public Sub SaveXMLTemp(strNodeName, strNodeText)
'Procedure to run a stored procedure which updates the summary details for a file just loaded, which is used where summary details are displayed

'If there is no node name (data) then exit the procedure
If IsNull(strNodeName) Then Exit Sub

If IsNull(strNodeText) Then strNodeText = ""

'response.write y & " "
  	With objCmd2
  	

	If y = 1 then
	
		.CommandType = 4
		.CommandText = "spCAPSXMLTempSave"
		
		.Parameters.Append objCmd2.CreateParameter("NodeName", adVarchar, adParamInput,200)
		.Parameters.Append objCmd2.CreateParameter("NodeText", adVarchar, adParamInput,1000)
	End if
	
		.Parameters("NodeName") = strNodeName
		.Parameters("NodeText") = Left(strNodeText,1000)

		.ActiveConnection = objCon
                
    End With
                
	objCmd2.Execute        
	
		
'Set ObjCmd2 = Nothing

End Sub


Public Sub SaveXMLApplication(strFileName,strFormType)
'Procedure to run a stored procedure which inserts all of the temporary data from the XMLTemp table into the application table (transpose)

  	With objCmd3
  	
		.CommandType = 4
		.CommandText = "spCAPSProcessXMLApplication"
		
		If intFiles = 1 Then
		
			.Parameters.Append objCmd3.CreateParameter("FileName", adVarchar, adParamInput,500)
			.Parameters.Append objCmd3.CreateParameter("LoadedBy", adInteger, adParamInput)
			.Parameters.Append objCmd3.CreateParameter("FormType", adVarchar, adParamInput,10)
			.Parameters.Append objCmd3.CreateParameter("XMLApplicationIDOutput", adInteger, adParamOutput)
		End If

		.Parameters("FileName") = strFileName
		.Parameters("LoadedBy") = Session("UserID")
		.Parameters("FormType") = strFormType
		
		
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd3.Execute        
	
	'Return the result of the Save Function.
     intRecord = objCmd3.Parameters.Item("XMLApplicationIDOutput")  

End Sub


Public Sub UpdateApplicationContacts()
'Procedure to run a stored procedure which updates all applications (based on their status [default="on hold"]) and adds CDMC details to them as they are not collected as part of the AE602 form.
Dim strStatusProcess

	'The status passed into the procedure determines which status of applications is processed (have CDMC details updated into their application)
	strStatusProcess = "Awaiting Review"
	
  	With objCmd3
  	
		.CommandType = 4
		.CommandText = "spCAPSApplicationXMLUpdateContact"
		
		.Parameters.Append objCmd3.CreateParameter("EIDInput", adVarchar, adParamInput,12)
		.Parameters.Append objCmd3.CreateParameter("StatusInput", adVarchar, adParamInput,20)
		.Parameters.Append objCmd3.CreateParameter("UpdateCountOutput", adInteger, adParamOutput)
		
		.Parameters("EIDInput") = "" 'Empty as this then processes all applications (AE602) rather than just the EID passed in
		.Parameters("StatusInput") = strStatusProcess 'This can be left empty to process "on hold" applications only or made another specific Status to process ("Awaiting review") 
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd3.Execute 

	'Return the result of the Save Function.
     intRecord = objCmd3.Parameters.Item("UpdateCountOutput")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Application contact details updated with status " & strStatusProcess & "!</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Applications with Status " & strStatusProcess & " updated with CDMC contact details!</div>"
	End If
	
End Sub

Public Sub UpdateApplicationCDMC(strEID)
'Procedure to run a stored procedure which checks all AE602 XML applications (on hold or awaiting review) for errors and adds errors or resolves them.
	
  	With objCmd4
  	
		.CommandType = 4
		.CommandText = "spCAPSCDMCProcessContactDetails"
		
		.Parameters.Append objCmd4.CreateParameter("UserID", adInteger, adParamInput)
		.Parameters.Append objCmd4.CreateParameter("EmployeeID", adVarchar, adParamInput,20)
		.Parameters.Append objCmd4.CreateParameter("CDMCProcessOutput", adInteger, adParamOutput)
		
		.Parameters("UserID") = Session("UserID") 'The User who had performed the error checks
		.Parameters("EmployeeID") = strEID'"0" 'Set EmployeeID to 0 so that procedure checks all applications
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd4.Execute        
	
	'Return the result of the Save Function.
     intRecord = objCmd4.Parameters.Item("CDMCProcessOutput")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Applications CDMC details updated!</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Application for " & strEID & " updated from CDMC details!</div>"
	End If
	
End Sub


Public Sub UpdateApplicationErrors()
'Procedure to run a stored procedure which checks all AE602 XML applications (on hold or awaiting review) for errors and adds errors or resolves them.
	
  	With objCmd2
  	
		.CommandType = 4
		.CommandText = "spCAPSApplicationErrors"
		
		.Parameters.Append objCmd2.CreateParameter("UserID", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarchar, adParamInput,20)
		.Parameters.Append objCmd2.CreateParameter("ApplicationErrorsOutput", adInteger, adParamOutput)
		
		.Parameters("UserID") = Session("UserID") 'The User who had performed the error checks
		.Parameters("EmployeeID") = "0" 'Set EmployeeID to 0 so that procedure checks all applications
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd2.Execute        
	
	'Return the result of the Save Function.
     intRecord = objCmd2.Parameters.Item("ApplicationErrorsOutput")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Applications checked for errors with status On Hold or Awaiting Review!</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Applications with Status On Hold or Awaiting Review checked for Errors!</div>"
	End If
	
End Sub

Public Sub UpdateEmailErrorTemplate()
'Procedure to run a stored procedure which checks all AE602 XML applications (on hold or awaiting review) for errors and adds errors or resolves them.
	
  	With objCmd5
  	
		.CommandType = 4
		.CommandText = "spCAPSProcessEmailErrorTemplate"
		
		.Parameters.Append objCmd5.CreateParameter("ApplicationID", adInteger, adParamInput)
				
		.Parameters("ApplicationID") = "0" 'Set EmployeeID to 0 so that procedure checks all applications
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd5.Execute        
	
End Sub

Public Sub AutoApproveApplications()
'Procedure to run a stored procedure which checks all AE602 XML applications (on hold or awaiting review) for errors and adds errors or resolves them.
	
  	With objCmd6
  	
		.CommandType = 4
		.CommandText = "spCAPSAutoApproveApplications"
		
		.Parameters.Append objCmd6.CreateParameter("ApplicationID", adInteger, adParamInput)
		.Parameters.Append objCmd6.CreateParameter("UserID", adInteger, adParamInput)
		.Parameters.Append objCmd6.CreateParameter("AutoApproveOutputID", adInteger, adParamOutput)
		
		.Parameters("ApplicationID") = 0 'Set Application to 0 so that procedure checks all applications
		.Parameters("UserID") = 0'Session("UserID") 'The User who had performed the error checks
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd6.Execute        
	
	'Return the result of the Save Function.
     intRecord = objCmd6.Parameters.Item("AutoApproveOutputID")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Applications Auto Approved from Awaiting Review!</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord - 1 & " Applications with Status Awaiting Review AUTO APPROVED!</div>"
	End If
	
End Sub


Public Sub AutoApproveLimitApplications()
'Procedure to run a stored procedure which checks all AE602 XML applications (on hold or awaiting review) for errors and adds errors or resolves them.
Dim strLimitParameter

	'First check to see if the parameter to Auto Reduce Limits is ActiveConnection
	strLimitParameter = GetSystemAdmin("LimitChangeAutoReduction")
	
	If IsNull(strLimitParameter) or strLimitParameter = "" Then
		Response.Write "<div class=""alert alert-warning"" role=""alert"">Limit Applications Auto Reduce NOT RUN! System Setting 'LimitChangeAutoReduction' is empty (Not 'Y')</div>"
		Exit Sub
	Else
		If strLimitParameter = "Y" Then
			'Continue on to execute the procedure
		Else
			Response.Write "<div class=""alert alert-warning"" role=""alert"">Limit Applications Auto Reduce NOT RUN! System Setting 'LimitChangeAutoReduction' is:" & strLimitParameter & " (Not 'Y')</div>"
			Exit Sub
		End If
	End If
	
	'Removed by AB 09/12/2021 and replaced with 
  	'With objCmd7
  	
		'.CommandType = 4
		'.CommandText = "spCAPSAutoApproveApplicationsLimits"
		
		'.Parameters.Append objCmd7.CreateParameter("ApplicationID", adInteger, adParamInput)
		'.Parameters.Append objCmd7.CreateParameter("UserID", adInteger, adParamInput)
		'.Parameters.Append objCmd7.CreateParameter("AutoApproveOutputID", adInteger, adParamOutput)
		
		'.Parameters("ApplicationID") = 0 'Set Application to 0 so that procedure checks all applications
		'.Parameters("UserID") = 0'Session("UserID") 'The User who had performed the error checks
		
		'.ActiveConnection = objCon
                
    'End With
	
	With objCmd7
  	
		.CommandType = 4
		.CommandText = "spCAPSTempLimitChangeApplicationsRevert"
		
		.Parameters.Append objCmd7.CreateParameter("RecordsUpdated", adInteger, adParamOutput)
		
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd7.Execute        
	
	'Return the result of the Save Function.
     intRecord = objCmd7.Parameters.Item("RecordsUpdated")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Limit Applications Auto Reduced!</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Applications with temporary Credit Limit Change Reduced!</div>"
	End If
	
End Sub

Public Sub DeleteWarningApplications()
'Procedure to run a stored procedure which checks all AE602 XML applications (awaiting review) where the warning date has passed and deletes them.
	
  	With objCmd9
  	
		.CommandType = 4
		.CommandText = "spCAPSApplicationWarningDateProcess"
		
		.Parameters.Append objCmd9.CreateParameter("ApplicationID", adInteger, adParamInput)
		.Parameters.Append objCmd9.CreateParameter("UserID", adInteger, adParamInput)
		.Parameters.Append objCmd9.CreateParameter("WarningOutputID", adInteger, adParamOutput)
		
		.Parameters("ApplicationID") = 0 'Set Application to 0 so that procedure checks all applications
		.Parameters("UserID") = 0'Session("UserID") 'The User who had performed the error checks
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd9.Execute        
	
	'Return the result of the Save Function.
     intRecord = objCmd9.Parameters.Item("WarningOutputID")  

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Applications Deleted after Warning Date Check!</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord - 1 & " Applications with Status Awaiting Review DELETED as warning date has passed!</div>"
	End If
	
End Sub


Public Sub PostErrorEmails()
'Procedure to run a stored procedure which checks all AE602 XML applications (on hold or awaiting review) for errors and adds errors or resolves them.
	
  	With objCmd8
  	
		.CommandType = 4
		.CommandText = "spCAPSPostErrorEmails"
		
		.Parameters.Append objCmd8.CreateParameter("ApplicationID", adInteger, adParamInput)
		.Parameters.Append objCmd8.CreateParameter("UpdatedBy", adInteger, adParamInput)
				
		.Parameters("ApplicationID") = "0" 'Set ApplicationID to 0 so that procedure checks all applications
		.Parameters("UpdatedBy") = Session("UserID")
		
		.ActiveConnection = objCon
                
    End With
                
	objCmd8.Execute        
	
	Response.Write "<div class=""alert alert-success"" role=""alert"">Application Error Emails have been generated.</div>"
	
End Sub


Public Function GetApplicationType()
'Procedure to get the Application details to see if the application is a current form or not
'The old CAPS used the XML field 'subACTUseOnly_ddlOffice' (which is from an application form dropped on 15-Sep-2015) so if an application contains this field it is too old to process
'The field being used to check the form currency may change.

Dim objRS3
Dim strSQL3

Set objRS3 = Server.CreateObject("ADODB.Recordset")

	strSQL3 = "SELECT * FROM tblCAPSXMLTemp WITH(NOLOCK) WHERE [NodeName] = 'subACTUseOnly_ddlOffice'"

	objRS3.Open strSQL3,objCon

		If objRS3.EOF Then
			GetApplicationType = 0
		Else
			GetApplicationType = 1'objRS3("XMLApplicationID")
		End If

	objRS3.Close
 
Set objRS3 = Nothing
	
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
    Response.Write "<font face=arial size=1><b>Error occurred while trying to process the request<b><br>Please contact system administrator</font><br>"
    Error.Clear
End If


Set ObjCon = Nothing
Set ObjRS = Nothing
Set ObjCmd = Nothing
Set ObjCmd2 = Nothing
Set ObjCmd3 = Nothing

%>

