<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #include file="upload.asp" -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #include file="../CC/CAPSFunctions.asp" -->
<%
'Description:	Old PDF (XML) AE602 Application Upload Administration screen
'Author:		Michael Giacomin
'Date:			April 2020

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

function UploadLocal()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="UploadApplications.asp?Action=SaveFileLocal"
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
	self.location="UploadApplications.asp?FileDate=" + document.getElementById("CSDate").value;
}
</script>
<script src="../js/jquery.js"></script>

<body>
<main class="main py-3">
      <div class="container">
<form action="UploadApplications.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	<div class="row" id="basic-table">
  <div class="col-6">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../images/defence_logo_dark.png" height="40px" width="110px" title="Diners"> Application File Load</h4>
        </div>
      <div class="card-content">
        <div class="card-body">
		

<div class="col-lg-12 col-md-12">
<fieldset class="form-group">
	<label for="basicInputFile">Select a File to Load</label>
	<div class="custom-file">
		<input type="file" class="custom-file-input" NAME="FILE1" id="FILE1" >
		<label class="custom-file-label" for="FILE1" >Choose file</label>
	</div>
</fieldset>
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
<div class="row col-11">
<div class="col-auto mr-auto">
	<div class="form-group">
	   <div class="checkbox">
		<input HIDDEN type="checkbox" class="checkbox-input" id="chkDelete" name="chkDelete">
		<label HIDDEN for="chkDelete">Overwrite Existing Batch</label>
	  </div>
	</div>
  </div>
  <div class="col-auto text-center">
	<button type="button" class="btn btn-primary btn-xs" onclick="UploadLocal();"><i class="fa fa-upload"></i> Load Local</button>
  </div>
  <div class="col-auto text-right">
	<button type="button" class="btn btn-primary btn-xs" onclick="upload();"><i class="fa fa-upload"></i> Upload</button>
  </div>
</div>
</div>

<div class="col-lg-12 col-md-12">
<p class="text-left">
<span id="Progress" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>
<br>
<font color="red" size="2"><b>NOTE: </Font><font color="black" size="2">Applications must be XML only not PDF
            <BR>* Files must be located in the "G:" drive folder
            </B></Font>
</p>
</div>

<!--<div class="col-lg-12 col-md-12">
<fieldset class="form-group">
    <button type="button" class="btn btn-primary btn-xs" onclick="window.open('CSFromDinersTemplateExcel.asp')"><i class="fa fa-file"></i> Application Template Excel </button>
</fieldset>
</div>-->

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

objRS.Open "SELECT TOP 50 * FROM qryCAPSApplications "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no Application data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Upload Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing Application Data already in CAPS</th></tr>" & _
                "<tr><td colspan=""16"" style=""text-align:left; color:red;font-size:20px""><B>WARNING! The data below will be deleted if you Upload a new Application file!</B></td></tr><tr>" & _
		        "<th>Application ID</th><th>Date Submitted</th>" & _
				"<th>Name On Card</th>" & _
				"<th>Card Type</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR><TD>" & objRS("ApplicationID") & "</TD><TD>" & objRS("DateSubmitted") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("NameOnCard") & "</TD><TD style=""text-align:center"">" & objRS("CardType") & "</TD>" & _
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

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		objRS.Open "SELECT TOP 6 * FROM qryCAPSApplicationSummary WITH(NOLOCK)",objCon
		'objRS.Open "SELECT TOP 20 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'Applications' ORDER BY FileSeqNum DESC",objCon

		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no Application data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""5"" style=""text-align:left"">Application Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
		        "<tr><th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th><th>Deleted</th></tr>" 
				
		    End If
		    
		    Do until objRS.eof
				
				x = x + 1
				
				Response.Write "<TR><TD><a href=""UploadApplications.asp?BatchNo=" & objRS("Applications") & """>" & objRS("Applications") & "</A></B></TD>" & _
							"<TD style=""text-align:center"">" & objRS("fldCardType") & "</TD><TD></TD><TD></TD><TD></TD></TR>"
							
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
					Response.Write "<b>ROMAN File Successfully Uploaded!!<b><br>"
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

		filePath = "D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\Attachments\Applications\Attachment1.xml"
		strFileName ="Attachment1.xml"

		'Call the procedure to read the file
		LoadXML filePath,strFileName
			
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

Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set outPut = objFSO.CreateTextFile("c:\\output.txt", true);
Set objTextFile = objFSO.OpenTextFile (strFileNamePath, ForReading)
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
			
			strFileSeqNum = GetFileLoadID("ROMAN",0,strFileName)
			'strFileSeqNum = GetFileLoadID("ROMAN",strFileSeqNum,strFileName)
			'''Set the File Number to 1 as the file will always be replaced by what is being loaded.
			'''Also avoid the check and always delete
			
			'Check to see if the same FileSeqNum for the same FileType has already been loaded
			lngFileID = GetFileLoadID("ROMANCostCentres","",strFileName)

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
						"<span>NOT LOADED! The ROMAN File Seq Num """ & strFileSeqNum & """ has already been loaded! <a href=""UploadApplications.asp?Reload=1&FileSeqNum=" & strFileSeqNum & """ Style=""font-weight:bold; color:white;""> Click here to Load anyway and overwrite the existing ROMAN file...</a></span></div></div></div>"
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
		lngFileLoadID = SaveFileLoadID ("ROMANCostCentres",strFileName,strFileNamePath,y,0,0,0,0,0,0,strFooterCount,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
		'response.write "fie=" & lngFileLoadID
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
		Call UpdateFileLoadSummary ("ROMANCostCentres",strFileSeqNum, strFileName,lngFileLoadID)
		'response.write "UpdateFileLoadSummary (""CSFRomDiners""," & strFileSeqNum & "," & lngFileLoadID & ")"
		
		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>ROMAN Cost Centre File Seq Num """ & strFileSeqNum & """ load COMPLETE!</span></div></div></div>"
						
						
	End If
	
'outPut.Close();

Set objFSO = Nothing
Set outPut = Nothing

End Sub


Public Sub LoadXML(filePath,strFileName)

Dim xmlDoc
Dim x
Dim Nodes

	Set xmlDoc=CreateObject("MSXML2.DOMDocument")
	'Set xmlDoc=CreateObject("Microsoft.XMLDOM")
		xmlDoc.async="false"
		xmlDoc.load(filePath)

Set Nodes = xmlDoc.selectNodes("//*")
	'response.write Nodes.length
	'response.write filePath
	
	for x = 0 to Nodes.length-1
		response.write nodes(x).nodeName
	next

	
'	for each x in xmlDoc.documentElement.childNodes
	'for each x in xmlDoc.nodes
'	  response.write(x.nodename)
'	  response.write(": ")
'	  response.write(x.text)
'	next


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


Set objRS = Nothing
Set objCon = Nothing

 %>


