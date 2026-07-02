<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #include file="upload.asp" -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #include file="../CC/CAPSFunctions.asp" -->
<%
'Description:	ANZ Cardlist Upload Administration screen
'Author:		Michael Giacomin
'Date:			May 2020

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

If request.QueryString("Action")="Save" Then

	Call StartLoad()
End If

If Not IsEmpty(Request.QueryString("Reload")) Then

	Call StartLoad()
	
End If

'If the Process button has been clicked next to a file, then call the Process procedure
If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "Process" Then
		Call ProcessFile(Request.QueryString("FileLoadID"))
	End If
End If

If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
End If

'If the local load has been clicked then call the procedure to load the network file rather than uploading it
If request.QueryString("Action")="SaveFileLocal" Then
	strDeleteCheck = Request.QueryString("Delete")
	Call StartLoadLocal()
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
        if(getFileExt(document.getElementById('FILE1').value)==".xls")
        {
            if(window.confirm('This will overwrite any existing ANZ Cardlist file data! \n \n Continue?')==true)
                {
                document.getElementById('Progress').style.display = "inline";
                frm.submit();
            }
        } 
        else
        {
            alert("Please enter a valid Excel(.xls) file");
        }
    }
}

function UploadLocal()
{
	document.getElementById('Progress').style.display = "inline";
	var check = false //frm.chkDelete.checked

  if ( check==true) {
        // Returns true if checked
   	self.location="UploadANZ.asp?Action=SaveFileLocal&Delete=on"
    } else {
        // Returns false if not checked
	self.location="UploadANZ.asp?Action=SaveFileLocal&Delete=off"
    }
}

function UploadLocalG()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="UploadANZ.asp?Action=SaveFileLocal&ActionType=Service&Delete=off"
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
	self.location="UploadANZ.asp?FileDate=" + document.getElementById("CSDate").value;
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
<!--<div class="modal fade" id="ModApp" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
		<div class="text-center">
		  <div class="spinner-border" role="status">
			<span class="sr-only">Loading...</span>
		  </div>
		</div>
    </div>
  </div>
</div>-->

<!-- Modal -->
	<div class="loader" id="ModApp">
        <div class="wrap" >
            <div class="spinner"></div>
            <span class="loading-message">Loading...</h6>
        </div>
    </div>

	<div class="loader" id="SpinnerMod">
        <div class="wrap">
            <div class="spinner"></div>
            <span class="loading-message">Loading...</h6>
        </div>
    </div>

<main class="main py-3">
      <div class="container">
<form action="UploadANZ.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	 <div class="row" id="basic-table">
  <div class="col-3">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../CC/img/ANZ2.png" height="30px" width="50px" title="ANZ"> ANZ Cardlist File Load</h4>
        </div>
      <div class="card-content">
        <div class="card-body">
		

<div class="col-lg-12 col-md-12">

</div>

<div class="form-body">
<div class="row col-12" style="margin:5px; padding:0px;">

	<button type="button" class="btn btn-primary btn-sm" onclick="UploadLocalG();" Title="Click to Load any existing file in the G Drive Imports Folder"><i class="fa fa-upload"></i> Load G</button>&nbsp;
	<button type="button" class="btn btn-outline-secondary btn-sm" onclick="UploadLocal();" Title="Click to Load any existing file in the ANZ Folder"><i class="fa fa-upload"></i> Load Local</button>

</div>
</div>

<div class="col-lg-12 col-md-12">
<p class="text-left">
<div class="py-3"> 
<span id="Progress" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Processing...</b></span>
<br>
</div>
</p>
</div>

<div class="col-lg-12 col-md-12">

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
			<%DisplayFileSummaryG()%>	
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
	
		If Request.QueryString("BatchNo") = 0 Then strWhere = strWhere & " OR FileID = 0"
	End If
Else
	strWhere = ""
End If

	objRS.Open "SELECT TOP 50 * FROM tblCAPSANZCardlist WITH(NOLOCK) "  & strWhere,objCon
	
		    If objRS.eof Then
		        Response.Write"<table Class=""table table-bordered table-hover"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no ANZ Cardlist data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Uplod Data</td></tr>"
		    Else
		    
		         Response.Write "<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing ANZ Cardlist Data already in CAPS</th></tr>" & _
		        "<tr><th Style=""width:20px;"">ANZ File ID</th>" & _
				"<th>Account</th>" & _
				"<th>Name on Card</th><th>Card No</th>" & _	
		        "<th>Phone</th>" & _
	 	        "<th>Card Type</th>" & _	
	 	        "<th>Expiry Date</th>" & _
		        "<th>Employee ID</th><th>OTC Limit</th>" & _
                "<th>Transaction</th><th>ATM</th>" & _
                "<th>Credit</th><th>Address 1</th><th>Address 2</th>" & _
                "<th>Address 3</th><th>File Seq No</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR class='clickable-row' data-href='UploadANZDetail.asp?BatchNo=" & objRS("FileID") & "&EIDNo=" & objRS("EmployeeID") & "' style=""cursor: pointer;""><TD style=""text-align:center; font-size:12px;"">" & objRS(0) & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS(1) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS(5) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS(3) & " " & objRS(4) & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS(6) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS(7) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS(8) & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS(9) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS(10) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS(11) & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS(12) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS(13) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS(14) & "</TD>" & _
								"<TD style=""text-align:center; font-size:12px;"">" & objRS(15) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS(16) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("FileID") & "</TD>" & _
			                    "</TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close

        Response.Write "</table>"
		
End Sub

Sub DisplaySummary()

Dim dteBatchDateFormat
Dim strAction
Dim strStatus
Dim dteDateLoaded
Dim strRecordCount
Dim strCardCount
Dim strEmployeeCount
Dim strRecordsLoaded
Dim strView

Dim strDateLoadColour
Dim strDateTitle

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		objRS.Open "SELECT TOP 6 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'ANZCardlist' ORDER BY FileLoadID DESC",objCon
	
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table table-compact"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""9"" style=""text-align:left"">There is no ANZ Cardlist data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write "<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""9"" style=""text-align:left"">ANZ Cardlist Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
		        "<th Style=""width:20px; font-size:12px;"">Batch No.</th><th Style=""font-size:12px;"">File</th>" & _
				"<th Style=""font-size:12px;"">Total Cards</th><th Style=""font-size:12px;"">Total Employees</th>" & _	
		        "<th Style=""font-size:12px;"">No. Loaded</th>" & _
	 	        "<th Style=""font-size:12px;"">Action</th>" & _	
				"<th Style=""font-size:12px;"">Status</th>" & _
	 	        "<th Style=""font-size:12px;"">Date Loaded</th>" & _
				"<th Style=""font-size:12px;"">View</th></tr>"
				
		    End If
		    
		    Do until objRS.eof
				
				If IsNull(objRS("Status")) Then
					strStatus = ""
				Else
					strStatus = objRS("Status")
				End If
				
				If strStatus = "Imported" Then
					strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='UploadANZ.asp?Action=Process&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to Process the ANZ File and update changes to cards in CAPS from the ANZ File loaded " & objRS("DateLoaded") & """>Process</button>"
					'strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='UploadANZ.asp?Action=Process&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to Process the ANZ File and update changes to cards in CAPS from the ANZ File loaded " & objRS("DateLoaded") & """><i class=""fa fa-cogs""></i> Process</button>"
				Else
					strAction = ""
				End If
				
				If DateDiff("d",objRS("DateLoaded"),Now()) = 0 Then
					strDateLoadColour = " Color:Green; font-weight:bold;"
					strDateTitle = " - ROMAN File for today has been LOADED"
				Else
					strDateLoadColour = ""
					strDateTitle = ""
				End If
				
				If IsNull(objRS("DateLoaded")) Then
					dteDateLoaded = ""
				Else
					dteDateLoaded = FormatDateTime(objRS("DateLoaded"),vbShortDate)
				End If
				
				'Format all numbers
				'If IsNull(objRS("RecordCount")) Then
				'	strRecordCount = ""
				'Else
				'	strRecordCount = FormatNumber(objRS("RecordCount"),0)
				'End If
				
				If IsNull(objRS("CardCount")) Then
					strCardCount = ""
				Else
					strCardCount = FormatNumber(objRS("CardCount"),0)
				End If
				
				If IsNull(objRS("EmployeeCount")) Then
					strEmployeeCount = ""
				Else
					strEmployeeCount = FormatNumber(objRS("EmployeeCount"),0)
				End If
				
				If IsNull(objRS("RecordsLoaded")) Then
					strRecordsLoaded = ""
				Else
					strRecordsLoaded = FormatNumber(objRS("RecordsLoaded"),0)
				End If
				
				'Create the View button detail
				strView = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='ANZTransactions.asp?FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to view details of the ANZ File loaded " & objRS("DateLoaded") & """>View</button>"
				

				Response.Write "<TR><TD style=""text-align:right; font-size:12px;""><a href=""UploadANZ.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD><TD style=""text-align:right; font-size:12px;""><a href=""UploadANZ.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileName") & "</A></B></TD>" & _
							"<TD style=""text-align:right; font-size:12px;"">" & strCardCount & "</TD><TD style=""text-align:right; font-size:12px;"">" & strEmployeeCount & "</TD>" & _
							"<TD style=""text-align:right; font-size:12px;"">" & strRecordsLoaded & "</TD><TD style=""text-align:center"">" & strAction & "</TD><TD style=""text-align:center; font-size:12px;"">" & strStatus & "</TD>" & _
							"<TD style=""text-align:right; font-size:12px; " & strDateLoadColour & """ title=""" & objRS("DateLoaded") & strDateTitle & """>" & dteDateLoaded & "</TD><TD style=""text-align:center; font-size:12px;"">" & strView & "</TD></TR>"
			
							'Removed to make room and as it is not used
							'<TD style=""text-align:right"">" & strRecordCount & "</TD>
							
    			objRS.Movenext			
		    Loop
    			
			
								
	    objRS.Close

        Response.Write "</table>"
		
End Sub

Sub StartLoad()

'On Error Resume Next 
Dim objExcelCon
Dim strUploadPath

Dim errors
Dim lineNo
Dim strFileName
Dim lngFileID
	
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
	' Check if any file is successfully uploaded
	    If Uploader.Files.Count = 0 Then
		    Response.Write "No File(s) uploaded."
	    Else
			'Set the global (page) variable to the Check Delete value for use when processing data
			strDeleteCheck = Uploader.Form("chkDelete")
			If strFileStatus = "Imported" Then
				strDeleteCheck = "on"
			End If
		    ' Loop through the uploaded file
		    For Each File In Uploader.Files.Items					  		    		    
			    'set the upload path
	            strUploadPath = Server.MapPath(GetFilePath()) & "\Attachments"
				File.SaveToDisk strUploadPath
			    filePath = Server.MapPath(GetFilePath()) & "\Attachments\" & File.FileName
				strFileName = File.FileName
				strFileDateTime = Mid(strFileName,14,8)
		    Next
		    
			'Check to see if the same FileSeqNum for the same FileType has already been loaded
			lngFileID = GetFileLoadID("ANZCardlist","",strFileName)
			
			If lngFileID = "" Then
				'Get the next fileID Number for the ANZCardlist File
				lngFileID = GetNextANZFileID
			Else
				'If the checkbox to overwrite is checked then load the data, otherwise do not load
				If strDeleteCheck = "on" Then
					'Delete any existing CS From Diners Records
					objCon.Execute "DELETE FROM tblCAPSANZCardlist WHERE [FileID] = " & lngFileID & ""
					'Response.Write "DELETE FROM tblCAPSANZCardlist WHERE [FileID] = " & lngFileID & ""
				Else
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! The ANZ Cardlist File """ & strFileName & """ has already been loaded! <a href=""UploadANZ.asp?Reload=1&FileSeqNum=" & lngFileID & """ Style=""font-weight:bold; color:white;""> Check the 'Overwrite Existing Batch' box and load again to overwrite...</a></span></div></div></div>"
					Exit Sub
				End If
			End If
			
			'Call the relevant procedure depending on whether the file is .xls or .txt
			If Right(filePath,3) = "xls" Then
			
				'After uploading, Read excel file
				Set objExcelCon = Server.CreateObject("ADODB.connection")     
				'objExcelCon.Open "DBQ=" & filePath & "; DRIVER={Microsoft Excel Driver (*.xls)};" 
				'objExcelCon.Open "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & filePath & ";DefaultDir=c:\Apps\CAPS2\ASP2\Admin\CAPSAdmin\Attachments;" 
				objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"";",3,3
				
				'objExcelCon.Open strUploadPath & "\CAPSCSFromDiners.dsn"
				
				'Write the SQL Query 
				objRS.open "SELECT * FROM [deptofdefence$A2:S]", objExcelCon  
		    
				'Run the procedure to read the Excel File and save to the database
				errors = ReadExcel 
				
				
		    
				'Check for errors
				If(errors<>"") Then
					'Print the errors and return
					Response.Write "<font face=arial size=1><b>File not uploaded. Please correct the followings errors and load again</b> <br> "& errors &"</font><br>"         
				Else  
					'If the no errors found in the ReadExcel method, then start uploading the records in the database       	    	    
					UploadExcel Uploader.Form("chkDelete"),lngFileID, strFileName,filePath,strFileDateTime
					'Response.Write "<b>ANZ Cardlist File Sucessfully Uploaded!!<b><br>"
				End if
				
				
				
				'Close the recordset/connection 
				objRS.Close 
				objExcelCon.Close 
		    
			Else
				'ANZ file is xls as provided by ANZ and  not a text file, so do not enable text load
				'ReadText(filePath)
			End If
			
	    'End If
	End IF
	
End if

End Sub

'Function for validating the excel file values
Function ReadExcel 


Dim strCardNumber
Dim strEID
Dim lineNo
Dim strFileSeqNum
Dim strSurname

    'First loop to check all the values
    Do until objRS.EOF 
        'Read each record present in the Excel file and check for the validation
 
        strCardNumber = objRS("Card Number") 'Should be integer and Not Null values       
        If IsNull(strFileSeqNum) Then
            errors = errors & "Error in line no. " & lineNo & ": Card Number should NOT be an Empty/Null value <br>" 
			
        End if
        
		strEID = objRS("Employee ID") 'Should be integer and Not Null values     
		
        If IsNull(strEID) or IsEmpty(strEID) Then
			'Response.Write IsNumeric(strEID) & " : " & strEID & "</BR>"
            errors = errors & "Error in line no. " & lineNo & ": Employee ID should NOT be an Empty/Null value <br>" 
			
		End if
       
	   strSurname = objRS("Cardholder Name") 'Should be integer and Not Null values       
        If IsNull(strSurname) Then
            errors = errors & "Error in line no. " & lineNo & ": Cardholder Name should NOT be an Empty/Null value <br>"
					
        End if
                          
            lineNo = lineNo + 1            
            
        objRS.movenext 
    
        If IsNull(objRS("Relationship")) AND IsNull(objRS("Relationship")) Then       
					
            ReadExcel = errors
			Exit Function
        End If
    
    Loop       
		
    
End Function	

'Function for validating the excel file values
Sub UploadExcel(chkDelete,lngFileID, strFileName,strFilePath,strFileDateTime)

Dim x

Dim lngANZCardlistID
Dim strRelationship
Dim strBillingAccount
Dim strCardNumber
Dim strStatus
Dim strNameOnCard
Dim strPhone
Dim strCardType
Dim strExpiry
Dim strEmployeeID
Dim strOTCLimit
Dim strTransactionLimit
Dim strATMLimit
Dim strCreditLimit
Dim strAddress1
Dim strAddress2
Dim strAddress3
Dim strSuburb
Dim strState
Dim strPostCode
Dim strDateLoaded
Dim strUpdatedBy
Dim strFileID
Dim strLoaded
Dim lngFileLoadID
Dim objStartFolder
Dim objTextFile
Dim objFSO
Dim strFileNamePath
Dim intRecordCount


Set objFSO = CreateObject("Scripting.FileSystemObject")

	objRS.MoveLast
	
    objRS.MoveFirst()
    intRecordCount = objRS.RecordCount
		
    Do until objRS.EOF 
        
		x = x + 1
		
		lngANZCardlistID = 0'objRS("ANZCardlistID") 
		strRelationship = objRS(0)'objRS("Relationship") 
		strBillingAccount = objRS(1)'objRS("Billing Account ") 
		strCardNumber = objRS(2)'objRS("Card Number") 
		strStatus = objRS(3)'objRS("Card Status") 
		strNameOnCard = objRS(4)'objRS("Cardholder Name") 
		strPhone = objRS(5)'objRS("Cardholder Phone Number") 
		strCardType = objRS(6)'objRS("Card Type") 
		strExpiry = objRS(7)'objRS("EXP Date") 
		strEmployeeID = objRS(8)'objRS("Employee ID") 
		strOTCLimit = objRS(9)'objRS("OTC Limit") 
		strTransactionLimit = objRS(10)'objRS("Transaction Limit") 
		strATMLimit = objRS(11)'objRS("ATM Limit") 
		strCreditLimit = objRS(12)'objRS("Credit Limit") 
		strAddress1 = objRS(13)'objRS("Address Line 1") 
		strAddress2 = objRS(14)'objRS("Address Line 2") 
		strAddress3 = objRS(15)'objRS("Address Line 3") 
		strSuburb = objRS(16)'objRS("Suburb") 
		strState = objRS(17)'objRS("State") 
		strPostCode = objRS(18)'objRS("Post Code") 
		strDateLoaded = Now()'objRS("DateLoaded") 
		strUpdatedBy = Session("UserID")'objRS("UpdatedBy") 
		strFileID = lngFileID'objRS("FileID") 
		'strLoaded = "N"'objRS("Loaded") 


        'The first 2 rows of the ANZ Cardlist Excel file has an image (ANZ logo) and header which are to be ignored
		'If x < 3 Then
			
			
		'Else
            
            'Call the procedure to save the record to SQL
			
			'Response.Write  "exec spGeneralExpensesSave ="& strCSFromDinersID & "," & strFileDateTime & "," & strFileSeqNum & "," & x
			
            SaveRecord lngANZCardlistID,strRelationship,strBillingAccount,strCardNumber,strStatus,strNameOnCard,strPhone,strCardType, _
					strExpiry,strEmployeeID,strOTCLimit,strTransactionLimit,strATMLimit,strCreditLimit,strAddress1,strAddress2,strAddress3, _
					strSuburb,strState,strPostCode,strDateLoaded,strUpdatedBy,strFileID, x
            
        
        'End If
        
		objRS.movenext 
		
		'If IsNull(strEmployeeID) and IsNull(strCardNumber) Then   
		If x = intRecordCount Then					
				'If x > 1 Then
					'The ANZ Cardlist contains 2 header rows, so remove then from the count
					'If x > 2 Then x = x -2
					
					'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
					lngFileLoadID = SaveFileLoadID ("ANZCardlist",strFileName, strFilePath,x,0,0,0,0,0,0,0,strFileDateTime,lngFileID,"Imported",Session("UserID"),"N")
					
					'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
					Call UpdateFileLoadSummary ("ANZCardlist",lngFileID, strFileName, lngFileLoadID)
					'response.write "UpdateFileLoadSummary (""ANZCardlist""," & lngFileID & "," & strFileName & "," & lngFileLoadID & ")"
					
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
									"<span aria-hidden=""true"">&times;</span></button>" & _
									"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
									"<span>ANZ Cardlist File """ & strFileName & """ load COMPLETE! " & intRecordCount & " records loaded.</span></div></div></div>"
									
				
					
					
				'End If
		   
		   Exit Sub
		   
		End If
        
    Loop 	
	
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
Dim x
Dim strFileDateTime
Dim lngFileLoadID
Dim errors
Dim intSuccess
Dim strFileNamePath
Dim strFileStatus

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass
Dim objStartFolderArchive
		
	'Check that the Save/Upload has been clicked/passed in
	If Request.QueryString("Action")="SaveFileLocal" Then

		Set objNetwork = CreateObject("WScript.Network")
		Set objFSO = CreateObject("Scripting.FileSystemObject")

		'If Request.QueryString("Action")="SaveFileLocal" Then

			'''---Start the New Service Account Login section
			If Request.QueryString("ActionType")="Service" Then

				'Get the System Parameter for the start of the Training File Location
				strServer = GetSystemAdmin("GDriveFilePath")

				'Get the System Parameter for the Service Account UserName and Password
				strUser = GetSystemAdmin("CAPSServiceAccountName")
				strPass = GetSystemAdmin("CAPSServiceAccountPassword")

				'Get the System Parameter for the fileName
				'strFileNameDefault = GetSystemAdmin("CSFromDinersFileName")
							
				objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

				objStartFolder = strServer
		
			Else
				'Set objFSO = CreateObject("Scripting.FileSystemObject")

				'objStartFolder = "D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\ANZ"
				objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\ANZ\ANZFrom"
				
			'''End of the new Service account section
			End If
		
			Set objFolder = objFSO.GetFolder(objStartFolder)
			Set colFiles = objFolder.Files

			'Get the System Parameter for the fileName
			strFileNameDefault = GetSystemAdmin("ANZCardlistFileName")
			
			If IsNull(strFileNameDefault) or strFileNameDefault = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! There is no System Parameter for ANZ File Names (""ANZCardlistFileName""). See System Admin.</span></div></div></div>"
					Exit Sub
			End If
			
			For Each objFile in colFiles

				If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) Then
					strFileName = objFile.Name
					filePath = objStartFolder & "\" & strFileName
					
					'Response.Write objFile.Name & "</br>"
				End If
				
			Next
			
			strFileStatus = CheckANZFileStatus(strFileName)
						
			If strFileStatus = "Processed" Then	
			
				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
				"<span aria-hidden=""true"">&times;</span></button>" & _
				"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
				"<span>" & strFileNameDefault & " FILE HAS ALREADY BEEN PROCESSED! File Status needs to be changed to re-load file.</span></div></div></div>"
				Exit Sub
					
			Else
			
				If strFileStatus = "Imported" AND strDeleteCheck <> "on" Then				
		
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
					"<span aria-hidden=""true"">&times;</span></button>" & _
					"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
					"<span>" & strFileNameDefault & " FILE HAS ALREADY BEEN LOADED! To reload file ensure 'Overwrite Existing Batch' is checked.</span></div></div></div>"
					Exit Sub
					
				End If				
			
			End If

			If IsNull(strFileName) or strFileName = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>" & strFileNameDefault & " NOT LOADED! There is no ANZ File in the Server Folder to Load! Copy the 'deptofdefenceDDMMYYY.xls' file to the location " & objStartFolder & "</span></div></div></div>"
					Exit Sub
			End If

		'<!--------- start the load procedure, which shares similar code as the upload procedure above
		'Could be re-written to reduce duplicated code --->
		
		'Check to see if the same FileSeqNum for the same FileType has already been loaded
			lngFileID = GetFileLoadID("ANZCardlist","",strFileName)

			If lngFileID = "" Then
				'Get the next fileID Number for the ANZCardlist File
				lngFileID = GetNextANZFileID
			Else
				'If the checkbox to overwrite is checked then load the data, otherwise do not load
				If strDeleteCheck = "on" Then
					'Delete any existing CS From Diners Records
					objCon.Execute "DELETE FROM tblCAPSANZCardlist WHERE [FileID] = " & lngFileID & ""
				Else
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! The ANZ Cardlist File """ & strFileName & """ has already been loaded! <a href=""UploadANZ.asp?Reload=1&FileSeqNum=" & lngFileID & """ Style=""font-weight:bold; color:white;""> Check the 'Overwrite Existing Batch' box and load again to overwrite...</a></span></div></div></div>"
					Exit Sub
				End If
			End If
			
			'Call the relevant procedure depending on whether the file is .xls or .txt
			If Right(filePath,3) = "xls" Then
			
				'After uploading, Read excel file
				Set objExcelCon = Server.CreateObject("ADODB.connection")     
				'objExcelCon.Open "DBQ=" & filePath & "; DRIVER={Microsoft Excel Driver (*.xls)};"
				'objExcelCon.Open "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & filePath & ";DefaultDir=c:\Apps\CAPS2\ASP2\Admin\CAPSAdmin\Attachments;" 
				'objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"";"
				'objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"";",3,3
				'objExcelCon.Open "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & filepath & "; Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""

				'objExcelCon.Open strUploadPath & "\CAPSCSFromDiners.dsn"
				
				objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"";"
				
				'Write the SQL Query 
				objRS.CursorType = adOpenStatic
				On error resume next				
				objRS.open "SELECT * FROM [deptofdefence$A2:S]", objExcelCon
				
				'If there is an error with Opening the record set (if the sheet is not named correctly) then exit and display a message
				If err.Number <>0 then
					Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR with ANZ Upload File. The worksheet is NOT named 'deptofdefence'. Please notify System Administrators to change the ANZ file sheet name.</div>"
					exit sub
				End If
				'Call the procedure to check the Excel fields and data are correct
				errors = ReadExcel 
					    
				'Check for errors
				If(errors<>"") Then
					'Print the errors and return
					Response.Write "<font face=arial size=1><b>File not uploaded. Please correct the followings errors and load again</b> <br> "& errors &"</font><br>"         
				Else  
					strFileDateTime = Mid(strFileName,14,8)
					
					'If the no errors found in the ReadExcel method, then start uploading the records in the database    
					objCon.Execute "DELETE FROM tblCAPSANZCardlist WHERE [FileID] = " & lngFileID & ""
					UploadExcel strDeleteCheck,lngFileID, strFileName,filePath,strFileDateTime
					'UploadExcel Uploader.Form("chkDelete"),lngFileID, strFileName,filePath,strFileDateTime
					
					'Response.Write "<b>ANZ Cardlist File Sucessfully Uploaded!!<b><br>"
					intSuccess = 1
					
				End If				
				
				'Close the recordset/connection 
				objRS.Close 
				objExcelCon.Close 
				
				If intSuccess = 1 Then
				
					'''---Start the New Service Account Login section
					If Request.QueryString("ActionType")="Service" Then
					
						objStartFolderArchive = GetSystemAdmin("GDriveArchiveFilePath") & "Imports\ANZ Cardlist\"
						strFileNamePath = objStartFolder & strFileName
						
						'Delete any files with the same name before moving
						Call DeleteExistingFile(objStartFolderArchive,strFileName)
						
						objFSO.MoveFile strFileNamePath,objStartFolderArchive & strFileName	
						
					Else
						objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\ANZ\"
						strFileNamePath = objStartFolder & "\ANZFrom\" & strFileName				
						objFSO.MoveFile strFileNamePath,objStartFolder & "Loaded\" & strFileName 
					End If
				End If
		    
			Else
				'ANZ file is xls as provided by ANZ and  not a text file, so do not enable text load
				'ReadText(filePath)
			End If
			
			'<!--------- End of duplicated code --->
		
	
		
		If x > 1 Then
			'The ANZ Cardlist contains 2 header rows, so remove then from the count
			If x > 2 Then x = x -2
			
			'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
			lngFileLoadID = SaveFileLoadID ("ANZCardlist",strFileName, filePath,x,0,0,0,0,0,0,0,strFileDateTime,lngFileID,"Imported",Session("UserID"),"N")
			
			'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			Call UpdateFileLoadSummary ("ANZCardlist",lngFileID, strFileName, lngFileLoadID)
			'response.write "UpdateFileLoadSummary (""ANZCardlist""," & lngFileID & "," & strFileName & "," & lngFileLoadID & ")"
			
			'Finally update the Active field in tblCAPSANZCardlist so that the FileID loaded is the only active file
			Call UpdateActive(lngFileLoadID)
			
			Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
							"<span aria-hidden=""true"">&times;</span></button>" & _
							"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
							"<span>ANZ Cardlist File """ & strFileName & """ load COMPLETE!</span></div></div></div>"
						
		End If
	
		'''---Start the New Service Account Login section
		If Request.QueryString("ActionType")="Service" Then
			
			objNetwork.RemoveNetworkDrive strServer, True, False
			 
			Set objFSO = Nothing
			Set objNetwork = Nothing
		End If
	
	
	'END of Check that the Save/Upload has been clicked/passed in
	End if

End Sub


Sub SaveRecord(lngANZCardlistID,strRelationship,strBillingAccount,strCardNumber,strStatus,strNameOnCard,strPhone,strCardType, _
					strExpiry,strEmployeeID,strOTCLimit,strTransactionLimit,strATMLimit,strCreditLimit,strAddress1,strAddress2,strAddress3, _
					strSuburb,strState,strPostCode,strDateLoaded,strUpdatedBy,strFileID, x)

Dim intRecord

'Make sure that the values being loaded are not too long (varchar(20))
If IsNull(strTransactionLimit) Then strTransactionLimit = ""
strTransactionLimit = trim(strTransactionLimit)
If Len(strTransactionLimit) > 10 then strTransactionLimit = left(strTransactionLimit,10)


  	With objCmd
  	
  	    'If the procedure has already run then don't create the parameter objects again (more than once)
  	    If x = 1 then
                .CommandType = 4
                .CommandText = "spCAPSANZCardlistSave"
                
				.Parameters.Append objCmd.CreateParameter("ANZCardlistID", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("Relationship", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("BillingAccount", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("CardNumber", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("Status", advarchar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("NameOnCard", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("Phone", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("CardType", advarchar, adParamInput,50)
				.Parameters.Append objCmd.CreateParameter("Expiry", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("EmployeeID", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("OTCLimit", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("TransactionLimit", advarchar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("ATMLimit", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("CreditLimit", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("Address1", advarchar, adParamInput,200)
				.Parameters.Append objCmd.CreateParameter("Address2", advarchar, adParamInput,200)
				.Parameters.Append objCmd.CreateParameter("Address3", advarchar, adParamInput,200)
				.Parameters.Append objCmd.CreateParameter("Suburb", advarchar, adParamInput,100)
				.Parameters.Append objCmd.CreateParameter("State", advarchar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("PostCode", advarchar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("DateLoaded", adDate, adParamInput)
				.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("FileID", adInteger, adParamInput)
				.Parameters.Append objCmd.CreateParameter("Loaded", adChar, adParamInput,1) 
				.Parameters.Append objCmd.CreateParameter("ProcessStatus", adVarChar, adParamInput,20)	
				.Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput,1)				
				.Parameters.Append objCmd.CreateParameter("ANZCardlistIDOutput", adInteger, adParamOutput)				
            
        End If
                 
				.Parameters("ANZCardlistID") = lngANZCardlistID
				.Parameters("Relationship") = strRelationship
				.Parameters("BillingAccount") = strBillingAccount
				.Parameters("CardNumber") = strCardNumber
				.Parameters("Status") = strStatus
				.Parameters("NameOnCard") = strNameOnCard
				.Parameters("Phone") = strPhone
				.Parameters("CardType") = strCardType
				.Parameters("Expiry") = strExpiry
				.Parameters("EmployeeID") = strEmployeeID
				.Parameters("OTCLimit") = strOTCLimit
				.Parameters("TransactionLimit") = strTransactionLimit
				.Parameters("ATMLimit") = strATMLimit
				'
				'If IsNull(strCreditLimit) Then strCreditLimit = 0 End If
				.Parameters("CreditLimit") = strCreditLimit
				.Parameters("Address1") = strAddress1
				.Parameters("Address2") = strAddress2
				.Parameters("Address3") = strAddress3
				.Parameters("Suburb") = strSuburb
				.Parameters("State") = strState
				.Parameters("PostCode") = strPostCode
				.Parameters("DateLoaded") = strDateLoaded
				.Parameters("UpdatedBy") = strUpdatedBy
				.Parameters("FileID") = strFileID
				.Parameters("ProcessStatus") = "Loaded"
				.Parameters("Loaded") = "N"'strLoaded
                .Parameters("Active") = "Y"
				
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute        
            
            'Return the result of the Save Function.
     		intRecord = objCmd.Parameters.Item("ANZCardlistIDOutput")    
     		                  			     				     		     		
       'response.write  "exec spGeneralExpensesSave =0," & Session("BudgetID") & "," & Session("VersionID") & "," & CostCentreID & ",'GEXP'" & GLCode & "," & BM1 & "," & BM2 & "," & BM3 & "," & BM4 & "," & BM5 & "," & _
       '                     BM6 & "," & BM7 & "," & BM8 & "," & BM9 & "," & BM10 & "," & BM11 & "," & BM12 & "," & OY1 & "," & OY2 & "," & OY3 & ",'" & Comments & "','" & UpdatedBy & "'," & Session("ColumnLock")
End Sub


Public Function ProcessFile(strFileID)
'Function to Process an ANZ File which has been loaded into the database (update all changes from the ANZ file and add changes to the audit log)
Dim intRecord

	With objCmd

		.CommandType = 4
		.CommandText = "spCAPSANZProcessCardlist"

		.Parameters.Append objCmd.CreateParameter("UserID", adInteger)
		.Parameters.Append objCmd.CreateParameter("FileID", adInteger)
		.Parameters.Append objCmd.CreateParameter("CDMCProcessOutput", adInteger, adParamOutput)
		
		.Parameters("UserID") = Session("UserID")
		.Parameters("FileID") = strFileID
		
		.ActiveConnection = objCon
		 
	End With
	
	Response.Write strFileID
   
	objCmd.Execute        
  
	'Return the result of the Save Function.
	intRecord = objCmd.Parameters.Item("CDMCProcessOutput") 

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Records Processed in ANZ Upload File " & strFileID & ". Please notify System Administrators.</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert""> " & intRecord & " ANZ Upload File records Processed in file " & strFileID & "</div>"
	End If
	
	ProcessFile = intRecord
	
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

	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPSDev\ASPNew\Admin\CAPSAdmin\Attachments\ANZ\ANZFrom")
	objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\ANZ\ANZFrom"

	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be loaded. Click 'Load Server' button to Load ANZ Cardlist File.""></i></th></tr>"
	
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
strFileNameDefault = GetSystemAdmin("ANZCardlistFileName")
			
objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

	objStartFolder = strServer
	
	'objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Diners\DinersFrom"

	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
'Set objFSO = CreateObject("Scripting.FileSystemObject")

	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPSDev\ASPNew\Admin\CAPSAdmin\Attachments\ANZ\ANZFrom")
	'objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\ANZ\ANZFrom"

	'Set objFolder = objFSO.GetFolder(objStartFolder)
		
	'Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">G Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the G Drive waiting to be loaded. Click 'Load ANZ G' button to Load ANZ Cardlist File.""></i></th></tr>"
	
	intCount = 0
	
	For Each objFile in colFiles
		
		'only display the files with the same name as the System Parameter
		If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) Then
		
			intCount = intCount + 1
			
			If intCount < 6 Then
				If IsNull(objFile.Name) or objFile.Name = "" Then
					strFile = ""
				Else
					strFile = Left(objFile.Name,10)
				End If
				
				Response.Write "<TR><TD>" & strFile & "...xls</TD></TR>"
			End If
		End If
	Next
	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"

	'Close the mapped network drive
	objNetwork.RemoveNetworkDrive strServer, True, False
	 
Set objFSO = Nothing
Set objNetwork = Nothing

'Set objFSO = Nothing

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
    Err.Clear
End If


Public Function GetNextANZFileID()

	'Description:	Gets the next FILE ID number for ANZ Cardlist files.
	objRS.Open "SELECT TOP 1 [FileSeqNum] FROM tblCAPSFileLoad WHERE FileType = 'ANZCardlist' AND [Deleted] = 'N' ORDER BY [FileLoadID] DESC ",objCon

		If Not objRS.EOF Then
			GetNextANZFileID = objRS("FileSeqNum") + 1
		Else
			GetNextANZFileID = 1
		End If

	objRS.Close
	
End Function

Public Sub UpdateActive(lngFileID)
'Procedure to update the Active field in the table tblCAPSANZCardlist so only the current fileID is active
Dim strSQL

	'If the Batch Number is a number then update the System Parameter, otherwise post an error to the screen
	If IsNumeric(lngFileID) Then
		
		strSQL = "UPDATE tblCAPSANZCardlist SET [Active] = 'N' WHERE [FileID] <> " & lngFileID & " AND [Active] = 'Y'"

		objCon.Execute strSQL
	
	Else
		
		Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR! ANZ Card List File ID: " & lngFileID & " is not a number. See System Admin.</div>"
		
	End If

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
			'Delete any files already in the folder passed in (Archive) for the current user
			objFSO.DeleteFile strFileName, True
				
		End If
	Next
	
End Sub



Set objRS = Nothing
Set objCon = Nothing

 %>


