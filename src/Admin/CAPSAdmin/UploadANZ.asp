<%@ Language=VBScript %>
<!-- #include file="../upload.asp" -->
<!-- #Include file=../../ADOVBS.inc -->
<!-- #include file="../../CC/CAPSFunctions.asp" -->
<%
'Description:	ANZ Cardlist Upload Administration screen
'Author:		Michael Giacomin
'Date:			May 2020

    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("UserID")) Then Response.Redirect("../../Timeout.asp")
	
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

If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
End If

'Response.Write "B=" & 	dteBatchDate
'dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
'Response.Write " 2B=" & 	dteBatchDate

'	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
'	Response.Write " 3B=" & 	dteBatchDateFormat
 %>

<html>
<head>
<link rel="stylesheet" type="text/css" href="../../BertStyle.css">
<!-- Bootstrap Core CSS -->
    <!--<link href="../css/bootstrap.min.css" rel="stylesheet">-->
	
	<!-- jQuery -->
    <script src="../../js/jquery.js"></script>
	
	  <!-- Custom fonts for this template-->
  <link href="../../vendor/font-awesome/css/font-awesome.min.css" rel="stylesheet" type="text/css">
  
  
  <!-- BEGIN: Vendor CSS-->
    <link rel="stylesheet" type="text/css" href="../../Frest/app-assets/vendors/css/vendors.min.css">
    <!-- END: Vendor CSS-->

    <!-- BEGIN: Theme CSS-->
    <link rel="stylesheet" type="text/css" href="../../Frest/app-assets/css/bootstrap.min.css">
    <link rel="stylesheet" type="text/css" href="../../Frest/app-assets/css/bootstrap-extended.min.css">
    <link rel="stylesheet" type="text/css" href="../../Frest/app-assets/css/colors.min.css">
    <link rel="stylesheet" type="text/css" href="../../Frest/app-assets/css/components.min.css">
    <link rel="stylesheet" type="text/css" href="../../Frest/app-assets/css/themes/dark-layout.min.css">
    <link rel="stylesheet" type="text/css" href="../../Frest/app-assets/css/themes/semi-dark-layout.min.css">
    <!-- END: Theme CSS-->

    <!-- BEGIN: Page CSS-->
    <link rel="stylesheet" type="text/css" href="../../Frest/app-assets/css/core/menu/menu-types/horizontal-menu.min.css">
    <!-- END: Page CSS-->



  <!-- BEGIN: Custom CSS-->
    <link rel="stylesheet" type="text/css" href="../../Frest/assets/css/style.css">
    <!-- END: Custom CSS-->
	
	
	
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

<form action="UploadANZ.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	 <div class="row" id="basic-table">
  <div class="col-6">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../../CC/img/ANZ2.png" height="30px" width="50px" title="ANZ"> ANZ Cardlist File Load</h4>
        </div>
      <div class="card-content">
        <div class="card-body">
		

<div class="col-lg-12 col-md-12">
<fieldset class="form-group">
	<label for="basicInputFile">Select a File to Load</label>
	<div class="custom-file">
		<input type="file" class="custom-file-input" NAME="FILE1" id="FILE1">
		<label class="custom-file-label" for="FILE1">Choose file</label>
	</div>
</fieldset>
</div>

<div class="form-body">
<div class="row col-11">
<div class="col-auto mr-auto">
	<div class="form-group">
	   <div class="checkbox">
		<input type="checkbox" class="checkbox-input" id="chkDelete" name="chkDelete">
		<label for="chkDelete">Overwrite Existing Batch</label>
	  </div>
	</div>
  </div>
  <div class="col-auto">
	<button type="button" class="btn btn-secondary btn-xs" onclick="upload();"><i class="fa fa-upload"></i> Upload</button>
  </div>
</div>
</div>
<p class="text-left">
<span id="Progress" style="display:none"><img src="../../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>

<font color="red" size="2"><b>NOTE: </Font><font color="black" size="2">When loading from Excel the Worksheet MUST be named 'Sheet1' (no spaces)
            <BR>* The file must be '.xls' only
            <BR>* Do not change the first row headers from the template files (below) or the ANZ provided workbook</B></Font>
</p>
<div class="col-lg-12 col-md-12">
<fieldset class="form-group">
    <button type="button" class="btn btn-secondary btn-xs" onclick="window.open('TemplateExcel.asp?T=tblCAPSANZCardlist')"><i class="fa fa-file"></i> ANZ Cardlist Template Excel </button>
</fieldset>
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

</body>
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
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no ANZ Cardlist data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Uplod Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing ANZ Cardlist Data already in CAPS2</th></tr>" & _
                "<tr><td colspan=""16"" style=""text-align:left; color:red;font-size:20px""><B>WARNING! The data below will be deleted if you Upload a new ANZ Cardlist file!</B></td></tr><tr>" & _
		        "<th Style=""width:20px;"">ApplicationID</th>" & _
				"<th>EmployeeID</th>" & _
				"<th>Card No</th><th>Card Type</th>" & _	
		        "<th>Title</th>" & _
	 	        "<th>FirstName</th>" & _	
	 	        "<th>Surname</th>" & _
		        "<th>Address1</th><th>Address2</th>" & _
                "<th>Address3</th><th>Suburb</th>" & _
                "<th>State</th><th>PostCode</th><th>EmailAddress</th>" & _
                "<th>Status</th><th>File Seq No</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR class='clickable-row' data-href='UploadANZDetail.asp?BatchNo=" & objRS("FileID") & "&EIDNo=" & objRS("EmployeeID") & "' style=""cursor: pointer;""><TD>" & objRS(0) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(1) & "</TD><TD style=""text-align:center"">" & objRS(5) & "</TD><TD style=""text-align:center"">" & objRS(3) & " " & objRS(4) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(6) & "</TD><TD style=""text-align:center"">" & objRS(7) & "</TD><TD style=""text-align:center"">" & objRS(8) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(9) & "</TD><TD style=""text-align:center"">" & objRS(10) & "</TD><TD style=""text-align:center"">" & objRS(11) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(12) & "</TD><TD style=""text-align:center"">" & objRS(13) & "</TD><TD style=""text-align:center"">" & objRS(14) & "</TD>" & _
								"<TD style=""text-align:center"">" & objRS(15) & "</TD><TD style=""text-align:center"">" & objRS(16) & "</TD><TD style=""text-align:center"">" & objRS("FileID") & "</TD>" & _
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

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		objRS.Open "SELECT TOP 20 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'ANZCardlist' ORDER BY FileSeqNum DESC",objCon
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no ANZ Cardlist data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">ANZ Cardlist Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
		        "<th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>Total Cards</th><th>Total Employees</th>" & _	
		        "<th>DTC</th>" & _
	 	        "<th>CMC</th>" & _	
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th></tr>" 
				
		    End If
		    
		    Do until objRS.eof
					
				Response.Write "<TR><TD><a href=""UploadANZ.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center"">" & objRS("RecordCount") & "</TD><TD style=""text-align:center"">" & objRS("CardCount") & "</TD><TD style=""text-align:center"">" & objRS("EmployeeCount") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("DTCCount") & "</TD><TD style=""text-align:center"">" & objRS("CMCCount") & "</TD><TD style=""text-align:center"">" & objRS("Status") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("DateLoaded") & "</TD></TR>"
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
				objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=YES"";"
				
				'objExcelCon.Open strUploadPath & "\CAPSCSFromDiners.dsn"
				
				'Write the SQL Query 
				objRS.open "SELECT * FROM [Sheet1$A2:S]", objExcelCon  
		    
				ReadExcel 
		    
				
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
Sub ReadExcel 

Dim errors
Dim strCardNumber
Dim strEID

    'First loop to check all the values
    Do until objRS.EOF 
        'Read each record present in the Excel file and check for the validation
 
        strCardNumber = objRS("Card Number") 'Should be integer and Not Null values       
        If IsNull(strFileSeqNum) Then
            errors = errors & "Error in line no. " & lineNo & ": Card Number should NOT be an Empty/Null value <br>" 
        End if
        
		strEID = objRS("Employee ID") 'Should be integer and Not Null values       
        If IsNull(strEID) Then
            errors = errors & "Error in line no. " & lineNo & ": Employee ID should NOT be an Empty/Null value <br>" 
        End if
       
	   strSurname = objRS("Cardholder Name") 'Should be integer and Not Null values       
        If IsNull(strSurname) Then
            errors = errors & "Error in line no. " & lineNo & ": Cardholder Name should NOT be an Empty/Null value <br>" 
        End if
                          
            lineNo = lineNo + 1            
            
        objRS.movenext 
    
        If IsNull(objRS("Relationship")) AND IsNull(objRS("Relationship")) Then       
        
            Exit Sub

        End If
    
    Loop        
    
End Sub	

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

    objRS.MoveFirst()
    
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
		
		If IsNull(strEmployeeID) Then               
		   
		   Exit Sub
		   
		End If
        
    Loop 
    
	
	If x > 1 Then
		'The ANZ Cardlist contains 2 header rows, so remove then from the count
		'If x > 2 Then x = x -2
		
			'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
			lngFileLoadID = SaveFileLoadID ("ANZCardlist",strFileName, strFilePath,x,0,0,0,0,0,0,0,strFileDateTime,lngFileID,"Imported",Session("UserID"),"N")
			
			'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			Call UpdateFileLoadSummary ("ANZCardlist",lngFileID, strFileName, lngFileLoadID)
			response.write "UpdateFileLoadSummary (""ANZCardlist""," & lngFileID & "," & strFileName & "," & lngFileLoadID & ")"
			
			Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
							"<span aria-hidden=""true"">&times;</span></button>" & _
							"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
							"<span>ANZ Cardlist File """ & strFileName & """ load COMPLETE!</span></div></div></div>"
							
	End If
	
End Sub	


Sub SaveRecord(lngANZCardlistID,strRelationship,strBillingAccount,strCardNumber,strStatus,strNameOnCard,strPhone,strCardType, _
					strExpiry,strEmployeeID,strOTCLimit,strTransactionLimit,strATMLimit,strCreditLimit,strAddress1,strAddress2,strAddress3, _
					strSuburb,strState,strPostCode,strDateLoaded,strUpdatedBy,strFileID, x)

Dim intRecord

  	With objCmd
  	
  	    'If the procedure has akready run then don't create the parameter objects again (more than once)
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
				.Parameters.Append objCmd.CreateParameter("OTCLimit", advarchar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("TransactionLimit", advarchar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("ATMLimit", advarchar, adParamInput,10)
				.Parameters.Append objCmd.CreateParameter("CreditLimit", advarchar, adParamInput,10)
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
				.Parameters("Loaded") = "N"'strLoaded
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute        
            
            'Return the result of the Save Function.
     		intRecord = objCmd.Parameters.Item("ANZCardlistIDOutput")    
     		                  			     				     		     		
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


Public Function GetNextANZFileID()

	'Description:	Gets the next FILE ID number for ANZ Cardlist files.
	objRS.Open "SELECT TOP 1 [FileSeqNum] FROM tblCAPSFileLoad WHERE FileType = 'ANZCardlist' AND [Deleted] = 'N' ORDER BY [FileSeqNum] DESC ",objCon

		If Not objRS.EOF Then
			GetNextANZFileID = objRS("FileSeqNum") + 1
		Else
			GetNextANZFileID = 1
		End If

	objRS.Close
	
End Function

Set objRS = Nothing
Set objCon = Nothing

 %>


