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
	self.location="UploadCDMC.asp?FileDate=" + document.getElementById("CSDate").value;
}
</script>

<style>

    table.newd th, table.newd td{

        padding: 4px; 

    }

</style>
</head>
<body>

<form action="UploadCDMC.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	 <div class="row" id="basic-table">
  <div class="col-6">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../../CC/img/DFG_Logo.png" height="40px" width="100px" title="Certificate and Directory Management Centre"> CDMC File Load</h4>
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
		<label for="chkDelete">Overwrite Existing File</label>
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
            <BR>* The file must be '.csv' only
            <BR>* Do not change the first row headers from the template files (below) or the CDMC provided CSV</B></Font>
</p>
<div class="col-lg-12 col-md-12">
<fieldset class="form-group">
    <button type="button" class="btn btn-secondary btn-xs" onclick="window.open('TemplateExcel.asp?T=tblCAPSANZCardlist')"><i class="fa fa-file"></i> CDMC Template Excel </button>
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

	objRS.Open "SELECT TOP 50 * FROM tblCAPSCDMC WITH(NOLOCK) "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no CDMC data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Uplod Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing CDMC Data already in CAPS</th></tr>" & _
                "<tr><td colspan=""16"" style=""text-align:left; color:red;font-size:20px""><B>WARNING! The data below will be deleted if you Upload a new CDMC file!</B></td></tr><tr>" & _
		        "<th Style=""width:20px;"">CDMCID</th>" & _
				"<th>EmployeeID</th>" & _
				"<th>Group</th><th>Rank</th>" & _	
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
			    Response.Write "<TR class='clickable-row' data-href='UploadCDMCDetail.asp?BatchNo=" & objRS("FileID") & "&EIDNo=" & objRS("EmployeeID") & "' style=""cursor: pointer;""><TD>" & objRS(0) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("EmployeeID") & "</TD><TD style=""text-align:center"">" & objRS("GroupName") & "</TD><TD style=""text-align:center"">" & objRS("ActualRankLvl") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("Title") & "</TD><TD style=""text-align:center"">" & objRS("FormalFirstName") & "</TD><TD style=""text-align:center"">" & objRS("FormalLastName") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("PostalAddress_Unit") & "</TD><TD style=""text-align:center"">" & objRS("PostalAddress_ClientLocation") & "</TD><TD style=""text-align:center"">" & objRS("PostalAddress_DeliveryLocation") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("PostalAddress_City") & "</TD><TD style=""text-align:center"">" & objRS("PostalAddress_State") & "</TD><TD style=""text-align:center"">" & objRS("PostalAddress_PostCode") & "</TD>" & _
								"<TD style=""text-align:center"">" & objRS("Email_Address") & "</TD><TD style=""text-align:center"">" & objRS("hasChanged") & "</TD><TD style=""text-align:center"">" & objRS("FileID") & "</TD>" & _
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

		objRS.Open "SELECT TOP 20 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'CDMCFile' ORDER BY FileSeqNum DESC",objCon
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no CDMC data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""6"" style=""text-align:left"">CDMC Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
		        "<tr><th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>No Address</th><th>Total Employees</th>" & _	
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th></tr>" 
				
		    End If
		    
		    Do until objRS.eof
					
				Response.Write "<TR><TD><a href=""UploadCDMC.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center"">" & objRS("RecordCount") & "</TD><TD style=""text-align:center"">" & objRS("CardCount") & "</TD><TD style=""text-align:center"">" & objRS("EmployeeCount") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("Status") & "</TD>" & _
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
				strFileDateTime = now()'Mid(strFileName,14,8)
		    Next
		    
			'Check to see if the same FileSeqNum for the same FileType has already been loaded
			lngFileID = GetFileLoadID("CDMCFile","",strFileName)
			
			If lngFileID = "" Then
				'Get the next fileID Number for the ANZCardlist File
				lngFileID = GetNextCDMCFileID
			Else
				'If the checkbox to overwrite is checked then load the data, otherwise do not load
				If strDeleteCheck = "on" Then
					'Delete any existing CS From Diners Records
					objCon.Execute "DELETE FROM tblCAPSCDMC WHERE [FileID] = " & lngFileID & ""
				Else
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! The CDMC File """ & strFileName & """ has already been loaded! <a href=""UploadCDMC.asp?Reload=1&FileSeqNum=" & lngFileID & """ Style=""font-weight:bold; color:white;""> Check the 'Overwrite Existing Batch' box and load again to overwrite...</a></span></div></div></div>"
					Exit Sub
				End If
			End If
			
			'Call the relevant procedure depending on whether the file is .xls or .txt
			If Right(filePath,3) = "csv" or Right(filePath,3) = "txt" then
			
				'After uploading, Read excel file
				Set objExcelCon = Server.CreateObject("ADODB.connection")     
				'objExcelCon.Open "DBQ=" & filePath & "; DRIVER={Microsoft Excel Driver (*.xls)};" 
				'objExcelCon.Open "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & filePath & ";DefaultDir=c:\Apps\CAPS2\ASP2\Admin\CAPSAdmin\Attachments;" 
		'		objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strUploadPath & "\" & ";Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
				'objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""text;HDR=Yes;FMT=Delimited(~)"";"
				
				objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strUploadPath & "\" & ";Extended Properties=""text;"";"
				
				'objExcelCon.Open strUploadPath & "\CAPSCSFromDiners.dsn"
				
				'Write the SQL Query 
				objRS.open "SELECT * FROM [" & strFileName & "]", objExcelCon  
		    
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
 
        'strCardNumber = objRS("Card Number") 'Should be integer and Not Null values       
        'If IsNull(strFileSeqNum) Then
        '    errors = errors & "Error in line no. " & lineNo & ": Card Number should NOT be an Empty/Null value <br>" 
        'End if

		strEID = objRS("EmployeeNumber") 'Should be integer and Not Null values       
        If IsNull(strEID) Then
            errors = errors & "Error in line no. " & lineNo & ": Employee ID should NOT be an Empty/Null value <br>" 
        End if
       
	   'strSurname = objRS("Cardholder Name") 'Should be integer and Not Null values       
        'If IsNull(strSurname) Then
        '    errors = errors & "Error in line no. " & lineNo & ": Cardholder Name should NOT be an Empty/Null value <br>" 
        'End if
                          
            lineNo = lineNo + 1            
            
        objRS.movenext 
    
        If IsNull(objRS("EmployeeNumber")) AND IsNull(objRS("EmployeeNumber")) Then       
        
            Exit Sub

        End If
    
    Loop        
    
End Sub	

'Function for validating the excel file values
Sub UploadExcel(chkDelete,lngFileID, strFileName,strFilePath,strFileDateTime)

Dim x

Dim lngCDMCID
Dim strGroupName
Dim strDivisionName
Dim strBranchName
Dim strDepartmentName
Dim strDepartmentNumber
Dim strCostCentre
Dim strEmployeeID
Dim strEmployeeType
Dim strFirstname
Dim strSurname
Dim strTitle
Dim strEmail_Address
Dim strTelephoneNumber
Dim strMobileNumber
Dim strDateofBirth
Dim strGender
Dim strActualRankLvl
Dim strSite
Dim strUnit
Dim strReportsTo
Dim strDCD_PostalAddress
Dim straddressline1
Dim straddressline2
Dim straddressline3
Dim straddressline4
Dim straddressline5
Dim straddressline6
Dim strPostalAddress_Unit
Dim strPostalAddress_ClientLocation
Dim strPostalAddress_DeliveryLocation
Dim strPostalAddress_City
Dim strPostalAddress_State
Dim strPostalAddress_PostCode
Dim strPostalAddress_Country
Dim strDCDProtectedIdentity
Dim strIsValidPostal
Dim strOutAddr1
Dim strOutAddr2
Dim strOutAddr3
Dim strOutSuburb
Dim strOutState
Dim strOutPostCode
Dim strPostalMessage
Dim strhasChanged
Dim strDCD_WorkAddress
Dim strClientLocation
Dim strStreetAddress
Dim strCity
Dim strState
Dim strPostCode
Dim strFormalFirstName
Dim strFormalLastName
Dim strFormalMiddleName
Dim strOutTitle
Dim strOutDinersWorkPhone
Dim strOutDinersMobilePhone
Dim strOutANZPhone
Dim strOutDinersAddress1
Dim strOutDinersAddress2
Dim strRemoveCountdown
Dim strFirstUpdated
Dim strLastUpdated
Dim strActive
Dim strUpdatedBy
Dim strDateUpdated
Dim strFileID
Dim strLoaded

    objRS.MoveFirst()
    
    Do until objRS.EOF 
        
		x = x + 1
		
		lngCDMCID = 0'objRS("CDMCID") 
		strGroupName = objRS("GroupName") 
		strDivisionName = objRS("DivisionName") 
		strBranchName = objRS("BranchName") 
		strDepartmentName = objRS("DepartmentName") 
		strDepartmentNumber = objRS("DepartmentNumber") 
		strCostCentre = objRS("CostCentre") 
		strEmployeeID = objRS("EmployeeNumber") 
		strEmployeeType = objRS("EmployeeType") 
		strFirstname = objRS("GivenName") 
		strSurname = objRS("Surname") 
		strTitle = objRS("Title") 
		strEmail_Address = objRS("Email_Address") 
		strTelephoneNumber = objRS("TelephoneNumber") 
		strMobileNumber = objRS("MobileNumber") 
		strDateofBirth = objRS("DateofBirth") 
		strGender = objRS("Gender") 
		strActualRankLvl = objRS("ActualRankLv1")
		strSite = objRS("Site")
		strUnit = objRS("Unit")
		strReportsTo = objRS("ReportsTo")
		strDCD_PostalAddress = objRS("DCD_PostalAddress")		
		straddressline1 = objRS("addressline1") 
		straddressline2 = objRS("addressline2") 
		straddressline3 = objRS("addressline3") 
		straddressline4 = objRS("addressline4") 
		straddressline5 = objRS("addressline5") 
		straddressline6 = objRS("addressline6") 
		strPostalAddress_Unit = objRS("PostalAddress_Unit") 
		strPostalAddress_ClientLocation = objRS("PostalAddress_ClientLocation") 
		strPostalAddress_DeliveryLocation = objRS("PostalAddress_DeliveryLocation") 
		strPostalAddress_City = objRS("PostalAddress_City") 
		strPostalAddress_State = objRS("PostalAddress_State") 
		strPostalAddress_PostCode = objRS("PostalAddress_PostCode") 
		strPostalAddress_Country = objRS("PostalAddress_Country") 
		strDCDProtectedIdentity = objRS("DCDProtectedIdentity") 
		strIsValidPostal = ""'""'objRS("IsValidPostal") 
		strOutAddr1 = ""'objRS("OutAddr1") 
		strOutAddr2 = ""'objRS("OutAddr2") 
		strOutAddr3 = ""'objRS("OutAddr3") 
		strOutSuburb = ""'objRS("OutSuburb") 
		strOutState = ""'objRS("OutState") 
		strOutPostCode = ""'objRS("OutPostCode") 
		strPostalMessage = ""'objRS("PostalMessage") 
		strhasChanged = "N"'objRS("hasChanged") 
		strDCD_WorkAddress = objRS("DCD_WorkAddress") 
		strClientLocation = objRS("ClientLocation") 
		strStreetAddress = objRS("StreetAddress") 
		strCity = objRS("City") 
		strState = objRS("State") 
		strPostCode = objRS("PostCode") 
		strFormalFirstName = objRS("FirstName") 
		strFormalLastName = objRS("LastName") 
		strFormalMiddleName = objRS("MiddleName") 
		strOutTitle = ""'objRS("OutTitle") 
		strOutDinersWorkPhone = ""'objRS("OutDinersWorkPhone") 
		strOutDinersMobilePhone = ""'objRS("OutDinersMobilePhone") 
		strOutANZPhone = ""'objRS("OutANZPhone") 
		strOutDinersAddress1 = ""'objRS("OutDinersAddress1") 
		strOutDinersAddress2 = ""'objRS("OutDinersAddress2") 
		strRemoveCountdown = 0'objRS("RemoveCountdown") 
		strFirstUpdated = "NULL"'objRS("FirstUpdated") 
		strLastUpdated = "NULL"'objRS("LastUpdated") 
		strActive = ""'objRS("Active") 
		strDateLoaded = Now()'objRS("DateLoaded") 
		strUpdatedBy = Session("UserID")'objRS("UpdatedBy") 
		strFileID = lngFileID'objRS("FileID") 
		strLoaded = "N"'objRS("Loaded") 


        'The first 2 rows of the ANZ Cardlist Excel file has an image (ANZ logo) and header which are to be ignored
		'If x < 3 Then
			
			
		'Else
            
            'Call the procedure to save the record to SQL
			
			'Response.Write  "exec spGeneralExpensesSave ="& strCSFromDinersID & "," & strFileDateTime & "," & strFileSeqNum & "," & x
			
            SaveRecord lngCDMCID,strGroupName,strDivisionName,strBranchName,strDepartmentName,strDepartmentNumber,strCostCentre,strEmployeeID,strEmployeeType, _
					strFirstname,strSurname,strTitle,strEmail_Address,strTelephoneNumber,strMobileNumber,strDateofBirth,strGender,strActualRankLvl,strSite,strUnit,strReportsTo,strDCD_PostalAddress, _
					straddressline1,straddressline2,straddressline3,straddressline4,straddressline5,straddressline6,strPostalAddress_Unit,strPostalAddress_ClientLocation, _
					strPostalAddress_DeliveryLocation,strPostalAddress_City,strPostalAddress_State,strPostalAddress_PostCode,strPostalAddress_Country,strDCDProtectedIdentity, _
					strIsValidPostal,strOutAddr1,strOutAddr2,strOutAddr3,strOutSuburb,strOutState,strOutPostCode,strPostalMessage,strhasChanged,strDCD_WorkAddress,strClientLocation, _
					strStreetAddress,strCity,strState,strPostCode,strFormalFirstName,strFormalLastName,strFormalMiddleName,strOutTitle,strOutDinersWorkPhone,strOutDinersMobilePhone, _
					strOutANZPhone,strOutDinersAddress1,strOutDinersAddress2,strRemoveCountdown,strFirstUpdated,strLastUpdated,strActive,strUpdatedBy,strDateUpdated,strFileID,strLoaded, x
            
        
        'End If
        
		objRS.movenext 
		
		If IsNull(strEmployeeID) Then               
		   
		   Exit Sub
		   
		End If
        
    Loop 
    
	
	If x > 1 Then
		
		'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
		lngFileLoadID = SaveFileLoadID ("CDMCFile",strFileName, strFilePath,x,0,0,0,0,0,0,0,strFileDateTime,lngFileID,"Imported",Session("UserID"),"N")
		
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
		Call UpdateFileLoadSummary ("CDMCFile",lngFileID, strFileName, lngFileLoadID)
		response.write "UpdateFileLoadSummary (""CDMCFile""," & lngFileID & "," & strFileName & "," & lngFileLoadID & ")"
		
		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>CDMC File """ & strFileName & """ load COMPLETE!</span></div></div></div>"
					
	End If
	
End Sub	


Sub SaveRecord(lngCDMCID,strGroupName,strDivisionName,strBranchName,strDepartmentName,strDepartmentNumber,strCostCentre,strEmployeeID,strEmployeeType, _
			strFirstname,strSurname,strTitle,strEmail_Address,strTelephoneNumber,strMobileNumber,strDateofBirth,strGender,strActualRankLvl,strSite,strUnit,strReportsTo,strDCD_PostalAddress, _
			straddressline1,straddressline2,straddressline3,straddressline4,straddressline5,straddressline6,strPostalAddress_Unit,strPostalAddress_ClientLocation, _
			strPostalAddress_DeliveryLocation,strPostalAddress_City,strPostalAddress_State,strPostalAddress_PostCode,strPostalAddress_Country,strDCDProtectedIdentity, _
			strIsValidPostal,strOutAddr1,strOutAddr2,strOutAddr3,strOutSuburb,strOutState,strOutPostCode,strPostalMessage,strhasChanged,strDCD_WorkAddress,strClientLocation, _
			strStreetAddress,strCity,strState,strPostCode,strFormalFirstName,strFormalLastName,strFormalMiddleName,strOutTitle,strOutDinersWorkPhone,strOutDinersMobilePhone, _
			strOutANZPhone,strOutDinersAddress1,strOutDinersAddress2,strRemoveCountdown,strFirstUpdated,strLastUpdated,strActive,strUpdatedBy,strDateUpdated,strFileID,strLoaded, x)

Dim intRecord

  	With objCmd
  	
  	    'If the procedure has akready run then don't create the parameter objects again (more than once)
  	    If x = 1 then
			.CommandType = 4
			.CommandText = "spCAPSCDMCSave"
			
			.Parameters.Append objCmd.CreateParameter("CDMCID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("GroupName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("DivisionName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("BranchName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("DepartmentName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("DepartmentNumber", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("CostCentre", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarWChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("EmployeeType", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("Firstname", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("Surname", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("Title", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("Email_Address", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("TelephoneNumber", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("MobileNumber", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("DateofBirth", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("Gender", adVarWChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("ActualRankLvl", adVarWChar, adParamInput,30)
			.Parameters.Append objCmd.CreateParameter("Site", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("Unit", adVarWChar, adParamInput,200)
			.Parameters.Append objCmd.CreateParameter("ReportsTo", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("DCD_PostalAddress", adVarWChar, adParamInput,500)
			.Parameters.Append objCmd.CreateParameter("addressline1", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline2", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline3", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline4", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline5", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addressline6", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_Unit", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_ClientLocation", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_DeliveryLocation", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_City", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_State", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_PostCode", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("PostalAddress_Country", adVarWChar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("ClientLocation", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("StreetAddress", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("City", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("State", adVarWChar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("PostCode", adVarWChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("DCDProtectedIdentity", adVarWChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("IsValidPostal", adVarWChar, adParamInput,3)
			.Parameters.Append objCmd.CreateParameter("OutAddr1", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutAddr2", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutAddr3", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutSuburb", adVarWChar, adParamInput,22)
			.Parameters.Append objCmd.CreateParameter("OutState", adVarWChar, adParamInput,3)
			.Parameters.Append objCmd.CreateParameter("OutPostCode", adVarWChar, adParamInput,4)
			.Parameters.Append objCmd.CreateParameter("PostalMessage", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("hasChanged", adVarWChar, adParamInput,3)
			.Parameters.Append objCmd.CreateParameter("FormalFirstName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("FormalLastName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("FormalMiddleName", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutTitle", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutDinersWorkPhone", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutDinersMobilePhone", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutANZPhone", adVarWChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("OutDinersAddress1", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("OutDinersAddress2", adVarWChar, adParamInput,40)
			.Parameters.Append objCmd.CreateParameter("RemoveCountdown", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("FirstUpdated", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("LastUpdated", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Active", adChar, adParamInput,1)

			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("FileID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Loaded", adChar, adParamInput,1)    
			.Parameters.Append objCmd.CreateParameter("CDMCIDOutput", adInteger, adParamOutput)				
            
        End If
                 
			.Parameters("CDMCID") = lngCDMCID
			.Parameters("GroupName") = strGroupName
			.Parameters("DivisionName") = strDivisionName
			.Parameters("BranchName") = strBranchName
			.Parameters("DepartmentName") = strDepartmentName
			.Parameters("DepartmentNumber") = strDepartmentNumber
			.Parameters("CostCentre") = strCostCentre
			.Parameters("EmployeeID") = strEmployeeID
			.Parameters("EmployeeType") = strEmployeeType
			.Parameters("Firstname") = strFirstname
			.Parameters("Surname") = strSurname
			.Parameters("Title") = strTitle
			.Parameters("Email_Address") = strEmail_Address
			.Parameters("TelephoneNumber") = strTelephoneNumber
			.Parameters("MobileNumber") = strMobileNumber
			.Parameters("DateofBirth") = strDateofBirth
			.Parameters("Gender") = strGender
			.Parameters("ActualRankLvl") = strActualRankLvl
			.Parameters("Site") = strSite
			.Parameters("Unit") = strUnit
			.Parameters("ReportsTo") = strReportsTo
			.Parameters("DCD_PostalAddress") = strDCD_PostalAddress
			.Parameters("addressline1") = straddressline1
			.Parameters("addressline2") = straddressline2
			.Parameters("addressline3") = straddressline3
			.Parameters("addressline4") = straddressline4
			.Parameters("addressline5") = straddressline5
			.Parameters("addressline6") = straddressline6
			.Parameters("PostalAddress_Unit") = strPostalAddress_Unit
			.Parameters("PostalAddress_ClientLocation") = strPostalAddress_ClientLocation
			.Parameters("PostalAddress_DeliveryLocation") = strPostalAddress_DeliveryLocation
			.Parameters("PostalAddress_City") = strPostalAddress_City
			.Parameters("PostalAddress_State") = strPostalAddress_State
			.Parameters("PostalAddress_PostCode") = strPostalAddress_PostCode
			.Parameters("PostalAddress_Country") = strPostalAddress_Country
			.Parameters("ClientLocation") = strClientLocation
			.Parameters("StreetAddress") = strStreetAddress
			.Parameters("City") = strCity
			.Parameters("State") = strState
			.Parameters("PostCode") = strPostCode
			.Parameters("DCDProtectedIdentity") = strDCDProtectedIdentity
			.Parameters("IsValidPostal") = strIsValidPostal
			.Parameters("OutAddr1") = strOutAddr1
			.Parameters("OutAddr2") = strOutAddr2
			.Parameters("OutAddr3") = strOutAddr3
			.Parameters("OutSuburb") = strOutSuburb
			.Parameters("OutState") = strOutState
			.Parameters("OutPostCode") = strOutPostCode
			.Parameters("PostalMessage") = strPostalMessage
			.Parameters("hasChanged") = strhasChanged
			.Parameters("FormalFirstName") = strFormalFirstName
			.Parameters("FormalLastName") = strFormalLastName
			.Parameters("FormalMiddleName") = strFormalMiddleName
			.Parameters("OutTitle") = strOutTitle
			.Parameters("OutDinersWorkPhone") = strOutDinersWorkPhone
			.Parameters("OutDinersMobilePhone") = strOutDinersMobilePhone
			.Parameters("OutANZPhone") = strOutANZPhone
			.Parameters("OutDinersAddress1") = strOutDinersAddress1
			.Parameters("OutDinersAddress2") = strOutDinersAddress2
			.Parameters("RemoveCountdown") = strRemoveCountdown
			.Parameters("FirstUpdated") = NULL 'strFirstUpdated
			.Parameters("LastUpdated") = NULL 'strLastUpdated
			.Parameters("Active") = strActive
			.Parameters("UpdatedBy") = Session("UserID")
			'.Parameters("DateUpdated") = strDateUpdated

			.Parameters("FileID") = strFileID
			.Parameters("Loaded") = "N"'strLoaded
					   
			.ActiveConnection = objCon
			
		End With
			
		objCmd.Execute        
		
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCIDOutput")    
     		                  			     				     		     		
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


Public Function GetNextCDMCFileID()

	'Description:	Gets the next FILE ID number for ANZ Cardlist files.
	objRS.Open "SELECT TOP 1 [FileSeqNum] FROM tblCAPSFileLoad WHERE FileType = 'CDMCFile' AND [Deleted] = 'N' ORDER BY [FileSeqNum] DESC ",objCon

		If Not objRS.EOF Then
			GetNextCDMCFileID = objRS("FileSeqNum") + 1
		Else
			GetNextCDMCFileID = 1
		End If

	objRS.Close
	
End Function

Set objRS = Nothing
Set objCon = Nothing

 %>


