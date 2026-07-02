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
Dim objRS1

Dim strDeleteCheck
Dim dteBatchDate

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")

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

'If the local load has been clicked then call the procedure to load the network file rather than uploading it -  this is for the Promaster Accounts load
If request.QueryString("Action")="SaveFileLocalAccounts" Then

	Call StartLoadLocalAccounts()
	
End If

'New (Feb 2025) Load ProMaster Accounts data from the SQL linked server
If request.QueryString("Action")="SaveFileLocalAccountsSQL" Then
	Call LoadProMasterAccountsLinked(0,"")
End If


'If the load decode has been clicked then call the procedure to load from CMS/ProMaster
If Request.QueryString("Action")="SaveDecode" Then
	Call StartLoadDecode()
End If


If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
End If


'Check the Status of CMS. If it is down then display a warning message to the user so they do not attempt to load PM data
on error resume next

Dim objCon2

	'ProMaster Connection details
	Set objCon2 = Server.CreateObject("ADODB.Connection")
	'Session("DBConnection2") = "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	objCon2.ConnectionTimeout=2
	'objCon2.Open Session("DBConnection2")
	objCon2.Open "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	
	objRS1.Open "SELECT TOP 1 account_ref_no FROM card_account WITH(nolock)",objCon2
			
	objRS1.Close

	If err.Number="3704" Then Response.Write "<div class=""alert alert-danger"" role=""alert""><i class=""fa fa-exclamation""></i> CMS is currently not available. DO NOT Load ProMaster Data until CMS is back up.</div>"

on error goto 0
	
				
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
	self.location="UploadPM.asp?Action=SaveFileLocal"
}
function UploadLocalAccounts()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="UploadPM.asp?Action=SaveFileLocalAccounts"
}

function UploadDecode()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="UploadPM.asp?Action=SaveDecode"
}


function UploadLocalAccountsSQL()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="UploadPM.asp?Action=SaveFileLocalAccountsSQL"
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
	self.location="UploadPM.asp?FileDate=" + document.getElementById("CSDate").value;
}
</script>
<script src="../js/jquery.js"></script>

<body>
<main class="main py-3">
      <div class="container">
<form action="UploadPM.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	<div class="row" id="basic-table">
  <div class="col-5">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../images/CMS.png" title="ProMaster (CMS)"> ProMaster File Load</h4>
        </div>
      <div class="card-content">
        <div class="card-body">
		

<div class="form-body">
<div class="row col-12">
<!--<div class="col-auto mr-auto">
	<div class="form-group">
	   <div class="checkbox">
		<input HIDDEN type="checkbox" class="checkbox-input" id="chkDelete" name="chkDelete">
		<label HIDDEN for="chkDelete">Overwrite Existing Batch</label>
	  </div>
	</div>
  </div>-->
  <div class="col-4 text-right">
	<button type="button" class="btn btn-primary btn-xs" onclick="UploadLocal();"><i class="fa fa-upload"></i> Update Users</button>
  </div>
    <div class="col-4 text-right">
	<button type="button" class="btn btn-primary btn-xs" onclick="UploadLocalAccounts();"><i class="fa fa-upload"></i> Update Accounts</button>
  </div>
  <div class="col-4 text-right">
	<button type="button" class="btn btn-primary btn-xs" onclick="UploadDecode();"><i class="fa fa-upload"></i> Update Decode</button>
  </div>
</div>
</div>

<div class="col-lg-12 col-md-12">
<p class="text-left">
<span id="Progress" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>
<br>
<font color="red" size="2"><b>NOTE: </Font><font color="black" size="2">Clicking 'Upload PM Data' button above will refresh CAPS with ProMaster data
           <!-- <BR>* This may take a few minutes and is best run outside normal hours, as best as possible-->
</B></Font>
</p>
</div>

<div class="row col-12">

 <div class="col-4 text-right">
	
  </div>
   <div class="col-4 text-right">
	<button type="button" class="btn btn-outline-primary btn-xs" onclick="UploadLocalAccountsSQL();"><i class="fa fa-upload"></i> Update Accounts (SQL)</button>
  </div>
   <div class="col-4 text-right">
	
  </div>
</div>
		</div>
	  </div>
    </div>
   </div>
  

  <div class="col-7">
    <div class="card">
     
      <div class="card-content">
        <div class="card-body">		
		<%DisplaySummary()%>	
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

	'Create the SQL statement
	
	objRS.Open "SELECT TOP 20 * FROM tblCAPSProMasterUser WITH(NOLOCK) ",objCon,0,1
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no ProMaster User data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Update Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 table-compact"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing ProMaster User Data already in CAPS</th></tr>" & _
		        "<tr><th style=""text-align:center"">File ID</th>" & _
				"<th style=""text-align:center"">extract_date</th>" & _
				"<th style=""text-align:center"">employee_id</th><th style=""text-align:center"">contractor_ind</th>" & _	
		        "<th style=""text-align:center"">user_name</th>" & _
	 	        "<th style=""text-align:center"">first_name</th>" & _	
	 	        "<th style=""text-align:center"">Date Updated</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR><TD>" & objRS("ProMasterUserID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("extract_date") & "</TD><TD style=""text-align:center"">" & objRS("employee_id") & "</TD><TD style=""text-align:center"">" & objRS("contractor_ind") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("user_name") & "</TD><TD style=""text-align:center"">" & objRS("first_name") & "</TD><TD style=""text-align:center""></TD>" & _
			                    "</TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close
		
		objRS.Open "SELECT TOP 20 * FROM tblCAPSProMasterAccount WITH(NOLOCK) ",objCon,0,1
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no ProMaster Account data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Update Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 table-compact"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing ProMaster Account Data already in CAPS</th></tr>" & _
		        "<tr><th style=""text-align:center"">File ID</th>" & _
				"<th style=""text-align:center"">user_name</th>" & _
				"<th style=""text-align:center"">unit_id</th><th style=""text-align:center"">cost_ctr</th>" & _	
		        "<th style=""text-align:center"">account_ref_no</th>" & _
	 	        "<th style=""text-align:center"">create_date</th>" & _	
	 	        "<th style=""text-align:center"">created_by</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR><TD>" & objRS("ProMasterAccountID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("user_name") & "</TD><TD style=""text-align:center"">" & objRS("unit_id") & "</TD><TD style=""text-align:center"">" & objRS("cost_ctr") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("account_ref_no") & "</TD><TD style=""text-align:center"">" & objRS("create_date") & "</TD><TD style=""text-align:center""></TD>" & _
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

		objRS.Open "SELECT TOP 6 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE DateLoaded > '" & dteBatchDate & "' AND FileType In('ProMasterUser','ProMasterAccount') ORDER BY DateLoaded DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no ProMaster User data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
				"<tr><th Style=""width:20px;"">File Type.</th>" & _
		        "<th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th><th>Deleted</th></tr>" 
				
				'Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                '"<tr><th colspan=""5"" style=""text-align:left"">ProMaster User Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
		        '"<tr><th Style=""width:20px;"">Batch No.</th>" & _
				'"<th>Total Records</th>" & _
				'"<th>Status</th>" & _
	 	        '"<th>Date Loaded</th><th>Deleted</th></tr>" 
				
		    End If
		    
		    Do until objRS.eof
				
				x = x + 1
				
				If IsNull(objRS("DateLoaded")) Then
					strDateUpdated = ""
				Else
					strDateUpdated = FormatDateTime(objRS("DateLoaded"),2)
				End If
				
				Response.Write "<TR><TD>" & objRS("FileType") & "</TD><TD><a href=""UploadPM.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center"">" & objRS("RecordCount") & "</TD><TD style=""text-align:center"">" & objRS("Status") & "</TD>" & _
							"<TD style=""text-align:center"" Title=""" & objRS("DateLoaded") & """>" & strDateUpdated & "</TD><TD style=""text-align:center"">" & objRS("Deleted") & "</TD></TR>"
							
    			objRS.Movenext			
		    Loop
    			
			
								
	    objRS.Close

        Response.Write "</table>"
		
End Sub

Sub StartLoadLocal()
'Procedure to load the local file from within the network, rather than loading the file to the server
Dim objCon2
Dim y
Dim strSQL
Dim strFileDateTime
Dim strFileSeqNum
Dim lngFileLoadID
Dim strFileName

Dim strProMasterUserID
Dim strextract_date
Dim stremployee_id
Dim strcontractor_ind
Dim struser_name
Dim strfirst_name
Dim strsurname
Dim strlocation_name
Dim stradmin_ctr
Dim stradmin_ctr_name
Dim stractive_indicator
Dim strlocked
Dim strinactive_reason
Dim stradmin_centre_controller
Dim strenterprise_controller
Dim stremail_address
Dim strWork_Phone
Dim strMobile
Dim strreview_date
Dim strcreate_date
Dim strcreated_by
Dim strlast_logon
Dim strunprocessed_transactions
Dim stractive_cards
Dim strSupervisor

on error resume next
	'ProMaster Connection details
	Set objCon2 = Server.CreateObject("ADODB.Connection")
	Session("DBConnection2") = "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	objCon2.ConnectionTimeout=2
	objCon2.Open Session("DBConnection2")
	
on error goto 0

	If Request.QueryString("Action")="SaveFileLocal" Then

		'If there is no connection to ProMaster (Card Management System) then do not try to use the connection
		If objCon2.State = 1 Then
		
			'response.write "objCon2.State = " & objCon2.State
			'exit sub
			'Clear the exiting data in the table tblCAPSProMasterUsers before inserting the new records
			objCon.Execute "TRUNCATE TABLE tblCAPSProMasterUser"
			
			strSQL = "set nocount on declare @timeoffset int select @timeoffset = isnull(times.offset + times.dlsoffset,10) from (select offset, case daylight_savings_flag when 'Y' then 1 else 0 " & _
				"end dlsoffset from company_unit c inner join timezones t on t.timezone_id = c.timezone_id where unit_id='COMPANY' and unit_type='COMPANY') times " & _
				"SELECT convert(varchar, getdate(), 103) as extract_date, dbo.ASCIICharOnly(pu.Employee_id) as Employee_id, contractor_ind, pu.User_name, " & _
				"dbo.ASCIICharOnly(First_Name) as First_Name, dbo.ASCIICharOnly(Surname) as Surname, Location_Name, Unit_ID as Admin_Ctr, " & _
				"(SELECT Name FROM Company_Unit CU (nolock) where CU.Unit_ID = PU.Unit_ID) as Admin_Ctr_Name, Active_Indicator, " & _
				"CASE WHEN (Active_indicator = 'N' and Inactive_Reason <> 'User has failed to log on 3 times and was made inactive.') THEN 'Y' Else 'N' END as Locked, " & _
				"isnull(Inactive_Reason,'') as 'Inactive_Reason', " & _
				"CASE WHEN (ISNULL((Select group_name from SECURITY_GROUPINGS (nolock) where user_name = pu.user_name and group_name = 'SA'),'N') = 'SA') THEN 'S' " & _
				"WHEN (ISNULL((Select group_name from SECURITY_GROUPINGS (nolock) where user_name = pu.user_name and group_name = 'AC'),'N') = 'AC') THEN 'Y' " & _
				"Else 'N' END as 'Admin_centre_controller', CASE WHEN (ISNULL((Select group_name from SECURITY_GROUPINGS (nolock) where user_name = pu.user_name and group_name = 'EC'),'N') = 'EC') THEN 'Y' " & _
				"WHEN (ISNULL((Select group_name from SECURITY_GROUPINGS (nolock) where user_name = pu.user_name and group_name = 'HD'),'N') = 'HD') THEN 'Y' Else 'N' END as 'Enterprise_Controller', " & _
				"email_address, dbo.ASCIICharOnly(Work_Phone) as Work_Phone, dbo.ASCIICharOnly(Work_Phone2) as Mobile, isnull(Contact2,'') as Review_date, " & _
				"convert(varchar, dateadd(hour, @timeoffset, pu.create_date), 103) as create_date, pu.Created_By, Isnull((SELECT convert(varchar, max(dateadd(hour, @timeoffset,logon_date)), 103) from Logon_History LH (nolock) where LH.user_name = PU.User_Name and Logon_Success = 'y'),'') as Last_Logon, " & _
				"(SELECT ISNULL(SUM(Reccount),0) from User_WF_Counts WFC (nolock) WHERE WFC.User_Name = PU.User_Name and Group_Name = 'username' and Activity_ID in (3,4,8) and Object_Group = 'T') as Unprocessed_Transactions, " & _
				"(SELECT COUNT(*) from card_account ca (nolock) where ca.user_name = PU.User_Name and Card_Status = 'N') as Active_Cards, " & _
				"isnull(ss.supervisor_id,'') as 'Supervisor' FROM procharge_user PU (nolock) left join supervisor_structure ss on pu.user_name = SS.user_name"
				
			
			
			'Open a recordset in the ProMaster (CMS) database to check the Employee has a CMS Account
			objRS.Open strSQL,objCon2,0,1
			'objRS.Open "SELECT top 100 * FROM procharge_user WITH(NoLock) WHERE [active_indicator ] = 'Y'",objCon2,0,1

				Do Until objRS.EOF
				
					y = y + 1
					
					'strProMasterUserID = objRS("ProMasterUserID")
					strextract_date = objRS("extract_date")
					stremployee_id = objRS("employee_id")
					strcontractor_ind = objRS("contractor_ind")
					struser_name = objRS("user_name")
					strfirst_name = objRS("first_name")
					strsurname = objRS("surname")
					strlocation_name = objRS("location_name")
					stradmin_ctr = objRS("admin_ctr")
					stradmin_ctr_name = objRS("admin_ctr_name")
					stractive_indicator = objRS("active_indicator")
					strlocked = objRS("locked")
					strinactive_reason = objRS("inactive_reason")
					stradmin_centre_controller = objRS("admin_centre_controller")
					strenterprise_controller = objRS("enterprise_controller")
					stremail_address = objRS("email_address")
					strWork_Phone = objRS("Work_Phone")
					strMobile = objRS("Mobile")
					strreview_date = objRS("review_date")
					strcreate_date = objRS("create_date")
					strcreated_by = objRS("created_by")
					strlast_logon = objRS("last_logon")
					strunprocessed_transactions = objRS("unprocessed_transactions")
					stractive_cards = objRS("active_cards")
					strSupervisor = objRS("Supervisor")

					'Call the procedure to save the record via a Command Object
					SaveRecord strextract_date,stremployee_id,strcontractor_ind,struser_name,strfirst_name,strsurname,strlocation_name,stradmin_ctr,stradmin_ctr_name,_ 
							stractive_indicator,strlocked,strinactive_reason,stradmin_centre_controller,strenterprise_controller,stremail_address,strWork_Phone,strMobile,strreview_date,_ 
							strcreate_date,strcreated_by,strlast_logon,strunprocessed_transactions,stractive_cards,strSupervisor, y
				
					objRS.Movenext
				Loop
				
			objRS.Close
			
			'If any records have been loaded then update the file load record
			If y > 0 Then
				
				'Get the File Date time, which is just today
				strFileDateTime = FormatDateTime(now(),2)
				'strFileSeqNum = Mid(strLine,11,12)
				
				'Set the filename to the server name
				strFileName = "SERVER=CMSPRXCMDBL,12805"
				
				'Get the next file number sequence (increments of 1 from the previous file)
				'strFileSeqNum = GetFileLoadID("ProMasterUser",0,strFileName)
				strFileSeqNum = GetLastFileLoadID("ProMasterUser",strFileName)
				strFileSeqNum = strFileSeqNum + 1
			
			
				'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
				lngFileLoadID = SaveFileLoadID ("ProMasterUser","CMS Database","SERVER=CMSPRXCMDBL,12805",y,0,0,0,0,0,0,0,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
				
				'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
				Call UpdateFileLoadSummary ("ProMasterUser",strFileSeqNum, strFileName,lngFileLoadID)
				'response.write "UpdateFileLoadSummary (""CSFRomDiners""," & strFileSeqNum & "," & lngFileLoadID & ")"
				
				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
								"<span aria-hidden=""true"">&times;</span></button>" & _
								"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
								"<span>ProMaster Users updated from CMS """ & strFileSeqNum & """ load COMPLETE!</span></div></div></div>"
								
								
			End If
	
		Else
			Response.write = "CMS database currently unavailable, please try again in 1 hour"
			
			
		End If	
	
		
	End if

End Sub

Sub StartLoadLocalAccounts()
'Procedure to load the local file from within the network, rather than loading the file to the server
Dim objCon2
Dim y
Dim strSQL
Dim strFileDateTime
Dim strFileSeqNum
Dim lngFileLoadID
Dim strFileName

Dim strProMasterAccountID
Dim strcard_type
Dim struser_name
Dim strunit_id
Dim strcompany
Dim strgl_code
Dim strcost_ctr
Dim strinternal_order
Dim strwbs_element
Dim strpost_code
Dim straccount_ref_no
Dim strcreate_date
Dim strcreated_by
Dim strissue_date
Dim strexpiry_date
Dim strcard_status
Dim strAttention
Dim straddr1
Dim straddr2
Dim straddr3
Dim straddr4
Dim straddr_postcode
Dim straddr_state
Dim strAuto_Approve
Dim strPlastic_Type
Dim strmonthlyspendlimit
Dim strtxnlimit
Dim strnameonaccount
Dim strcardholdereid
Dim strcardholderemail
Dim strcardholdermobile
Dim strcardholderwork
Dim strreportgroup
Dim strExtract_date
Dim strTrue_Account_Ref


on error resume next
	'ProMaster Connection details
	Set objCon2 = Server.CreateObject("ADODB.Connection")
	Session("DBConnection2") = "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	objCon2.ConnectionTimeout=2
	objCon2.Open Session("DBConnection2")
	
on error goto 0

	If Request.QueryString("Action")="SaveFileLocalAccounts" Then

		'If there is no connection to ProMaster (Card Management System) then do not try to use the connection
		If objCon2.State = 1 Then
		
			'Clear the exiting data in the table tblCAPSProMasterUsers before inserting the new records
			'objCon.Execute "TRUNCATE TABLE tblCAPSProMasterAccount"									------MOVED BELOW to after opening recordset to avoid leaving an empty table if ProMaster down
			
			strSQL = "set nocount on " & _
"declare @timeoffset int " & _
"select @timeoffset = isnull(times.offset + times.dlsoffset,10) " & _
"from (select offset, case daylight_savings_flag when 'Y' then 1 else 0 end dlsoffset " & _
"from company_unit c inner join timezones t on t.timezone_id = c.timezone_id " & _
"where unit_id='COMPANY' and unit_type='COMPANY') times " & _
"select " & _
"ca.card_type, " & _
"isnull(UPPER(ca.user_name),'') as 'User_name', " & _
"isnull(pu.unit_id,'') as 'unit_id', " & _
"isnull(UPPER(dbo.getglfromcode(gl_code,1)),'') as 'Company', " & _
"isnull(UPPER(dbo.getglfromcode(gl_code,2)),'') as 'GL_Code', " & _
"isnull(UPPER(dbo.getglfromcode(gl_code,3)),'')  as 'Cost_CTR', " & _
"isnull(UPPER(dbo.getglfromcode(gl_code,4)),'') as 'Internal_Order', " & _
"isnull(UPPER(dbo.getglfromcode(gl_code,5)),'') as 'WBS_Element', " & _
"isnull(UPPER(dbo.getglfromcode(gl_code,6)),'') as Post_Code, " & _
"case when ca.pan != pc.pan " & _
"then ca.user_name + '03' + (convert(varchar,pc.payment_card_id)) " & _
"else isnull(UPPER(ca.account_ref_no),'') " & _
"end as 'Account_ref_no'," & _
"case when pc.pan is not null " & _
"then isnull(convert(varchar, dateadd(hour, @timeoffset, pc.create_date), 103),'') " & _
"else isnull(convert(varchar, dateadd(hour, @timeoffset, ca.create_date), 103),'') " & _
"end as 'create_date', " & _
"case when pc.pan is not null " & _
"then isnull(pc.created_by,'') " & _
"else isnull(ca.created_by,'') " & _
"end as 'created_by', " & _
"case when pc.pan is not null  " & _
"then isnull(convert(varchar, dateadd(hour, @timeoffset, pc.create_date), 103),'') " & _
"else isnull(convert(varchar, dateadd(hour, @timeoffset, ca.create_date), 103),'') " & _
"end as 'issue_date', " & _
"case when pc.pan is not null " & _
"then isnull(convert(varchar, (pc.expiry_date), 103),'') " & _
"else isnull(convert(varchar, (ca.expiry_date), 103),'') " & _
"end as 'expiry_date', " & _  
"case when pc.pan is not null " & _
"then isnull(pc.card_status,'') " & _
"else isnull(ca.card_status,'') " & _
"end as 'card_status', " & _   
"isnull((Select value_string " & _ 
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = '*attention'),'') " & _
"as 'Attention', " & _
"isnull((Select value_string " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Address 1'),'') " & _
"as 'Addr1' , " & _
"isnull((Select value_string " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Address 2'),'') " & _
"as 'Addr2' , " & _
"isnull((Select value_string " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and (value_label = 'Address 3' or value_label = 'Address 3 (City / Suburb)')),'') " & _
"as 'Addr3' , " & _
"isnull(UPPER(LEFT((Select value_string " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Address Postcode'),10)),'') " & _
"as 'Addr_postcode' , " & _
"isnull(UPPER(LEFT((Select value_string " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Address State'),3)),'') " & _
"as 'Addr_State' , " & _
"isnull((Select value_decimal " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Auto-approve'),0) " & _
"as 'Auto_approve' , " & _
"isnull(UPPER((Select value_string " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Card Type')),'') " & _
"as 'Plastic_Type' , " & _
"isnull((Select value_decimal " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Monthly Spend Limit'),0) " & _
"as 'Monthly Spend Limit', " & _
"isnull((Select value_String " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Name On Account'),'') " & _
"as 'Name On Account', " & _
"isnull((Select value_String " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Cardholder EID'),'') " & _
"as 'Cardholder EID', " & _
"isnull((Select value_String " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Cardholder Email'),'') " & _
"as 'Cardholder Email', " & _
"isnull((Select value_String " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Report Group'),'') " & _
"as 'Report Group', " & _
"convert(varchar,getdate(),103) as 'Extract_date', " & _
"ca.account_ref_no as 'True_Account_Ref', " & _
"isnull((Select value_String " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Address 4'),'') " & _
"as 'Addr4', " & _
"isnull((Select value_decimal " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Transaction Limit'),0) " & _
"as 'Transaction Limit', " & _
"isnull((Select value_String " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Phone - Mobile'),'') " & _
"as 'cardholder mobile', " & _
"isnull((Select value_String " & _
"from card_account_param (nolock) " & _
"where card_type = ca.card_type and card_account_number = ca.card_account_number and value_label = 'Phone - Work'),'') " & _
"as 'cardholder work' " & _
"from " & _
"card_account ca (nolock) " & _
"left join payment_cards pc (nolock) on pc.card_type = ca.card_type and pc.card_account_number = ca.card_account_number " & _
"left join procharge_user pu (nolock) on ca.user_name = pu.user_name "
			'response.write strSQL
			'Open a recordset in the ProMaster (CMS) database to check the Employee has a CMS Account
			objRS.Open strSQL,objCon2,0,1
			'objRS.Open "SELECT top 100 * FROM procharge_user WITH(NoLock) WHERE [active_indicator ] = 'Y'",objCon2,0,1

				''Moved from above to avoid the table being deleted/truncated before a connection is made to ProMaster, as it times out when ProMaster is down and leaves an empty table
				'Clear the exiting data in the table tblCAPSProMasterUsers before inserting the new records
				objCon.Execute "TRUNCATE TABLE tblCAPSProMasterAccount"
			
				Do Until objRS.EOF
				
					y = y + 1
					
					'strProMasterUserID = objRS("ProMasterUserID")
					'strProMasterAccountID = objRS("ProMasterAccountID")
					strcard_type = objRS("card_type")
					struser_name = objRS("user_name")
					strunit_id = objRS("unit_id")
					strcompany = objRS("company")
					strgl_code = objRS("gl_code")
					strcost_ctr = objRS("cost_ctr")
					strinternal_order = objRS("internal_order")
					strwbs_element = objRS("wbs_element")
					strpost_code = objRS("post_code")
					straccount_ref_no = objRS("account_ref_no")
					strcreate_date = objRS("create_date")
					strcreated_by = objRS("created_by")
					strissue_date = objRS("issue_date")
					strexpiry_date = objRS("expiry_date")
					strcard_status = objRS("card_status")
					strAttention = objRS("Attention")
					straddr1 = objRS("addr1")
					straddr2 = objRS("addr2")
					straddr3 = objRS("addr3")
					straddr4 = objRS("addr4")
					straddr_postcode = objRS("addr_postcode")
					straddr_state = objRS("addr_state")
					strAuto_Approve = objRS("Auto_Approve")
					strPlastic_Type = objRS("Plastic_Type")
					strmonthlyspendlimit = objRS("monthly spend limit")
					strtxnlimit = objRS("transaction limit")
					strnameonaccount = objRS("name on account")
					strcardholdereid = objRS("cardholder eid")
					strcardholderemail = objRS("cardholder email")
					strcardholdermobile = objRS("cardholder mobile")
					strcardholderwork = objRS("cardholder work")
					strreportgroup = objRS("report group")
					strExtract_date = objRS("Extract_date")
					strTrue_Account_Ref = objRS("True_Account_Ref")

					'Call the procedure to save the record via a Command Object
					SaveRecordAccount 0,strcard_type,struser_name,strunit_id,strcompany,strgl_code,strcost_ctr,strinternal_order,strwbs_element,strpost_code,straccount_ref_no,strcreate_date,strcreated_by,strissue_date,_
strexpiry_date,strcard_status,strAttention,straddr1,straddr2,straddr3,straddr4,straddr_postcode,straddr_state,strAuto_Approve,strPlastic_Type,strmonthlyspendlimit,strtxnlimit,strnameonaccount,strcardholdereid,strcardholderemail,strcardholdermobile,strcardholderwork,strreportgroup,_
strExtract_date,strTrue_Account_Ref,y
					
					
				
					objRS.Movenext
				Loop
				
			objRS.Close
			
			'If any records have been loaded then update the file load record
			If y > 0 Then
				
				'Get the File Date time, which is just today
				strFileDateTime = FormatDateTime(now(),2)
				'strFileSeqNum = Mid(strLine,11,12)
				
				'Set the filename to the server name
				strFileName = "SERVER=CMSPRXCMDBL,12805"
				
				'Get the next file number sequence (increments of 1 from the previous file)
				strFileSeqNum = GetLastFileLoadID("ProMasterAccount",strFileName)
				strFileSeqNum = strFileSeqNum + 1
	
				
			
			
			
				'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
				lngFileLoadID = SaveFileLoadID ("ProMasterAccount","CMS Database","SERVER=CMSPRXCMDBL,12805",y,0,0,0,0,0,0,0,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
				
				'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
				Call UpdateFileLoadSummary ("ProMasterAccount",strFileSeqNum, strFileName,lngFileLoadID)
				'response.write "UpdateFileLoadSummary (""CSFRomDiners""," & strFileSeqNum & "," & lngFileLoadID & ")"
				
				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
								"<span aria-hidden=""true"">&times;</span></button>" & _
								"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
								"<span>ProMaster Users updated from CMS """ & strFileSeqNum & """ load COMPLETE!</span></div></div></div>"
								
								
			End If
	
		Else
			Response.write = "CMS database currently unavailable, please try again in 1 hour"
			
		End If	
	
		
	End if

End Sub



Sub SaveRecord(strextract_date,stremployee_id,strcontractor_ind,struser_name,strfirst_name,strsurname,strlocation_name,stradmin_ctr,stradmin_ctr_name,_ 
				stractive_indicator,strlocked,strinactive_reason,stradmin_centre_controller,strenterprise_controller,stremail_address,strWork_Phone,strMobile,strreview_date,_ 
				strcreate_date,strcreated_by,strlast_logon,strunprocessed_transactions,stractive_cards,strSupervisor, x)

Dim intRecord

  	With objCmd
  	
  	    'If the procedure has akready run then don't create the parameter objects again (more than once)
  	    If x = 1 then
			.CommandType = 4
			.CommandText = "spCAPSProMasterUserSave"
			
			.Parameters.Append objCmd.CreateParameter("ProMasterUserID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("extract_date", adDate, adParamInput,0)
			.Parameters.Append objCmd.CreateParameter("employee_id", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("contractor_ind", adChar, adParamInput,1)
			.Parameters.Append objCmd.CreateParameter("user_name", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("first_name", adVarchar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("surname", adVarchar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("location_name", adVarchar, adParamInput,75)
			.Parameters.Append objCmd.CreateParameter("admin_ctr", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("admin_ctr_name", adVarchar, adParamInput,75)
			.Parameters.Append objCmd.CreateParameter("active_indicator", adChar, adParamInput,1)
			.Parameters.Append objCmd.CreateParameter("locked", adChar, adParamInput,1)
			.Parameters.Append objCmd.CreateParameter("inactive_reason", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("admin_centre_controller", adChar, adParamInput,1)
			.Parameters.Append objCmd.CreateParameter("enterprise_controller", adChar, adParamInput,1)
			.Parameters.Append objCmd.CreateParameter("email_address", adVarchar, adParamInput,150)
			.Parameters.Append objCmd.CreateParameter("Work_Phone", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("Mobile", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("review_date", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("create_date", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("created_by", adVarchar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("last_logon", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("unprocessed_transactions", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("active_cards", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Supervisor", advarchar, adParamInput,50)

			.Parameters.Append objCmd.CreateParameter("ProMasterUserIDOutput", adInteger, adParamOutput)				
            
        End If
		
		If strlast_logon = "" Then strlast_logon = NULL
		
                 'response.write "strlast_logon" = strlast_logon
			.Parameters("ProMasterUserID") = 0
			.Parameters("extract_date") = strextract_date
			.Parameters("employee_id") = stremployee_id
			.Parameters("contractor_ind") = strcontractor_ind
			.Parameters("user_name") = struser_name
			.Parameters("first_name") = strfirst_name
			.Parameters("surname") = strsurname
			.Parameters("location_name") = strlocation_name
			.Parameters("admin_ctr") = stradmin_ctr
			.Parameters("admin_ctr_name") = stradmin_ctr_name
			.Parameters("active_indicator") = stractive_indicator
			.Parameters("locked") = strlocked
			.Parameters("inactive_reason") = strinactive_reason
			.Parameters("admin_centre_controller") = stradmin_centre_controller
			.Parameters("enterprise_controller") = strenterprise_controller
			.Parameters("email_address") = stremail_address
			.Parameters("Work_Phone") = strWork_Phone
			.Parameters("Mobile") = strMobile
			.Parameters("review_date") = strreview_date
			.Parameters("create_date") = strcreate_date
			.Parameters("created_by") = strcreated_by
			.Parameters("last_logon") = strlast_logon
			.Parameters("unprocessed_transactions") = strunprocessed_transactions
			.Parameters("active_cards") = stractive_cards
			.Parameters("Supervisor") = strSupervisor
   
		.ActiveConnection = objCon
                
       End With
                
		objCmd.Execute        
		
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("ProMasterUserIDOutput")    
     		                  			     				     		     		
       'response.write  "exec spGeneralExpensesSave =0," & Session("BudgetID") & "," & Session("VersionID") & "," & CostCentreID & ",'GEXP'" & GLCode & "," & BM1 & "," & BM2 & "," & BM3 & "," & BM4 & "," & BM5 & "," & _
       '                     BM6 & "," & BM7 & "," & BM8 & "," & BM9 & "," & BM10 & "," & BM11 & "," & BM12 & "," & OY1 & "," & OY2 & "," & OY3 & ",'" & Comments & "','" & UpdatedBy & "'," & Session("ColumnLock")
End Sub

Sub SaveRecordAccount(strProMasterAccountID,strcard_type,struser_name,strunit_id,strcompany,strgl_code,strcost_ctr,strinternal_order,strwbs_element,strpost_code,straccount_ref_no,strcreate_date,strcreated_by,strissue_date,_
strexpiry_date,strcard_status,strAttention,straddr1,straddr2,straddr3,straddr4,straddr_postcode,straddr_state,strAuto_Approve,strPlastic_Type,strmonthlyspendlimit,strtxnlimit,strnameonaccount,strcardholdereid,strcardholderemail,strcardholdermobile,strcardholderwork,strreportgroup,_
strExtract_date,strTrue_Account_Ref,x)

Dim intRecord

  	With objCmd
  	
  	    'If the procedure has akready run then don't create the parameter objects again (more than once)
  	    If x = 1 then
			.CommandType = 4
			.CommandText = "spCAPSProMasterAccountSave"
			
			.Parameters.Append objCmd.CreateParameter("ProMasterAccountID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("card_type", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("user_name", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("unit_id",  adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("company", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("gl_code", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("cost_ctr", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("internal_order", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("wbs_element", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("post_code", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("account_ref_no", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("create_date", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("created_by", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("issue_date",  adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("expiry_date", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("card_status", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("Attention", adVarchar, adParamInput,100)
			.Parameters.Append objCmd.CreateParameter("addr1", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("addr2", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("addr3", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("addr4", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("addr_postcode", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("addr_state", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("Auto_Approve", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("Plastic_Type", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("monthlyspendlimit", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("transactionlimit", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("nameonaccount", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("cardholdereid", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("cardholderemail", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("cardholdermobile", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("cardholderwork", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("reportgroup", adVarchar, adParamInput,255)
			.Parameters.Append objCmd.CreateParameter("Extract_date", adDate, adParamInput)
			.Parameters.Append objCmd.CreateParameter("True_Account_Ref", adVarchar, adParamInput,255)

			.Parameters.Append objCmd.CreateParameter("ProMasterAccountIDOutput", adInteger, adParamOutput)				
            
        End If
		
		'If strlast_logon = "" Then strlast_logon = NULL
		
                 'response.write "strlast_logon" = strlast_logon
			.Parameters("ProMasterAccountID") = 0
			.Parameters("card_type") = strcard_type
			.Parameters("user_name") = struser_name
			.Parameters("unit_id") = strunit_id
			.Parameters("company") = strcompany
			.Parameters("gl_code") = strgl_code
			.Parameters("cost_ctr") = strcost_ctr
			.Parameters("internal_order") = strinternal_order
			.Parameters("wbs_element") = strwbs_element
			.Parameters("post_code") = strpost_code
			.Parameters("account_ref_no") = straccount_ref_no
			
			If IsDate(strcreate_date) Then 
				.Parameters("create_date") = strcreate_date
			Else
				.Parameters("create_date") = Null
			End If
			
			.Parameters("created_by") = strcreated_by
		  
			If IsDate(strissue_date) Then 
				.Parameters("issue_date") = strissue_date
			Else
				.Parameters("issue_date") = Null
			End If
						 
			If IsDate(strexpiry_date) Then 
				.Parameters("expiry_date") = strexpiry_date
			Else
				 'Response.Write strexpiry_date & " : " & straccount_ref_no & "<BR>"
				.Parameters("expiry_date") = Null
			End If
			
			.Parameters("card_status") = strcard_status
			.Parameters("Attention") = strAttention
			.Parameters("addr1") = straddr1
			.Parameters("addr2") = straddr2
			.Parameters("addr3") = straddr3
			.Parameters("addr4") = straddr4
			.Parameters("addr_postcode") = straddr_postcode
			.Parameters("addr_state") = straddr_state
			.Parameters("Auto_Approve") = strAuto_Approve
			.Parameters("Plastic_Type") = strPlastic_Type
			.Parameters("monthlyspendlimit") = strmonthlyspendlimit
			.Parameters("transactionlimit") = strtxnlimit
			.Parameters("nameonaccount") = strnameonaccount
			.Parameters("cardholdereid") = strcardholdereid
			.Parameters("cardholderemail") = strcardholderemail
			.Parameters("cardholdermobile") = strcardholdermobile
			.Parameters("cardholderwork") = strcardholderwork
			.Parameters("reportgroup") = strreportgroup
			.Parameters("Extract_date") = strExtract_date
			.Parameters("True_Account_Ref") = strTrue_Account_Ref
   
		.ActiveConnection = objCon
                
       End With
                
		objCmd.Execute        
		
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("ProMasterAccountIDOutput")    
     		                  			     				     		     		
       'response.write  "exec spGeneralExpensesSave =0," & Session("BudgetID") & "," & Session("VersionID") & "," & CostCentreID & ",'GEXP'" & GLCode & "," & BM1 & "," & BM2 & "," & BM3 & "," & BM4 & "," & BM5 & "," & _
       '                     BM6 & "," & BM7 & "," & BM8 & "," & BM9 & "," & BM10 & "," & BM11 & "," & BM12 & "," & OY1 & "," & OY2 & "," & OY3 & ",'" & Comments & "','" & UpdatedBy & "'," & Session("ColumnLock")
End Sub


Sub StartLoadDecode()
'Procedure to load the local file from within the network, rather than loading the file to the server
Dim objCon2
Dim y
Dim strSQL
Dim strFileDateTime
Dim strFileSeqNum
Dim lngFileLoadID
Dim strFileName

Dim strAccountRefNo
Dim strUserName
Dim strCardType
Dim strCardAccountNumber


on error resume next
	'ProMaster Connection details
	Set objCon2 = Server.CreateObject("ADODB.Connection")
	Session("DBConnection2") = "File Name=" & Server.MapPath("../Database/ProMaster.udl") & ";"
	objCon2.ConnectionTimeout=2
	objCon2.Open Session("DBConnection2")
	
on error goto 0

	If Request.QueryString("Action")="SaveDecode" Then

		'If there is no connection to ProMaster (Card Management System) then do not try to use the connection
		If objCon2.State = 1 Then
		
			'Clear the exiting data in the table tblCAPSProMasterUsers before inserting the new records
			objCon.Execute "TRUNCATE TABLE tblCAPSDecodedCardAccounts"
			
			strSQL = "SELECT REVERSE(PARSENAME(REPLACE(REVERSE(decode_data.data_line), CHAR(9), '.'), 1)) AS [User], " & _
				"REVERSE(PARSENAME(REPLACE(REVERSE(decode_data.data_line), CHAR(9), '.'), 2)) AS CardType, " & _
				"REVERSE(PARSENAME(REPLACE(REVERSE(decode_data.data_line), CHAR(9), '.'), 3)) AS CardNo, " & _
				"REVERSE(PARSENAME(REPLACE(REVERSE(decode_data.data_line), CHAR(9), '.'), 4)) AS UserName2 " & _
				"FROM decode_data WITH(NOLOCK)"
				
			
			'Open a recordset in the ProMaster (CMS) database to check the Employee has a CMS Account
			objRS.Open strSQL,objCon2,0,1
			'objRS.Open "SELECT top 100 * FROM procharge_user WITH(NoLock) WHERE [active_indicator ] = 'Y'",objCon2,0,1

				Do Until objRS.EOF
				
					y = y + 1
					
					strUserName = objRS("User")
					strCardType = objRS("CardType")
					strCardAccountNumber = objRS("CardNo")
					strAccountRefNo = objRS("UserName2")
					
					'Call the procedure to save the record via a Command Object
					SaveDecodeRecord strUserName, strCardType, strCardAccountNumber, strAccountRefNo, y
				
					objRS.Movenext
				Loop
				
			objRS.Close
			
			'If any records have been loaded then update the file load record
			If y > 0 Then
				
				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
								"<span aria-hidden=""true"">&times;</span></button>" & _
								"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
								"<span>ProMaster Decode Accounts updated from CMS load COMPLETE!</span></div></div></div>"
								
								
			End If
	
		Else
			Response.write = "CMS database currently unavailable, please try again in 1 hour"
			
		End If	
	
	End if

End Sub


Sub SaveDecodeRecord(strUserName, strCardType, strCardAccountNumber, strAccountRefNo, x)

Dim intRecord

  	With objCmd
  	
  	    'If the procedure has already run then don't create the parameter objects again (more than once)
  	    If x = 1 then
			.CommandType = 4
			.CommandText = "spCAPSProMasterDecodeSave"
			
			.Parameters.Append objCmd.CreateParameter("AccountRefNo", adVarChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("UserName", adVarChar, adParamInput,50)
			.Parameters.Append objCmd.CreateParameter("CardType", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("CardAccountNumber", adVarChar, adParamInput,50)

			.Parameters.Append objCmd.CreateParameter("DecodedCardIDOutput", adInteger, adParamOutput)				
            
        End If
		
		.Parameters("AccountRefNo") = strAccountRefNo
		.Parameters("UserName") = strUserName
		.Parameters("CardType") = strCardType
		.Parameters("CardAccountNumber") = strCardAccountNumber
			
		.ActiveConnection = objCon
                
       End With
                
		objCmd.Execute        
		
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("DecodedCardIDOutput")    

End Sub




Public Sub LoadProMasterAccountsLinked(intAccountRefNo,strTrueAccountRef)
'''Procedure to load the ProMAster accounts to CAPS from a stored procedure, which uses a linked SQL server (between CAPS and CMS/ProMAster)
'''rather than connecting to CMS, downloading the data to a flat file then loading back into CAPS, as currently happens in procedure StartLoadLocalAccounts() above
Dim intRecord
Dim strFileSeqNum
Dim strFileDateTime
Dim strFileName
Dim lngFileLoadID

  	With objCmd
		''Set the Command Object Parameters/Settings
		.CommandType = 4
		.CommandText = "spCAPSProMasterAccountsLinkedImport"
		
		.Parameters.Append objCmd.CreateParameter("AccountRefNo", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("TrueAccountRef", adVarChar, adParamInput, 50)
		.Parameters.Append objCmd.CreateParameter("ProMasterAccountIDOutput", adInteger, adParamOutput)				
        
		.Parameters("AccountRefNo") = intAccountRefNo
		.Parameters("TrueAccountRef") = strTrueAccountRef
			
		.ActiveConnection = objCon
                
       End With
                
		objCmd.Execute        
		
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("ProMasterAccountIDOutput")    

		If intRecord > 0 Then
		
			'Get the File Date time, which is just today
				strFileDateTime = FormatDateTime(now(),2)
				'strFileSeqNum = Mid(strLine,11,12)
				
				'Set the filename to the server name
				strFileName = "SERVER=CMSPRXCMDBL,12805"
				
				'Get the next file number sequence (increments of 1 from the previous file)
				strFileSeqNum = GetLastFileLoadID("ProMasterAccount",strFileName)
				strFileSeqNum = strFileSeqNum + 1
	
				'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
				lngFileLoadID = SaveFileLoadID ("ProMasterAccount","CMS Database","SERVER=CMSPRXCMDBL,12805",intRecord,0,0,0,0,0,0,0,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
				
				'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
				Call UpdateFileLoadSummary ("ProMasterAccount",strFileSeqNum, strFileName,lngFileLoadID)
				'response.write "UpdateFileLoadSummary (""CSFRomDiners""," & strFileSeqNum & "," & lngFileLoadID & ")"
				
				
				'''Display a Success message to the user
				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
					"<span aria-hidden=""true"">&times;</span></button>" & _
					"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
					"<span>ProMaster Accounts updated from CMS """ & strFileSeqNum & """ load COMPLETE!</span></div></div></div>"
		Else
			''Display an error message if an error is returned
			Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
				"<span aria-hidden=""true"">&times;</span></button>" & _
				"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
				"<span>ERROR! ProMaster Accounts NOT updated from CMS """ & strFileSeqNum & """. Please contact CAPS System Admin</span></div></div></div>"
			
		End If
								
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
Set objRS1 = Nothing
Set objCon = Nothing
Set objCon2 = Nothing
 %>


