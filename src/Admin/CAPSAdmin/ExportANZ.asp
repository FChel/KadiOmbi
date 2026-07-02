<!-- #Include file=../../CC/CAPSHeader.asp -->
<!-- #Include file=../../ADOVBS.inc -->
<!-- #Include file=../../CC/CAPSFunctions.asp -->
<%
'Description:	New Card file export for Diners
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
Dim strNextANZFileNumber

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

If request.QueryString("Action")="ExportANZ" Then

	Call ExportANZFile()
	
End If

If Not IsEmpty(Request.QueryString("Reload")) Then

	Call StartLoad()
End If

If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
End If

	If Not IsEmpty(Request.QueryString("BatchNumber")) Then
		strNextANZFileNumber = Request.QueryString("BatchNumber")
		If IsNumeric(strNextANZFileNumber) Then strNextANZFileNumber = PadDigits(strNextANZFileNumber,6)
	Else
		strNextANZFileNumber = GetSystemAdmin("ANZFileNumberTo")
		If IsNumeric(strNextANZFileNumber) Then strNextANZFileNumber = PadDigits(strNextANZFileNumber,6)
		
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
        if(getFileExt(document.getElementById('FILE1').value)==".csv" || getFileExt(document.getElementById('FILE1').value)=="..csv")
        {
            if(window.confirm('This will overwrite any existing ANZ To file data! \n \n Continue?')==true)
                {
                document.getElementById('Progress').style.display = "inline";
                frm.submit();
            }
        } 
        else
        {
            alert("Please enter a valid Excel(.csv) or TEXT (..csv) file");
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
	self.location="ExportANZ.asp?FileDate=" + document.getElementById("ANZDate").value;
}
</script>

<body>



<main class="main py-3">
    <div class="container">
	
 <!-- Approve Modal -->
<div class="modal fade" id="ModalApprove" tabindex="-1" role="dialog" aria-labelledby="ModalApprove" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalApproveTitle" style="font-weight:bold;">ANZ File Export Confirmation</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <span>Export ANZ File Number: <input type="text" name="ANZFileExport" id="ANZFileExport" style="border:0; font-weight:bold; width:100px; text-align:right;" value="<%=strNextANZFileNumber%>" >?</span><br><br>
      </div>
      <div class="modal-footer">
		<button type="button" class="btn btn-primary" onClick='window.location="ExportANZ.asp?Action=ExportANZ"'><i class="fa fa-check"></i> Yes</button>
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
        
      </div>
    </div>
  </div>
</div>

<!-- Modal Detail End-->

<form action="ExportANZ.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	 <div class="row" id="basic-table">
  <div class="col-3">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../../CC/img/ANZ2.png" height="40px" width="70px" title="Diners"> ANZ File Export</h4>
        </div>
      <div class="card-content">
        <div class="card-body">
			<p class="card-text">
				New ANZ Applications which have been processed in CAPS but not yet exported to ANZ will appear below.  
				Click Export ANZ File to create the export file and process all records below.
			</p>
<div class="col-lg-12 col-md-12">
<fieldset class="form-group">
    <button type="button" class="btn btn-outline-secondary btn-sm" data-toggle="modal" data-target="#ModalApprove" HREF="#" ><i class="fa fa-file"></i> Export </button>
</fieldset>
	<!-- <button type="button" class="btn btn-primary btn-xs" onclick="window.open('TemplateExcel.asp?T=tblCAPSApplication&w=WHERE [CardType] = 'DTC' AND [Status] = 'Awaiting Export' AND [DateExported] IS NULL')"><i class="fa fa-file"></i> View ANZ File in Excel </button>-->
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
			<%DisplaySummary()%>	
		</div>
	  </div>
    </div>
   </div>
   

  
  <div class="col-3">
		<div class="card">
           <div class="card-content">
				<div class="card-body">
					<%DisplayFileSummary()%>	
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
<!-- #Include file=../../CC/CAPSFooter.asp -->
</body>
</html>

<%
Sub DisplayTableDetails()

Dim strWhere

'If Not IsEmpty(Request.QueryString("BatchNo")) Then
'	If IsNull(Request.QueryString("BatchNo")) or Request.QueryString("BatchNo")= "" Then 
		
'	Else
		strWhere = "WHERE [CardType] = 'DPC' AND [Status] = 'Awaiting Export'"
	
		'If Request.QueryString("BatchNo") = 0 Then strWhere = strWhere & " OR BatchNo IS NULL"
'	End If
'Else
'	strWhere = ""
'End If

objRS.Open "SELECT TOP 50 * FROM qryCAPSANZApplicationsExport WITH(NOLOCK) "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no ANZ data ready for export</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Files will be added as part of the admin and load functions, and will appear here when appropriate</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing ANZ Data ready for export</th></tr>" & _
		        "<tr><th Style=""width:20px;"">ApplicationID</th>" & _
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
			    Response.Write "<TR class='clickable-row' data-href='Displaydataset.asp?tbl=qryCAPSApplications&W=WHERE ApplicationID=" & objRS("ApplicationID") & "' style=""cursor: pointer;""><TD>" & objRS("ApplicationID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("Security") & "</TD><TD style=""text-align:center"">" & objRS(5) & "</TD><TD style=""text-align:center"">" & objRS(3) & " " & objRS(4) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(6) & "</TD><TD style=""text-align:center"">" & objRS(7) & "</TD><TD style=""text-align:center"">" & objRS(8) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(9) & "</TD><TD style=""text-align:center"">" & objRS(10) & "</TD><TD style=""text-align:center"">" & objRS(11) & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS(12) & "</TD><TD style=""text-align:center"">" & objRS(13) & "</TD><TD style=""text-align:center"">" & objRS(14) & "</TD>" & _
								"<TD style=""text-align:center"">" & objRS(15) & "</TD><TD style=""text-align:center"">" & objRS(16) & "</TD><TD style=""text-align:center""></TD>" & _
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
Dim strDateUpdatedTitle

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		objRS.Open "SELECT TOP 6 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE DateLoaded > '" & dteBatchDate & "' AND FileType = 'ANZAppExport' ORDER BY FileSeqNum DESC",objCon
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no ANZ data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""ANZDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-compact mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">ANZ Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""ANZDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
		        "<th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>Total Cards</th><th>Total Employees</th>" & _	
		      	"<th>Status</th>" & _
	 	        "<th>Date Loaded</th></tr>" 
				
		    End If
		    
		    Do until objRS.eof
				
				If IsNull(objRS("DateLoaded")) or objRS("DateLoaded") = "" Then
					strDateUpdated = ""
					strDateUpdatedTitle = ""
				Else
					strDateUpdated = FormatDatetime(objRS("DateLoaded"),vbShortDate)
					strDateUpdatedTitle = "Title=""" & objRS("DateLoaded") & """"
				End If

				Response.Write "<TR><TD><a href=""ExportANZ.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center"">" & objRS("RecordCount") & "</TD><TD style=""text-align:center"">" & objRS("CardCount") & "</TD><TD style=""text-align:center"">" & objRS("EmployeeCount") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("Status") & "</TD>" & _
							"<TD style=""text-align:center"" " & strDateUpdatedTitle & ">" & strDateUpdated & "</TD></TR>"
    			objRS.Movenext			
		    Loop
    			
			
								
	    objRS.Close

        Response.Write "</table>"
		
End Sub

Public Sub ExportANZFile()
'Procedure to Export the NA File including all of the records which are yet to be exported
Const fsoForWriting = 2
Dim strFilePath
Dim strFileNameStart
Dim strNextFileNumber
Dim intRecordCount
Dim strRecordCount
Dim strFileDateTime
Dim lngFileLoadID
Dim strFileDateTimeSec
Dim strError

'Get the Default file location for the server then add the filepath and name for the NA File
strFilePath = GetSystemAdmin("ServerFilePath")
strFileNameStart = GetSystemAdmin("ANZFileStart")
strNextFileNumber = GetSystemAdmin("ANZFileNumberTo")



'Pad the number out to 6 digits
strNextFileNumber = PadDigits(strNextFileNumber,6)
strNextFileNumber = PadDigits(strNextFileNumber,6)

'Compile the File name and path from the variables above
strFilePath = strFilePath & "\Admin\CAPSAdmin\Attachments\ANZ\ANZTo\" & strFileNameStart & PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & ".csv"

strFileDateTime = PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2)
strFileDateTimeSec = PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & Hour(Now()) & Minute(Now()) & Second(Now())
'response.write 	"lngBatchNumber=" & strNextFileNumber & " strNextCSFileNumber= " &  strNextCSFileNumber

'Get the FileLoad details
'lngBatchNumber = GetFileLoadID("NAFile",strNextFileNumber,"")

'If IsNull(lngBatchNumber) OR lngBatchNumber = "" then
'	lngBatchNumber = strNextCSFileNumber
'End If

Dim objFSO
Dim strFileName
Dim lngANZTxnLimit
Dim lngANZDailyCashLimit
Dim lngANZOTCCashLimit

'Get System Parameter Values for ANZ Export
lngANZTxnLimit = GetSystemAdmin("ANZTxnLimit")
lngANZDailyCashLimit = GetSystemAdmin("ANZDailyCashLimit")
lngANZOTCCashLimit = GetSystemAdmin("ANZOTCCashLimit")

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

'Open the text file
Dim objTextStream


	'Open a recordset of all of the NA File records yet to be exported
	
	objRS.Open "SELECT * FROM qryCAPSANZApplicationsExport WITH(NOLOCK) WHERE CardType = 'DPC'  AND Status = 'Awaiting Export'",objCon
	'objRS.Open "SELECT * FROM tblCAPSApplication WITH(NOLOCK) WHERE CardType = 'DPC'",objCon
''Awaiting Export'
		If objRS.EOF Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No records to write to the ANZ File.</div>"
		Else
			lngFileLoadID=0	
			'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
			lngFileLoadID = SaveFileLoadID ("ANZAppExport",strFileNameStart & strFileDateTime & ".csv", strFilePath,-1,0,0,0,0,0,0,0,strFileDateTime,strNextFileNumber,"Exported",Session("UserID"),"N")
			Set objTextStream = objFSO.OpenTextFile(strFilePath, fsoForWriting, True)
			objTextStream.WriteLine "Name_On_Card,Title,GivenName,Initial,Surname,F,G,Address_1,Address_2,Address_3,K,Suburb,State,Postcode,Limit,txnLimit,DailyCash,OTCCash,DOB,Security,EmployeeID,Contact_Number,Type,Delivery,Email"
			
			'format(Now(),"yyyymmddhhmmss")+Format(DLookup("Next_File_No","Next_File_No","File_Type = 'NA_Out'"),"000000");\
			
			'Write each record to the text file
			Do Until objRS.EOF
			
				intRecordCount = intRecordCount + 1
				'Display the contents of the text file
				'strError = strError & CheckForNull(objRS("EIDNo"),"EIDNo",intRecordCount)
				'strError = strError & CheckForNull(objRS("CardNo"),"Card No",intRecordCount)
				'strError = strError & CheckForNull(objRS("CardUpdateInd"),"Card Update Ind",intRecordCount)
				'strError = strError & CheckForNull(objRS("CardExpiryDate"),"Card Expiry Date",intRecordCount)
				'strError = strError & CheckForNull(objRS("CardStatus"),"Card Status",intRecordCount)
				'strError = strError & CheckForNull(objRS("Title"),"Title",intRecordCount)
				'strError = strError & CheckForNull(objRS("Surname"),"Surname",intRecordCount)
				'strError = strError & CheckForNull(objRS("GivenNames"),"Given Names",intRecordCount)
				'strError = strError & CheckForNull(objRS("NameOnCard"),"Name on Card",intRecordCount)
				'strError = strError & CheckForNull(objRS("Address1"),"Address 1",intRecordCount)
				'strError = strError & CheckForNull(objRS("Address2"),"Address 2",intRecordCount)
				'strError = strError & CheckForNull(objRS("Address3"),"Address 3",intRecordCount)
				'strError = strError & CheckForNull(objRS("Suburb"),"Suburb",intRecordCount)
				'strError = strError & CheckForNull(objRS("State"),"State",intRecordCount)
				'strError = strError & CheckForNull(objRS("Postcode"),"Postcode",intRecordCount)
				'strError = strError & CheckForNull(objRS("HomePhone"),"Home Phone",intRecordCount)
				'strError = strError & CheckForNull(objRS("WorkPhone"),"Work Phone",intRecordCount)
				'strError = strError & CheckForNull(objRS("MobilePhone"),"Mobile Phone",intRecordCount)
				'strError = strError & CheckForNull(objRS("Email"),"Email",intRecordCount)
				'strError = strError & CheckForNull(objRS("ReportGroup"),"Report Group",intRecordCount)
				'strError = strError & CheckForNull(objRS("CreditLimit"),"Credit Limit",intRecordCount)				
				
				If IsNull(objRS("ApplicationID")) = False Then
					'objTextStream.WriteLine "D " & PadSpaceLeft(objRS("NameOnCard"),10) & PadDigits(objRS("CardNo"),19) & PadSpaceLeft(objRS("CardUpdateInd"),2) & PadSpaceLeft(objRS("CardExpiryDate"),8) & PadSpaceLeft(objRS("CardStatus"),2) & PadSpaceLeft(objRS("Title"),12) & PadDigitsLeft(objRS("Surname"),25) & PadSpaceLeft(objRS("GivenNames"),30) & PadSpaceLeft(objRS("NameOnCard"),26) & PadSpaceLeft(objRS("Address1"),40) & PadSpaceLeft(objRS("Address2"),40) & PadSpaceLeft(objRS("Address3"),40) & PadSpaceLeft(objRS("Suburb"),25) &  PadSpaceLeft(objRS("State"),4) & PadSpaceLeft(objRS("Postcode"),12) &  PadSpaceLeft(objRS("HomePhone"),12) &  PadSpaceLeft(objRS("WorkPhone"),12) & PadSpaceLeft(objRS("MobilePhone"),12) & PadSpaceLeft(objRS("Email"),70) & PadSpaceLeft(objRS("ReportGroup"),8) &  PadSpaceLeft(objRS("CreditLimit"),11)                         
					objTextStream.WriteLine objRS("NameOnCard") & ",""" & objRS("Title") & """,""" & objRS("FirstName") & """,""" & objRS("Initial") & """,""" & objRS("Surname") & """,,,""" & objRS("Address1") & """,""" & objRS("Address2") & """,""" & objRS("Address3") & """,,""" & objRS("Suburb") & """,""" & objRS("State")& """," & objRS("PostCode")& "," & objRS("Limit") & "," & lngANZTxnLimit & "," & lngANZDailyCashLimit & "," & lngANZOTCCashLimit & ",""" & objRS("DateOfBirth") & """,""" & objRS("Security") & """,""" & objRS("Security") & """," & objRS("Contact_Number") & ",""" & objRS("Type") & """,""" & objRS("Delivery") & """," & objRS("Email")        
							
				End If
				'Call the procedure to update each record as exported once added to the CS File -- USE the File style Batch Number not FileLoadID
				'Call ExportCSRecord (objRS("CSToDinersID"),lngFileLoadID,intRecordCount)
				''Call ExportCSRecord (objRS("ANZToDinersID"),strNextFileNumber,intRecordCount)
				
			objRS.Movenext
			Loop
			
			strRecordCount = PadDigits(intRecordCount,6)
			
			'objTextStream.WriteLine "T" & strRecordCount		
			
			
			'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			strFileName = strFileNameStart & strFileDateTime & ".csv"
			'Response.Write "strFileName = "
			'Response.Write strFileName & "<BR>"
			
			''''Call UpdateFileLoadSummary ("CSToDiners",strNextFileNumber, strFileName, lngFileLoadID)
			'Call the procedure to update the System Parameter CSFileNumber. Increment the Number by 1.
			'''Call UpdateBatchNumber(strNextFileNumber)
			
			If strError = "" Then
				Response.Write "<div class=""alert alert-success"" role=""alert"">ANZ File " & "AUDC_INTOECS_DODNA_D" & strFileDateTime & ".csv" & " ADDED to the ANZ file export folder!</div>"
			Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">ANZ File " & "AUDC_INTOECS_DODNA_D" & strFileDateTime & ".csv" & " has errors : " & strError & "</div>"
			End If
				'Close the file and clean up
				objTextStream.Close
				
			'Response.Write "strNextFileNumber<BR>"
			'Response.Write strNextFileNumber & "<BR>"
			'Response.Write "strFileName<BR>"
			'Response.Write strFileName & "<BR>"
			'Response.Write "lngFileLoadID<BR>"
			'Response.Write lngFileLoadID & "<BR>"

			Call UpdateFileLoadSummary ("ANZAppExport",strNextFileNumber, strFileName, lngFileLoadID)
			
			'Call the procedure to update the System Parameter CSFileNumber. Increment the Number by 1.
			Call UpdateBatchNumber(strNextFileNumber)

			'Call the procedure to update the Exported ANZ Cards as exported after the export process
			Call UpdateANZExports("'" & strNextFileNumber & "'")	
				
		End If

	objRS.Close
	


Set objTextStream = Nothing
Set objFSO = Nothing
			
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

Public Sub UpdateBatchNumber(lngBatchNumber)
'Procedure to update the BatchNumber field in the System Parameters table with the next number
Dim strSQL

	'If the Batch Number is a number then update the System Parameter, otherwise post an error to the screen
	If IsNumeric(lngBatchNumber) Then
		lngBatchNumber = lngBatchNumber + 1
		
		strSQL = "UPDATE tblCAPSSystemParameters SET [ParameterValue] = '" & lngBatchNumber & "' WHERE [ParameterName] = 'ANZFileNumberTo'"
		
		objCon.Execute strSQL
	
	Else
		
		Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR! ANZ File Batch Number: " & lngBatchNumber & " is not a number. See System Admin.</div>"
		
	End If

End Sub


Public Sub UpdateANZExports(BatchNumber)
'Procedure to update the ANZ Applications to Exported 
Dim strSQL

	Response.Write "UPDATE tblCAPSApplication SET ExportBatch = " & BatchNumber & ", [Status] = 'Exported', [DateExported] = GetDate() WHERE [CardType] = 'DPC' AND [Status] = 'Awaiting Export'"
	strSQL = "UPDATE tblCAPSApplication SET ExportBatch = " & BatchNumber & ", [Status] = 'Exported', [DateExported] = GetDate() WHERE [CardType] = 'DPC' AND [Status] = 'Awaiting Export'"
	
	objCon.Execute strSQL

	Response.Write "<div class=""alert alert-success"" role=""alert"">ANZ Applications updated to 'Exported'</div>"

End Sub


Public Sub DisplayFileSummary()

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim strExtension
Dim objFSO
Dim objFolder
Dim objFile

Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\ANZ\ANZTo"

	Set objFolder = objFSO.GetFolder(objStartFolder)	
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">Files Exported <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be sent to ANZ.""></i></th></tr>"
			
	For Each objFile in colFiles

		intCount = intCount + 1
		
		If intCount < 6 Then
			If IsNull(objFile.Name) or objFile.Name = "" Then
				strFile = ""
			Else
				strFile = Left(objFile.Name,8)
				strExtension = "<A HREF=""Attachments/ANZ/ANZTo/" & objFile.Name & """>" & objFSO.GetExtensionName(objStartFolder & "/" & objFile.Name) & "</A>"
				
			End If
			
			Response.Write "<TR><TD>" & strFile & "..." & strExtension & "</TD></TR>"
		End If
		
	Next
	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"
	 
Set objFSO = Nothing

End Sub



Set objRS = Nothing
Set objCon = Nothing

 %>


