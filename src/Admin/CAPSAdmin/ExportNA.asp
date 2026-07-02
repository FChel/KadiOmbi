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
Dim strNextNAFileNumber
Dim strCardType
Dim strNATransactionPage
Dim strCardTypeButton
Dim strServer

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("CardType")) Then

	strCardType = Request.QueryString("CardType")
	
End If

'Response.Write "Card Type = " & strCardType & "<BR>"

If request.QueryString("Action")="Save" Then

	Call StartLoad()
End If

If Not IsEmpty(Request.QueryString("Reload")) Then

	Call StartLoad()
End If

If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
End If

'If the Cancel/Remove has been clicked on the NA File Modal (within the file GetNAfile.asp) then flag the NA File record as removed
If Request.QueryString("Action") = "CancelNA" Then

	Call RemoveNARecord(Request.QueryString("NAToDinersID"),Request.QueryString("NAEID"),Request.QueryString("Status"))

End If

	If Not IsEmpty(Request.QueryString("BatchNumber")) Then
		strNextNAFileNumber = Request.QueryString("BatchNumber")
		If IsNumeric(strNextNAFileNumber) Then strNextNAFileNumber = PadDigits(strNextNAFileNumber,6)
	Else
		If strCardType = "DTC" Then
			strNextNAFileNumber = GetSystemAdmin("NAFileNumber")
			
			'Set the Transaction page variable where this page will redirect to when NA transastions are exported (NATrasnactions.asp or NATransactionsDPC.asp)
			strNATransactionPage = "NATransactions.asp"
		Else
			strNextNAFileNumber = GetSystemAdmin("NAFileNumberDPC")
			'Set the Transaction page variable where this page will redirect to when NA transastions are exported (NATrasnactions.asp or NATransactionsDPC.asp)
			strNATransactionPage = "NATransactionsDPC.asp"
		End If
		If IsNumeric(strNextNAFileNumber) Then strNextNAFileNumber = PadDigits(strNextNAFileNumber,6)
	
	End If	

	'Get the card type button to toggle between DPC and DTC
	If strCardType = "DTC" Then
		strCardTypeButton = "<button type=""button"" class=""btn btn-outline-secondary"" title=""DTC Selected.  Click to Change to DPC"" onClick='window.location=""ExportNA.asp?CardType=DPC""'><i class=""fa fa-plane""></i> DTC</button>"
	Else
		strCardTypeButton = "<button type=""button"" class=""btn btn-outline-secondary"" title=""DPC Selected.  Click to Change to DTC"" onClick='window.location=""ExportNA.asp?CardType=DTC""'><i class=""fa fa-dollar-sign""></i> DPC</button>"
	End If
	
	
	'Get the File Export Location to display to users
	strServer = GetSystemAdmin("GDriveExportFilePath")
	
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
        if(getFileExt(document.getElementById('FILE1').value)==".xls" || getFileExt(document.getElementById('FILE1').value)==".txt")
        {
            if(window.confirm('This will overwrite any existing CS From Diners file data! \n \n Continue?')==true)
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

function loadNA(CardType) {

  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("NADetail").innerHTML = this.responseText;
    }
  };
  
  xhttp.open("GET", "../../CC/AJAX/GetNAFile.asp?CardType=" + CardType, true);
  xhttp.send();
}


function DatePickChange(CardType) {
	
	self.location="ExportNA.asp?CardType=" + CardType + "&FileDate=" + document.getElementById("CSDate").value;
}

function ConfirmExport(cb) {
	
	var id = cb.getAttribute('data-NANumber');
	document.getElementById('NAFileExport').value=id;

}
</script>

<body>

<!-- Modal -->
<div class="modal fade" id="LoadMod" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
   <div class="loader">
        <div class="wrap">
            <div class="spinner"></div>
            <span class="loading-message">Loading...</h6>
        </div>
    </div>
</div>

<!-- Modal -->
<div class="modal fade" id="NAModal" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-large modal-dialog-right modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  NA File to Diners - <%=strCardType & " : Previous File Number: " & strNextNAFileNumber%>
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="NADetail">
               
				  
                
            </div>

			<div class="modal-footer">
				<!--<button type="button" class="btn btn-primary" onclick="window.open('NAToDinersExportExcel.asp')"><i class="fa fa-file"></i> Export NA File to Excel</button>-->
				<button type="button" class="btn btn-primary float-right" data-toggle="modal" data-target="#ModalApprove" data-NANumber="<%=strNextNAFileNumber%>" onClick="ConfirmExport(this);" Title="Click to send the Next NA File"><i class="fa fa-file"></i> Export NA File</button>
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
 
 <!-- Approve Modal -->
<div class="modal fade" id="ModalApprove" tabindex="-1" role="dialog" aria-labelledby="ModalApprove" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalApproveTitle" style="font-weight:bold;">NA File Export Confirmation</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <span>Export NA File Number: <input type="text" name="NAFileExport" id="NAFileExport" style="border:0; font-weight:bold; width:100px; text-align:right;" value="<%=strNextNAFileNumber%>" >?</span><br>
		<span>Export NA File Location (after move from server): <input type="text" name="NAFileExport" id="NAFileExport" style="border:0; font-weight:bold; width:400px; text-align:left;" value="<%=strServer%>" ></span><br><br>
	  </div>
      <div class="modal-footer">
		<button type="button" class="btn btn-primary" onClick='window.location="<%=strNATransactionPage%>?CardType=<%=strCardType%>&Action=ExportNA"'><i class="fa fa-check"></i> Yes</button>
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
        
      </div>
    </div>
  </div>
</div>
<!-- End Approve Modal -->
<main class="main py-3">
    <div class="container">
<form action="ExportCS.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	 <div class="row" id="basic-table">
  <div class="col-4">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../../CC/img/Diners2.png" height="40px" width="50px" title="Diners"> NA To Diners File Export <%=strCardTypeButton%></h4>
        </div>
      <div class="card-content">
        <div class="card-body">
			<p class="card-text">
				New Diners Applications which have been processed in CAPS but not yet exported to Diners (in the NA file) will appear below.  
				Click Export NA File to create the export file and process all records below.
			</p>
<div class="col-lg-12 col-md-12">
<fieldset class="form-group">
    <button type="button" class="btn btn-outline-secondary btn-sm" data-toggle="modal" data-target="#NAModal" HREF="#" onClick="loadNA('<%=strCardType%>');"><i class="fa fa-file"></i> Export <%=strCardType%></button>
	<button type="button" class="btn btn-outline-secondary btn-sm" onClick="self.location='NATransactions.asp?CardType=<%=strCardType%>'"><i class="fa fa-file"></i> View Details </button>
	<!--<button type="button" class="btn btn-outline-secondary btn-sm" onclick="window.open('CSFromDinersTemplateExcel.asp')"><i class="fa fa-file"></i> Export NA File </button>-->
	<!--<button type="button" class="btn btn-outline-secondary btn-sm" onclick="window.open('TemplateExcel.asp?T=qryCAPSNAToDiners')"><i class="fa fa-file"></i> Excel </button>-->
	<!--<button type="button" class="btn btn-primary btn-xs" onclick="window.open('TemplateExcel.asp?T=qryCAPSNAToDiners&w=<% ="WHERE FileSeqNum = " & Request.QueryString("BatchNo") & ""%>')"><i class="fa fa-file"></i> View NA File in Excel </button> -->
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
        
      DisplaySummary(strCardType)
        
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
<!-- #Include file=../../CC/CAPSFooter.asp -->
</body>
</html>

<%
Sub DisplayTableDetails()

Dim strWhere

If Not IsEmpty(Request.QueryString("BatchNo")) Then
	If IsNull(Request.QueryString("BatchNo")) or Request.QueryString("BatchNo")= "" Then 
		strWhere = "WHERE CardType = '" & strCardType & "' AND [BatchNumber] = 0"
	Else
		strWhere = "WHERE CardType = '" & strCardType & "' AND FileSeqNum = " & Request.QueryString("BatchNo") & ""
	
		If Request.QueryString("BatchNo") = 0 Then strWhere = strWhere & " OR BatchNumber IS NULL"
	End If
Else
	strWhere = "WHERE CardType = '" & strCardType & "' AND [BatchNumber] = 0 AND [Status] <> 'Deleted'"
End If

Response.Write strWhere 

objRS.Open "SELECT TOP 50 * FROM qryCAPSNAToDiners "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no NA To Diners data ready for export</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Files will be added as part of the admin and load functions and will appear here when appropriate</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover table-compact"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing NA To Diners Data ready for export</th></tr>" & _
		        "<tr><th Style=""width:20px;"">NA To Diners ID</th>" & _
				"<th>EmployeeID</th>" & _
				"<th>Record Type</th><th>Record Text</th>" & _	
		        "<th>Status</th>" & _
	 	        "<th>Batch Number</th>" & _	
	 	        "<th>Date Updated</th>" & _
		        "<th>Updated By</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR class='clickable-row' data-href='../DisplayDataset.asp?tbl=qryCAPSNAToDiners&W=WHERE NAToDinersID=" & objRS("NAToDinersID") & "' style=""cursor: pointer;""><TD>" & objRS("NAToDinersID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("EmployeeID") & "</TD><TD style=""text-align:center"">" & objRS("RecordType") & "</TD><TD style=""text-align:center"">" & objRS("RecordText") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("Status") & "</TD><TD style=""text-align:center"">" & objRS("BatchNumber") & "</TD><TD style=""text-align:center"">" & objRS("DateUpdated") & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS("UpdatedByName") & "</TD>" & _
			                    "</TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close

        Response.Write "</table>"
		
End Sub

Sub DisplaySummary(strCardType)

Dim lngCards
Dim lngEmployees
Dim lngBatchNo
Dim lngDTC
Dim lngCMC
Dim lngOther
Dim lngTotalRecords
Dim lngBatchNo1

Dim dteBatchDateFormat
Dim srStatus1
Dim strDateUpdated
Dim dteDateLoaded
Dim strAction
Dim strView
Dim strStatus

'Response.Write "Card Type2 = " & strCardType & "<BR>"


	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		'Change the select depending onthe Card Type currently selected
		If strCardType = "DTC" Then
			objRS.Open "SELECT TOP 6 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'NAFile' ORDER BY FileSeqNum DESC",objCon
		Else
			objRS.Open "SELECT TOP 6 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'NAFileDPC' ORDER BY FileSeqNum DESC",objCon
		End If
		
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""11"" style=""text-align:left"">There is no " & strCardType & " NA File data exported (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange('" & strCardType & "');""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-compact mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""11"" style=""text-align:left""><b>" & strCardType & "</b> NA File Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange('" & strCardType & "');"" />)</th></tr>" & _
		        "<th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Cards</th>" & _	
		        "<th>DTC</th>" & _
	 	        "<th>CMC</th>" & _	
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th>" & _
				"<th>View</th></tr>"
				
		    End If
		    
		    Do until objRS.eof
				
				If IsNull(objRS("Status")) Then
					strStatus = ""
				Else
					strStatus = objRS("Status")
				End If
				
				If strStatus = "Imported" Then
					strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='../NATransactions.asp?CardType=" & strCardType & "&Action=Process&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to Process the NA File and update changes to cards in CAPS from the CS File loaded " & objRS("DateLoaded") & """>Process</button>"
					'strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='UploadANZ.asp?Action=Process&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to Process the ANZ File and update changes to cards in CAPS from the ANZ File loaded " & objRS("DateLoaded") & """><i class=""fa fa-cogs""></i> Process</button>"
				Else
					strAction = ""
				End If
				
				If IsNull(objRS("DateLoaded")) Then
					dteDateLoaded = ""
				Else
					dteDateLoaded = FormatDateTime(objRS("DateLoaded"),vbShortDate)
				End If
				
				'Create the View button detail
				strView = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='NATransactions.asp?CardType=" & strCardType & "&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to view details of the NA File exported " & objRS("DateLoaded") & """>View</button>"
				
				
				Response.Write "<TR><TD><a href=""NATransactions.asp?CardType=" & strCardType & "&BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center"">" & objRS("CardCount") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("DTCCount") & "</TD><TD style=""text-align:center"">" & objRS("CMCCount") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("Status") & "</TD><TD style=""text-align:center"" Title=""" & objRS("DateLoaded") & """>" & dteDateLoaded & "</TD><TD style=""text-align:center"">" & strView & "</TD></TR>"
    			objRS.Movenext			
		    Loop
    			
			
								
	    objRS.Close

        Response.Write "</table>"
		
End Sub


Sub DisplaySummary2()

Dim lngCards
Dim lngEmployees
Dim lngBatchNo
Dim lngDTC
Dim lngCMC
Dim lngOther
'Dim strCardType
Dim lngTotalRecords
Dim lngBatchNo1

Dim dteBatchDateFormat
Dim srStatus1
Dim strDateUpdated

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		objRS.Open "SELECT TOP 20 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'CSToDiners' ORDER BY FileSeqNum DESC",objCon
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no NA To Diners data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange('" & strCardType & "');""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""8"" style=""text-align:left""> NA To Diners Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange('" & strCardType & "');"" />)</th></tr>" & _
		        "<th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>Total Cards</th><th>Total Employees</th>" & _	
		        "<th>DTC</th>" & _
	 	        "<th>CMC</th>" & _	
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th></tr>" 
				
		    End If
		    
		    Do until objRS.eof
					
				Response.Write "<TR><TD><a href=""ExportCS.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center"">" & objRS("RecordCount") & "</TD><TD style=""text-align:center"">" & objRS("CardCount") & "</TD><TD style=""text-align:center"">" & objRS("EmployeeCount") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("DTCCount") & "</TD><TD style=""text-align:center"">" & objRS("CMCCount") & "</TD><TD style=""text-align:center"">" & objRS("Status") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("DateLoaded") & "</TD></TR>"
    			objRS.Movenext			
		    Loop
    			
			
								
	    objRS.Close

        Response.Write "</table>"
		
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
Dim strFileNameDefault

'Get the System Parameter for the fileName
If strCardType = "DTC" Then
	strFileNameDefault = GetSystemAdmin("NAFileStart")
Else
	strFileNameDefault = GetSystemAdmin("NAFileStartDPC")
End If

'Exit the procedure if there is no file name start in system settings
If IsNull(strFileNameDefault) Then
	Response.Write "<div class=""alert alert-danger"" role=""alert"" style=""position: absolute; top:0px; left:0px; z-index:100;"">Error! NA File Start System Parameter (NAFileStart) for " & strCardType & " invalid or empty/blank: " & strFileNameDefault & ". Check System Parameters.</div>"			

	Exit Sub
End If

	
Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Diners\DinersTo"

	Set objFolder = objFSO.GetFolder(objStartFolder)	
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"" Title=""Final Export Location for NA File: " & strServer & """>Files Exported <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be sent to Diners.""></i></th></tr>"
			
	For Each objFile in colFiles
		'Check the file name start to see if it is relevant to the Card Type currently selected
		If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) Then
		
			intCount = intCount + 1
			
			If intCount < 6 Then
				If IsNull(objFile.Name) or objFile.Name = "" Then
					strFile = ""
				Else
					strFile = Left(objFile.Name,8)
					strExtension = objFSO.GetExtensionName(objStartFolder & "/" & objFile.Name)
				End If
				
				Response.Write "<TR><TD>" & strFile & "..." & strExtension & "</TD></TR>"
			End If
		'---End check of file name start for Card Type
		End If
		
	Next
	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"
	 
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


Public Sub RemoveNARecord(lngNAToDinersID, strEmployeeID, strStatus)

Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSNAFileRemoveCard"

			.Parameters.Append objCmd.CreateParameter("NAToDinersID", adVarChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("NAFileRemoveOutput", adInteger, adParamOutput)
			
			.Parameters("NAToDinersID") = lngNAToDinersID
			.Parameters("UpdatedBy") = Session("UserID")
			.Parameters("Status") = strStatus
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("NAFileRemoveOutput") 
	 
		If intRecord = 0 Then
			If strStatus = "Deleted" Then
				Response.Write "<div class=""alert alert-danger"" role=""alert"">Application for " & strEmployeeID & " NOT Removed from NA File! An Error has occurred. See System Admin with NA File ID: " & lngNAToDinersID & " </div>"
			Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">Application for " & strEmployeeID & " NOT Added to the NA File! An Error has occurred. See System Admin with NA File ID: " & lngNAToDinersID & " </div>"
			End If
		Else
			If strStatus = "Deleted" Then
				Response.Write "<div class=""alert alert-success"" role=""alert"">Application for " & strEmployeeID & " REMOVED from the NA file!</div>"
			Else
				Response.Write "<div class=""alert alert-success"" role=""alert"">Application for " & strEmployeeID & " ADDED to the NA file!</div>"
			End If
		End If
		
	
End Sub

Function PadDigits(val, digits)
  PadDigits = Right(String(digits,"0") & val, digits)
End Function

Set objRS = Nothing
Set objCon = Nothing

 %>


