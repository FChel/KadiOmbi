<!-- #Include file=../../CC/CAPSHeader.asp -->
<!-- #Include file=../../ADOVBS.inc -->
<!-- #include file="../../CC/CAPSFunctions.asp" -->
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
Dim strExcelLink
Dim strTitle
Dim strQuery

Dim strUniqueID
Dim strBank

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

If Not IsEmpty(Request.QueryString("PMFileSelect")) Then
	Session("PMFileSelect") = Request.QueryString("PMFileSelect")
End If

strBank = "NAB"

If Not IsEmpty(Request.QueryString("Bank")) Then
	strBank = Request.QueryString("Bank")
End If

'If theExport button has been clicked then call the procedure to export the detail to the server folder
If Request.QueryString("Action") = "ExportPM" Then
	Call ExportPMFile()
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

function loadNA(varDetails) {
//  var xhttp = new XMLHttpRequest();
//  xhttp.onreadystatechange = function() {
//    if (this.readyState == 4 && this.status == 200) {
//     document.getElementById("NADetail").innerHTML = this.responseText;
//    }
//  };
 // xhttp.open("GET", "../../CC/AJAX/GetNAFile.asp", true);
//  xhttp.send();
  //document.getElementById('LoadButton').disabled=true;
  document.getElementById("NADetail").innerHTML = varDetails
}


function DatePickChange() {
	self.location="ExportNA.asp?FileDate=" + document.getElementById("CSDate").value;
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
     <div class="modal-dialog modal-medium modal-dialog-right modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  PM File Detail To Be Exported
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="NADetail">
               
				  
         
            </div>
			<div class="modal-body" id="ExportConfirm">
				Continue Export?
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-primary" onClick='window.location="ExportPM.asp?Action=ExportPM"'><i class="fa fa-check"></i> Yes</button>
				<button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
			</div>
		</div>
	</div>
 </div>
 
 
<main class="main py-3">
    <div class="container">
<form action="ExportPM.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	 <div class="row" id="basic-table">
  <div class="col-4">
    <div class="card">
     <div class="card-header">
          <h5 class="card-title"><img src="../../images/CMS.png" height="40px" width="60px" title="Diners"> ProMaster (CMS) File Export</h5>
        </div>
      <div class="card-content">
        <div class="card-body">
			
			  
			<%Call LoadButtons()%>
			
		</div>

	<!--
	<fieldset class="form-group">
		<button type="button" class="btn btn-outline-secondary btn-sm" data-toggle="modal" data-target="#NAModal" HREF="#" onClick="loadNA();"><i class="fa fa-file"></i> Export </button>
		<button type="button" class="btn btn-outline-secondary btn-sm" onClick="self.location='../ExportPM.asp'"><i class="fa fa-file"></i> View Details </button>
		<!--<button type="button" class="btn btn-outline-secondary btn-sm" onclick="window.open('CSFromDinersTemplateExcel.asp')"><i class="fa fa-file"></i> Export NA File </button>-->
	<!--	<button type="button" class="btn btn-outline-secondary btn-sm" onclick="window.open('TemplateExcel.asp?T=qryCAPSNAToDiners')"><i class="fa fa-file"></i> Excel </button>
		<!--<button type="button" class="btn btn-primary btn-xs" onclick="window.open('TemplateExcel.asp?T=qryCAPSNAToDiners&w=<% ="WHERE FileSeqNum = " & Request.QueryString("BatchNo") & ""%>')"><i class="fa fa-file"></i> View NA File in Excel </button> -->
	<!--</fieldset>
	-->


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
	  
	  <div class="card-content">
        <div class="card-body">
		
		<%
        
      DisplayFileSummaryProMaster()
        
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
Dim fld
Dim intCount
Dim intRecs
Dim intFields
Dim intTotalRecs

'The query is now set below in the LoadButtons Procedure, so below can be removed
'If Session("PMFileSelect") = "ANZAccountUpdates" Then
'	strQuery = "qryCAPSPMANZUpdateExport"
'End If

'If Session("PMFileSelect") = "DinersAccountUpdates" Then
'	strQuery = "qryCAPSPMDinersUpdateExport"
'End If

'If Session("PMFileSelect") = "NewDinersAccounts" Then 
'	strQuery = "qryCAPSPMExportNewCards"
'End If

'If Session("PMFileSelect") = "AddAccountCards" Then 
'	strQuery = "qryCAPSPMAdditionalCardsExport"
'End If

'If Session("PMFileSelect") = "UserUpdates" Then 
'	strQuery = "qryCAPSPMUpdateExport"
'End If

'If Session("PMFileSelect") = "AccountHolder" Then 
'	strQuery = "qryCAPSPMCATExport"
'End If


'If Not IsEmpty(Request.QueryString("BatchNo")) Then
'	If IsNull(Request.QueryString("BatchNo")) or Request.QueryString("BatchNo")= "" Then 
'		strWhere = "WHERE [BatchNumber] = 0"
'	Else
'		strWhere = "WHERE FileSeqNum = " & Request.QueryString("BatchNo") & ""
'	
'		If Request.QueryString("BatchNo") = 0 Then strWhere = strWhere & " OR BatchNumber IS NULL"
'	End If
'Else
'	strWhere = "WHERE [BatchNumber] = 0"
'End If

'If IsNull(strExcelLink) or strExcelLink = "" Then
'	objRS.Open "SELECT TOP 50 * FROM qryCAPSNAToDiners "  & strWhere,objCon
'Else
'	objRS.Open "SELECT TOP 50 * FROM " & strExcelLink & " "  & strWhere,objCon

'End If

	If IsNull(strQuery) or strQuery = "" Then
	
		Response.Write"<table Class=""table table-compact"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""7"" style=""text-align:left"">No ProMaster (CMS) Export file type selected</B></th></tr>" & _
		        "<tr><td colspan=""7"" style=""text-align:left"">Select a ProMaster (CMS) Export file to view above</td></tr></table>"
	
	Else
	
		strWhere = "SELECT * FROM  " & strQuery & " WITH(NOLOCK) "
		'Response.Write "xx"
	'response.write "SELECT * FROM  " & strQuery & " WITH(NOLOCK) " & strWhere
		objRS.Open strWhere,objCon,3

				If objRS.EOF Then
					Response.Write"<table Class=""table table-compact"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
					"<tr><th colspan=""7"" style=""text-align:left"">There is no " & strTitle & " data ready for export</B></th></tr>" & _
					"<tr><td colspan=""7"" style=""text-align:left"">Files will be added as part of the admin and load functions and will appear here when appropriate</td></tr>"
				Else
					
					'Get the number of fields in the table/query
					intFields = objRS.Fields.count - 1
					
					'Limit the display to 10 fields
					If intFields > 10 Then intFields = 10
					
					 Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
					"<tr><th colspan=""" & intFields + 1 & """ style=""text-align:left"">Sample of Existing " & strTitle & " Data ready for export</th></tr>" & _
					"<tr>"
					
					For intCount = 0 To intFields
						Response.Write "<th>" & objRS.Fields.Item(intCount).name & "</th>"
					Next
					
					Response.Write "</tr>"
					
				End If
				
				Do until objRS.EOF

					Response.Write "<TR class='clickable-row' data-href='../DisplayDataset.asp?tbl=" & strQuery & "' style=""cursor: pointer;""><TD>" & objRS(0) & "</TD>"
						
						For intCount = 1 To intFields
							Response.Write "<TD style=""text-align:center; font-size:14px;"">" & objRS(intCount) & "</TD>"
						Next

						Response.Write "</tr>"
					
					objRS.movenext
					
					intRecs = intRecs + 1
					
					'Only show the first 50 records
					If intRecs > 48 Then 
						
						If Not objRS.EOF Then objRS.Movelast
						intTotalRecs = objRS.Recordcount
					End If
					
				Loop
					
			objRS.Close

			Response.Write "<tr><th colspan=""" & intFields + 1 & """ style=""text-align:right"">Displaying " & intRecs & " of " & intTotalRecs & " Records</th></tr></table>"
		
		'End of nothing selected check
		End If
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

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		objRS.Open "SELECT TOP 6 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = 'PrMaster' ORDER BY FileSeqNum DESC",objCon
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""11"" style=""text-align:left"">There is no ProMaster data exported (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""11"" style=""text-align:left"">ProMaster Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
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
					strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='../ExportPM.asp?Action=Process&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to Process the NA File and update changes to cards in CAPS from the CS File loaded " & objRS("DateLoaded") & """>Process</button>"
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
				strView = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='../ExportPM.asp?FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to view details of the NA File exported " & objRS("DateLoaded") & """>View</button>"
				
				
				Response.Write "<TR><TD><a href=""../ExportPM.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center"">" & objRS("CardCount") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("DTCCount") & "</TD><TD style=""text-align:center"">" & objRS("CMCCount") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("Status") & "</TD><TD style=""text-align:center"" Title=""" & objRS("DateLoaded") & """>" & dteDateLoaded & "</TD><TD style=""text-align:center"">" & strView & "</TD></TR>"
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

Set objFSO = CreateObject("Scripting.FileSystemObject")

	objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\PM\")

	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">Files Exported <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be sent to ProMaster.""></i></th></tr>"
	
	intCount = 0
	
	For Each objFile in colFiles

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
		
	Next
	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"
	 
Set objFSO = Nothing

End Sub

Public Sub LoadButtons()
'Loads the file buttons depending on what has been selected
Dim arrActive(16)
Dim x, a
Dim strTitleSpace
Dim strFileDetails
Dim strBankButton

'Clear all of the active variables so it can be set below
For x = 0 to 16
	arrActive(x) = ""
Next 

Response.Write Session("PMFileSelect")

'If the user has selected a PM file then make the section active
If Session("PMFileSelect") = "ANZAccountUpdates" Then 
	arrActive(0) = "active"
	strQuery = "qryCAPSPMANZUpdateExport"
End If

If Session("PMFileSelect") = "DinersAccountUpdates" Then 
	arrActive(1) = "active"
	strQuery = "qryCAPSPMDinersUpdateExport"
End If

If Session("PMFileSelect") = "DPCDinersAccountUpd" Then 
	arrActive(9) = "active"
	strQuery = "qryCAPSPMDinersUpdateExportDPC"
End If

If Session("PMFileSelect") = "DTCNewDinersAccounts" Then 
	arrActive(2) = "active"
	strQuery = "qryCAPSPMExportNewCards"
End If

If Session("PMFileSelect") = "DPCNewDinersAccounts" Then 
	arrActive(8) = "active"
	strQuery = "qryCAPSPMExportNewCardsDPC"
End If

If Session("PMFileSelect") = "AddAccountCards" Then 
	arrActive(3) = "active"
	strQuery = "qryCAPSPMAdditionalCardsExport"
End If

If Session("PMFileSelect") = "DPCAddAccountCards" Then 
	arrActive(10) = "active"
	strQuery = "qryCAPSPMAdditionalCardsExportDPC"
End If


If Session("PMFileSelect") = "UserUpdates" Then 
	arrActive(4) = "active"
	strQuery = "qryCAPSPMUpdateExport"
End If

If Session("PMFileSelect") = "AccountHolder" Then 
	arrActive(5) = "active"
	strQuery = "qryCAPSPMCATExport"
End If

If Session("PMFileSelect") = "NewANZAccounts" Then 
	arrActive(6) = "active"
	strQuery = "qryCAPSPMExportNewCardsANZ"
End If

If Session("PMFileSelect") = "CDMCCancelUsers" Then 
	arrActive(7) = "active"
	strQuery = "qryCAPSPMExportCDMCCancels"
End If

If Session("PMFileSelect") = "DTCNewNABAccounts" Then 
	arrActive(11) = "active"
	strQuery = "qryCAPSPMExportNewCardsNAB"
End If

If Session("PMFileSelect") = "DTCNewNABLodgeAccts" Then 
	arrActive(12) = "active"
	strQuery = "qryCAPSPMExportNewCardsNABLodge"
End If

If Session("PMFileSelect") = "DTCNewNABLodgeAccts" Then 
	arrActive(12) = "active"
	strQuery = "qryCAPSPMExportNewCardsNABLodge"
End If

If Session("PMFileSelect") = "DPCNewNABAccounts" Then 
	arrActive(13) = "active"
	strQuery = "qryCAPSPMExportNewCardsDPCNAB"
End If

If Session("PMFileSelect") = "DTCNABAccountUpdates" Then 
	arrActive(14) = "active"
	strQuery = "qryCAPSPMDinersUpdateExportNAB"
End If

If Session("PMFileSelect") = "DTCNABLodgeAccntUpdt" Then 
	arrActive(15) = "active"
	strQuery = "qryCAPSPMDinersUpdateExportNABLodge"
End If

If Session("PMFileSelect") = "DPCNABAccountUpdates" Then 
	arrActive(16) = "active"
	strQuery = "qryCAPSPMDinersUpdateExportDPCNAB"
End If

	If Session("PMFileSelect") = "" Then
		strTitle = "Select a ProMaster File"
	Else
		strTitleSpace = Session("PMFileSelect")
		strTitle = ""
		
		'Loop through the title characters to add a space before each Capital letter, for user niceness!
		For a = 1 to Len(strTitleSpace)
			'If the character is Uppercase then add a space before the character
			If Asc(Mid(strTitleSpace,a,1)) < 91 Then
				If Asc(Mid(strTitleSpace,a+1,1)) < 91 Then
					strTitle = strTitle & Mid(strTitleSpace,a,1)
				Else
					strTitle = strTitle & " " & Mid(strTitleSpace,a,1)
				End If
			Else
				strTitle = strTitle & Mid(strTitleSpace,a,1)
			End If
		Next
		'strTitle = Session("PMFileSelect") & " selected"
	End If
	
	If Session("PMFileSelect") = "" Then
	Else
	'Open the related recordset and count the number of records to display when confirming on Click of the Export Button
	'Response.Write "SELECT Count(*) As CountRecs FROM " & strQuery & " WITH(NOLOCK)"
	objRS.Open "SELECT Count(*) As CountRecs FROM " & strQuery & " WITH(NOLOCK)",objCon
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		       strFileDetails = "No Records for " & Session("PMFileSelect") & " To be Exported"
		    Else
				strFileDetails = "<b>" & objRS("CountRecs") & "</b> Records for <b>" & Session("PMFileSelect") & "</b> To be Exported"
		    End If
			
	objRS.Close
	End If
	
	If strBank = "NAB" Then
		strBankButton = "<button class=""btn btn-outline-secondary btn-sm "" onClick=""self.location='ExportPM.asp?Bank=Diners'"" type=""button"" id=""Bank""><i class=""fa fa-building""></i> NAB</button> &nbsp;"
	Else
		strBankButton = "<button class=""btn btn-outline-secondary btn-sm "" onClick=""self.location='ExportPM.asp?Bank=NAB'"" type=""button"" id=""Bank""><i class=""fa fa-building""></i> Diners</button> &nbsp;"
	End If
	
	'Write the header for the buttons section
	'Response.write "<p class=""card-text"">" & strTitle & "</p><div class=""panel-header text-right my-auto"">" & _
	Response.write "<div style=""font-size:12px; font-weight:bold;"">" & strTitle & "</div><div class=""panel-header text-right my-auto"">" & _
		strBankButton & _
		"<button type=""button"" class=""btn btn-outline-secondary btn-sm"" onclick=""window.open('TemplateExcel.asp?T="" & strQuery & ""')""><i class=""fa fa-file-excel""></i> Excel </button> &nbsp;" & _
		"<button type=""button"" name=""LoadButton"" id=""LoadButton"" class=""btn btn-outline-secondary btn-sm"" data-toggle=""modal"" data-target=""#NAModal"" HREF=""#"" onClick=""loadNA('" & strFileDetails & "');""><i class=""fa fa-file""></i> Export </button>" & _
		"<h4></h4></div><div class=""panel-content"">"

	If strBank="Diners" Then
	Response.Write "<a href=""ExportPM.asp?PMFileSelect=ANZAccountUpdates"" class=""section-link " & arrActive(0) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">ANZ Account Updates</span><span class=""description"">Account Updates</span></div></a>"
	
	Response.Write "<a href=""ExportPM.asp?PMFileSelect=DinersAccountUpdates"" class=""section-link " & arrActive(1) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">Diners Account Updates</span><span class=""description"">Account Updates</span></div></a>"
	
	Response.Write "<a href=""ExportPM.asp?PMFileSelect=DPCDinersAccountUpd"" class=""section-link " & arrActive(9) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">DPC Diners Account Updates</span><span class=""description"">DPC Account Updates</span></div></a>"
		
	Response.Write "<a href=""ExportPM.asp?PMFileSelect=DTCNewDinersAccounts"" class=""section-link " & arrActive(2) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">DTC New Diners Accounts</span><span class=""description"">New Diners Accounts/Employees</span></div></a>"
		
	Response.Write "<a href=""ExportPM.asp?PMFileSelect=DPCNewDinersAccounts"" class=""section-link " & arrActive(8) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">DPC New Diners Accounts</span><span class=""description"">New Diners Accounts/Employees</span></div></a>"
		
	Response.Write "<a href=""ExportPM.asp?PMFileSelect=NewANZAccounts"" class=""section-link " & arrActive(6) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">New ANZ Accounts</span><span class=""description"">New ANZ Accounts</span></div></a>"

	Response.Write "<a href=""ExportPM.asp?PMFileSelect=AddAccountCards"" class=""section-link " & arrActive(3) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">Additional Account Cards</span><span class=""description"">Additional Cards for Diners Accounts</span></div></a>"
	
	'''New DPC Additional Cards Added Sep 2023
	Response.Write "<a href=""ExportPM.asp?PMFileSelect=DPCAddAccountCards"" class=""section-link " & arrActive(10) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">DPC Additional Account Cards</span><span class=""description"">Additional Cards for Diners Accounts DPC</span></div></a>"
		
	'Response.Write "<a href=""ExportPM.asp?PMFileSelect=UserUpdates"" class=""section-link " & arrActive(4) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
	'	"<div class=""content""><span class=""title"">User Updates</span><span class=""description"">CMS User Updates</span></div></a>"		
	
	Response.Write "<a href=""ExportPM.asp?PMFileSelect=AccountHolder"" class=""section-link " & arrActive(5) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">Account Holder Transfers</span><span class=""description"">Account Holder Updates</span></div></a>"
	
	
	
	Response.Write "<a href=""ExportPM.asp?PMFileSelect=CDMCCancelUsers"" class=""section-link " & arrActive(7) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">CDMC Cancel Users</span><span class=""description"">Cancel Users in ProMaster no longer on CDMC</span></div></a>"
		
				'<a href="ExportPM.asp?ApplicationChecks=CMS" class="section-link Active"><div class="status"><i class="fa fa-cogs"></i></div>
				'	<div class=""content""><span class="title">CMS User Account</span><span class="description">Diners updates</span></div></a>
	Else

		Response.Write "<a href=""ExportPM.asp?PMFileSelect=DTCNewNABAccounts"" class=""section-link " & arrActive(11) & """><div class=""status""><i class=""fa fa-plane""></i></div>" & _
		"<div class=""content"" title=""Currently Selected: " & strQuery & """><span class=""title"">DTC New NAB Accounts</span><span class=""description"">New NAB Accounts/Cards</span></div></a>"
		
		Response.Write "<a href=""ExportPM.asp?PMFileSelect=DTCNewNABLodgeAccts"" class=""section-link " & arrActive(12) & """><div class=""status""><i class=""fa fa-credit-card""></i></div>" & _
		"<div class=""content"" title=""Currently Selected: " & strQuery & """><span class=""title"">DTC New NAB Lodge Accounts</span><span class=""description"">New NAB Lodge Accounts/Cards</span></div></a>"
		
		Response.Write "<a href=""ExportPM.asp?PMFileSelect=DPCNewNABAccounts"" class=""section-link " & arrActive(13) & """><div class=""status""><i class=""fa fa-dollar-sign""></i></div>" & _
		"<div class=""content"" title=""Currently Selected: " & strQuery & """><span class=""title"">DPC New NAB Accounts</span><span class=""description"">New NAB DPC Accounts/Cards</span></div></a>"

		Response.Write "<a href=""ExportPM.asp?PMFileSelect=DTCNABAccountUpdates"" class=""section-link " & arrActive(14) & """><div class=""status""><i class=""fa fa-plane""></i></div>" & _
		"<div class=""content"" title=""Currently Selected: " & strQuery & """><span class=""title"">DTC NAB Account Updates</span><span class=""description"">NAB DTC Account Updates</span></div></a>"
		
		Response.Write "<a href=""ExportPM.asp?PMFileSelect=DTCNABLodgeAccntUpdt"" class=""section-link " & arrActive(15) & """><div class=""status""><i class=""fa fa-credit-card""></i></div>" & _
		"<div class=""content"" title=""Currently Selected: " & strQuery & """><span class=""title"">DTC NAB Lodge Account Updates</span><span class=""description"">NAB DTC Lodge Account Updates</span></div></a>"

		Response.Write "<a href=""ExportPM.asp?PMFileSelect=DPCNABAccountUpdates"" class=""section-link " & arrActive(16) & """><div class=""status""><i class=""fa fa-dollar-sign""></i></div>" & _
		"<div class=""content"" title=""Currently Selected: " & strQuery & """><span class=""title"">DPC NAB Account Updates</span><span class=""description"">NAB DPC Account Updates</span></div></a>"

		'''New DPC Additional Cards Added Sep 2023
		Response.Write "<a href=""ExportPM.asp?PMFileSelect=DPCAddAccountCards"" class=""section-link " & arrActive(10) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">DPC Additional Account Cards</span><span class=""description"">Additional Cards for Diners Accounts DPC</span></div></a>"
	
		Response.Write "<a href=""ExportPM.asp?PMFileSelect=AccountHolder"" class=""section-link " & arrActive(5) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">Account Holder Transfers</span><span class=""description"">Account Holder Updates</span></div></a>"
		
		Response.Write "<a href=""ExportPM.asp?PMFileSelect=CDMCCancelUsers"" class=""section-link " & arrActive(7) & """><div class=""status""><i class=""fa fa-cogs""></i></div>" & _
		"<div class=""content""><span class=""title"">CDMC Cancel Users</span><span class=""description"">Cancel Users in ProMaster no longer on CDMC</span></div></a>"

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

Public Sub ExportPMFile()
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
Dim strQuery
Dim intFieldCount
Dim strLine

Dim fldField
Dim strFilePathOnly
	
'Get the Default file location for the server then add the filepath and name for the NA File
strFilePath = GetSystemAdmin("ServerFilePath")

'Get the Filename start Depending on the file

'40
If Session("PMFileSelect") = "ANZAccountUpdates" Then 
	strFileNameStart = GetSystemAdmin("PMExportANZUpdates")
	strQuery = "qryCAPSPMANZUpdateExport"
	strUniqueID = 1
End If

'41
If Session("PMFileSelect") = "DinersAccountUpdates" Then 
	strFileNameStart = GetSystemAdmin("PMExportDinersAccountUpdates")
	strQuery = "qryCAPSPMDinersUpdateExport"
	strUniqueID = 1
End If

'41 ----- Haven't updated
If Session("PMFileSelect") = "DPCDinersAccountUpd" Then 
	strFileNameStart = GetSystemAdmin("PMExportDinersAccountUpdatesDPC")
	strQuery = "qryCAPSPMDinersUpdateExportDPC"
	strUniqueID = 1
End If

'39
If Session("PMFileSelect") = "DTCNewDinersAccounts" Then 
	strFileNameStart = GetSystemAdmin("PMExportDinersNewCards")
	strQuery = "qryCAPSPMExportNewCards"
	strUniqueID = 1
End If

If Session("PMFileSelect") = "DPCNewDinersAccounts" Then 
	strFileNameStart = GetSystemAdmin("PMExportDinersNewCardsDPC")
	strQuery = "qryCAPSPMExportNewCardsDPC"
	strUniqueID = 1
End If

'42
If Session("PMFileSelect") = "AddAccountCards" Then 
	strFileNameStart = GetSystemAdmin("PMExportPaymentCards")
	strQuery = "qryCAPSPMAdditionalCardsExport"
	strUniqueID = 1
End If

If Session("PMFileSelect") = "DPCAddAccountCards" Then 
	strFileNameStart = GetSystemAdmin("PMExportPaymentCardsDPC")
	strQuery = "qryCAPSPMAdditionalCardsExportDPC"
	strUniqueID = 1
End If


If Session("PMFileSelect") = "UserUpdates" Then 
	strFileNameStart = GetSystemAdmin("PMExportANZUpdates")
	strQuery = "qryCAPSPMUpdateExport"
	strUniqueID = 1
	
	'Not yet complete so exit and display an error message
	Response.Write "<div class=""alert alert-danger"" role=""alert"">ProMaster Export File " & Session("PMFileSelect") & " Not Exported as set-up is not complete. See System Admin</div>"
	exit sub
End If

'43
If Session("PMFileSelect") = "AccountHolder" Then 
	strFileNameStart = GetSystemAdmin("PMExportAccountHolders")
	strQuery = "qryCAPSPMCATExport"
	strUniqueID = 2
End If

If Session("PMFileSelect") = "NewANZAccounts" Then 
	strFileNameStart = GetSystemAdmin("PMExportANZNewCards")
	strQuery = "qryCAPSPMExportNewCardsANZ"
	strUniqueID = 1
End If

If Session("PMFileSelect") = "CDMCCancelUsers" Then 
	strFileNameStart = GetSystemAdmin("PMCDMCCancelUsers")
	strQuery = "qryCAPSPMExportCDMCCancels"
	strUniqueID = 2
End If

''''''Start of the new NAB file queries and details
'''DTC NAB New Cards
If Session("PMFileSelect") = "DTCNewNABAccounts" Then 
	strFileNameStart = GetSystemAdmin("PMExportNABDTCNewCards")
	strQuery = "qryCAPSPMExportNewCardsNAB"
	strUniqueID = 1
End If

'''DTC NAB Lodge New Cards
If Session("PMFileSelect") = "DTCNewNABLodgeAccts" Then 
	strFileNameStart = GetSystemAdmin("PMExportNABDTCLodgeNewCards")
	strQuery = "qryCAPSPMExportNewCardsNABLodge"
	strUniqueID = 1
End If

'''DPC NAB New Cards
If Session("PMFileSelect") = "DPCNewNABAccounts" Then 
	strFileNameStart = GetSystemAdmin("PMExportNABDPCNewCards")
	strQuery = "qryCAPSPMExportNewCardsDPCNAB"
	strUniqueID = 1
End If

'''DTC NAB Existing Card Updates
If Session("PMFileSelect") = "DTCNABAccountUpdates" Then 
	strFileNameStart = GetSystemAdmin("PMExportNABDTCCardUpdate")
	strQuery = "qryCAPSPMDinersUpdateExportNAB"
	strUniqueID = 1
End If

'''DTC NAB Lodge Existing Card Updates
If Session("PMFileSelect") = "DTCNABLodgeAccntUpdt" Then 
	strFileNameStart = GetSystemAdmin("PMExportNABDTCLodgeCardUpdate")
	strQuery = "qryCAPSPMDinersUpdateExportNABLodge"
	strUniqueID = 1
End If

'''DPC NAB Existing Card Updates
If Session("PMFileSelect") = "DPCNABAccountUpdates" Then 
	strFileNameStart = GetSystemAdmin("PMExportNABDPCCardUpdate")
	strQuery = "qryCAPSPMDinersUpdateExportDPCNAB"
	strUniqueID = 1
End If


strNextFileNumber = GetSystemAdmin(Session("PMFileSelect"))

'Pad the number out to 6 digits
'strNextFileNumber = PadDigits(strNextFileNumber,6)
'strNextFileNumber = PadDigits(strNextFileNumber,6)

'Get the Last part of the file name (dynamic based on the current date)
If Session("PMFileSelect") = "AccountHolder" Then 
	strFileDateTime = PadDigits(Right(Year(Now()),4),4) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2)
	strFileDateTimeSec = PadDigits(Right(Year(Now()),4),4) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) &  "_" & PadDigits(Hour(Now()),2) & PadDigits(Minute(Now()),2) & Second(Now())

Else
	strFileDateTime = PadDigits(Right(Year(Now()),4),4) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2)
	strFileDateTimeSec = PadDigits(Right(Year(Now()),4),4) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) &  "_" & PadDigits(Hour(Now()),2) & PadDigits(Minute(Now()),2) & PadDigits(Second(Now()),2)
End If

'Compile the File name and path from the variables above
strFilePath = strFilePath & "\Admin\CAPSAdmin\Attachments\PM\" & strFileNameStart & strFileDateTimeSec & ".txt"
'strFilePath = strFilePath & "\Admin\CAPSAdmin\Attachments\Diners\DinersTo\" & strFileNameStart & PadDigits(Right(Year(Now()),2),2) & PadDigits(Month(Now()),2) & PadDigits(Day(Now()),2) & ".txt"

'Get the file path only for the file transfer function
strFilePathOnly = strFilePath & "\Admin\CAPSAdmin\Attachments\PM\"

'response.write 	"lngBatchNumber=" & strNextFileNumber & " strNextCSFileNumber= " &  strNextCSFileNumber

'Get the FileLoad details
'lngBatchNumber = GetFileLoadID("NAFile",strNextFileNumber,"")

'If IsNull(lngBatchNumber) OR lngBatchNumber = "" then
'	lngBatchNumber = strNextCSFileNumber
'End If

Dim objFSO
Dim strFileName
Dim intUniqueID
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

'Open the text file
Dim objTextStream

Dim strFieldValue

Response.Write strFilePath

	'Open the relevant recordset of records to be exported
	'Response.Write "SELECT * FROM " & strQuery & " WITH(NOLOCK)"
	objRS.Open "SELECT * FROM " & strQuery & " WITH(NOLOCK)",objCon

		If objRS.EOF Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No records to write to the PM Export File " & Session("PMFileSelect") & "</div>"
		Else
		'response.write strFilePath
			'Response.Write "strFileType = " & Session("PMFileSelect") & "<BR>"
			'Response.Write "strFilePath = " & strFilePath & "<BR>"
			'response.write strQuery
			'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
			'response.write "Session(PMFileSelect)=" & Session("PMFileSelect") & ", strFileNameStart=" & strFileNameStart & strFileDateTimeSec & ".txt" 
			lngFileLoadID = SaveFileLoadID (Session("PMFileSelect"),strFileNameStart & strFileDateTimeSec & ".txt", strFilePath,-1,0,0,0,0,0,0,0,strFileDateTimeSec,0,"Exported",Session("UserID"),"N")
	'lngFileLoadID = 0		
			Set objTextStream = objFSO.OpenTextFile(strFilePath, fsoForWriting, True)
			'objTextStream.WriteLine "H CS" & strFileDateTimeSec & strNextFileNumber
			
			'format(Now(),"yyyymmddhhmmss")+Format(DLookup("Next_File_No","Next_File_No","File_Type = 'NA_Out'"),"000000");\
			
			'Write each record to the text file
			Do Until objRS.EOF
			
				intRecordCount = intRecordCount + 1
				strLine = ""
			
				intUniqueID = objRS(strUniqueID)					
				
				intFieldCount = 0
				
				'Loop through each field and check for errors
				For each fldField in objRS.fields
			
					intFieldCount = intFieldCount + 1
					'Temporary for the Training Report to show formatting (change cell colours)
				'	If strTable = "qryCAPSTrainingReport" AND intFieldCount = 33 Then
				'		strDetail = strDetail & "<td style=""color:white; background-color:blue;"">" & fldField.value & "</td>"
						
				'	ElseIf strTable = "qryCAPSTrainingReport" AND intFieldCount = 32 Then
				'		strDetail = strDetail & "<td style=""color:red; font-weight:bold;"">" & fldField.value & "</td>"
				'	Else
				'		strDetail = strDetail & "<td>" & fldField.value & "</td>"
				'	End If
				'	
					'Only count the fields if this is the first record
				'	If intFields = 0 Then x = x + 1
					
					strFieldValue = fldField.value
					strError = strError & CheckForNull(fldField.value,fldField.name,intRecordCount)
					'strLine = strLine & trim(fldField.value) & "" & chr(9)
					
					'Make the Credit Limit a value with two decimals and no comma delimeter
					If intFieldCount = 8 AND (strQuery = "qryCAPSPMExportNewCards" OR strQuery = "qryCAPSPMDinersUpdateExport") Then
						strFieldValue = fldField.value
						strFieldValue = strFieldValue & ".00"
						strLine = strLine & strFieldValue & "" & chr(9)
						'Remove trim as the fields have some leading and trailing spaces in the SQL view
						'strLine = strLine & trim(strFieldValue) & "" & chr(9)
					Else	
						strLine = strLine & strFieldValue & "" & chr(9)
						'Remove trim as the fields have some leading and trailing spaces in the SQL view
						'strLine = strLine & trim(strFieldValue) & "" & chr(9)
					End If
					
				Next
			
				'Remove the last comma
				strLine = Left(strLine,Len(strLine)-1)
				
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
				
				'Response.write intRecordCount & " = " & objRS(0) & "<BR>"
				
				If IsNull(objRS(0)) = False Then
					'Response.write " " & intRecordCount & " = " & strLine & "<BR>"
					objTextStream.WriteLine strLine
					'objTextStream.WriteLine "D " & PadSpaceLeft(objRS("EIDNo"),10) & PadDigits(objRS("CardNo"),19) & PadSpaceLeft(objRS("CardUpdateInd"),2) & PadSpaceLeft(objRS("CardExpiryDate"),8) & PadSpaceLeft(objRS("CardStatus"),2) & PadSpaceLeft(objRS("Title"),12) & PadDigitsLeft(objRS("Surname"),25) & PadSpaceLeft(objRS("GivenNames"),30) & PadSpaceLeft(objRS("NameOnCard"),26) & PadSpaceLeft(objRS("Address1"),40) & PadSpaceLeft(objRS("Address2"),40) & PadSpaceLeft(objRS("Address3"),40) & PadSpaceLeft(objRS("Suburb"),25) &  PadSpaceLeft(objRS("State"),4) & PadSpaceLeft(objRS("Postcode"),12) &  PadSpaceLeft(objRS("HomePhone"),12) &  PadSpaceLeft(objRS("WorkPhone"),12) & PadSpaceLeft(objRS("MobilePhone"),12) & PadSpaceLeft(objRS("Email"),70) & PadSpaceLeft(objRS("ReportGroup"),8) &  PadSpaceLeft(objRS("CreditLimit"),11)                         
				End If
				
				'Call the procedure to update each record as exported once added to the CS File -- USE the File style Batch Number not FileLoadID
				'Call ExportCSRecord (objRS("CSToDinersID"),lngFileLoadID,intRecordCount)
				
				'Updates records as they are exported (changes Status based on table/query)
				If strQuery <> "qryCAPSPMExportCDMCCancels" Then
					If strQuery <> "qryCAPSPMANZUpdateExport" Then
						''Removed this procedure as there is an error when too many calls are made to the database to update individual records
						Call ExportPMRecord (intUniqueID,strNextFileNumber,intRecordCount)
					End If
				End If
				
			objRS.Movenext
			Loop
			
			'strRecordCount = PadDigits(intRecordCount,6)
			
			'objTextStream.WriteLine "T" & strRecordCount		
			
			
			'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			strFileName = strFileNameStart & strFileDateTimeSec & ".txt"
			
			'''**************************************
			''''UPDATE THIS procedure later to assist with updating TOTAL records exported
			'Call UpdateFileLoadSummary (Session("PMFileSelect"),strNextFileNumber, strFileName, lngFileLoadID)
			
			'Call the procedure to update the table so exported records are marked as exported
			'Call ExportPMRecord()
			
			If strError = "" Then
				Response.Write "<div class=""alert alert-success"" role=""alert"">ProMaster Export File " & Session("PMFileSelect") & " " & strFileNameStart & strFileDateTimeSec & ".txt" & " ADDED to the CS file export folder!</div>"
			Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">ProMaster Export File " & Session("PMFileSelect") & " " & strFileNameStart & strFileDateTimeSec & ".txt" & " has errors : " & strError & "</div>"
			End If
				'Close the file and clean up
				objTextStream.Close
		End If

	objRS.Close
	

	''New Change 28th Oct 2021 -- Update all ANZ records at the end of the process otherwise it times out and does not export all records
	If strQuery = "qryCAPSPMANZUpdateExport" Then
		''Removed this procedure as there is an error when too many calls are made to the database to update individual records
		Call ExportPMRecord (intUniqueID,strNextFileNumber,intRecordCount)
	End If

	'Call the procedure to move the file created to the G/ProMaster Drive
	Call MoveExportFiles(strFilePath, strFileName, strFilePathOnly)

Set objTextStream = Nothing
Set objFSO = Nothing
					
End Sub

'Public Sub ExportPMRecord()
Public Sub ExportPMRecord(strUniqID,lngBatchNumber,x)
'Procedure to Change the Status of CS file records being exported and adds an Audit Log record
Dim intRecord


	'Update records in the File just exported all together at then end of the export (as trying to update each individual record with objCon causes an error
	
	'Get the Filename start Depending on the file
	
	If Session("PMFileSelect") = "ANZAccountUpdates" Then 
		'objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardNumber = '" & strUniqID & "'"
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DPC' AND PMLoadStatus = 'Pending update acc'"
		'Response.Write "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DPC' AND PMLoadStatus = 'Pending'" & "<BR>"
	End If

	If Session("PMFileSelect") = "DinersAccountUpdates" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardNumberShort = '" & strUniqID & "'"
	End If
	
	If Session("PMFileSelect") = "DPCDinersAccountUpd" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardNumberShort = '" & strUniqID & "'"
	End If

	If Session("PMFileSelect") = "DTCNewDinersAccounts" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DTC' AND CardNumberShort = '" & strUniqID & "'"
	End If
	
	If Session("PMFileSelect") = "DPCNewDinersAccounts" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DPC' AND CardNumberShort = '" & strUniqID & "'"
	End If

	If Session("PMFileSelect") = "DPCNewNABAccounts" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DPC' AND CardNumberShort = '" & strUniqID & "'"
	End If

	If Session("PMFileSelect") = "AddAccountCards" Then 
	
		'''****--Add in update for card numbers short and full with leading zeroes so they are marked as Exported
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardNumber = '" & strUniqID & "'"
		'Response.Write "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardNumber = '" & strUniqID & "'" & "<BR>"
	End If
	
	If Session("PMFileSelect") = "DPCAddAccountCards" Then 
	
		'''****--Add in update for card numbers short and full with leading zeroes so they are marked as Exported
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardNumber = '" & strUniqID & "'"
		'Response.Write "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardNumber = '" & strUniqID & "'" & "<BR>"
	End If

	If Session("PMFileSelect") = "UserUpdates" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardNumber = '" & strUniqID & "'"
	End If

	If Session("PMFileSelect") = "AccountHolder" Then 
		objCon.Execute "UPDATE tblCAPSCardAccountTransfer SET Exported = 'Y' WHERE Exported = 'N'"
	End If

	If Session("PMFileSelect") = "NewANZAccounts" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardNumberShort = '" & strUniqID & "'"
	End If
	
	'''New NAB Card Types
	If Session("PMFileSelect") = "DTCNewNABAccounts" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DTC' AND CardNumberShort = '" & strUniqID & "'"
	End If

	'''DTC NAB Lodge New Cards
	If Session("PMFileSelect") = "DTCNewNABLodgeAccts" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DTC' AND CardNumberShort = '" & strUniqID & "'"
	End If

	'''DPC NAB New Cards
	If Session("PMFileSelect") = "DPCNewNABAccounts" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DPC' AND CardNumberShort = '" & strUniqID & "'"
	End If

	'''DTC NAB Existing Card Updates
	If Session("PMFileSelect") = "DTCNABAccountUpdates" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DTC' AND CardNumberShort = '" & strUniqID & "'"
	End If

	'''DTC NAB Lodge Existing Card Updates
	If Session("PMFileSelect") = "DTCNABLodgeAccntUpdt" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DTC' AND CardNumberShort = '" & strUniqID & "'"
	End If

	'''DPC NAB Existing Card Updates
	If Session("PMFileSelect") = "DPCNABAccountUpdates" Then 
		objCon.Execute "UPDATE tblCAPSCard SET PMLoadStatus = 'Exported', PMLoadDate = GetDate() WHERE CardType = 'DPC' AND CardNumberShort = '" & strUniqID & "'"
	End If
	
  '	With objCmd
'
'		.CommandType = 4
'		.CommandText = "spCAPSCSFileExportCard"
'		
'		'Only create the parameters the first time the procedure is created otherwise there will be an error
'		If x = 1 Then
'			.Parameters.Append objCmd.CreateParameter("CSToDinersID", adVarChar, adParamInput,10)
'			.Parameters.Append objCmd.CreateParameter("BatchNumber", adVarChar, adParamInput, 20)
'			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
'			.Parameters.Append objCmd.CreateParameter("CSFileExportOutput", adInteger, adParamOutput)
'		End If	
'	
'		.Parameters("CSToDinersID") = lngCSToDinersID
'		.Parameters("BatchNumber") = lngBatchNumber
'		.Parameters("UpdatedBy") = Session("UserID")
'		
'		.ActiveConnection = objCon
'		 
'	End With
 '  
'	objCmd.Execute        
  
	'Return the result of the Save Function.
'	intRecord = objCmd.Parameters.Item("CSFileExportOutput") 
 
End Sub

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



Public Function CheckForNull(strValue,strField,intRow)
'Function to check for null values in the field passed in and return an error message if there is
	If IsNull(strValue) Then 
		CheckForNull = "</BR> Error in field " & strField & " is Null at row " & intRow & " : "
	Else
		CheckForNull = ""
	End If

End Function


Public Sub DisplayFileSummaryProMaster()
'Procedure to display files on the ProMaster server

exit sub


Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim strExtension
Dim objFSO
Dim objFolder
Dim objFile

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass

Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get the System Parameter for the start of the Training File Location
strServer = GetSystemAdmin("ProMasterFilePath")
'strServer = "\\groupdata.rus.car.drn.defence.mil.au\groupdata\CFO\CFO\CMS Admin\CAPS\Import Files\Training\"

'Get the System Parameter for the Service Account UserName and Password
strUser = GetSystemAdmin("CAPSServiceAccountName")
strPass = GetSystemAdmin("CAPSServiceAccountPassword")
'strUser = "DRN\svc_CAPS_VBMRSN05_Ad"
'strPass = "$dw2zt%2V9D2"

objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

	objStartFolder = strServer & "Caps\Cards\"

'Set objFSO = CreateObject("Scripting.FileSystemObject")

	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\PM\")

	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"" title=""" & objFolder & """>Files Exported <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be sent to ProMaster. - " & objFolder & """></i></th></tr>"
	
	intCount = 0
	
	For Each objFile in colFiles

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
		
	Next
	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"
	 
objNetwork.RemoveNetworkDrive strServer, True, False
	 
Set objFSO = Nothing
Set objNetwork = Nothing

	
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

	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\PM\")

	'Get the Training Data File Starting folder from the System Parameters
	objStartFolder = GetSystemAdmin("ServerFilePath")
							
	'If there is no System Parameter for the Training Data starting folder then set to the VBMRSN05 server
	If IsNull(objStartFolder) Or objStartFolder = "" Then			
		'objStartFolder = objStartFolder & "\Admin\CAPSAdmin\Attachments\PM\"		

		'Write an error to the top of the screen
		'Response.Write "<div class=""container"" style=""position:relative; z-index:100; top:40px; left:40px;""><div class=""alert alert-danger"" role=""alert"" style=""position: absolute; top:40px; left:40px; z-index:100;"">Error! Server path not found: " & strServer & "</div></div>"
		'Write a message in the G Drive Div area
		Response.Write "<div class=""alert alert-danger"" role=""alert"" style=""position: absolute; top:0px; left:0px; z-index:100;"">Error! No folder to Export Files To. Server path not found: " & ServerFilePath & "</div>"
		
	Else
		objStartFolder = objStartFolder & "\Admin\CAPSAdmin\Attachments\PM\"
	End If

On Error Resume Next

	Set objFolder = objFSO.GetFolder(objStartFolder)
	
If Err.Number <>0 Then
	'Write an error to the top of the screen
	'Response.Write "<div class=""container"" style=""position:relative; z-index:100; top:40px; left:40px;""><div class=""alert alert-danger"" role=""alert"" style=""position: absolute; top:40px; left:40px; z-index:100;"">Error! Server path not found: " & strServer & "</div></div>"
	'Write a message in the G Drive Div area
	Response.Write "<div class=""alert alert-danger"" role=""alert"" style=""position: absolute; top:0px; left:0px; z-index:100;"">Error! G Drive path not found: " & objStartFolder & "</div>"			

	Err.Clear
	On Error Goto 0
	Exit Sub
	
End If

On Error Goto 0

	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"" title=""" & objFolder & """>Files Exported <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be sent to ProMaster.""></i></th></tr>"
	
	intCount = 0
	
	For Each objFile in colFiles

		intCount = intCount + 1
		
		If intCount < 6 Then
			If IsNull(objFile.Name) or objFile.Name = "" Then
				strFile = ""
			Else
				strFile = Left(objFile.Name,8)
				strExtension = objFSO.GetExtensionName(objStartFolder & "/" & objFile.Name)
			End If
			
			Response.Write "<TR><TD title=""" & objFile.Name & """>" & strFile & "..." & strExtension & "</TD></TR>"
		End If
		
	Next
	
	 Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr></table>"
	 
Set objFSO = Nothing

End Sub


Public Sub MoveExportFiles(strFilePathFrom, strFileName, strFilePathOnly)
'Procedure to move the exported files to the G Drive once the files have been produced.
Dim objNetwork
Dim strServer
Dim strUser
Dim strPass
'Dim strFileNameDefault
Dim strFileExtension
Dim objFSO
Dim objStartFolder
Dim strFilePathTo
Dim strFileNameNoExt
Dim x

Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")

	'Get the System Parameter for the start of the Training File Location
	strServer = GetSystemAdmin("GDriveExportFilePath")
	''''*** Replace the above with the System Parameter - ProMasterFilePath
	''Once the CAPS service account has been given access to the Promaster drive - Nov 2021
		
	'Get the System Parameter for the Service Account UserName and Password
	strUser = GetSystemAdmin("CAPSServiceAccountName")
	strPass = GetSystemAdmin("CAPSServiceAccountPassword")

	'Get the System Parameter for the fileName
	'strFileNameDefault = GetSystemAdmin("CSFromDinersFileName")
				
	objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

		objStartFolder = strServer
	
	'strFilePathTo = strServer & strFileName
	strFilePathTo = strServer & "\" & strFileName
		
'Update to incliude trailing backslah after mapping drive (where trailing backslash causes an error)
	strServer = strServer & "\"

	'response.write "<Br>strFilePathFrom=" & strFilePathFrom
	'response.write "<Br>strFilePathTo=" & strFilePathTo
	
	'response.write "<Br>strServer=" & strServer & strFileNameNoExt & x & strFileExtension
	
	'response.write "<Br>strFilePathOnly=" & strFilePathOnly
	
	If objFSO.FileExists(strFilePathTo) Then
		
		strFileExtension = objFSO.GetExtensionName(strFilePathTo)
		strFileNameNoExt = Left(strFilePathTo,Len(strFilePathTo)-Len(strFileExtension)-1)
		
		For x = 1 to 10
			If objFSO.FileExists(strFilePathFrom) Then
				strFilePathTo = strFileNameNoExt & x & "." & strFileExtension
			Else
				strFilePathTo = strFileNameNoExt & x + 1 & "." & strFileExtension
				'strFilePathTo = strServer & strFileNameNoExt & x & strFileExtension
				'strFilePathFrom = strFilePathOnly & strFileName & x & strFileExtension
				
				'Move the file to the Loaded folder
				objFSO.MoveFile strFilePathFrom, strFilePathTo
				'objFSO.MoveFile strFilePathFrom, strServer & strFileNameNoExt & x & strFileExtension
				'objFSO.MoveFile strFilePathFrom & strFileName & x & strFileExtension
				
				x = 10
			End If
		Next
	Else
		'Response.Write "<BR>strFilePathFrom=" & strFilePathFrom
		'Response.Write "<BR>strFilePathTo=" & strFilePathTo

		'Move the file to the Loaded folder
		'objFSO.MoveFile strFilePathOnly & strFileName,strServer & strFileName
		objFSO.MoveFile strFilePathFrom,strFilePathTo'strServer & strFileName
		
	End If
	
	'Remove the trailing backslash as the FSO object doesn;t like this on the new DPE server
	strServer = Left(strServer, len(strServer)- 1)	
	objNetwork.RemoveNetworkDrive strServer, True, False
		 
	Set objFSO = Nothing
	Set objNetwork = Nothing

	'If strError = "" Then
	'	Response.Write "<div class=""alert alert-success"" role=""alert"">ProMaster Export File " & Session("PMFileSelect") & " " & strFileNameStart & strFileDateTimeSec & ".txt" & " ADDED to the CS file export folder!</div>"
	'Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">File Moved FROM: " & strFilePathFrom & " - To: " & strFilePathTo & "</div>"
	'End If						
		
End Sub


Set objRS = Nothing
Set objCon = Nothing

 %>


