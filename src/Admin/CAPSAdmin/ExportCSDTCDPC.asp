<!-- #Include file=../../CC/CAPSHeader.asp -->
<!-- #Include file=../../ADOVBS.inc -->
<!-- #Include file=../../CC/CAPSFunctions.asp -->
<%
'Description:	Genera Expenses Upload Administration screen
'Author:		MG
'Date:			April 2013

    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("UserID")) Then Response.Redirect("../../Timeout.asp")
	
'Instantiate Common Page Variables.
Dim objCon
Dim objCmd
Dim objRS

Dim strDeleteCheck
Dim dteBatchDate
Dim strNextCSFileNumber
Dim strCardType
Dim strFileType

If Not IsEmpty(Request.QueryString("CardType")) Then

	strCardType = Request.QueryString("CardType")
	
End If

If strCardType = "DTC" Then
	strFileType = "CSDinersTo"
Else
	strFileType = "CSDinersToDPC"
End If

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

'If the Cancel/Remove has been clicked on the CS File Modal (within the file GetNAfile.asp) then flag the CS File record as removed
If Request.QueryString("Action") = "CancelCS" Then

	Call RemoveCSRecord(Request.QueryString("CSToDinersID"),Request.QueryString("CSEID"),Request.QueryString("Status"))

End If

	If Not IsEmpty(Request.QueryString("BatchNumber")) Then
		strNextCSFileNumber = Request.QueryString("BatchNumber")
		If IsNumeric(strNextCSFileNumber) Then strNextCSFileNumber = PadDigits(strNextCSFileNumber,6)
	Else
		strNextCSFileNumber = GetSystemAdmin("CSFileNumberTo")
		If IsNumeric(strNextCSFileNumber) Then strNextCSFileNumber = PadDigits(strNextCSFileNumber,6)
	
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
	self.location="ExportCSDTCDPC.asp?FileDate=" + document.getElementById("CSDate").value;
}

function loadCS(CardType) {

	//document.getElementById('Progress').style.display = "inline";
	alert(CardType);
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CSDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../../CC/AJAX/GetCSFile.asp?CardType=" + CardType + "&Deleted=N", true);
  xhttp.send();
  
  //document.getElementById('Progress').style.display = "none";
}

function loadCSDeleted() {

	//document.getElementById('Progress').style.display = "inline";
	//document.getElementById('Progress').style.display = "inline";
	
	document.getElementById("CSDetail").innerHTML = '<span id="Progress" style="display:inline; padding-left:20px; padding-bottom:20px;"><img src="../../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>'
	
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("CSDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../../CC/AJAX/GetCSFile.asp?Deleted=Y", true);
  xhttp.send();
  
  //document.getElementById('Progress').style.display = "none";
}

function HideWaiting() {

document.getElementById('Progress').style.display = "none";
}

function ConfirmExport(cb) {
	
	var id = cb.getAttribute('data-CSNumber');
	document.getElementById('CSFileExport').value=id;

}

</script>

<main class="main py-3">
    <div class="container">
<body>

<!-- Modal Detail Start-->
<div class="modal fade" id="CSModal" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-large modal-dialog-right modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  CS File to Diners - <%=strCardType%>
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="CSDetail">
               
			   <span id="Progress" style="display:inline; padding-left:20px; padding-bottom:20px;"><img src="../../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>
				  
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-primary" onclick="window.open('CSToDinersExportExcel.asp')"><i class="fa fa-file"></i> Export CS File to Excel</button>
				<button type="button" class="btn btn-primary float-right" data-toggle="modal" data-target="#ModalApprove" data-CSNumber="<%=strNextCSFileNumber%>" onClick="ConfirmExport(this);" Title="Click to send the Next CS File"><i class="fa fa-file"></i> Export CS File</button>
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
<!-- Modal Detail End-->

 <!-- Approve Modal -->
<div class="modal fade" id="ModalApprove" tabindex="-1" role="dialog" aria-labelledby="ModalApprove" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalApproveTitle" style="font-weight:bold;">CS File Export Confirmation</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <span>Export CS File Number: <input type="text" name="CSFileExport" id="CSFileExport" style="border:0; font-weight:bold; width:100px; text-align:right;" value="<%=strNextCSFileNumber%>" >?</span><br><br>
      </div>
      <div class="modal-footer">
		<button type="button" class="btn btn-primary" onClick='window.location="../CSTransactionsTo.asp?cardType=<%=strCardType%>&Action=ExportCS"'><i class="fa fa-check"></i> Yes</button>
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
        
      </div>
    </div>
  </div>
</div>
<!-- End Approve Modal -->


<form action="ExportCSDTCDPC.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	 <div class="row" id="basic-table">
  <div class="col-3">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../../CC/img/Diners2.png" height="40px" width="50px" title="Diners"> CS To Diners File Export - <%=strCardType%></h4>
        </div>
      <div class="card-content">
        <div class="card-body">
			<p class="card-text">
				Changes to Diners cards which have been processed in CAPS but not yet exported to Diners (in the CS file) will appear below.  
				Click Export CS File to create the export file and process all records below.
			</p>
<!--<div class="col-lg-12 col-md-12">-->
<p class="card-text">
<!-- <fieldset class="form-group">
    <button type="button" class="btn btn-primary btn-xs" onclick="window.open('CSFromDinersTemplateExcel.asp')"><i class="fa fa-file"></i> Export CS File</button>
	<button type="button" class="btn btn-primary btn-xs" onclick="window.open('TemplateExcel.asp?T=tblCAPSCSToDiners&w=WHERE [FileSeqNum] IS NULL')"><i class="fa fa-file"></i> View CS File in Excel </button>
</fieldset> -->
<fieldset class="form-group">
    <button type="button" class="btn btn-outline-secondary btn-sm" data-toggle="modal" data-target="#CSModal" HREF="#" onClick="loadCS('<%=strCardType%>');"><i class="fa fa-file"></i> Export </button>
	<button type="button" class="btn btn-outline-secondary btn-sm" onClick="self.location='../CSTransactionsTo.asp?CardType=<%=strCardType%>'"><i class="fa fa-file"></i> View Details </button>
</fieldset>
<!--</div>-->
</p>
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
Dim lngRecords

If Not IsEmpty(Request.QueryString("BatchNo")) Then
	If IsNull(Request.QueryString("BatchNo")) or Request.QueryString("BatchNo")= "" Then 
		
	Else
		'strWhere = "WHERE FileSeqNum = '" & Request.QueryString("BatchNo") & "'"
		strWhere = "WHERE CardType = '" & strCardType & "' AND (FileSeqNum = " & Request.QueryString("BatchNo") & ")"
	
		If Request.QueryString("BatchNo") = 0 Then strWhere = strWhere & " OR BatchNo IS NULL"
	End If
Else
	strWhere = "WHERE CardType = '" & strCardType & "' AND (FileSeqNum Is NULL OR FileSeqNum = '')"
End If

'First Get the Total records Ready to be exported for display
objRS.Open "SELECT Count(*) AS [CountCS] FROM qryCAPSCSToDiners "  & strWhere,objCon

	If objRS.EOf Then
		lngRecords = 0
	Else
		lngRecords = objRS("CountCS")
	End If
	
objRS.Close
Response.Write "SELECT TOP 50 * FROM qryCAPSCSToDiners "  & strWhere

objRS.Open "SELECT TOP 50 * FROM qryCAPSCSToDiners "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no CS To Diners data ready for export</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Files will be added as part of the admin and load functions and will appear here when appropriate</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 table-compact"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing CS To Diners Data ready for export of <b>" & lngRecords & "</b> Total to be exported.</th></tr>" & _
		        "<tr><th Style=""width:20px;"">CSToDinersID</th>" & _
				"<th>EmployeeID</th>" & _
				"<th>Card No</th><th>Statusxxx</th>" & _	
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
			    Response.Write "<TR class='clickable-row' data-href='../CSTransactionsTo.asp?CardType=" & objRS("CardType") & " &BatchNo=" & objRS("FileSeqNum") & "&EIDNo=" & objRS("EIDNo") & "' style=""cursor: pointer;""><TD>" & objRS("CSToDinersID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("EIDNo") & "</TD><TD style=""text-align:center"">" & objRS("CardNo") & "</TD><TD style=""text-align:center"">" & objRS("CardStatus") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("Title") & "</TD><TD style=""text-align:center"">" & objRS("GivenNames") & "</TD><TD style=""text-align:center"">" & objRS("Surname") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("Address1") & "</TD><TD style=""text-align:center"">" & objRS("Address2") & "</TD><TD style=""text-align:center"">" & objRS("Address3") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("Suburb") & "</TD><TD style=""text-align:center"">" & objRS("State") & "</TD><TD style=""text-align:center"">" & objRS("PostCode") & "</TD>" & _
								"<TD style=""text-align:center"">" & objRS("Email") & "</TD><TD style=""text-align:center"">" & objRS("Status") & "</TD><TD style=""text-align:center"">" & objRS("FileSeqNum") & "</TD>" & _
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

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", +20, Now())
	'If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", +200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		objRS.Open "SELECT TOP 10 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded < '" & dteBatchDate & "' AND FileType = '" & strFileType & "' ORDER BY FileSeqNum DESC",objCon
		
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no CS To Diners data loaded (Before <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-compact mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">CS To Diners Summary (Before <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)</th></tr>" & _
		        "<th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>Total Cards</th><th>Total Employees</th>" & _	
		        "<th>DTC</th>" & _
	 	        "<th>CMC</th>" & _	
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th></tr>" 
				
		    End If
		    
		    Do until objRS.eof
				
				If IsNull(objRS("DateLoaded")) Then
					strDateUpdatedTitle = ""
					strDateUpdated = ""
				Else
					strDateUpdated = FormatDateTime(objRS("DateLoaded"),vbShortDate)
					strDateUpdatedTitle = "title=""" & objRS("DateLoaded") & """"
				End If
				
				
				Response.Write "<TR><TD><a href=""ExportCSDTCDPC.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center"">" & objRS("RecordCount") & "</TD><TD style=""text-align:center"">" & objRS("CardCount") & "</TD><TD style=""text-align:center"">" & objRS("EmployeeCount") & "</TD>" & _
							"<TD style=""text-align:center"">" & objRS("DTCCount") & "</TD><TD style=""text-align:center"">" & objRS("CMCCount") & "</TD><TD style=""text-align:center"">" & objRS("Status") & "</TD>" & _
							"<TD style=""text-align:center"" " & strDateUpdatedTitle & ">" & strDateUpdated & "</TD></TR>"
    			objRS.Movenext			
		    Loop
    			
			
								
	    objRS.Close

        Response.Write "</table>"
		
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

Public Sub RemoveCSRecord(lngCSToDinersID, strEmployeeID, strStatus)

Dim intRecord

  	With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCSFileRemoveCard"

			.Parameters.Append objCmd.CreateParameter("CSToDinersID", adVarChar, adParamInput,10)
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Status", adVarChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("CSFileRemoveOutput", adInteger, adParamOutput)
			
			.Parameters("CSToDinersID") = lngCSToDinersID
			.Parameters("UpdatedBy") = Session("UserID")
			.Parameters("Status") = strStatus
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CSFileRemoveOutput") 
	 
		If intRecord = 0 Then
			If strStatus = "Deleted" Then
				Response.Write "<div class=""alert alert-danger"" role=""alert"">CS Record for " & strEmployeeID & " NOT Removed from CS File! An Error has occurred. See System Admin with CS File ID: " & lngCSToDinersID & " </div>"
			Else
				Response.Write "<div class=""alert alert-danger"" role=""alert"">CS Record for " & strEmployeeID & " NOT Added to the CS File! An Error has occurred. See System Admin with CS File ID: " & lngCSToDinersID & " </div>"
			End If
		Else
			If strStatus = "Deleted" Then
				Response.Write "<div class=""alert alert-success"" role=""alert"">CS Record for " & strEmployeeID & " REMOVED from the CS file!</div>"
			Else
				Response.Write "<div class=""alert alert-success"" role=""alert"">CS Record for " & strEmployeeID & " ADDED to the CS file!</div>"
			End If
		End If
		
	
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
	
	objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Diners\DinersTo"

	Set objFolder = objFSO.GetFolder(objStartFolder)	
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered table-compact mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">Files Exported <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be sent to Diners.""></i></th></tr>"
			
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


Set objRS = Nothing
Set objCon = Nothing

 %>


