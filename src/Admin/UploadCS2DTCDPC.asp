<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #include file="upload.asp" -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #include file="../CC/CAPSFunctions.asp" -->
<%
'Description:	CS From Diners file upload Administration screen
'Author:		MG
'Date:			October 2020

    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")
	
'Instantiate Common Page Variables.
Dim objCon
Dim objCmd
Dim objRS

Dim strDeleteCheck
Dim dteBatchDate
Dim strCardType
Dim strCardTypeSub
Dim strFileType
Dim strFileNameDefault
Dim strFileTypeColour
Dim strFileTypeImage

'strCardType = "Diners"

'strFileType = "CSFromDiners"

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")


If Not IsEmpty(Request.QueryString("CardType")) Then

	strCardType = Request.QueryString("CardType")
	
End If

If strCardType = "DPC" Then

	strFileType = "CSFromDinersDPC"
	strCardTypeSub = ""
	strFileTypeColour = "red"
	strFileTypeImage = "Mastercard2.png"
	
End If

If strCardType = "DPCU" Then

	strFileType = "CSFromDinersDPCU"
	strCardTypeSub = ""
	strFileTypeColour = "red"
	strFileTypeImage = "Mastercard2.png"
	
End If
	
If strCardType = "DTC" Then

	strFileType = "CSFromDiners"
	strCardTypeSub = ""
	strFileTypeColour = "blue"
	strFileTypeImage = "Diners2.png"
	
End If



If Request.QueryString("Action")="Save" Then

	Call StartLoad()
	
End If



'If the Process button has been clicked next to a file, then call the Process procedure
If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "Process" Then
		Call ProcessFile(Request.QueryString("FileSeqNum"),Request.QueryString("FileLoadID"),strCardType,strFileType)
	End If
End If

If Not IsEmpty(Request.QueryString("Reload")) Then

	Call StartLoad()
End If

If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
End If

'If the local load has been clicked then call the procedure to load the network file rather than uploading it
If Request.QueryString("Action")="SaveFileLocal" Then

	'Call StartLoadLocal()
	strDeleteCheck = Request.QueryString("chkDelete")
	Call StartLoadLocal()
End If

'Response.Write "B=" & 	dteBatchDate
'dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
'Response.Write " 2B=" & 	dteBatchDate

'	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
'	Response.Write " 3B=" & 	dteBatchDateFormat

'Get the System Parameter for the fileName
If strCardType = "DPC" Then
	strFileNameDefault = GetSystemAdmin("CSFromDinersFileNameDPC")
End If

If strCardType = "DTC" Then
	strFileNameDefault = GetSystemAdmin("CSFromDinersFileNameDTC")
End If

If strCardType = "DPCU" Then
	strFileNameDefault = GetSystemAdmin("CSFromDinersFileNameDPCU")
End If

 %>

<html>
<head>

	
	
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

function UploadLocal(CardType)
{

	document.getElementById('Progress').style.display = "inline";
	self.location="UploadCS2DTCDPC.asp?Action=SaveFileLocal&chkDelete=no&CardType=" + CardType
}

function UploadLocalG(CardType)

{

	document.getElementById('Progress').style.display = "inline";
	self.location="UploadCS2DTCDPC.asp?Action=SaveFileLocal&chkDelete=no&ActionType=Service&CardType=" + CardType
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

function DatePickChange(CardType) {
	
	self.location="UploadCS2DTCDPC.asp?CardType=" + CardType + "&FileDate=" + document.getElementById("CSDate").value;
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
	<div class="loader" id="ModApp">
        <div class="wrap" >
            <div class="spinner" ></div>
            <span class="loading-message">Processing...</span>
        </div>
    </div>
	
<main class="main py-3">
      <div class="container">
<form  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 <div class="row" id="basic-table">

  <div class="col-3">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title"><img src="../CC/img/<%=strFileTypeImage%>" height="40px" width="50px" title="Diners"> CS From Diners File Load - <span style="color:<%=strFileTypeColour%>";><%=strCardType%></span></h4>
        </div>
      <div class="card-content">
        <div class="card-body">
		


<div class="form-body">
<div class="row col-12" style="padding-left:10px; padding-right:10px;">

   <!--<div class="col-auto text-center">-->
	<button type="button" class="btn btn-primary btn-xs" onclick="UploadLocalG('<%=strCardType%>');" Title="Click to Load any existing CS From Diners file in the G Drive Folder"><i class="fa fa-upload"></i> Load CS G</button>&nbsp;
	<button type="button" class="btn btn-outline-secondary btn-xs" onclick="UploadLocal('<%=strCardType%>');" Title="Click to Load any existing CS From Diners file in the CAPS Server Folder"><i class="fa fa-upload"></i> Load CS File</button>
  <!--</div>-->
  <!--<div class="col-auto">
	<button type="button" class="btn btn-primary btn-xs" onclick="upload();"><i class="fa fa-upload"></i> Upload</button>
  </div>-->
</div>
</div>
<p class="text-left">
<div class="py-3"> 
	<span id="Progress" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Processing...</b></span>
</div>
<!--<font color="red" size="2"><b>NOTE: </Font><font color="black" size="2">When loading from Excel the Worksheet MUST be named 'CSData' (no spaces)
            <BR>* The file must be '.xls' or '.txt' only
            <BR>* Do not change the first row headers from the template files (below)</B></Font>-->
</p>
<!--<div class="col-lg-12 col-md-12">
<fieldset class="form-group">
    <button type="button" class="btn btn-outline-primary btn-xs" onclick="window.open('CSFromDinersTemplateExcel.asp')"><i class="fa fa-file"></i> CS From Diners Template Excel </button>
</fieldset>
</div>-->

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
        <div class="card-body py-1 md-1" style="padding-right:5px; padding-left:5px;">		
		<%DisplaySummary(strCardType)%>	
		</div>
	  </div>
    </div>
   </div>
   
   
    <div class="col-2">
    <div class="card">     
      <div class="card-content">
        <div class="card-body">
			<%DisplayFileSummary(strFileNameDefault)%>	
		</div>
	  </div>
    </div>
	
	<div class="card">     
      <div class="card-content">
        <div class="card-body">
			<%DisplayFileSummaryG(strFileNameDefault)%>	
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


<!-- #Include file=../CC/CAPSFooter.asp -->
</body>
</html>

<%
Sub DisplayTableDetails()

Dim strWhere

If Not IsEmpty(Request.QueryString("BatchNo")) Then
	If IsNull(Request.QueryString("BatchNo")) or Request.QueryString("BatchNo")= "" Then 
		
	Else
		strWhere = "WHERE CardType = '" & strCardType & "' AND FileSeqNum = " & Request.QueryString("BatchNo") & ""
	
		If Request.QueryString("BatchNo") = 0 Then strWhere = strWhere & " OR BatchNo IS NULL"
	End If
Else
	strWhere = "WHERE CardType = '" & strCardType & "' AND CardTypeSub = '" & strCardTypeSub & "' AND FileSeqNum = '0'"
End If

objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDiners WITH(NOLOCK) "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no CS From Diners data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Upload Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing CS From Diners Data Batch No. " & Request.QueryString("BatchNo") & " already in CAPS</th></tr>" & _
                "<tr><td colspan=""16"" style=""text-align:left; color:red;font-size:20px""><B>WARNING! The data below will be deleted if you Upload a new CS From Diners file!</B></td></tr><tr>" & _
		        "<th>File Seq No</th><th Style=""width:20px;"">App ID</th>" & _
				"<th>Employee ID</th>" & _
				"<th>Card No</th><th>Card Type</th>" & _	
		        "<th>Title</th>" & _
	 	        "<th>First Names</th>" & _	
	 	        "<th>Surname</th>" & _
		        "<th>Address 1</th><th>Suburb</th>" & _
                "<th>State</th><th>Post Code</th><th>Email</th>" & _
                "<th>Status</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR class='clickable-row' data-href='UploadCSDetail.asp?BatchNo=" & objRS("FileSeqNum") & "&EIDNo=" & objRS("EIDNo") & "' style=""cursor: pointer;""><TD style=""text-align:center; font-size:12px;"">" & objRS("FileSeqNum") & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS("EIDNo") & "</TD><TD style=""text-align:center; font-size:12px;"">" & MaskCard(objRS("CardNo")) & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("PlasticID") & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS("Title") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("GivenNames") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("Surname") & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS("Address1") & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS("Suburb") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("State") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("PostCode") & "</TD>" & _
								"<TD style=""text-align:center; font-size:12px;"">" & objRS("Email") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("Status") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS(0) & "</TD>" & _
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
Dim strDPCTypeFileName
Dim strDPCType
Dim strBranded

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -10, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		objRS.Open "SELECT TOP 10 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = '" & strFileType & "' ORDER BY FileSeqNum DESC",objCon
		'Response.Write "SELECT TOP 10 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE Deleted = 'N' AND DateLoaded > '" & dteBatchDate & "' AND FileType = '" & strFileType & "' ORDER BY FileSeqNum DESC"
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""11"" style=""text-align:left"">There is no CS From Diners data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange('" & strCardType & "');""/>)</th></tr>" 
		        
		    Else
				'Write a column for the card file type (Branded or Unbranded - with a U in the filename) if the screen is for DPC Diners cards
				If strCardType = "DPC" or strCardType = "DPCU" Then
					strBranded = "<th>Card File</th>"
				Else
					strBranded = ""
				End If
				
		         Response.Write "<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""12"" style=""text-align:left"">CS From Diners Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange('" & strCardType & "');"" />)</th></tr>" & _
		        "<th Style=""width:20px;"">Batch No.</th><th>File Date</th>" & _
				"<th>Total Records</th>" & _
				strBranded & _
		        "<th>" & strCardType & "</th>" & _
	 	        "<th>CMC</th>" & _	
				"<th>No. Loaded</th>" & _
				"<th>Action</th>" & _
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
					strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='UploadCS2DTCDPC.asp?Action=Process&CardType=" & strCardType & "&FileSeqNum=" & objRS("FileSeqNum") & "&FileLoadID=" & objRS("FileLoadID") & "'""; title=""Click to Process the CS File and update changes to cards in CAPS from the CS File loaded " & objRS("DateLoaded") & """>Process" & strCardType & "</button>"
					'strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='UploadANZ.asp?Action=Process&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to Process the ANZ File and update changes to cards in CAPS from the ANZ File loaded " & objRS("DateLoaded") & """><i class=""fa fa-cogs""></i> Process</button>"
				Else
					strAction = strStatus
				End If
				
				If IsNull(objRS("DateLoaded")) Then
					dteDateLoaded = ""
				Else
					dteDateLoaded = FormatDateTime(objRS("DateLoaded"),vbShortDate)
				End If
				
				'APRIL 2023 -- New for DPC Diners cards to get the type of CS file (Branded or Unbranded)
				strDPCTypeFileName = objRS("FileName")
				
				'Get the System Parameter for the fileName
				If strCardType = "DTC" Then
					strDPCType = GetSystemAdmin("CSFromDinersFileNameDTC")
				End If
				
				If strCardType = "DPC" Then
					strDPCType = GetSystemAdmin("CSFromDinersFileNameDPC")
				End If
				
				If strCardType = "DPCU" Then
					strDPCType = GetSystemAdmin("CSFromDinersFileNameDPCU")
				End If
			
				'Write a column for the card file type (Branded or Unbranded - with a U in the filename) if the screen is for DPC Diners cards
				If strCardType = "DPC" or strCardType = "DPCU" Then
					
						If strCardType = "DPCU" Then
							strDPCType = "<span class=""badge badge-pill badge-secondary"" style=""font-size:10px;"">Un-branded</span>"
						Else
							strDPCType = "<span class=""badge badge-pill badge-info"" style=""font-size:10px;"">Branded</span>"
						End If
					
					strDPCType = "<TD style=""text-align:center; font-size:12px;"" Title=""" & strDPCTypeFileName & """>" & strDPCType & "</TD>"
				Else
					strDPCType = ""
				End If
					
				'Create the View button detail
				strView = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='CSTransactions.asp?CardType=" & strCardType & "&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to view details of the CS File loaded " & objRS("DateLoaded") & """>View</button>"
				
				
				Response.Write "<TR><TD><a href=""UploadCS2DTCDPC.asp?CardType=" & strCardType & "&BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center; font-size:12px;"">" & objRS("FileDatetime") & "</TD><TD style=""text-align:center"">" & objRS("RecordCount") & "</TD>" & strDPCType & _
							"<TD style=""text-align:center"">" & objRS("DTCCount") & "</TD><TD style=""text-align:center"">" & objRS("CMCCount") & "</TD><TD style=""text-align:center"">" & objRS("RecordsLoaded") & "</TD><TD style=""text-align:center"">" & strAction & "</TD>" & _
							"<TD style=""text-align:center; font-size:12px;"" Title=""" & objRS("DateLoaded") & """>" & dteDateLoaded & "</TD><TD style=""text-align:center"">" & strView & "</TD></TR>"
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
Dim strCardType
Dim lngTotalRecords
Dim lngBatchNo1
Dim dteBatchDate
Dim dteBatchDateFormat
Dim srStatus1
Dim strDateUpdated

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -10, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)

objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary  ORDER BY FileSeqNum DESC",objCon
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">There is no CS From Diners data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" />)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""8"" style=""text-align:left"">CS From Diners Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" />)</th></tr>" & _
		        "<th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>Total Cards</th><th>Total Employees</th>" & _	
		        "<th>DTC</th>" & _
	 	        "<th>CMC</th>" & _	
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th></tr>" 
				
		        If IsNull(objRS("FileSeqNum")) or objRS("FileSeqNum") = "" Then 
					lngBatchNo1 = 0
				Else
					lngBatchNo1 = objRS("FileSeqNum")
				End If
				
				'clng(lngBatchNo) = clng(lngBatchNo1)
				
				'Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                '"<tr><th colspan=""7"" style=""text-align:left"">CS From Diners Summary</th></tr>" & _
		        '"<th Style=""width:20px;"">Cards (Count)</th>" & _
				'"<th>Card Type</th>" & _
				'"<th>Card Type Sub</th><th>Employees (Count)</th>" & _	
		        '"<th>Batch No.</th>" & _
	 	        '"<th>Status</th>" & _	
	 	        '"<th>Date Loaded</th></tr>" 
				
		    End If
		    
		    Do until objRS.eof
			
				If IsNull(objRS("FileSeqNum")) or objRS("FileSeqNum") = "" Then 
					lngBatchNo = 0
				Else
					lngBatchNo = objRS("FileSeqNum")
				End If
					
				If clng(lngBatchNo) = clng(lngBatchNo1) Then
					
					lngCards = lngCards + objRS("CardCount")
					lngEmployees = lngEmployees + objRS("EmployeeCount")
					lngTotalRecords = lngTotalRecords + objRS("TotalRecords")
					
					'Get the Card Type value and make sure it is not null
					If isNull(objRS("CardType")) Then 
						strCardType = ""
					Else	
						strCardType = objRS("CardType")
					End If
					
					Select Case strCardType
						
						Case "DTC"
							lngDTC = lngDTC + lngCards
						Case "Mastercard"
							lngCMC = lngCMC + lngCards
						Case Else
							lngOther = lngOther + 1
					End Select
					
				
				Else
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR><TD><a href=""UploadCS2DTCDPC.asp?BatchNo=" & lngBatchNo1 & """>" & lngBatchNo1 & "</a></B></TD>" & _
			                    "<TD style=""text-align:center"">" & lngTotalRecords & "</TD><TD style=""text-align:center"">" & lngCards & "</TD><TD style=""text-align:center"">" & lngEmployees & "</TD>" & _
			                    "<TD style=""text-align:center"">" & lngDTC & "</TD><TD style=""text-align:center"">" & lngCMC & "</TD><TD style=""text-align:center"">" & objRS("Status") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("FileDateTime") & "</TD></TR>"
					
					'Response.Write "<TR><TD>&nbsp;" & objRS(0) & "</A></B></TD>" & _
			         '           "<TD style=""text-align:center"">" & objRS(1) & "</TD><TD style=""text-align:center"">" & objRS(2) & "</TD><TD style=""text-align:center"">" & objRS(3) & "</TD>" & _
			         '           "<TD style=""text-align:center"">" & objRS(4) & "</TD><TD style=""text-align:center"">" & objRS(5) & "</TD><TD style=""text-align:center"">" & objRS(6) & "</TD>" & _
			         '           "</TR>"
    			'End If
				
				lngBatchNo = 0
				
				If IsNull(objRS("FileSeqNum")) or objRS("FileSeqNum") = "" Then 
					lngBatchNo1 = 0
				Else
					lngBatchNo1 = objRS("FileSeqNum")
				End If
				
    			End If
				
				srStatus1 = objRS("Status")
				strDateUpdated = objRS("FileDateTime")
				
			    objRS.movenext
				
				If objRS.Eof Then
					
					Response.Write "<TR><TD><a href=""UploadCS2DTCDPC.asp?BatchNo=" & lngBatchNo1 & """>" & lngBatchNo1 & "</a></B></TD>" & _
			                    "<TD style=""text-align:center"">" & lngTotalRecords & "</TD><TD style=""text-align:center"">" & lngCards & "</TD><TD style=""text-align:center"">" & lngEmployees & "</TD>" & _
			                    "<TD style=""text-align:center"">" & lngDTC & "</TD><TD style=""text-align:center"">" & lngCMC & "</TD><TD style=""text-align:center"">" & srStatus1 & "</TD>" & _
			                    "<TD style=""text-align:center"">" & strDateUpdated & "</TD></TR>"
					
				End If
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
Dim strUpdatedBy

				
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
					Response.Write "<b>CS From Diners File Successfully Uploaded!!<b><br>"
				End if
				
				
				
				'Close the recordset/connection 
				objRS.Close 
				objExcelCon.Close 
		    
			Else
			
			
				ReadText filePath,strFileName,strFileType,strCardType
			End If
			
	    'End If
	End IF
	
End if

End Sub


Sub StartLoadLocal()
'Procedure to load the local file from within the network, rather than loading the file to the server
Dim strFileName
Dim File, filePath
Dim objFSO
Dim objStartFolder
Dim objFolder
Dim colFiles
Dim objFile
Dim lngFileSize

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass
'Dim strFileNameDefault

Dim strDPCFullFileName
Dim strFileTypeDPCWithU

Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")

	If Request.QueryString("Action")="SaveFileLocal" Then

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
	'If Request.QueryString("Action")="SaveFileLocal" Then
		
		'Set objFSO = CreateObject("Scripting.FileSystemObject")

			'objStartFolder = "D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\Diners"
			objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Diners\DinersFrom\"	
			'objStartFolder = GetSystemAdmin("ServerFilePath")

		End If
		
			Set objFolder = objFSO.GetFolder(objStartFolder)
			Set colFiles = objFolder.Files

			'Get the System Parameter for the fileName
			If strCardType = "DTC" Then
				strFileNameDefault = GetSystemAdmin("CSFromDinersFileNameDTC")
			End If
			
			If strCardType = "DPC" Then
				strFileNameDefault = GetSystemAdmin("CSFromDinersFileNameDPC")
			End If
			
			If strCardType = "DPCU" Then
				strFileNameDefault = GetSystemAdmin("CSFromDinersFileNameDPCU")
			End If
			
			
			If IsNull(strFileNameDefault) or strFileNameDefault = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! There is no System Parameter for CS From Diners File Names (""CSFromDinersFileName""). See System Admin.</span></div></div></div>"
					Exit Sub
			End If
			
			
			' Load DPC Unbranded
			
			If strFileType = "CSFromDinersDPCU" Then
		
				For Each objFile in colFiles

					If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) AND Mid(objFile.Name,24,1) = "U" Then
				
						''---START New DPC filename section
						'Get the DPC file name as branded is the system default name whereas Unbranded is the DPC default with the letter U after it.  They need to be processed and logged separately.
						
						'Response.write "<br>Mid=" & Mid(objFile.Name,Len(strFileNameDefault)+1,1)
						'Response.write "<br>objFile.Name=" & objFile.Name
						If Mid(objFile.Name,Len(strFileNameDefault)+1,1) = "U" Then
							strDPCFullFileName = "U"
						Else
							strDPCFullFileName = ""
						End If
						''---END New DPC filename section
						
						strFileName = objFile.Name
						filePath = objStartFolder & "\" & strFileName
						'filePath = objStartFolder & "\" & strFileName
						lngFileSize = objFile.Size
						
						'Response.Write objFile.Name & "</br>"
					
					'Moved the End If and Next (1) below to (2) so thatthe rocedure loops and loads both Brnaded and Unbranded files if they exist in the folder
					''End If (1)
					
				''Next (1)

						If IsNull(strFileName) or strFileName = "" Then

							Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
									"<span aria-hidden=""true"">&times;</span></button>" & _
									"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
									"<span>" & strFileNameDefault & " NOT LOADED! There is no CS FROM DINERS File in the Server Folder to Load! Copy the 'AUDC_FROMECS_CS_DO_xxxx.txt' file to the location " & objStartFolder & "</span></div></div></div>"
								Exit Sub
								
						'Moved the End if below to after the Read text file to 
						End If
						
						''New DPC section to add the U for unbranded to the FileType for the ReadText procedure so the branded and unbranded can be proceessed seperately
						If strFileType = "CSFromDinersDPC" AND strDPCFullFileName = "U" Then 
							'strFileType = strFileType & "U"
							strFileTypeDPCWithU = strFileType & "U"
						Else
							strFileTypeDPCWithU = strFileType
						End If
						
						'response.write "<BR>strFileType=" & strFileTypeDPCWithU
						'The CS File From Diners is a text file so run the ReadText procedure -- this lods the data from the text fiel into the database
						ReadText filePath,strFileName,strFileTypeDPCWithU,strCardType
				
					End If '(2)
					
				Next '(2)
			Else
			' Load DPC Branded and DTC
		
			For Each objFile in colFiles
			
			
					If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) AND Mid(objFile.Name,24,1) <> "U" Then
						
						''---START New DPC filename section
						'Get the DPC file name as branded is the system default name whereas Unbranded is the DPC default with the letter U after it.  They need to be processed and logged separately.
						
						'Response.write "<br>Mid=" & Mid(objFile.Name,Len(strFileNameDefault)+1,1)
						'Response.write "<br>objFile.Name=" & objFile.Name
						If Mid(objFile.Name,Len(strFileNameDefault)+1,1) = "U" Then
							strDPCFullFileName = "U"
						Else
							strDPCFullFileName = ""
						End If
						''---END New DPC filename section
						
						strFileName = objFile.Name
						filePath = objStartFolder & "\" & strFileName
						'filePath = objStartFolder & "\" & strFileName
						lngFileSize = objFile.Size
						
						'Response.Write objFile.Name & "</br>"
					
					'Moved the End If and Next (1) below to (2) so thatthe rocedure loops and loads both Brnaded and Unbranded files if they exist in the folder
					''End If (1)
					
				''Next (1)

						If IsNull(strFileName) or strFileName = "" Then

							Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
									"<span aria-hidden=""true"">&times;</span></button>" & _
									"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
									"<span>" & strFileNameDefault & " NOT LOADED! There is no CS FROM DINERS File in the Server Folder to Load! Copy the 'AUDC_FROMECS_CS_DO_xxxx.txt' file to the location " & objStartFolder & "</span></div></div></div>"
								Exit Sub
								
						'Moved the End if below to after the Read text file to 
						End If
						
						''New DPC section to add the U for unbranded to the FileType for the ReadText procedure so the branded and unbranded can be proceessed seperately
						If strFileType = "CSFromDinersDPC" AND strDPCFullFileName = "U" Then 
							'strFileType = strFileType & "U"
							strFileTypeDPCWithU = strFileType & "U"
						Else
							strFileTypeDPCWithU = strFileType
						End If
						
						'response.write "<BR>strFileType=" & strFileTypeDPCWithU
						'The CS File From Diners is a text file so run the ReadText procedure -- this lods the data from the text fiel into the database
					
						ReadText filePath,strFileName,strFileTypeDPCWithU,strCardType
				
					End If '(2)
					
				Next '(2)
			
			
			End If
	
	End if

	
	'''---Start the New Service Account Login section
	If Request.QueryString("ActionType")="Service" Then
		
		objNetwork.RemoveNetworkDrive strServer, True, False
		 
	Set objFSO = Nothing
	Set objNetwork = Nothing

	End If
	
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


Sub ReadText(strFileNamePath,strFileName,strFileType1,strCardType)

Const ForReading = 1
Dim strLine
Dim strRow
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

Dim strFooterCount
Dim lngFileLoadID
Dim objFSO
Dim objTextFile
Dim strFileSeqNumLast
Dim strFileSeqNumNew
Dim objStartFolder
Dim objFolder
Dim objFile
Dim colFiles

Dim strArchiveFolder
Dim strNEFileName
Dim strFileNamePathNE
Dim strNEFileName2

Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set outPut = objFSO.CreateTextFile("c:\\output.txt", true);
Set objTextFile = objFSO.OpenTextFile (strFileNamePath, ForReading)
'Set objTextFile = objFSO.OpenTextFile ("c:\mytextfile.txt", ForReading)


x = 0 
 
	Do Until objTextFile.AtEndOfStream
	
		'Count the rows for use in line counts, summary and for getting header
		x = x + 1
		
		strLine = objTextFile.Readline
		
		'The first row of the CS file has a header with FileDateTime and  FileSequenceNumber
		If x = 1 Then
			'The fileDate and Number are only in the header row
			strFileDateTime = Mid(strLine,5,14)
			strFileSeqNum = Mid(strLine,19,6)
			'strFileSeqNumLast = GetLastCSFileNumber()
			'strFileSeqNumNew = clng(strFileSeqNumLast) + 1
			strFileSeqNumNew = GetSystemAdmin(strFileType1)
			
			'Check to see if the same FileSeqNum for the same FileType has already been loaded
			If GetFileLoadID(" & strFileType1 & ",strFileSeqNum,strFileName) = "" Then		
			'response.write "<br>strFileSeqNumNew=" & strFileSeqNumNew
				If clng(strFileSeqNum) <> clng(strFileSeqNumNew) Then
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! The CS From Diners (" & strFileType1 & " : " & strFileName & ") File Seq Num """ & PadDigits(strFileSeqNum,6) & """ is not the correct next Batch No to be loaded.  It should be Seq Num """ & strFileSeqNumNew & """</span></div></div></div>"
					Exit Sub
				End If
			Else
				'If the checkbox to overwrite is checked then load the data, otherwise do not load
				If strDeleteCheck = "true" Then
					'Delete any existing CS From Diners Records
					objCon.Execute "DELETE FROM tblCAPSCSFromDiners WHERE CardType = '" & strCardType & "' AND FileSeqNum] = " & strFileSeqNum & ""
				Else
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! The CS From Diners File (" & strFileType1 & " : " & strFileName & ") Seq Num """ & strFileSeqNum & """ has already been loaded! </span></div></div></div>"
					Exit Sub
				End If
			End If
			
		Else
			
			'If the first character is a T then it is the final row
			If Mid(strLine, 1, 1) = "T" Then
				strFooterCount = Mid(strLine, 3, 6)
			Else
				' parse strLine
				'*** Removed By AB *** strCardType = Mid(strLine, 1, 1)
				'strEID = Mid(strLine, 3, 10)
				'strRow = Right(strLine,589)
		  
				strCSFromDinersID = 0'objRS("CSFromDinersID") 
				'strFileDateTime = objRS("FileDateTime") 
				'strFileSeqNum = objRS("FileSeqNum") 
				strEIDNo = Mid(strLine, 3, 10) 
				strCardNo = Mid(strLine, 13, 19) 
				strCardUpdateInd = Mid(strLine, 32, 2) 
				strCardExpiryDate = Mid(strLine, 40, 2) & "-" & Monthname(Mid(strLine, 38, 2)) & "-" & Mid(strLine, 34, 4)'Mid(strLine, 34, 4) & Mid(strLine, 38, 2) & Mid(strLine, 40, 2) 
				strCardStatus = Mid(strLine, 42, 2) 
				strTitle = Mid(strLine, 44, 12) 
				strSurname = Mid(strLine, 56, 25) 
				strGivenNames = Mid(strLine, 81, 30) 
				strNameOnCard = Mid(strLine, 111, 26) 
				strAddress1 = Mid(strLine, 137, 40) 
				strAddress2 = Mid(strLine, 177, 40)  
				strAddress3 = Mid(strLine, 217, 40)  
				strSuburb = Mid(strLine, 257, 25)  
				strState = Mid(strLine, 282, 4) 
				strPostCode = Mid(strLine, 286, 12) 
				strHomePhone = Mid(strLine, 300, 10) 
				strWorkPhone = Mid(strLine, 312, 10) 
				strMobilePhone = Mid(strLine, 324, 10) 
				strEmail = Mid(strLine, 334, 70) 
				strReportGroup = Mid(strLine, 404, 8) 
				strCreditLimit = Mid(strLine, 412, 11) 
				strRelationship = Mid(strLine, 423, 19) 
				strCat2 = Mid(strLine, 442, 19) 
				strAccountNumber = Mid(strLine, 461, 19) 
				strActivationFlag = Mid(strLine, 480, 1) 
				strPlasticID = Mid(strLine, 481, 3) 
				strCompanion = Mid(strLine, 484, 19) 
				strStatus = "Imported" 
				strNotes = ""'Mid(strLine, 3, 10) 
				intCardUpdated = 1'Mid(strLine, 3, 10) 
				strAccountBlockCode1 = Mid(strLine, 503, 1) 
				strAccountBlockCode2 = Mid(strLine, 504, 1) 
				strCardLevelBlockCode = Mid(strLine, 505, 1) 
				strCardLevelCreditLimit = Mid(strLine, 506, 13) 
				strCashHoldFlag = Mid(strLine, 519, 1) 
				strCashAllowFlag = Mid(strLine, 520, 1) 
				strZeroes = Mid(strLine, 521, 12) 
				
				
				'Response.Write  "exec spGeneralExpensesSave ="& strCSFromDinersID & "," & strFileDateTime & "," & strFileSeqNum & "," & strCardExpiryDate & "," & x
				
				SaveRecord strCSFromDinersID,strFileDateTime,strFileSeqNum,strEIDNo,strCardNo,strCardUpdateInd,strCardExpiryDate,strCardStatus,strTitle,strSurname, _
							strGivenNames,strNameOnCard,strAddress1,strAddress2,strAddress3,strSuburb,strState,strPostCode,strHomePhone,strWorkPhone,strMobilePhone,  _
							strEmail,strReportGroup,strCreditLimit,strRelationship,strCat2,strAccountNumber,strActivationFlag,strPlasticID,strCompanion,strStatus,  _
							strNotes,intCardUpdated,strAccountBlockCode1,strAccountBlockCode2,strCardLevelBlockCode,strCardLevelCreditLimit,strCashHoldFlag,strCashAllowFlag,strZeroes, strCardType,"",strFileType,x-1
				
				
				'response.write strRow & "," & strEID & "," & strRow
		  
				'outPut.WriteLine(id_no & "_" & strEID);
			End If
		End If
	Loop


	If x > 1 Then
		'the CS FRom Diners contains a header and footer row, so remove then from the count
		If x > 2 Then x = x -2
		
		'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
		lngFileLoadID = SaveFileLoadID (strFileType,strFileName,strFileNamePath,x,0,0,0,0,0,0,strFooterCount,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
		
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
		Call UpdateFileLoadSummary (strFileType,strFileSeqNum, strFileName,lngFileLoadID)
		'Response.write "UpdateFileLoadSummary ("CSFromDiners"," & strFileSeqNum & "," & strFileName &"," & lngFileLoadID & ")"
		
		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>CS From Diners (" & strFileType & " : " & strFileName & ") File Seq Num """ & strFileSeqNum & """ load COMPLETE!</span></div></div></div>"
		
		'Set the reference to the File to nothing so that it can be moved
		Set objTextFile = Nothing
		
		'''---Start the New Service Account Login section
		If Request.QueryString("ActionType")="Service" Then
		
			'Get the System Parameter for the start of the Training File Location
			objStartFolder = GetSystemAdmin("GDriveFilePath")
			
			strArchiveFolder = GetSystemAdmin("GDriveArchiveFilePath")
			
			'Move the file to the Archive folder

		Response.Write strFileNamePath & "<BR>"
		Response.write strArchiveFolder & "Imports\CS File\" & strFileName & "<BR>"

			objFSO.MoveFile strFileNamePath,strArchiveFolder & "Imports\CS File\" & strFileName
			
			'Also move the NE file if it is in the imports folder
			strNEFileName = GetSystemAdmin("NEFromDinersFileName")
			
			Set objFolder = objFSO.GetFolder(objStartFolder)
			Set colFiles = objFolder.Files
			
			For Each objFile in colFiles

				If Left(objFile.Name,Len(strNEFileName)) = Trim(strNEFileName) Then
					strFileNamePathNE = objFile.Name
					strNEFileName2 = objFile.Name
					strFileNamePathNE = objStartFolder & "\" & strFileNamePathNE
					
				End If
				
			Next
			
			'Make sure that teh NE file esists (if there are 2 files to load it may have been moved with the previous load)
			If IsNull(strFileNamePathNE) or strFileNamePathNE= "" Then
			Else
				If isnull(strNEFileName2) or strNEFileName2 ="" Then
					Response.Write "<div class=""alert alert-danger"" role=""alert"">ERROR! NE File Not Found: " & strNEFileName2 & ". See System Admin.</div>"
				Else
				'Move the file to the Archive folder
				objFSO.MoveFile strFileNamePathNE,strArchiveFolder & "Imports\CS File\" & strNEFileName2
				End IF
			End If
			
			Set objFolder = Nothing
			Set colFiles = Nothing
			
		Else
			objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\DINERS\"		
			'Move the file
			objFSO.MoveFile strFileNamePath,objStartFolder & "Loaded\" & strFileName
		End If
		
		'Set objTextFile = Nothing
		
		 						
						
	End If
	
'outPut.Close();

Set objFSO = Nothing
'Set outPut = Nothing

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
Dim strCardTypeSub

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
		'strCardType = objRS("CardType")
		'strCardTypeSub = ""


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
						strNotes,intCardUpdated,strAccountBlockCode1,strAccountBlockCode2,strCardLevelBlockCode,strCardLevelCreditLimit,strCashHoldFlag,strCashAllowFlag,strZeroes, strCardType, "",strFileType, x
            
        
        'End IF
        
     objRS.movenext 
    
    If IsNull(objRS("EIDNo")) Then               
       
       exit Sub
       
    End If
        
    Loop 
    
End Sub	


Sub SaveRecord(strCSFromDinersID,strFileDateTime,strFileSeqNum,strEIDNo,strCardNo,strCardUpdateInd,strCardExpiryDate,strCardStatus,strTitle,strSurname, _
						strGivenNames,strNameOnCard,strAddress1,strAddress2,strAddress3,strSuburb,strState,strPostCode,strHomePhone,strWorkPhone,strMobilePhone, _
						strEmail,strReportGroup,strCreditLimit,strRelationship,strCat2,strAccountNumber,strActivationFlag,strPlasticID,strCompanion,strStatus, _
						strNotes,intCardUpdated,strAccountBlockCode1,strAccountBlockCode2,strCardLevelBlockCode,strCardLevelCreditLimit,strCashHoldFlag,strCashAllowFlag,strZeroes, strCardType, strCardTypeSub,strFileType,x)
						
Dim intRecord

  	With objCmd
  	
  	    'If the procedure has akready run then don't create the parameter objects again (more than once)
  	    If x = 1 then
                .CommandType = 4
                .CommandText = "spCAPSCSFromDinersSave"
                
				.Parameters.Append objCmd.CreateParameter("CSFromDinersID", adInteger)
				.Parameters.Append objCmd.CreateParameter("FileDateTime", advarchar, adParamInput,14)
				.Parameters.Append objCmd.CreateParameter("FileSeqNum", advarchar, adParamInput,6)
				.Parameters.Append objCmd.CreateParameter("EIDNo", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("CardNo", advarchar, adParamInput,19)
				.Parameters.Append objCmd.CreateParameter("CardUpdateInd", advarchar, adParamInput,2)
				.Parameters.Append objCmd.CreateParameter("CardExpiryDate", adDate)
				.Parameters.Append objCmd.CreateParameter("CardStatus", advarchar, adParamInput,2)
				.Parameters.Append objCmd.CreateParameter("Title", advarchar, adParamInput,12)
				.Parameters.Append objCmd.CreateParameter("Surname", advarchar, adParamInput,25)
				.Parameters.Append objCmd.CreateParameter("GivenNames", advarchar, adParamInput,30)
				.Parameters.Append objCmd.CreateParameter("NameOnCard", advarchar, adParamInput,26)
				.Parameters.Append objCmd.CreateParameter("Address1", advarchar, adParamInput,40)
				.Parameters.Append objCmd.CreateParameter("Address2", advarchar, adParamInput,40)
				.Parameters.Append objCmd.CreateParameter("Address3", advarchar, adParamInput,40)
				.Parameters.Append objCmd.CreateParameter("Suburb", advarchar, adParamInput,25)
				.Parameters.Append objCmd.CreateParameter("State", advarchar, adParamInput,4)
				.Parameters.Append objCmd.CreateParameter("PostCode", advarchar, adParamInput,12)
				.Parameters.Append objCmd.CreateParameter("HomePhone", advarchar, adParamInput,12)
				.Parameters.Append objCmd.CreateParameter("WorkPhone", advarchar, adParamInput,12)
				.Parameters.Append objCmd.CreateParameter("MobilePhone", advarchar, adParamInput,12)
				.Parameters.Append objCmd.CreateParameter("Email", advarchar, adParamInput,70)
				.Parameters.Append objCmd.CreateParameter("ReportGroup", advarchar, adParamInput,8)
				.Parameters.Append objCmd.CreateParameter("CreditLimit", advarchar, adParamInput,11)
				.Parameters.Append objCmd.CreateParameter("Relationship", advarchar, adParamInput,19)
				.Parameters.Append objCmd.CreateParameter("Cat2", advarchar, adParamInput,19)
				.Parameters.Append objCmd.CreateParameter("AccountNumber", advarchar, adParamInput,19)
				.Parameters.Append objCmd.CreateParameter("ActivationFlag", advarchar, adParamInput,1)
				.Parameters.Append objCmd.CreateParameter("PlasticID", advarchar, adParamInput,3)
				.Parameters.Append objCmd.CreateParameter("Companion", advarchar, adParamInput,19)
				.Parameters.Append objCmd.CreateParameter("Status", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("Notes", advarchar, adParamInput,100)
				.Parameters.Append objCmd.CreateParameter("CardUpdated", adInteger)
				.Parameters.Append objCmd.CreateParameter("AccountBlockCode1", advarchar, adParamInput,1)
				.Parameters.Append objCmd.CreateParameter("AccountBlockCode2", advarchar, adParamInput,1)
				.Parameters.Append objCmd.CreateParameter("CardLevelBlockCode", advarchar, adParamInput,1)
				.Parameters.Append objCmd.CreateParameter("CardLevelCreditLimit", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("CashHoldFlag", advarchar, adParamInput,1)
				.Parameters.Append objCmd.CreateParameter("CashAllowFlag", advarchar, adParamInput,1)
				.Parameters.Append objCmd.CreateParameter("Zeroes", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("CardType", advarchar, adParamInput,20) 
				.Parameters.Append objCmd.CreateParameter("CardTypeSub", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("FileType", advarchar, adParamInput,20)
				.Parameters.Append objCmd.CreateParameter("CSFromDinersIDOutput", adInteger, adParamOutput)				
            
        End If
                 
				.Parameters("CSFromDinersID") = strCSFromDinersID
				.Parameters("FileDateTime") = strFileDateTime
				.Parameters("FileSeqNum") = strFileSeqNum
				.Parameters("EIDNo") = strEIDNo
				.Parameters("CardNo") = strCardNo
				.Parameters("CardUpdateInd") = strCardUpdateInd
				.Parameters("CardExpiryDate") = strCardExpiryDate
				.Parameters("CardStatus") = strCardStatus
				.Parameters("Title") = strTitle
				.Parameters("Surname") = strSurname
				.Parameters("GivenNames") = strGivenNames
				.Parameters("NameOnCard") = strNameOnCard
				.Parameters("Address1") = strAddress1
				.Parameters("Address2") = strAddress2
				.Parameters("Address3") = strAddress3
				.Parameters("Suburb") = strSuburb
				.Parameters("State") = strState
				.Parameters("PostCode") = strPostCode
				.Parameters("HomePhone") = strHomePhone
				.Parameters("WorkPhone") = strWorkPhone
				.Parameters("MobilePhone") = strMobilePhone
				.Parameters("Email") = strEmail
				.Parameters("ReportGroup") = strReportGroup
				.Parameters("CreditLimit") = strCreditLimit
				.Parameters("Relationship") = strRelationship
				.Parameters("Cat2") = strCat2
				.Parameters("AccountNumber") = strAccountNumber
				.Parameters("ActivationFlag") = strActivationFlag
				.Parameters("PlasticID") = strPlasticID
				.Parameters("Companion") = strCompanion
				.Parameters("Status") = strStatus
				.Parameters("Notes") = strNotes
				.Parameters("CardUpdated") = intCardUpdated
				.Parameters("AccountBlockCode1") = ""'strAccountBlockCode1
				.Parameters("AccountBlockCode2") = ""'strAccountBlockCode2
				.Parameters("CardLevelBlockCode") = ""'strCardLevelBlockCode
				.Parameters("CardLevelCreditLimit") = strCardLevelCreditLimit
				.Parameters("CashHoldFlag") = strCashHoldFlag
				.Parameters("CashAllowFlag") = strCashAllowFlag
				.Parameters("Zeroes") = strZeroes
				.Parameters("CardType") = Left(strCardType,3)
				.Parameters("CardTypeSub") = ""
				.Parameters("FileType") = strFileType
                           
                .ActiveConnection = objCon
                
            End With
                
            objCmd.Execute        
            
            'Return the result of the Save Function.
     		intRecord = objCmd.Parameters.Item("CSFromDinersIDOutput")    
     		                  			     				     		     		
       'response.write  "exec spGeneralExpensesSave =0," & Session("BudgetID") & "," & Session("VersionID") & "," & CostCentreID & ",'GEXP'" & GLCode & "," & BM1 & "," & BM2 & "," & BM3 & "," & BM4 & "," & BM5 & "," & _
       '                     BM6 & "," & BM7 & "," & BM8 & "," & BM9 & "," & BM10 & "," & BM11 & "," & BM12 & "," & OY1 & "," & OY2 & "," & OY3 & ",'" & Comments & "','" & UpdatedBy & "'," & Session("ColumnLock")
End Sub

Public Function ProcessFile(strFileSeqNum,lngFileID,strCardType,strFileType)
'Function to Process an ANZ File which has been loaded into the database (update all changes from the ANZ file and add changes to the audit log)
Dim intRecord
Dim strFileID

	With objCmd

		.CommandType = 4
		If strCardType = "DTC" Then
			.CommandText = "spCAPSDinersProcessCSFile"		
		Else
			.CommandText = "spCAPSDinersProcessCSFileDPC"	
		End If

		.Parameters.Append objCmd.CreateParameter("UserID", adInteger)
		.Parameters.Append objCmd.CreateParameter("FileID", advarchar,adParamInput,6)
		.Parameters.Append objCmd.CreateParameter("CardType", advarchar,adParamInput,20)
		.Parameters.Append objCmd.CreateParameter("CardTypeSub", advarchar,adParamInput,20)
		.Parameters.Append objCmd.CreateParameter("FileType", advarchar,adParamInput,20)
		.Parameters.Append objCmd.CreateParameter("CAPSDinersProcessCSFileOutput", adInteger, adParamOutput)
		
		.Parameters("UserID") = Session("UserID")
		.Parameters("FileID") = strFileSeqNum
		.Parameters("CardType") = Left(strCardType,3)
		.Parameters("CardTypeSub") = ""
		.Parameters("FileType") = strFileType
		
		.ActiveConnection = objCon
		 
	End With
   
	objCmd.Execute        
  
	'Return the result of the Save Function.
	intRecord = objCmd.Parameters.Item("CAPSDinersProcessCSFileOutput") 

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Records Processed in CS Upload File " & strFileID & ". Please notify System Administrators.</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert""> " & intRecord & " CS Upload File records Processed in file " & strFileSeqNum & "</div>"
		'Call UpdateFileLoadSummary ("CSFromDiners",strFileSeqNum, "", lngFileID)
	End If
	
	ProcessFile = intRecord
	
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

Public Sub DisplayFileSummary(strFileNameDefault)

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile
Dim strFileSize

Set objFSO = CreateObject("Scripting.FileSystemObject")

	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\CDMC\")
	objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Diners\DinersFrom"
	
	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"">Files To Be Loaded<i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be loaded. Click 'Load Server' button to Load CS File.""></i></th></tr>"
	
	intCount = 0
	
	'Response.Write "<BR>##" & Mid(strFileNameDefault,24,1) & "##<BR>"
	Dim intFileLen
	
	intFileLen = Len(strFileNameDefault) 
	
	If strFileType = "CSFromDinersDPCU" Then
	
		For Each objFile in colFiles
			
			If Left(objFile.Name,intFileLen) = Trim(strFileNameDefault) AND Mid(objFile.Name,24,1) = "U" Then			
			
				intCount = intCount + 1
			
				If intCount < 6 Then
					If IsNull(objFile.Name) or objFile.Name = "" Then
						strFile = ""
						strFileSize = 0
					Else
						strFile = Left(objFile.Name,10) & ".." & Right(objFile.Name,4)
						strFileSize = Round(objFile.Size/1024000,4)
					End If
					
					Response.Write "<TR><TD Title=""" & objFile.Name & " - Size: " & strFileSize & " MB"">" & strFile & "</TD></TR>"
				End If
				
			End If
			
		Next
		
	Else
	
		For Each objFile in colFiles
			
			If Left(objFile.Name,intFileLen) = Trim(strFileNameDefault) AND Mid(objFile.Name,24,1) <> "U"  Then			
			
				intCount = intCount + 1
			
				If intCount < 6 Then
					If IsNull(objFile.Name) or objFile.Name = "" Then
						strFile = ""
						strFileSize = 0
					Else
						strFile = Left(objFile.Name,10) & ".." & Right(objFile.Name,4)
						strFileSize = Round(objFile.Size/1024000,4)
					End If
					
					Response.Write "<TR><TD Title=""" & objFile.Name & " - Size: " & strFileSize & " MB"">" & strFile & "</TD></TR>"
				End If
				
			End If
			
		Next	
	
	End If
	
	 If intCount > 1 And strCardType = "DTC" Then
		Response.Write "<tr><th style=""text-align:left""><FONT Color=""RED"">WARNING MORE THAN 1 " & strCardType & " FILE EXISTS IN LOAD FOLDER: " & intCount & "</FONT>&nbsp;<i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""There should only be 1 " & strCardType & " file in the load folder.  This can cause issues with the loading sequence.""></i></th></tr></table>"
	 Else
		Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr>"
		Response.Write "<tr><th style=""text-align:left"">Size: " & strFileSize & " MB</th></tr></table>"
	 End If
	 
Set objFSO = Nothing
'Set outPut = Nothing

End Sub


Public Sub DisplayFileSummaryG(strFileNameDefault)

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile
Dim strFileSize

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass


Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get the System Parameter for the start of the Training File Location
strServer = GetSystemAdmin("GDriveFilePath")

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
	
	'objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\Diners\DinersFrom"

	Set objFolder = objFSO.GetFolder(objStartFolder)
	
	
		
	Set colFiles = objFolder.Files
	
	Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
			"<tr><th style=""text-align:left"" Title=""" & strServer & """>G Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the G Drive waiting to be loaded. Click 'Load G' button to Load CS File.""></i></th></tr>"
	
	intCount = 0
	
	Dim intFileLen
	
	intFileLen = Len(strFileNameDefault) 
	
	If strFileType = "CSFromDinersDPCU" Then
	
		For Each objFile in colFiles
			
			If Left(objFile.Name,intFileLen) = Trim(strFileNameDefault) AND Mid(objFile.Name,24,1) = "U" Then			
			
				intCount = intCount + 1
			
				If intCount < 6 Then
					If IsNull(objFile.Name) or objFile.Name = "" Then
						strFile = ""
						strFileSize = 0
					Else
						strFile = Left(objFile.Name,10) & ".." & Right(objFile.Name,4)
						strFileSize = Round(objFile.Size/1024000,4)
					End If
					
					Response.Write "<TR><TD Title=""" & objFile.Name & " - Size: " & strFileSize & " MB"">" & strFile & "</TD></TR>"
				End If
				
			End If
			
		Next
		
	Else
	
		For Each objFile in colFiles
			
			If Left(objFile.Name,intFileLen) = Trim(strFileNameDefault) AND Mid(objFile.Name,24,1) <> "U"  Then			
			
				intCount = intCount + 1
			
				If intCount < 6 Then
					If IsNull(objFile.Name) or objFile.Name = "" Then
						strFile = ""
						strFileSize = 0
					Else
						strFile = Left(objFile.Name,10) & ".." & Right(objFile.Name,4)
						strFileSize = Round(objFile.Size/1024000,4)
					End If
					
					Response.Write "<TR><TD Title=""" & objFile.Name & " - Size: " & strFileSize & " MB"">" & strFile & "</TD></TR>"
				End If
				
			End If
			
		Next	
	
	End If
	 If intCount > 1 And strCardType = "DTC" Then
		Response.Write "<tr><th style=""text-align:left""><FONT Color=""RED"">WARNING MORE THAN 1 " & strCardType & " FILE EXISTS IN LOAD FOLDER: " & intCount & "</FONT>&nbsp;<i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""There should only be 1 " & strCardType & " file in the load folder.  This can cause issues with the loading sequence.""></i></th></tr></table>"
	 Else
		Response.Write "<tr><th style=""text-align:left"">Total: " & intCount & "</th></tr>"
		Response.Write "<tr><th style=""text-align:left"">Size: " & strFileSize & " MB</th></tr></table>"
	 End If

	 
objNetwork.RemoveNetworkDrive strServer, True, False
	 
Set objFSO = Nothing
Set objNetwork = Nothing

'Set objFSO = Nothing
'Set outPut = Nothing

End Sub

Set objRS = Nothing
Set objCon = Nothing

 %>


