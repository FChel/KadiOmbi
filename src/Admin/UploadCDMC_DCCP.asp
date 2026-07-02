<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #include file="upload.asp" -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #include file="../CC/CAPSFunctions.asp" -->
<%

'On Error Resume Next
'Description:	ANZ Cardlist Upload Administration screen
'Author:		Michael Giacomin
'Date:			May 2020

	Server.ScriptTimeout = 6000
	
    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")
	
'Instantiate Common Page Variables.
Dim objCon
Dim objCmd
Dim objCmd1
Dim objCmd2
Dim objCmd3
Dim objRS

Dim strDeleteCheck
Dim dteBatchDate
Dim strLoadType

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objCmd1 = Server.CreateObject("ADODB.Command")
Set objCmd2 = Server.CreateObject("ADODB.Command")
Set objCmd3 = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("ServerLocationCDMC")) Then

	Session("ServerLocationCDMC") = Request.QueryString("ServerLocationCDMC")
	
End If



If IsNull(Session("ServerLocationCDMC")) OR Session("ServerLocationCDMC") = "" Then Session("ServerLocationCDMC") = "CDMC"

If request.QueryString("Action")="Save" Then

	Call StartLoad()
End If

If Not IsEmpty(Request.QueryString("Reload")) Then

	Call StartLoad()
End If

'If the local load has been clicked then call the procedure to load the network file rather than uploading it
If Request.QueryString("Action")="SaveFileLocal" Then
	
	Call StartLoad()

End If

If request.QueryString("Action")="Save12" Then

	Response.write "Loaded from...."
End If


'If the Process CDMC has been clicked then call the procedure to Process the Current CDMC against the Diners Cards to add to the CS To Diners file
If Request.QueryString("Action")="ProcessCDMC" Then
	
	Call ProcessCDMC()

End If

'If refresh CDMC button has been clicked then rn the procedure
If Request.QueryString("Action")="SynchCDMC" Then
	
	Call SynchCDMC()

End If


If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
	
End If

'If the Process button has been clicked next to a file, then call the Process procedure
If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "ProcessA" Then
		Call ProcessFile("A",Request.QueryString("FileSeqNum"),Request.QueryString("FileName"),Request.QueryString("FileLoadID"))
	End If
End If

'If the Process button has been clicked next to a file, then call the Process procedure
If Not IsEmpty(Request.QueryString("Action")) Then
	If Request.QueryString("Action") = "ProcessB" Then
		Call ProcessFile("B",Request.QueryString("FileSeqNum"),Request.QueryString("FileName"),Request.QueryString("FileLoadID"))
	End If
End If

 %>

<html>
<head>
	
	
<script language=javascript>

var varProgTime;
var varProgNum;

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

//Set a timer to call the Get Progress Num function which will return the Number of records loaded to the screen
//varProgTime = setTimeout(GetProgressNum, 3000);
//window.setInterval(GetProgressNum, 2000);

//GetProgressNum();

//window.setInterval(loadDoc("../CC/AJAX/GetProgressNum.asp", myFunction1), 2000);
//varProgTime = setTimeout(GetProgressNum, 3000);

var check = false

  if ( check==true) {
        // Returns true if checked
	//LoadCDMCAJAX('on');
   	self.location="UploadCDMC_DCCP.asp?Action=SaveFileLocal&Delete=on"
    } else {
        // Returns false if not checked
	//LoadCDMCAJAX('off');
	
	self.location="UploadCDMC_DCCP.asp?Action=SaveFileLocal&Delete=off"
    }

	//clearTimeout(varProgTime);
}

function UploadLocalG()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="UploadCDMC_DCCP.asp?Action=SaveFileLocal&Delete=off&ActionType=Service"
}

function synchCDMC()
{
	document.getElementById('Progress3').style.display = "inline";
	self.location="UploadCDMC_DCCP.asp?Action=SynchCDMC"
}



function ConfirmCDMC()
{

var varActionType = document.getElementById("LoadCDMCConfirmType").value

	if(varActionType===null || varActionType===""){
		varActionType="Service";
};
document.getElementById('Progress4').style.display = "inline";

	//self.location="UploadCDMC_DCCP.asp?Action=Save12"
	//self.location="UploadCDMC_DCCP.asp?Action=SaveFileLocal&Delete=off&ActionType=Service"
	self.location="UploadCDMC_DCCP.asp?Action=SaveFileLocal&Delete=off&ActionType=" + varActionType



}

function OpenCDMC(varPath)
{
	//Replace slases with multiples otherwise they will not display
	varPath = str.replace(/\\/g, replacement)
document.getElementById('LoadCDMCConfirmLocationPath').value = varPath;

}

function ProcessCDMC()
{

document.getElementById('Progress3').style.display = "inline";

	self.location="UploadCDMC_DCCP.asp?Action=ProcessCDMC"

}

function LoadProcessA(FileSeqNum,FileName,FileLoadID)
{

	document.getElementById("FileSeqNum").value = FileSeqNum;
	document.getElementById("FileName").value = FileName;
	document.getElementById("FileLoadID").value = FileLoadID;
   
}

function LoadProcessB(FileSeqNum,FileName,FileLoadID)
{

	document.getElementById("FileSeqNumB").value = FileSeqNum;
	document.getElementById("FileNameB").value = FileName;
	document.getElementById("FileLoadIDB").value = FileLoadID;
   
}

function RunProcessA()
{

	var FileSeqNum = document.getElementById("FileSeqNum").value
	var FileName = document.getElementById("FileName").value
	var FileLoadID = document.getElementById("FileLoadID").value
	
   document.getElementById('Progress1').style.display = "inline";

   self.location="UploadCDMC_DCCP.asp?Action=ProcessA&FileSeqNum=" + FileSeqNum + "&FileName=" + FileName + "&FileLoadID=" + FileLoadID 

}

function RunProcessB()
{

	var FileSeqNum = document.getElementById("FileSeqNumB").value
	var FileName = document.getElementById("FileNameB").value
	var FileLoadID = document.getElementById("FileLoadIDB").value
	
   document.getElementById('Progress2').style.display = "inline";

   self.location="UploadCDMC_DCCP.asp?Action=ProcessB&FileSeqNum=" + FileSeqNum + "&FileName=" + FileName + "&FileLoadID=" + FileLoadID 

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


function loadDoc(url, cFunction) {
  var xhttp;
  xhttp=new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
      cFunction(this);
    }
 };
  xhttp.open("GET", url, true);
  xhttp.send();
}

function myFunction1(xhttp) {
  // action goes here
  document.getElementById("ProgressNum").innerHTML = xhttp.responseText;
}
function myFunction2(xhttp) {
  // action goes here
  document.getElementById("LoadCDMCAJAXDiv").innerHTML = xhttp.responseText;
}


function GetProgressNum() {
//loadDoc("../CC/AJAX/GetProgressNum.asp", myFunction1);

  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {

    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ProgressNum").innerHTML = this.responseText;
	 
	 //varProgNum = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetProgressNum.asp", true);
  xhttp.send();
}


function LoadCDMCAJAX(varDelete) {

loadDoc("../CC/AJAX/LoadCDMC.asp?Action=SaveFileLocal&Delete="+varDelete, myFunction2);

//  var xhttp = new XMLHttpRequest();
//  xhttp.onreadystatechange = function() {
  
//    if (this.readyState == 4 && this.status == 200) {
//     document.getElementById("LoadCDMCAJAXDiv").innerHTML = this.responseText;
//    }
//  };
//  xhttp.open("GET", "../CC/AJAX/LoadCDMC.asp?Action=SaveFileLocal&Delete="+varDelete, true);
//  xhttp.send();
}



function DatePickChange() {
	self.location="UploadCDMC_DCCP.asp?FileDate=" + document.getElementById("CSDate").value;
}
</script>

<style>

    table.newd th, table.newd td{

        padding: 4px; 

    }

</style>
</head>
<body>

<main class="main py-3">
      <div class="container">
<form action="UploadCDMC_DCCP.asp?Action=Save" method="POST" enctype="multipart/form-data" id="frm" name="frm">
<!-- Select Batch Number Modal -->
<div class="modal fade" id="ProcessARun" tabindex="-1" role="dialog" aria-labelledby="ProcessARun" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalReleaseTitle" style="font-weight:bold;">Execute the 'Process History' procedure for the specified Batch No?</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
	  <div class="modal-header">
	  <h6 style="color:navy;">Batch&nbsp;No:&nbsp;&nbsp;</h6>
	  <div><input type="text" name="FileSeqNum" id="FileSeqNum" class="form-control input-md" READONLY value="">
	  <input type="hidden" name="FileName" id="FileName" class="form-control input-md" value="">
	  <input type="hidden" name="FileLoadID" id="FileLoadID" class="form-control input-md" value=""></div>
	  </div>
      <div class="modal-body" id="ErrorDetailsMod">
         <span id="Progress1" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Processing CDMC History...</b></span>
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
        <button type="button" class="btn btn-primary" onClick="RunProcessA();" ><i class="fa fa-check"></i> Yes</button>
		<input type="hidden" id="NewStatus" name="NewStatus" value=""/>
      </div>
    </div>
  </div>
</div>
<!-- End Select Batch Number Modal -->

<!-- Select Batch Number Modal -->
<div class="modal fade" id="ProcessBRun" tabindex="-1" role="dialog" aria-labelledby="ProcessBRun" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="ModalReleaseTitle" style="font-weight:bold;">Execute the 'Process Details' procedure for the specified Batch No?</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
	  <div class="modal-header">
	  <h6 style="color:navy;">Batch&nbsp;No:&nbsp;&nbsp;</h6>
	  <div><input type="text" name="FileSeqNumB" id="FileSeqNumB" class="form-control input-md" READONLY value="">
	  <input type="hidden" name="FileNameB" id="FileNameB" class="form-control input-md" value="">
	  <input type="hidden" name="FileLoadIDB" id="FileLoadIDB" class="form-control input-md" value=""></div>
	  </div>
      <div class="modal-body" id="ErrorDetailsMod">
         <span id="Progress2" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Processing CDMC Details...</b></span>
		
		 
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
        <button type="button" class="btn btn-primary" onClick="RunProcessB();" ><i class="fa fa-check"></i> Yes</button>
		<input type="hidden" id="NewStatus" name="NewStatus" value=""/>
      </div>
    </div>
  </div>
</div>
<!-- End Select Batch Number Modal -->


<!-- Select Synch CDMC  Modal -->
<div class="modal fade" id="SynchCDMCMod" tabindex="-1" role="dialog" aria-labelledby="SynchCDMCMod" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="SynchCDMCModTitle" style="font-weight:bold;">Copy CDMC Load records to CDMC table (refresh)?</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
	  <div class="modal-header">
	  <h6 style="color:navy;">Admin CDMC Refresh. Run procedure?</h6>
	  <H7> - current Records in tblCapsCDMCLoad: 

		<b> <%=GetCMDCLoad() %> </b>
	</H7>
	
	  </div>
      <div class="modal-body" id="ErrorDetailsMod">
	<p>
	This procedure will remove all records in tblCAPSCDMCPortal and tblCAPSCDMCLoad and replace them with records in tblCAPSCDMC.<br><br>
	In order to assist with data loading of CDMC and to avoid empty tables due to slow loading, data will be loaded to tblCAPSCDMC as per normal CAPS process, then the Portal and Load tables updated for DCCP via this refresh.
	<br><br>
	<b>NOTE:</b> This refresh is only needed if the existing load process (Load CDMC and Process CDMC buttons on this screen) don't refresh tblCAPSCDMC.<br>You can see if Total CDMC Load and Portal Records doesn't match Total CDMC records at the top of this screen.<br>
	In this case click Yes below.
	</p>
         <span id="ProgressSynch" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Synching CDMC tables...</b></span>
		
		 
      </div>
      
       <div class="modal-footer d-flex justify-content-end" style="padding:5px; margin:0px;">
		<button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
        <button type="button" class="btn btn-primary" onClick="synchCDMC();" ><i class="fa fa-check"></i> Yes</button> </div>
		
     
    </div>
  </div>
</div>
<!-- End Synch CDMC Number Modal -->


<!-- Load Selected CDMC File from Location Modal -->
<div class="modal fade" id="LoadCDMCConfirm" tabindex="-1" role="dialog" aria-labelledby="LoadCDMCConfirm" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="LoadCDMCConfirmTitle" style="font-weight:bold;">Load CDMC File From?</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
	 <div class="modal-header">
		<div class="col-lg-12 col-md-12">
			<div class="row" id="basic-table">
			<div class="col-lg-4 col-md-4">
	  		Selected CDMC File
			</div>
			<div class="col-lg-8 col-md-8">
				<input type="text" name="LoadCDMCConfirmLocation" id="LoadCDMCConfirmLocation" class="form-control input-md" READONLY value="<%=Session("ServerLocationCDMC")%>">
			</div>
			</div>
				<div class="row" id="basic-table">
				<div class="col-lg-4 col-md-4">Selected CDMC File Path
				</div>
				  <div class="col-lg-8 col-md-8">
					<input type="text" name="LoadCDMCConfirmLocationPath" id="LoadCDMCConfirmLocationPath" class="form-control input-md" READONLY value="<%=Session("ServerLocationCDMCPath")%>">
				  </div>
			</div>

				<div class="row" id="basic-table">
				<div class="col-lg-4 col-md-4">Load Type
				</div>
				  <div class="col-lg-8 col-md-8">
					<input type="text" name="LoadCDMCConfirmType" id="LoadCDMCConfirmType" class="form-control input-md" READONLY value="">
				  </div>
			</div>
		</div>
</div>
      <div class="modal-body" id="LoadCDMCConfirmLocationMod">
         <span id="Progress4" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading CDMC File...</b></span>
		
		 
      </div>
      <div class="modal-footer">
		<button type="button" class="btn btn-primary" onClick="ConfirmCDMC();" ><i class="fa fa-check"></i> Yes</button>
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
        
      </div>
    </div>
  </div>
</div>
<!-- End Load Selected CDMC File from Location Modal -->

<!-- Modal -->
	<div class="loader" id="ModalWait">
        <div class="wrap" >
            <div class="spinner"></div>
            <span class="loading-message">Loading...</h6>
        </div>
    </div>
	
<div class="content-wrapper">
    <div class="container-fluid">
	 <div class="col-12" id="LoadCDMCAJAXDiv"></div>
	 <div class="row" id="basic-table">
  <div class="col-3">
    <div class="card">
     <div class="card-header" style="background-color: #E9EFFC;">
          <h4 class="card-title"><img src="../CC/img/DFG_Logo.png" height="40px" width="100px" title="Certificate and Directory Management Centre"><br><br> DCCP CDMC Load </h4>
		 
        </div>
      <div class="card-content">
        <div class="card-body">
		

<div class="col-lg-12 col-md-12">

</div>

<div class="form-body">
<div class="col-lg-12 col-md-12" >
<!--
<div class="row" style="text-align:right; d-flex justify-content-end">
<button type="button" class="btn btn-outline-primary btn-sm" onclick="UploadLocalG();" Title="Click to Load any existing file in the CDMC G Drive Folder"><i class="fa fa-upload"></i> Load G</button>
<button type="button" class="btn btn-outline-secondary btn-sm" onclick="UploadLocal();" Title="Click to Load any existing file in the CDMC Server (Local) Folder"><i class="fa fa-upload"></i> Load Local</button>
</div>
-->
<div class="row" style="text-align:right; d-flex justify-content-end; padding-top:5px;">
<!--<button type="button" class="btn btn-primary btn-sm" data-toggle=""modal"" data-target=""#LoadCDMCModal"" Title="Click to Load any existing file in the selected Folder"><i class="fa fa-upload"></i> Load CDMC File</button>-->
<button type="button" class="btn btn-primary btn-sm" data-toggle="modal" data-target="#LoadCDMCConfirm" onClick="OpenCDMC('<%=Session("ServerLocationCDMCPath")%>')" Title="Click to Load any existing file in the CDMC Drive Folder Selected"><i class="fa fa-upload"></i> Load CDMC</button>
</div>
</div>
</div>
<div class="col-lg-12 col-md-12">


</div>
<p class="text-left">
<div>
<span id="Progress" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b><DIV id="ProgressNum" style="display: inline; font-weight: bold;"></DIV></span>
</div>


<div class="col-lg-12 col-md-12">

</div>
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
		
		
	<div class="card">
      <div class="card-content">
	   <div class="card-header">
          <h6 class="card-title" style="text-align:center;">Process CDMC Updates for CS File To Diners</h6>
        </div>
        <div class="card-body">
		

			<div class="col-lg-12 col-md-12">
				<div class="row" style="text-align:right; d-flex justify-content-end">
				<button type="button" class="btn btn-outline-secondary btn-sm" onclick="ProcessCDMC();" Title="Click to Process the current CDMC File against the Current Diners Card list to add changes to the CS File To Diners"><i class="fa fa-cogs"></i> Process CDMC</button>
					&nbsp;&nbsp;
				<button type="button" class="btn btn-outline-info btn-sm" data-toggle="modal" data-target="#SynchCDMCMod" Title="NOTE: Admin function only -- to refresh tblCDMC from tblCDMCLoad if normal load preocedures did not complete"><i class="fa fa-recycle"></i> </button>

				</div>
			</div>
			<p class="text-left">
			<div>
			<span id="Progress3" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b><DIV id="ProgressNum" style="display: inline; font-weight: bold;"></DIV></span>
			</div>
			</p>
		</div>
	  </div>
    </div>
		
<!--END of 3 Column  <div class="col-3"> -->	
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

  <div class="col-2">
    <!--<div class="card">     
      <div class="card-content">
        <div class="card-body">
			<%DisplayFileSummary()%>	
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

<div class="card">     
      <div class="card-content">
        <div class="card-body">
			<%DisplayFileSummaryCDMC()%>	
		</div>
	  </div>
    </div>
  </div>-->

<div class="card">     
      <div class="card-content">
        <div class="card-body" style="padding:1px; margin:1px;">
		<table Class="table table-bordered mb-0 newd" cellspacing="0" cellpadding="0">
	<tr><th style="text-align:left">Files To Be Loaded (select location below)<i class="fa help-tooltip fa-question-circle" data-toggle="tooltip" title="Files on the CAPS Server (Local), G Drive or CDMC server waiting to be loaded. Click to change selection of where to load file from."></i></th></tr>
	
			<%DisplayFileSummary()%>	
		
			<%DisplayFileSummaryG()%>	
		
			<%DisplayFileSummaryCDMC()%>	
	</table>
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
</body>
<script>
$(document).ready(function(){
	//alert($('#LoadCDMCConfirmLocationPath').val());
	if (!$('#LoadCDMCConfirmLocationPath').val()) {
	var text = '<%=Session("ServerLocationCDMCPath")%>';
	var result  = text.replace(/\//g,'-');

    $("#LoadCDMCConfirmLocationPath").val(result);

	}
	//alert("<%=strLoadType%>");
	$("#LoadCDMCConfirmType").val("<%=strLoadType%>");
});
</script>
<!-- #Include file=../CC/CAPSFooter.asp -->
</html>

<%
Sub DisplayTableDetails()

Dim strWhere
Dim strAddress1
Dim strAddress2
Dim strAddress3

					
If Not IsEmpty(Request.QueryString("BatchNo")) Then
	If IsNull(Request.QueryString("BatchNo")) or Request.QueryString("BatchNo")= "" Then 
		
	Else
		strWhere = "WHERE FileID = " & Request.QueryString("BatchNo") & ""
	
		If Request.QueryString("BatchNo") = 0 Then strWhere = strWhere & " OR FileID = 0"
	End If
Else
	strWhere = ""
End If

	objRS.Open "SELECT TOP 50 * FROM tblCAPSCDMCLoad WITH(NOLOCK) "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no CDMC Load data loaded</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Upload Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of Existing CDMC Load Data already in CAPS</th></tr>" & _
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
				
				If IsNull(objRS("PostalAddress_Unit")) Then
					strAddress1 = ""
				Else
					If Len(objRS("PostalAddress_Unit")) > 10 Then
						strAddress1 = Left(objRS("PostalAddress_Unit"),10)
					Else
						strAddress1 = objRS("PostalAddress_Unit")
					End If
				End If
				
				If IsNull(objRS("PostalAddress_ClientLocation")) Then
					strAddress2 = ""
				Else
					If Len(objRS("PostalAddress_ClientLocation")) > 10 Then
						strAddress2 = Left(objRS("PostalAddress_ClientLocation"),10)
					Else
						strAddress2 = objRS("PostalAddress_ClientLocation")
					End If
				End If
				
				If IsNull(objRS("PostalAddress_DeliveryLocation")) Then
					strAddress3 = ""
				Else
					If Len(objRS("PostalAddress_DeliveryLocation")) > 10 Then
						strAddress3 = Left(objRS("PostalAddress_DeliveryLocation"),10)
					Else
						strAddress3 = objRS("PostalAddress_DeliveryLocation")
					End If
				End If
				
				
			    Response.Write "<TR class='clickable-row' data-href='UploadCDMCDetail.asp?BatchNo=" & objRS("FileID") & "&EIDNo=" & objRS("EmployeeID") & "' style=""cursor: pointer;""><TD style=""text-align:center; font-size:12px;"">" & objRS(0) & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS("EmployeeID") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("GroupName") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("ActualRankLvl") & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS("Title") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("FormalFirstName") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("FormalLastName") & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & strAddress1 & "</TD><TD style=""text-align:center; font-size:12px;"">" & strAddress2 & "</TD><TD style=""text-align:center; font-size:12px;"">" & strAddress3 & "</TD>" & _
			                    "<TD style=""text-align:center; font-size:12px;"">" & objRS("PostalAddress_City") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("PostalAddress_State") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("PostalAddress_PostCode") & "</TD>" & _
								"<TD style=""text-align:center; font-size:12px;"">" & objRS("Email_Address") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("hasChanged") & "</TD><TD style=""text-align:center; font-size:12px;"">" & objRS("FileID") & "</TD>" & _
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

Dim strDateLoaded
Dim strDateLoadedTitle

Dim dteBatchDateFormat
Dim srStatus1
Dim strDateUpdated
Dim strStatus
Dim strAction

Dim strDateLoadColour
Dim strDateTitle

Dim strRecordCount
Dim strAddressCount
Dim strEmployeeCount

Dim strRecordCountMessage
Dim strRecordCountColour

Dim strSameRecordCountMessage
Dim strSameRecordCountColour
Dim lngTotalRecordsPrevious

Dim strDupeEmployeeMessage
Dim strDupeEmployeeColour

Dim strCountRec

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)
	
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Right("0" & Month(dteBatchDate), 2) & "-" & Right("0" & Day(dteBatchDate), 2)

		'''Get the Count for tblCapsCDMCLoad
		objRS.Open "SELECT Count(*) AS Count, MAX(DateUpdated) AS DateLoaded FROM tblCAPSCDMCLoad WITH(NOLOCK)",objCon

		 If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th style=""text-align:left"">There is no CDMC Load data loaded </th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
		        "<tr><th >Total CDMC Load Records</th>" & _
	 	        "<th>Date Loaded</th></tr>" 
			
			strCountRec = objRS("Count")
			If IsNumeric(strCountRec) Then strCountRec = FormatNumber(strCountRec,0)

			Response.Write "<TR><TD style=""text-align:center;"">" & strCountRec & "</TD>" & _
					"<TD style=""text-align:center;"">" & objRS("DateLoaded") & "</TD></TR>"

		    End If
		
		objRS.Close


		'''Get the Count for tblCapsCDMCPortal
		objRS.Open "SELECT Count(*) AS Count, MAX(DateUpdated) AS DateLoaded FROM tblCAPSCDMCPortal WITH(NOLOCK)",objCon

		 If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th style=""text-align:left"">There is no CDMC Portal data loaded </th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
		        "<tr><th >Total CDMC Portal Records <br><span style=""font-size:10px;""></span></th>" & _
	 	        "<th>Date Loaded</th></tr>" 
			
			strCountRec = objRS("Count")
			If IsNumeric(strCountRec) Then strCountRec = FormatNumber(strCountRec,0)

			Response.Write "<TR><TD style=""text-align:center;"">" & strCountRec & "</TD>" & _
					"<TD style=""text-align:center;"">" & objRS("DateLoaded") & "</TD></TR>"

		    End If
		
		objRS.Close


		'''Get the Count for tblCapsCDMCLoad
		objRS.Open "SELECT Count(*) AS Count, MAX(DateUpdated) AS DateLoaded FROM tblCAPSCDMC WITH(NOLOCK)",objCon

		 If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th style=""text-align:left"">There is no CDMC data loaded </th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
		        "<tr><th >Total CDMC Records <br><span style=""font-size:10px;"">(should match above Total CDMC Load Records, if not click refresh button) - <i class=""fa fa-recycle""></i></span></th>" & _
	 	        "<th>Date Loaded</th></tr>" 
			
			strCountRec = objRS("Count")
			If IsNumeric(strCountRec) Then strCountRec = FormatNumber(strCountRec,0)

			Response.Write "<TR><TD style=""text-align:center;"">" & strCountRec & "</TD>" & _
					"<TD style=""text-align:center;"">" & objRS("DateLoaded") & "</TD></TR>"

		    End If
		
		objRS.Close



		objRS.Open "SELECT TOP 6 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE DateLoaded > '" & dteBatchDate & "' AND FileType = 'CDMC' ORDER BY DateLoaded DESC",objCon
		'objRS.Open "SELECT TOP 50 * FROM qryCAPSCSFromDinersSummary WHERE DateUpdated > '" & dteBatchDate & "' ORDER BY BatchNo DESC",objCon
		
		    If objRS.EOF Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                	"<tr><th colspan=""8"" style=""text-align:left"">There is no CDMC data loaded (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();""/>)</th></tr>" 
		        
		    Else
		    
		         Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                	"<tr><th colspan=""6"" style=""text-align:left"">CDMC Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" onChange=""DatePickChange();"" />)&nbsp;<i class=""fa fa-user-secret"" onclick=""self.location='UploadCDMCOld.asp'""></i></th></tr>" & _
		        "<tr><th Style=""width:20px;"">Batch No.</th>" & _
				"<th>Total Records</th>" & _
				"<th>No Address</th><th>Total Employees</th>" & _	
				"<th>Status</th>" & _
	 	        "<th>Date Loaded</th></tr>" 
			
			

		    End If
		    

			'New variable to keep the previous record record count to see of the next os the same to alert the user of file issues
			lngTotalRecordsPrevious = 0
			
		    Do until objRS.eof
			
				If IsNull(objRS("Status")) Then
					strStatus = ""
				Else
					strStatus = objRS("Status")
				End If
				
				If DateDiff("d",objRS("DateLoaded"),Now()) = 0 Then
					strDateLoadColour = " Color:Green; font-weight:bold;"
					strDateTitle = " - CDMC File for today has been LOADED "
				Else
					strDateLoadColour = ""
					strDateTitle = ""
				End If
				
				'Get the Date Loaded and Format for display with the full date as ther title
				If IsNull(objRS("DateLoaded")) Then
					strDateLoaded = ""
				Else
					strDateLoaded = FormatDateTime(objRS("DateLoaded"),vbShortDate)
					strDateLoadedTitle = "title=""" & objRS("DateLoaded") & " " & strDateTitle & " """
				End If
				
				If strStatus = "Imported" AND objRS("Deleted") = "N" Then
					'strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ProcessARun"" onclick=""ProcessA('" & objRS("FileSeqNum") & "','" & objRS("FileName") & "','" & objRS("FileLoadID") & "');"">Process History</button>"
					strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ProcessARun"" onclick=""LoadProcessA('" & objRS("FileSeqNum") & "','" & objRS("FileName") & "','" & objRS("FileLoadID") & "');"" title=""New process related to the load table tblCAPSCDMCLoad which will run the procedure spCAPSCDMCUpdateFromTempLoad on the data loaded to the table in this screen."">Process History Load</button>"
					
					'strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='UploadANZ.asp?Action=Process&FileLoadID=" & objRS("FileSeqNum") & "'""; title=""Click to Process the ANZ File and update changes to cards in CAPS from the ANZ File loaded " & objRS("DateLoaded") & """><i class=""fa fa-cogs""></i> Process</button>"
				ElseIf strStatus = "History Processed" AND objRS("Deleted") = "N" Then
					'strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModApp"" onclick=""self.location='UploadCDMC_DCCP.asp?Action=ProcessB&FileSeqNum=" & objRS("FileSeqNum") & "&FileName=" & objRS("FileName") & "&FileLoadID=" & objRS("FileLoadID") & "'""; title=""Click to Process the CDMC File and update changes to Corporate Directory " & objRS("DateLoaded") & ". "">Process Details</button>"
					strAction = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ProcessBRun"" onclick=""LoadProcessB('" & objRS("FileSeqNum") & "','" & objRS("FileName") & "','" & objRS("FileLoadID") & "');"" title=""This is back to the normal CAPS process in tblCAPSCDMC, not temp load table. Can be run here or the normal Load CDMC screen."">Process Details</button>"
				Else
					strAction = objRS("Status")
				End If
				
				'Get the Total records and format
				If IsNull(objRS("RecordCount")) or objRS("RecordCount") = "" Then
					strRecordCount = 0
				Else
					If IsNumeric(objRS("RecordCount")) Then
						strRecordCount = FormatNumber(objRS("RecordCount"),0)
						'check to make sure that the number of records loaded is about normal (above previous numbers) as loading an incomplete file will cause processing issues and cancel cards
						'If red text is displayed then the file should be checked and loaded again if it is not the full file count in CAPS (CAPS and CSV do not reconcile)
						If strRecordCount < 150000 Then
							strRecordCountMessage = "title=""The Number of records loaded looks to be below normal. Please SEE SYSTEM ADMINISTRATORS before continuing!"""
							strRecordCountColour = "color:red; font-weight:bold;"
						Else
							strRecordCountMessage = ""
							strRecordCountColour = ""
						End If
						
						'''New April 2023 -- Check current record count with previous record count due to common issue of CDMC files not being updated, to warn the user
						If lngTotalRecordsPrevious = strRecordCount Then
							strSameRecordCountMessage = "title=""The Number of records loaded is the same as the previous file. Please SEE SYSTEM ADMINISTRATORS before continuing!"""
							strSameRecordCountColour = "color:red; font-weight:bold;"
						Else
							strSameRecordCountMessage = ""
							strSameRecordCountColour = ""
						End If
						
						strEmployeeCount = FormatNumber(objRS("EmployeeCount"),0)
						If strRecordCount <> strEmployeeCount Then
							strSameRecordCountMessage = "title=""The number of employees does not match the number of records. Please SEE SYSTEM ADMINISTRATORS before continuing!"""
							strSameRecordCountColour = "color:red; font-weight:bold;"
						Else
							strSameRecordCountMessage = ""
							strSameRecordCountColour = ""
						End If	
					Else
						strRecordCount = 0
					End If
				End If
				
				If IsNull(objRS("CardCount")) or objRS("CardCount") = "" Then
					strAddressCount = 0
				Else
					If IsNumeric(objRS("CardCount")) Then
						strAddressCount = FormatNumber(objRS("CardCount"),0)
					Else
						strAddressCount=0
					End If
				End If
			
				If IsNull(objRS("EmployeeCount")) or objRS("EmployeeCount") = "" Then
					strEmployeeCount = 0
				Else
					If IsNumeric(objRS("EmployeeCount")) Then
						strEmployeeCount = FormatNumber(objRS("EmployeeCount"),0)
					Else
						strEmployeeCount = 0
					End If
				End If

				'Update the previous record count to the current record before moving to the next recordset
				lngTotalRecordsPrevious = strRecordCount


				Response.Write "<TR><TD><a href=""UploadCDMC_DCCP.asp?BatchNo=" & objRS("FileSeqNum") & """>" & objRS("FileSeqNum") & "</A></B></TD>" & _
							"<TD style=""text-align:center; " & strRecordCountColour & " " & strSameRecordCountColour & """ " & strRecordCountMessage & " " & strSameRecordCountMessage & ">" & strRecordCount & "</TD><TD style=""text-align:center"">" & strAddressCount & "</TD><TD style=""text-align:center; " & strSameRecordCountColour & " " & """ " & strSameRecordCountMessage & ">" & strEmployeeCount & "</TD>" & _
							"<TD style=""text-align:center;"">" & strAction & "</TD>" & _
							"<TD style=""text-align:center; " & strDateLoadColour & """ " & strDateLoadedTitle & ">" & strDateLoaded & "</TD></TR>"
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

Dim objStartFolder
Dim strFileNameDefault
Dim strUpdatedBy
Dim objFSO
Dim colFiles
Dim objFolder
Dim objFile

Dim strFileDateTime

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass
'Dim strFileNameDefault
Dim strExtension

Set objNetwork = CreateObject("WScript.Network")
Set objFSO = CreateObject("Scripting.FileSystemObject")

		
errors = "" 
lineNo = 1    
strUpdatedBy = Session("UserID")

Dim Uploader, File, filePath
Set Uploader = New FileUploader

'Run the file load code if a client side file has been selected (otherwise use the local file location)
If Request.QueryString("Action")="Save" Then

	' This starts the upload process
	Uploader.Upload()
	
	    If Uploader.Files.Count = 0 Then
		    Response.Write "No File(s) uploaded."
	    Else
			'Set the global (page) variable to the Check Delete value for use when processing data
			
		    ' Loop through the uploaded file
		    For Each File In Uploader.Files.Items					  		    		    
			    'set the upload path
	            strUploadPath = Server.MapPath(GetFilePath()) & "\Attachments"
				File.SaveToDisk strUploadPath
			    filePath = Server.MapPath(GetFilePath()) & "\Attachments\" & File.FileName
				strFileName = File.FileName
				strFileDateTime = now()'Mid(strFileName,14,8)
		    Next
		End If
		
End If

If Request.QueryString("Action")="SaveFileLocal" Then

		strDeleteCheck = Request.QueryString("Delete")
		
		'If the user has clicked on the Load G then get the file details from the G Drive settings
		'If Request.QueryString("Action")="SaveFileLocal" Then

			'''---Start the New Service Account Login section
			If Request.QueryString("ActionType")="Service" Then

				'Get the System Parameter for the start of the Training File Location ----Change 21/11/2022 to use the global variable from the file location selected
				strServer = Session("ServerLocationCDMCPath")
				
				'****** AB Move file from service account location to G:Drive ******
				
				
				
				strServer = GetSystemAdmin("GDriveFilePath")
						
				If Session("ServerLocationCDMC") = "CDMC" Then				
					Dim strCDMCFileFrom 
					Dim strCDMCFileTo 
					
					strCDMCFileFrom = GetSystemAdmin("CDMCDriveFilePath") & "\" & GetSystemAdmin("CDMCCardlistFileName") & ".csv"
					strCDMCFileTo = GetSystemAdmin("GDriveFilePath") & "\" & GetSystemAdmin("CDMCCardlistFileName") & ".csv"
					
					'Response.Write "<BR>From" & strCDMCFileFrom
					'Response.Write "<BR>To" & strCDMCFileTo & "<BR>"
					'Response.Write "<BR>Server Location " & Session("ServerLocationCDMC") & "<BR>"
					'Response.Write "<BR>Action Type " & Request.QueryString("ActionType")
					
					If MoveFile(strCDMCFileFrom, strCDMCFileTo) <> "OK" Then				
						
						Response.Write "<div class=""alert alert-danger"" role=""alert""><B>CDMC File is not in folder " & strCDMCFileFrom & ".</B></div>"
	   

						Exit Sub
					
					End If
				End If
				
				'Get the System Parameter for the Service Account UserName and Password
				strUser = GetSystemAdmin("CAPSServiceAccountName")
				strPass = GetSystemAdmin("CAPSServiceAccountPassword")

				'Get the System Parameter for the fileName
				'strFileNameDefault = GetSystemAdmin("CSFromDinersFileName")
							
				objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

				objStartFolder = strServer

			Else
		
				'Set objFSO = CreateObject("Scripting.FileSystemObject")

				'objStartFolder = "D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\CDMC"
				objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\CDMC\"		

			'End of the SaveFileLocal section to load from G drive or Server (local)
			End If
			
			'Set objFolder = objFSO.GetFolder(objStartFolder)
			Set objFolder = objFSO.GetFolder(objStartFolder)
			Set colFiles = objFolder.Files

			'Get the System Parameter for the fileName
			strFileNameDefault = GetSystemAdmin("CDMCCardlistFileName")
			
			If IsNull(strFileNameDefault) or strFileNameDefault = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>NOT LOADED! There is no System Parameter for CDMC File Names (""CDMCCardlistFileName""). See System Admin.</span></div></div></div>"
					Exit Sub
			End If

			For Each objFile in colFiles			

				If Left(objFile.Name,11) = Trim(strFileNameDefault) Then
					'Get the file extension as there can be a number of file types with the same name
					strExtension = objFSO.GetExtensionName(objFile.Name)
					
					'Only get the file name if the file is a csv file
					If LCase(strExtension) = "csv" Then
						strFileName = objFile.Name
						''----New 23/11/2022 added the check to add a final backslash to the starting folder if there isn't one
						If Right(objStartFolder,1) = "\" Then
							filePath = objStartFolder & "" & strFileName
						Else
							filePath = objStartFolder & "\" & strFileName
						End If
					End If
				End If
				
			Next

			If IsNull(strFileName) or strFileName = "" Then

				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>" & strFileNameDefault & " NOT LOADED! There is no CDMC File in the Server Folder to Load! Copy the 'CMS_AllData.csv' file to the location " & objStartFolder & "</span></div></div></div>"
					Exit Sub
			End If

'Response.Write "File=" & filePath
'exit sub
		'strDeleteCheck = "on"

'---Start Commented out to test the Excel driver upload
	'		Call ReadText(filePath,strFileName)
			
	'		Exit Sub
''End If
'---End Comment out

'---Start Excel driver upload

			'Check to see if the same FileSeqNum for the same FileType has already been loaded
			'lngFileID = GetFileLoadID("CDMC","",strFileName)
			lngFileID = GetSystemAdmin("CDMC")
			
			'Response.Write lngFileID
			
			If lngFileID = "" Then
				'Get the next fileID Number for the ANZCardlist File
				lngFileID = GetSystemAdmin("CDMC")
			'Else
				'If the checkbox to overwrite is checked then load the data, otherwise do not load
				'If strDeleteCheck = "on" Then
				'If lngFileID = 1 or strDeleteCheck = "on" Then
				
					'Delete any existing CS From Diners Records
					objCon.Execute "TRUNCATE TABLE tblCAPSCDMCLoad"
					'Response.Write "TRUNCATE TABLE tblCAPSCDMCLoad"
				'Else
					'Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
					'	"<span aria-hidden=""true"">&times;</span></button>" & _
					'	"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
					'	"<span>NOT LOADED! The CDMC File """ & strFileName & """ has already been loaded! <a href=""UploadCDMC_DCCP.asp?Reload=1&FileSeqNum=" & lngFileID & """ Style=""font-weight:bold; color:white;""> Check the 'Overwrite Existing Batch' box and load again to overwrite...</a></span></div></div></div>"
					'Exit Sub
				'End If
			End If
			
			'Call the relevant procedure depending on whether the file is .xls or .txt
			If Right(filePath,3) = "csv" or Right(filePath,3) = "txt" then
			
			'New Error check to make sure the file is found
			On Error Resume Next

				'After uploading, Read excel file
				Set objExcelCon = Server.CreateObject("ADODB.connection")     
				
		'		objExcelCon.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & objStartFolder & ";Extensions=asc,csv,tab,txt;ColNameHeader=Yes;"
				
				'objExcelCon.Open "DBQ=" & filePath & "; DRIVER={Microsoft Excel Driver (*.xls)};" 
				'objExcelCon.Open "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & filePath & ";DefaultDir=c:\Apps\CAPS2\ASP2\Admin\CAPSAdmin\Attachments;" 
		'		objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strUploadPath & "\" & ";Extended Properties=""text;HDR=Yes;FMT=Delimited"";"
				'objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""text;HDR=Yes;FMT=Delimited(~)"";"
				
				'-----This line below was working with a server with OLEDB drivers ****
				'objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strUploadPath & "\" & ";Extended Properties=""text;"";"
				
				''--New 23/11/2022 -- changed data source to filePath instead of objStartFolder
				'objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath & ";Extended Properties=""text;"";"
				'Reponse.Write "<BR>Loaded From = " & objStartFolder & "<BR>"
				objExcelCon.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & objStartFolder & ";Extended Properties=""text;"";"
				
			''New Error Capture for file not found --added 23/11/2022
			If Err.Number <> 0 Then
  				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>ERROR!!! CDMC File """ & filePath & """ Not Loaded! File not Found. objExcelCon.Open. See System Admin " & Err.Number & " : " & Err.Description & "</span></div></div></div>"
  				Err.Clear
				Exit Sub
			End If

			On Error GoTo 0
				
				'objExcelCon.Open strUploadPath & "\CAPSCSFromDiners.dsn"
				'response.write objStartFolder
				'response.write strFileName
				'Write the SQL Query 
				objRS.open "SELECT * FROM [" & strFileName & "]", objExcelCon  
		    
				ReadExcel		    
				
				'Check for errors
				If(errors<>"") Then
					'Print the errors and return
					Response.Write "<font face=arial size=1><b>File not uploaded. Please correct the following errors and load again</b> <br> "& errors &"</font><br>"         
				Else  
					'If the no errors found in the ReadExcel method, then start uploading the records in the database 
					objCon.Execute "TRUNCATE TABLE tblCAPSCDMCLoad"

					UploadExcel strDeleteCheck,lngFileID, strFileName,filePath,strFileDateTime
					'UploadExcel Uploader.Form("chkDelete"),lngFileID, strFileName,filePath,strFileDateTime
					'Response.Write "<b>ANZ Cardlist File Sucessfully Uploaded!!<b><br>"
				End If
				
				'Close the recordset/connection 
				objRS.Close 
				objExcelCon.Close 
		    
				'Move the file to the Loaded folder
				Set objFSO = CreateObject("Scripting.FileSystemObject")
				'Set outPut = objFSO.CreateTextFile("c:\\output.txt", true);
				'Set objTextFile = objFSO.OpenTextFile (strFileNamePath, ForReading)

				objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\CDMC\"		
				
				'Below line (File Name change) moved to within the If for Action Type, so as to reflect correct name
				'Change the file name to add the date so it does not overwrite existing files in the same folder
				'strFileName = Left(strFileName,Len(strFileName)-4) & Year(now()) & PadDigits(Month(now()),2) & Day(now()) & Hour(now()) & Minute(now()) & Right(strFileName,4)
				
				'''---Start the New Service Account Login section  ---MOVING THE FILE AFTER LOADING
				If Request.QueryString("ActionType")="Service" Then
					'Do not move the file if it was loaded from the G Drive
				Else
					'Change the file name to add the date so it does not overwrite existing files in the same folder
					strFileName = Left(strFileName,Len(strFileName)-4) & Year(now()) & PadDigits(Month(now()),2) & Day(now()) & Hour(now()) & Minute(now()) & Right(strFileName,4)

					'Set objTextFile = Nothing
					objFSO.MoveFile filePath,objStartFolder & "Loaded\" & strFileName 
				End If
				
				'objCon.Execute "spLoadCDMCHistoryLog"
				
				With objCmd1

					.CommandType = 4
					.CommandText = "spLoadCDMCHistoryLog"
					.CommandTimeout = 300

					.ActiveConnection = objCon
		 
				End With	
   
				objCmd1.Execute       
		
					Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>CDMC File """ & strFileName & """ load COMPLETE!</span></div></div></div>"

		
			Else
				'ANZ file is xls as provided by ANZ and  not a text file, so do not enable text load
				'ReadText(filePath)
			End If
			
	    'End If
	'End of Start load for Client side loading of files (which has been merged with server file load procedure at the top of this process)
	'End IF

	'''---Start the New Service Account Login section
	If Request.QueryString("ActionType")="Service" Then
		
		objNetwork.RemoveNetworkDrive strServer, True, False
		Set objNetwork = Nothing
	End If
	
	Set objFSO = Nothing
	
'---End New Excel Section	
End if

End Sub


Sub ReadText(strFileNamePath,strFileName)

Const ForReading = 1
Dim strLine
Dim strCardType
Dim strRow
Dim x, y

Dim strFooterCount
Dim lngFileLoadID
Dim arrValues

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
Dim objFSO
Dim objTextFile
Dim lngFileID
Dim strFileDateTime
Dim strFileSeqNum
Dim strDateLoaded
Dim objStartFolder
Dim strFileType
dim objFile

Dim arrValues2
Dim intOne
Dim intPreviousVals
Dim intArrayCount
Dim intArrTotal

Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set outPut = objFSO.CreateTextFile("c:\\output.txt", true);
Set objTextFile = objFSO.OpenTextFile (strFileNamePath, ForReading)
'Set objTextFile = objFSO.OpenTextFile ("c:\mytextfile.txt", ForReading)

x = 0 
'Set the variable which normally gets the checkbox value to overwrite, to "on" so previous loads are always deleted before loading the current file
'strDeleteCheck = "on"
lngFileID = 0

	Do Until objTextFile.AtEndOfStream
		
		'Count the rows for use in line counts, summary and for getting header
		x = x + 1
		
		'If this is the first row being loaded the delete any existing CDMC values in the table, as this is a temporary table
		If x = 1 Then

			'The fileDate and Number are only in the header row
			strFileDateTime = Mid(strLine,3,8)
			'strFileSeqNum = Mid(strLine,11,12)
			
			strFileSeqNum = GetLastFileLoadID("CDMC",strFileName)
			'lngFileID = GetFileLoadID("CDMC",strFileSeqNum,strFileName)
			lngFileID = GetSystemAdmin("CDMC")
			
		
			
			'Check to see if the same FileSeqNum for the same FileType has already been loaded
			'lngFileID = GetFileLoadID("CDMC","",strFileName)
			'lngFileID = strFileSeqNum
			
			'If the checkbox to overwrite is checked then load the data, otherwise do not load
			'If strDeleteCheck = "on" Then
			
			
			
			If lngFileID = 0 or lngFileID = 1 or strDeleteCheck = "on" Then
				'Delete any existing CS From Diners Records
			    strFileSeqNum = strFileSeqNum + 1
				objCon.Execute "DELETE FROM tblCAPSCDMC"
			Else
				Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-danger alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
					"<span aria-hidden=""true"">&times;</span></button>" & _
					"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
					"<span>NOT LOADED! The CDMC File Seq Num """ & strFileSeqNum & """ has already been loaded!</span></div></div></div>"
				Exit Sub
			End If
						
		End If
		
		'Get the next line (to the end of the row)			
		strLine = objTextFile.Readline
		'Replace [LF] with space
		'for y =1 to 50
		
			'response.write mid(strLine,y+1,1) & " - " & mid(strLine,y+1,1) & "</br>"
			'response.write mid(strLine,y+1,1) & " - " & asc(mid(strLine,y+1,1)) & "</br>"
		'next 
		'strLine = Replace(strLine,"LF"," ")
		'Replace [CR]space with [CR][LF]
		'strLine = Replace(strLine,"[CR]","[CR][LF]")
		
		'Response.Write strLine
		
		'Split the row into fields by the (~) delimiter
		arrValues = Split(strLine,"~")	
		
		'response.write "ubound= " & UBound(arrValues)
		
	
		'Response.Write x & "<BR>"
		For y = 0 to 43
			'If y > 40 Then
			'response.write mid(strLine,y+1,1) & " - " & asc(mid(strLine,y+1,1)) & "</br>"
			'end if
		'response.write "</br>" & y & "=" & arrValues(y)
			'Remove the text qualifier (") from the start and end of each field
			If Len(arrValues(y)) > 2 Then
				arrValues(y) = Left(arrValues(y),Len(arrValues(y))-1)
				arrValues(y) = Right(arrValues(y),Len(arrValues(y))-1)
			Else
				arrValues(y) = ""
			End If
			'response.write "</br>" & arrValues(y)			
		Next 
		
		lngCDMCID = 0
		strGroupName = arrValues(0)
		strDivisionName = arrValues(1)
		strBranchName = arrValues(2)
		strDepartmentName = arrValues(3)
		strDepartmentNumber = arrValues(4)
		strCostCentre = arrValues(5)
		strEmployeeID = arrValues(6)
		strEmployeeType = arrValues(7)
		strFirstname = arrValues(8)
		strSurname = arrValues(9)
		strTitle = arrValues(10) 
		strEmail_Address = arrValues(11)
		strTelephoneNumber = arrValues(12)
		strMobileNumber = arrValues(13)
		strDateofBirth = arrValues(14)
		strGender = arrValues(15)
		strActualRankLvl = arrValues(16)
		strSite = arrValues(17)
		strUnit = arrValues(18)
		strReportsTo = arrValues(19)
		strDCD_PostalAddress = arrValues(20)	
		straddressline1 = arrValues(21)
		straddressline2 = arrValues(22)
		straddressline3 = arrValues(23)
		straddressline4 = arrValues(24)
		straddressline5 = arrValues(25)
		straddressline6 = arrValues(26)
		strPostalAddress_Unit = arrValues(27)
		strPostalAddress_ClientLocation = arrValues(28)
		strPostalAddress_DeliveryLocation = arrValues(29)
		strPostalAddress_City = arrValues(30)
		strPostalAddress_State = arrValues(31)
		strPostalAddress_PostCode = arrValues(32)
		strPostalAddress_Country = arrValues(33)
		strDCDProtectedIdentity = arrValues(40)'arrValues(34)
		strIsValidPostal = ""'""'objRS("IsValidPostal") 
		strOutAddr1 = ""'objRS("OutAddr1") 
		strOutAddr2 = ""'objRS("OutAddr2") 
		strOutAddr3 = ""'objRS("OutAddr3") 
		strOutSuburb = ""'objRS("OutSuburb") 
		strOutState = ""'objRS("OutState") 
		strOutPostCode = ""'objRS("OutPostCode") 
		strPostalMessage = ""'objRS("PostalMessage") 
		strhasChanged = "N"'objRS("hasChanged") 
		strDCD_WorkAddress = arrValues(34)'arrValues(35)
		strClientLocation = arrValues(35)'arrValues(36)
		strStreetAddress = arrValues(36)'arrValues(37)
		strCity = arrValues(37)'arrValues(38)
		strState = arrValues(38)'arrValues(39)
		strPostCode = arrValues(39)'arrValues(40)
		strFormalFirstName = arrValues(41)
		strFormalLastName = arrValues(42)
		strFormalMiddleName = arrValues(43)
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
		
			'Call the procedure to save the record to the Database

	    			If x > 1 Then
						SaveRecord lngCDMCID,strGroupName,strDivisionName,strBranchName,strDepartmentName,strDepartmentNumber,strCostCentre,strEmployeeID,strEmployeeType, _
						strFirstname,strSurname,strTitle,strEmail_Address,strTelephoneNumber,strMobileNumber,strDateofBirth,strGender,strActualRankLvl,strSite,strUnit,strReportsTo,strDCD_PostalAddress, _
						straddressline1,straddressline2,straddressline3,straddressline4,straddressline5,straddressline6,strPostalAddress_Unit,strPostalAddress_ClientLocation, _
						strPostalAddress_DeliveryLocation,strPostalAddress_City,strPostalAddress_State,strPostalAddress_PostCode,strPostalAddress_Country,strDCDProtectedIdentity, _
						strIsValidPostal,strOutAddr1,strOutAddr2,strOutAddr3,strOutSuburb,strOutState,strOutPostCode,strPostalMessage,strhasChanged,strDCD_WorkAddress,strClientLocation, _
						strStreetAddress,strCity,strState,strPostCode,strFormalFirstName,strFormalLastName,strFormalMiddleName,strOutTitle,strOutDinersWorkPhone,strOutDinersMobilePhone, _
						strOutANZPhone,strOutDinersAddress1,strOutDinersAddress2,strRemoveCountdown,strFirstUpdated,strLastUpdated,strActive,strUpdatedBy,strDateUpdated,strFileID,strLoaded, x
            		End If

		Loop	

		
	If x > 1 Then
		
		'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
		lngFileLoadID = SaveFileLoadID ("CDMC",strFileName,strFileNamePath,x,0,0,0,0,0,0,0,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
		'response.write "SaveFileLoadID (CDMC," & strFileName & "," & strFileNamePath & "," & x & "," & strFileDateTime & "," & strFileSeqNum
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
		Call UpdateFileLoadSummary ("CDMC",strFileSeqNum, strFileName,lngFileLoadID)
		'response.write "UpdateFileLoadSummary (""CSFRomDiners""," & strFileSeqNum & "," & lngFileLoadID & ")"
		
		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
						"<span aria-hidden=""true"">&times;</span></button>" & _
						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
						"<span>CDMC File Seq Num """ & strFileSeqNum & """ load COMPLETE!</span></div></div></div>"

		objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\CDMC\"		
		
		Set objTextFile = Nothing
		
		objFSO.MoveFile strFileNamePath,objStartFolder & "Loaded\" & strFileName 
		
	End If

Set objFSO = Nothing

End Sub

'Function for validating the excel file values
Sub ReadExcel 

Dim errors
Dim strCardNumber
Dim strEID
Dim lineNo

    'First loop to check all the values
    Do until objRS.EOF 
        'Read each record present in the Excel file and check for the validation
 
        'strCardNumber = objRS("Card Number") 'Should be integer and Not Null values       
        'If IsNull(strFileSeqNum) Then
        '    errors = errors & "Error in line no. " & lineNo & ": Card Number should NOT be an Empty/Null value <br>" 
        'End if

		'response.write "fields=" & objRS(0) & "," '& objRS(1) & "," & objRS(2) & "," & objRS(3) & "," & objRS(4) & "," & objRS(5) & "," & objRS(6) & ","
		
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
Dim strDateLoaded

Dim strFileSeqNum
Dim lngFileLoadID

Dim objStartFolder
Dim objFSO

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
			
			'response.write "</br>" & lngCDMCID & "," & strGroupName & "," & strDivisionName & "," & strBranchName & "," & strEmployeeID & "," & strEmployeeType
			
			'response.write "</br>" & lngCDMCID & "," & strEmployeeID & "," & strDateofBirth
			
            SaveRecord lngCDMCID,strGroupName,strDivisionName,strBranchName,strDepartmentName,strDepartmentNumber,strCostCentre,strEmployeeID,strEmployeeType, _
					strFirstname,strSurname,strTitle,strEmail_Address,strTelephoneNumber,strMobileNumber,strDateofBirth,strGender,strActualRankLvl,strSite,strUnit,strReportsTo,strDCD_PostalAddress, _
					straddressline1,straddressline2,straddressline3,straddressline4,straddressline5,straddressline6,strPostalAddress_Unit,strPostalAddress_ClientLocation, _
					strPostalAddress_DeliveryLocation,strPostalAddress_City,strPostalAddress_State,strPostalAddress_PostCode,strPostalAddress_Country,strDCDProtectedIdentity, _
					strIsValidPostal,strOutAddr1,strOutAddr2,strOutAddr3,strOutSuburb,strOutState,strOutPostCode,strPostalMessage,strhasChanged,strDCD_WorkAddress,strClientLocation, _
					strStreetAddress,strCity,strState,strPostCode,strFormalFirstName,strFormalLastName,strFormalMiddleName,strOutTitle,strOutDinersWorkPhone,strOutDinersMobilePhone, _
					strOutANZPhone,strOutDinersAddress1,strOutDinersAddress2,strRemoveCountdown,strFirstUpdated,strLastUpdated,strActive,strUpdatedBy,strDateUpdated,strFileID,strLoaded, x
            
        
        'End If
	'If x > 1000 then exit sub
		objRS.movenext 
		
		If IsNull(strEmployeeID) Then               
		   
		   Exit Sub
		   
		End If
        
    Loop 
    
	If x > 1 Then
		
		'The fileDate and Number are only in the header row
		strFileDateTime = MediumDate(Now())
		'strFileSeqNum = Mid(strLine,11,12)
		
		'strFileSeqNum = GetLastFileLoadID("CDMC",strFileName)
		strFileSeqNum = GetSystemAdmin("CDMC")
		'The FileID is passed in
		'lngFileID = GetFileLoadID("CDMC",strFileSeqNum,strFileName)
		
		'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
'		lngFileLoadID = SaveFileLoadID ("CDMC",strFileName, strFilePath,x,0,0,0,0,0,0,0,strFileDateTime,lngFileID,"Imported",Session("UserID"),"N")
		
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
'		Call UpdateFileLoadSummary ("CDMC",lngFileID, strFileName, lngFileLoadID)
		'response.write "UpdateFileLoadSummary (""CDMCFile""," & lngFileID & "," & strFileName & "," & lngFileLoadID & ")"
		
'		Response.write "<div class=""content-body""><section id=""basic-alerts""><div class=""alert alert-success alert-dismissible mb-2"" role=""alert""><button type=""button"" class=""close"" data-dismiss=""alert"" aria-label=""Close"">" & _
'						"<span aria-hidden=""true"">&times;</span></button>" & _
'						"<div class=""d-flex align-items-center""><i class=""bx bx-error""></i>" & _
'						"<span>CDMC File """ & strFileName & """ load COMPLETE!</span></div></div></div>"
		
		'If the file has records in it then call the save File Load Procedure to save summary details (CAPSFunctions.asp)
		lngFileLoadID = SaveFileLoadID ("CDMC",strFileName,strFilePath,x,0,0,0,0,0,0,0,strFileDateTime,strFileSeqNum,"Imported",Session("UserID"),"N")
		'response.write "SaveFileLoadID (CDMC," & strFileName & "," & strFileNamePath & "," & x & "," & strFileDateTime & "," & strFileSeqNum
		'Call the procedure to update summary information for the file just loaded (CAPSFunctions.asp)
			
		Call UpdateFileLoadSummary ("CDMC",strFileSeqNum, strFileName,lngFileLoadID)
		'response.write "UpdateFileLoadSummary (""CSFRomDiners""," & strFileSeqNum & "," & lngFileLoadID & ")"
		
					
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

	'Make sure there are no long field names (where some data exists with long data)
	If IsNull(strFormalMiddleName) or strFormalMiddleName = "" Then
		strFormalMiddleName = ""
	Else
		If Len(strFormalMiddleName) > 50 Then strFormalMiddleName = Left(strFormalMiddleName,50)
	End If
	
  	With objCmd
  	
  	    'If the procedure has already run then don't create the parameter objects again (more than once)
  	    If x = 1 then
			.CommandType = 4
			.CommandText = "spCAPSCDMCLoad"
			
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

		'response.write "lngCDMCID=" & lngCDMCID
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
			.Parameters("TelephoneNumber") = Left(strTelephoneNumber,30)
			.Parameters("MobileNumber") = Left(strMobileNumber,30)
			'Response.Write strDateofBirth
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
			.Parameters("ClientLocation") = Left(strClientLocation,100)
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
                    			     				     		     		
    
End Sub

Sub SaveRecordCon(lngCDMCID,strGroupName,strDivisionName,strBranchName,strDepartmentName,strDepartmentNumber,strCostCentre,strEmployeeID,strEmployeeType, _
			strFirstname,strSurname,strTitle,strEmail_Address,strTelephoneNumber,strMobileNumber,strDateofBirth,strGender,strActualRankLvl,strSite,strUnit,strReportsTo,strDCD_PostalAddress, _
			straddressline1,straddressline2,straddressline3,straddressline4,straddressline5,straddressline6,strPostalAddress_Unit,strPostalAddress_ClientLocation, _
			strPostalAddress_DeliveryLocation,strPostalAddress_City,strPostalAddress_State,strPostalAddress_PostCode,strPostalAddress_Country,strDCDProtectedIdentity, _
			strIsValidPostal,strOutAddr1,strOutAddr2,strOutAddr3,strOutSuburb,strOutState,strOutPostCode,strPostalMessage,strhasChanged,strDCD_WorkAddress,strClientLocation, _
			strStreetAddress,strCity,strState,strPostCode,strFormalFirstName,strFormalLastName,strFormalMiddleName,strOutTitle,strOutDinersWorkPhone,strOutDinersMobilePhone, _
			strOutANZPhone,strOutDinersAddress1,strOutDinersAddress2,strRemoveCountdown,strFirstUpdated,strLastUpdated,strActive,strUpdatedBy,strDateUpdated,strFileID,strLoaded, x)

Dim intRecord

	'Make sure there are no long field names (where some data exists with long data)
	If IsNull(strFormalMiddleName) or strFormalMiddleName = "" Then
		strFormalMiddleName = ""
	Else
		If Len(strFormalMiddleName) > 50 Then strFormalMiddleName = Left(strFormalMiddleName,50)
		
	End If
	
		'If Not IsNull(strDCD_PostalAddress) Then strDCD_PostalAddress = Replace(strDCD_PostalAddress,"'","''")  End If
		If Not IsNull(strGroupName) Then strGroupName = Replace(strGroupName,"'","''")  End If
		If Not IsNull(strDivisionName) Then strDivisionName = Replace(strDivisionName,"'","''")  End If
		If Not IsNull(strBranchName) Then strBranchName = Replace(strBranchName,"'","''")  End If
		If Not IsNull(strDepartmentName) Then strDepartmentName = Replace(strDepartmentName,"'","''")  End If
		If Not IsNull(strDepartmentNumber) Then strDepartmentNumber = Replace(strDepartmentNumber,"'","''")  End If
		If Not IsNull(strCostCentre) Then strCostCentre = Replace(strCostCentre,"'","''")  End If
		If Not IsNull(strEmployeeID) Then strEmployeeID = Replace(strEmployeeID,"'","''")  End If
		If Not IsNull(strEmployeeType) Then strEmployeeType = Replace(strEmployeeType,"'","''")  End If
		If Not IsNull(strFirstname) Then strFirstname = Replace(strFirstname,"'","''")  End If
		If Not IsNull(strSurname) Then strSurname = Replace(strSurname,"'","''")  End If
		If Not IsNull(strTitle) Then strTitle = Replace(strTitle,"'","''")  End If
		If Not IsNull(strEmail_Address) Then strEmail_Address = Replace(strEmail_Address,"'","''")  End If
		If Not IsNull(strTelephoneNumber) Then strTelephoneNumber = Replace(strTelephoneNumber,"'","''")  End If
		If Not IsNull(strMobileNumber) Then strMobileNumber = Replace(strMobileNumber,"'","''")  End If
		If Not IsNull(strDateofBirth) Then strDateofBirth = Replace(strDateofBirth,"'","''")  End If
		If Not IsNull(strGender) Then strGender = Replace(strGender,"'","''")  End If
		If Not IsNull(strActualRankLvl) Then strActualRankLvl = Replace(strActualRankLvl,"'","''")  End If
		If Not IsNull(strSite) Then strSite = Replace(strSite,"'","''")  End If
		If Not IsNull(strUnit) Then strUnit = Replace(strUnit,"'","''")  End If
		If Not IsNull(strReportsTo) Then strReportsTo = Replace(strReportsTo,"'","''")  End If
		If Not IsNull(strDCD_PostalAddress) Then strDCD_PostalAddress = Replace(strDCD_PostalAddress,"'","''")  End If
		If Not IsNull(straddressline1) Then straddressline1 = Replace(straddressline1,"'","''")  End If
		If Not IsNull(straddressline2) Then straddressline2 = Replace(straddressline2,"'","''")  End If
		If Not IsNull(straddressline3) Then straddressline3 = Replace(straddressline3,"'","''")  End If
		If Not IsNull(straddressline4) Then straddressline4 = Replace(straddressline4,"'","''")  End If
		If Not IsNull(straddressline5) Then straddressline5 = Replace(straddressline5,"'","''")  End If
		If Not IsNull(straddressline6) Then straddressline6 = Replace(straddressline6,"'","''")  End If
		If Not IsNull(strPostalAddress_Unit) Then strPostalAddress_Unit = Replace(strPostalAddress_Unit,"'","''")  End If
		If Not IsNull(strPostalAddress_ClientLocation) Then strPostalAddress_ClientLocation = Replace(strPostalAddress_ClientLocation,"'","''")  End If
		If Not IsNull(strPostalAddress_DeliveryLocation) Then strPostalAddress_DeliveryLocation = Replace(strPostalAddress_DeliveryLocation,"'","''")  End If
		If Not IsNull(strPostalAddress_City) Then strPostalAddress_City = Replace(strPostalAddress_City,"'","''")  End If
		If Not IsNull(strPostalAddress_State) Then strPostalAddress_State = Replace(strPostalAddress_State,"'","''")  End If
		If Not IsNull(strPostalAddress_PostCode) Then strPostalAddress_PostCode = Replace(strPostalAddress_PostCode,"'","''")  End If
		If Not IsNull(strPostalAddress_Country) Then strPostalAddress_Country = Replace(strPostalAddress_Country,"'","''")  End If
		If Not IsNull(strClientLocation) Then strClientLocation = Replace(strClientLocation,"'","''")  End If
		If Not IsNull(strStreetAddress) Then strStreetAddress = Replace(strStreetAddress,"'","''")  End If
		If Not IsNull(strCity) Then strCity = Replace(strCity,"'","''")  End If
		If Not IsNull(strState) Then strState = Replace(strState,"'","''")  End If
		If Not IsNull(strPostCode) Then strPostCode = Replace(strPostCode,"'","''")  End If
		If Not IsNull(strDCDProtectedIdentity) Then strDCDProtectedIdentity = Replace(strDCDProtectedIdentity,"'","''")  End If
		If Not IsNull(strIsValidPostal) Then strIsValidPostal = Replace(strIsValidPostal,"'","''")  End If
		If Not IsNull(strOutAddr1) Then strOutAddr1 = Replace(strOutAddr1,"'","''")  End If
		If Not IsNull(strOutAddr2) Then strOutAddr2 = Replace(strOutAddr2,"'","''")  End If
		If Not IsNull(strOutAddr3) Then strOutAddr3 = Replace(strOutAddr3,"'","''")  End If
		If Not IsNull(strOutSuburb) Then strOutSuburb = Replace(strOutSuburb,"'","''")  End If
		If Not IsNull(strOutState) Then strOutState = Replace(strOutState,"'","''")  End If
		If Not IsNull(strOutPostCode) Then strOutPostCode = Replace(strOutPostCode,"'","''")  End If
		If Not IsNull(strPostalMessage) Then strPostalMessage = Replace(strPostalMessage,"'","''")  End If
		If Not IsNull(strhasChanged) Then strhasChanged = Replace(strhasChanged,"'","''")  End If
		If Not IsNull(strFormalFirstName) Then strFormalFirstName = Replace(strFormalFirstName,"'","''")  End If
		If Not IsNull(strFormalLastName) Then strFormalLastName = Replace(strFormalLastName,"'","''")  End If
		If Not IsNull(strFormalMiddleName) Then strFormalMiddleName = Replace(strFormalMiddleName,"'","''")  End If
		If Not IsNull(strOutTitle) Then strOutTitle = Replace(strOutTitle,"'","''")  End If
		If Not IsNull(strOutDinersWorkPhone) Then strOutDinersWorkPhone = Replace(strOutDinersWorkPhone,"'","''")  End If
		If Not IsNull(strOutDinersMobilePhone) Then strOutDinersMobilePhone = Replace(strOutDinersMobilePhone,"'","''")  End If
		If Not IsNull(strOutANZPhone) Then strOutANZPhone = Replace(strOutANZPhone,"'","''")  End If
		If Not IsNull(strOutDinersAddress1) Then strOutDinersAddress1 = Replace(strOutDinersAddress1,"'","''")  End If
		If Not IsNull(strOutDinersAddress2) Then strOutDinersAddress2 = Replace(strOutDinersAddress2,"'","''")  End If	
	
		'Response.Write "EXEC spCAPSCDMCLoad " & lngCDMCID & ",'" & strGroupName & "','" & strDivisionName & "','" & strBranchName & "','" & strDepartmentName & "','" & strDepartmentNumber & "','" & strCostCentre & "','" & strEmployeeID & "','" & strEmployeeType & "','" & strFirstname & "','" & strSurname & "','" & strTitle & "','" & strEmail_Address & "','" & Left(strTelephoneNumber,30) & "','" & Left(strMobileNumber,30) & "','" & strDateofBirth & "','" & strGender & "','" & strActualRankLvl & "','" & strSite & "','" & strUnit & "','" & strReportsTo & "','" & strDCD_PostalAddress & "','" & straddressline1 & "','" & straddressline2 & "','" & straddressline3 & "','" & straddressline4 & "','" & straddressline5 & "','" & straddressline6 & "','" & strPostalAddress_Unit & "','" & strPostalAddress_ClientLocation & "','" & strPostalAddress_DeliveryLocation & "','" & strPostalAddress_City & "','" & strPostalAddress_State & "','" & strPostalAddress_PostCode & "','" & strPostalAddress_Country & "','" & strClientLocation & "','" & strStreetAddress & "','" & strCity & "','" & strState & "','" & strPostCode & "','" & strDCDProtectedIdentity & "','" & strIsValidPostal & "','" & strOutAddr1 & "','" & strOutAddr2 & "','" & strOutAddr3 & "','" & strOutSuburb & "','" & strOutState & "','" & strOutPostCode & "','" & strPostalMessage & "','" & strhasChanged & "','" & strFormalFirstName & "','" & strFormalLastName & "','" & strFormalMiddleName & "','" & strOutTitle & "','" & strOutDinersWorkPhone & "','" & strOutDinersMobilePhone & "','" & strOutANZPhone & "','" & strOutDinersAddress1 & "','" & strOutDinersAddress2 & "'," & strRemoveCountdown & ",'" & NULL & "','" & NULL  & "','" & strActive & "'," & Session("UserID") & "," & strFileID & ",'N',0" & "<BR>"	   
	
		objCon.Execute "EXEC spCAPSCDMCLoad " & lngCDMCID & ",'" & strGroupName & "','" & strDivisionName & "','" & strBranchName & "','" & strDepartmentName & "','" & strDepartmentNumber & "','" & strCostCentre & "','" & strEmployeeID & "','" & strEmployeeType & "','" & strFirstname & "','" & strSurname & "','" & strTitle & "','" & strEmail_Address & "','" & Left(strTelephoneNumber,30) & "','" & Left(strMobileNumber,30) & "','" & strDateofBirth & "','" & strGender & "','" & strActualRankLvl & "','" & strSite & "','" & strUnit & "','" & strReportsTo & "','" & strDCD_PostalAddress & "','" & straddressline1 & "','" & straddressline2 & "','" & straddressline3 & "','" & straddressline4 & "','" & straddressline5 & "','" & straddressline6 & "','" & strPostalAddress_Unit & "','" & strPostalAddress_ClientLocation & "','" & strPostalAddress_DeliveryLocation & "','" & strPostalAddress_City & "','" & strPostalAddress_State & "','" & strPostalAddress_PostCode & "','" & strPostalAddress_Country & "','" & strClientLocation & "','" & strStreetAddress & "','" & strCity & "','" & strState & "','" & strPostCode & "','" & strDCDProtectedIdentity & "','" & strIsValidPostal & "','" & strOutAddr1 & "','" & strOutAddr2 & "','" & strOutAddr3 & "','" & strOutSuburb & "','" & strOutState & "','" & strOutPostCode & "','" & strPostalMessage & "','" & strhasChanged & "','" & strFormalFirstName & "','" & strFormalLastName & "','" & strFormalMiddleName & "','" & strOutTitle & "','" & strOutDinersWorkPhone & "','" & strOutDinersMobilePhone & "','" & strOutANZPhone & "','" & strOutDinersAddress1 & "','" & strOutDinersAddress2 & "'," & strRemoveCountdown & ",'" & NULL & "','" & NULL  & "','" & strActive & "'," & Session("UserID") & "," & strFileID & ",'N',0"	   
	
	
			
		
                    			     				     		     		
    
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
	objRS.Open "SELECT TOP 1 [FileSeqNum] FROM tblCAPSFileLoad WHERE FileType = 'CDMC' AND [Deleted] = 'N' ORDER BY [FileSeqNum] DESC ",objCon

		If Not objRS.EOF Then
			GetNextCDMCFileID = objRS("FileSeqNum") + 1
		Else
			GetNextCDMCFileID = 1
		End If

	objRS.Close
	
End Function

Set objRS = Nothing
Set objCon = Nothing

Public Sub DisplayFileSummary()

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile
Dim strFileSize
Dim strFileAttributes
Dim strServerSelect
Dim strServerSelectStyle
Dim strTitle
Dim strFa

Set objFSO = CreateObject("Scripting.FileSystemObject")

If Session("ServerLocationCDMC") = "LocalDrive" Then
	strServerSelect = " Style=""background-color:#d6f5d6;"" "  '#F5F5F6;
	strServerSelectStyle = " background-color:#d6f5d6;"" "
	strTitle = "Title=""" & Session("ServerLocationCDMC") & " currently SELECTED to be loaded!"""
	strFa = "<i class=""fa fa-check""></i>"
	strLoadType = "LocalDrive"
Else
	strTitle = "Title=""Click to select for loading"""
End If

	'objStartFolder = objFSO.GetAbsolutePathName("D:\Inetpub\WWWRoot\CAPS\ASPNew\Admin\CAPSAdmin\Attachments\CDMC\")
	objStartFolder = GetSystemAdmin("ServerFilePath") & "\Admin\CAPSAdmin\Attachments\CDMC\"

	'Set the Session path variable to this local path if it has been selected for loading
	If Session("ServerLocationCDMC") = "LocalDrive" Then
		'Set the global variable to the server in this procedure
		Session("ServerLocationCDMCPath") = objStartFolder
	End If

	Set objFolder = objFSO.GetFolder(objStartFolder)
		
	Set colFiles = objFolder.Files
	
	'Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
	Response.Write "<tr " & strTitle & "><th style=""text-align:left; " & strServerSelectStyle & """><a href=""UploadCDMC_DCCP.asp?ServerLocationCDMC=LocalDrive"">" & strFa & " CAPS Server (local)</a><i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server (Local) waiting to be loaded. Click 'Load Local' button to Load CDMC File. (" & objStartFolder & ")""></i></th></tr>"
	'Response.Write "<tr><th style=""text-align:left"">Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server (Local) waiting to be loaded. Click 'Load Local' button to Load CDMC File.""></i></th></tr>"

	intCount = 0
	
	For Each objFile in colFiles		
		
		If Right(objFile.Name,3) <> "ini" Then
		
		intCount = intCount + 1
		
			If intCount < 6 Then
				If IsNull(objFile.Name) or objFile.Name = "" Then
					strFile = ""
					strFileSize = 0
				Else
					strFile = Left(objFile.Name,10) & ".." & Right(objFile.Name,4)
					strFileSize = Round(objFile.Size/1024000,2)
					
					strFileAttributes =  "Created: " & objFile.DateCreated
					strFileAttributes = strFileAttributes & " Last Accessed: " & objFile.DateLastAccessed
					strFileAttributes = strFileAttributes & " Last Modified: " & objFile.DateLastModified  
	
				End If
				
				Response.Write "<TR><TD Title=""" & objFile.Name & " " & strFileAttributes & """ " & strServerSelect & ">" & strFile & "</TD></TR>"
			End If
			
		End If
		
	Next
	
	 Response.Write "<tr><td style=""text-align:left; font-size:14px;"">Total: " & intCount & " Size: " & strFileSize & " MB</td></tr>"
	 'Response.Write "<tr><th style=""text-align:left"">Size: " & strFileSize & " MB</th></tr>"'</table>"
	 
Set objFSO = Nothing
'Set outPut = Nothing

End Sub

Public Sub DisplayFileSummaryG()

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile
Dim strFileSize
Dim strFileAttributes

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass
Dim strFileNameDefault
Dim strFileExt
Dim strServerSelect
Dim strServerSelectStyle
Dim strTitle
Dim strFa

Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get the System Parameter for the start of the Training File Location
strServer = GetSystemAdmin("GDriveFilePath")

'Get the System Parameter for the Service Account UserName and Password
strUser = GetSystemAdmin("CAPSServiceAccountName")
strPass = GetSystemAdmin("CAPSServiceAccountPassword")

'Get the System Parameter for the fileName
strFileNameDefault = GetSystemAdmin("CDMCCardlistFileName")

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

If Session("ServerLocationCDMC") = "GDrive" Then
	strServerSelect = " Style=""background-color:#d6f5d6;"" "
	strServerSelectStyle = " background-color:#d6f5d6;"" "
	strTitle = "Title=""" & Session("ServerLocationCDMC") & " currently SELECTED to be loaded!"""
	strFa = "<i class=""fa fa-check""></i>"

	'Set the global variable to the server in this procedure
	Session("ServerLocationCDMCPath") = strServer

	strLoadType = "Service"
Else
	strTitle = "Title=""Click to select for loading"""
End If

	objStartFolder = strServer

	Set objFolder = objFSO.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files
	
	'Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
	Response.Write "<tr " & strTitle & "><th style=""text-align:left; " & strServerSelectStyle & """><a href=""UploadCDMC_DCCP.asp?ServerLocationCDMC=GDrive"">" & strFa & " G drive Files </a><i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be loaded. Click 'Load Server' button to Load CDMC File. (" & objStartFolder & ")""></i></th></tr>"
	'Response.Write "<tr><th style=""text-align:left"">G Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CAPS Server waiting to be loaded. Click 'Load Server' button to Load CDMC File.""></i></th></tr>"
	
	intCount = 0
	
	For Each objFile in colFiles		
		
		'Get the files with the starting name the same as the system default/setting name
		If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) Then
		
			strFileExt = objFSO.GetExtensionName(objFile.Name) 
			
			If IsNull(strFileExt) Then
				strFileExt = ""
			Else
				strFileExt = LCase(strFileExt)
			End If
			
			If strFileExt = "csv" Then
			intCount = intCount + 1
			
				If intCount < 6 Then
					If IsNull(objFile.Name) or objFile.Name = "" Then
						strFile = ""
						strFileSize = 0
					Else
						strFile = Left(objFile.Name,10) & ".." & strFileExt'Right(objFile.Name,4)
						strFileSize = Round(objFile.Size/1024000,2)
						
						strFileAttributes =  "Created: " & objFile.DateCreated
						strFileAttributes = strFileAttributes & " Last Accessed: " & objFile.DateLastAccessed
						strFileAttributes = strFileAttributes & " Last Modified: " & objFile.DateLastModified  
		
					End If
					
					Response.Write "<TR><TD Title=""" & objFile.Name & " " & strFileAttributes & """ " & strServerSelect & ">" & strFile & "</TD></TR>"
				End If
			End If
		End If
		
	Next
	
	
	 Response.Write "<tr><td style=""text-align:left"">Total: " & intCount & " Size: " & strFileSize & " MB</td></tr>"
	 'Response.Write "<tr><th style=""text-align:left"">Size: " & strFileSize & " MB</th></tr>"'</table>"

	If strFile <> "" Then
		If strFileSize < 80 Then
			Response.Write "<div class=""alert alert-danger alert-dismissible"" style=""z-index:1; position: absolute; left: -290px; top: -180px;"" role=""alert""><a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a> The CDMC file looks a bit small (" & strFileSize & " MB). Please notify System Administrators before loading.</div>"
		End If
	End If
	If intCount > 1 Then
		Response.Write "<div class=""alert alert-danger alert-dismissible"" style=""z-index:1; position: absolute; left: -280px; top: -170px;"" role=""alert""><a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a> There is more than 1 CDMC file in the Import folder. Please notify System Administrators before loading or remove all CMS_AllData.csv files not for today.</div>"
	End If
 
objNetwork.RemoveNetworkDrive strServer, True, False
	 
Set objFSO = Nothing
Set objNetwork = Nothing
	 
End Sub

Public Sub DisplayFileSummaryCDMC()

Dim objStartFolder
Dim colFiles
Dim strFile
Dim intCount
Dim objFSO
Dim objFolder
Dim objFile
Dim strFileSize
Dim strFileAttributes

Dim objNetwork
Dim strServer
Dim strUser
Dim strPass
Dim strFileNameDefault
Dim strFileExt
Dim strServerSelect
Dim strServerSelectStyle
Dim strTitle
Dim strFa

Set objNetwork = CreateObject("WScript.Network")

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Get the System Parameter for the start of the Training File Location
strServer = GetSystemAdmin("CDMCDriveFilePath")

'Get the System Parameter for the Service Account UserName and Password
strUser = GetSystemAdmin("CAPSServiceAccountName")
strPass = GetSystemAdmin("CAPSServiceAccountPassword")

On Error Resume Next
'Get the System Parameter for the fileName
strFileNameDefault = GetSystemAdmin("CDMCCardlistFileName")
			
objNetwork.MapNetworkDrive "",strServer, False, strUser, strPass

If Err.Number <>0 Then
	'Write an error to the top of the screen
	'Response.Write "<div class=""container"" style=""position:relative; z-index:100; top:80px; left:0px;""><div class=""alert alert-danger"" role=""alert"" style=""position: absolute; top:40px; left:40px; z-index:100;"">Error! CDMC Server path not found: " & strFileNameDefault & "</div></div>"
	'Write a message in the G Drive Div area
	Response.Write "<div class=""alert alert-danger"" role=""alert"" style=""position: absolute; top:0px; left:0px; z-index:100;"">Error! CDMC Drive path not found: " & strServer & "</div>"			

	Err.Clear
	On Error Goto 0
	Exit Sub
	
End If

On Error Goto 0

If Session("ServerLocationCDMC") = "CDMC" Then
	strServerSelect = " Style=""background-color:#d6f5d6;"" "
	strServerSelectStyle = " background-color:#d6f5d6;"" "
	strTitle = "Title=""" & Session("ServerLocationCDMC") & " currently SELECTED to be loaded!"""
	strFa = "<i class=""fa fa-check""></i>"

	'Set the global variable to the server in this procedure
	Session("ServerLocationCDMCPath") = strServer

	strLoadType = "Service"
Else
	strTitle = "Title=""Click to select for loading"""
End If

	objStartFolder = strServer

	Set objFolder = objFSO.GetFolder(objStartFolder)
	Set colFiles = objFolder.Files
	
	'Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
	Response.Write	"<tr " & strTitle & "><th style=""text-align:left; " & strServerSelectStyle & """><a href=""UploadCDMC_DCCP.asp?ServerLocationCDMC=CDMC"">" & strFa & " CDMC Server Files </a><i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CDMC Server waiting to be loaded. When selected (highlighted below), click 'Load Files' button to Load CDMC File. (" & objStartFolder & ")""></i></th></tr>"
	'Response.Write	"<tr><th style=""text-align:left"">CDMC Files To Be Loaded <i class=""fa help-tooltip fa-question-circle"" data-toggle=""tooltip"" title=""Files on the CDMC Server waiting to be loaded. When selected (highlighted below), click 'Load Files' button to Load CDMC File.""></i></th></tr>"
	
	intCount = 0
	
	For Each objFile in colFiles		
		
		'Get the files with the starting name the same as the system default/setting name
		If Left(objFile.Name,Len(strFileNameDefault)) = Trim(strFileNameDefault) Then
		
			strFileExt = objFSO.GetExtensionName(objFile.Name) 
			
			If IsNull(strFileExt) Then
				strFileExt = ""
			Else
				strFileExt = LCase(strFileExt)
			End If
			
			If strFileExt = "csv" Then
			intCount = intCount + 1
			
				If intCount < 6 Then
					If IsNull(objFile.Name) or objFile.Name = "" Then
						strFile = ""
						strFileSize = 0
					Else
						strFile = Left(objFile.Name,10) & ".." & strFileExt'Right(objFile.Name,4)
						strFileSize = Round(objFile.Size/1024000,2)
						
						strFileAttributes =  "Created: " & objFile.DateCreated
						strFileAttributes = strFileAttributes & " Last Accessed: " & objFile.DateLastAccessed
						strFileAttributes = strFileAttributes & " Last Modified: " & objFile.DateLastModified  
		
					End If
					
					Response.Write "<TR><TD Title=""" & objFile.Name & " " & strFileAttributes & """ " & strServerSelect & ">" & strFile & "</TD></TR>"
				End If
			End If
		End If
		
	Next
	
	 Response.Write "<tr><td style=""text-align:left"">Total: " & intCount & " Size: " & strFileSize & " MB</td></tr>"
	 'Response.Write "<tr><td style=""text-align:left"">Size: " & strFileSize & " MB</td></tr>"'</table>"

	If strFile <> "" Then
		If strFileSize < 80 Then
			Response.Write "<div class=""alert alert-danger alert-dismissible"" style=""z-index:1; position: absolute; left: -290px; top: -180px;"" role=""alert""><a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a> The CDMC file looks a bit small (" & strFileSize & " MB). Please notify System Administrators before loading.</div>"
		End If
	End If
	If intCount > 1 Then
		Response.Write "<div class=""alert alert-danger alert-dismissible"" style=""z-index:1; position: absolute; left: -280px; top: -170px;"" role=""alert""><a href=""#"" class=""close"" data-dismiss=""alert"" aria-label=""close"">&times;</a> There is more than 1 CDMC file in the Import folder. Please notify System Administrators before loading or remove all CMS_AllData.csv files not for today.</div>"
	End If
 
objNetwork.RemoveNetworkDrive strServer, True, False
	 
Set objFSO = Nothing
Set objNetwork = Nothing
	 
End Sub


Function MediumDate(str)
'Function to change all date formats to medium date to avoid American storage challenge!	
Dim aDay
Dim aMonth
Dim aYear

	'Check to see whether the date passed in uses dashes (-) or slashes (/)
	If Instr(1,str,"/") = 0 Then
		If Mid(str,2,1) = "-" Then
			aDay = (Left((str),InStr(1,(str),"-")-1))
			aMonth = Mid(str,(InStr(1,(str),"-")+1),2)
		Else
			aDay = Mid((str),9,2)
			aMonth = Mid(str,(InStr(1,(str),"-")+1),2)
		End If
		
		If Right(aMonth,1) = "-" Then
			aMonth = Left(aMonth,1)
		End If
	Else
		If Mid(str,2,1) = "/" Then
			aDay = (Left((str),InStr(1,(str),"/")-1))
			aMonth = Mid(str,(InStr(1,(str),"/")+1),2)
		Else
			'aDay = Mid((str),9,2)
			aMonth = Mid(str,(InStr(1,(str),"/")+1),2)
			aDay = (Left((str),InStr(1,(str),"/")-1))
			
		End If
		
		If Right(aMonth,1) = "/" Then
			aMonth = Left(aMonth,1)
		End If

	End If
	
	aMonth = MonthName(aMonth)
	aYear = Year(str)
	
	If Len(aDay) = 1 Then aDay = "0" & aDay
	
	MediumDate = aDay & "-" & aMonth & "-" & aYear

End Function

Public Function ProcessFile(strProcessSeq,strFileSeqNum,strFileName,lngFileID)
'Function to Process an CDMC File which has been loaded into the database 
Dim intRecord
	'Process History
	
	If strProcessSeq = "A" Then		
	
		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCDMCUpdateFromTempLoad"
			.CommandTimeout = 300

			.Parameters.Append objCmd.CreateParameter("UserID", adInteger)
			
			.Parameters.Append objCmd.CreateParameter("CDMCOutput", adInteger, adParamOutput)
			
			.Parameters("UserID") = Session("UserID")	
			'Set the EmployeeId passed in to Zero(0) so all Employees are processed (otherwise only the EmployeeID will be processed)
			
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute 

		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCOutput")
	
	End If
	
	If strProcessSeq = "B" Then		
	
		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCDMCProcessContactDetails"
			.CommandTimeout = 300

			.Parameters.Append objCmd.CreateParameter("UserID", adInteger)
			.Parameters.Append objCmd.CreateParameter("EmployeeID", adVarChar, adParamInput, 20) 
			.Parameters.Append objCmd.CreateParameter("CDMCProcessOutput", adInteger, adParamOutput)
			
			.Parameters("UserID") = Session("UserID")	
			.Parameters("EmployeeID") = "0"	
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute 

		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCProcessOutput")
	
	End If
	
	Call UpdateCDMCFileLoadSummary(strProcessSeq,"CDMC", strFileSeqNum, strFileName, lngFileID) 

	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Records Processed in CDMC Upload File " & strFileSeqNum & ". Please notify System Administrators.</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert""> " & intRecord & " CDMC Upload File records Processed in file " & strFileSeqNum & "</div>"
	End If
	
	ProcessFile = intRecord
	
End Function


Public Function ProcessCDMC()
'Function to Process the CDMC file against the card list for Diners and add differences to the CS To Diners file 
' Also runs the SPROC spCAPSCancelBatchRemoveCountDown which cancels cards that have been removed from the CDMC file for longer than 5 days.
Dim intRecord
Dim intRecordDPC

		With objCmd2

			.CommandType = 4
			.CommandText = "spCAPSCDMCProcessCSFileToLoad"
			.CommandTimeout = 300

			.Parameters.Append objCmd2.CreateParameter("UserID", adInteger)
			.Parameters.Append objCmd2.CreateParameter("EmployeeID",  adVarChar, adParamInput,20)
			.Parameters.Append objCmd2.CreateParameter("CardType",  adVarChar, adParamInput,20)			
			.Parameters.Append objCmd2.CreateParameter("CDMCProcessCSFileToOutput", adVarChar, adParamOutput, 200)
			
			.Parameters("UserID") = Session("UserID")	
			'Set the EmployeeID passed in to Zero(0) so all Employees are processed (otherwise only the EmployeeID will be processed)
			.Parameters("EmployeeID") ="0"
			.Parameters("CardType") = "DTC"  'Does both DPC and DTC even though DTC is passed in.
		
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd2.Execute 
		
		intRecord = objCmd2.Parameters.Item("CDMCProcessCSFileToOutput")
		
		
		With objCmd3

			.CommandType = 4
			.CommandText = "spCAPSCDMCProcessCSFileTo"
			.CommandTimeout = 300

			.Parameters.Append objCmd3.CreateParameter("UserID", adInteger)
			.Parameters.Append objCmd3.CreateParameter("EmployeeID",  adVarChar, adParamInput,20)
			.Parameters.Append objCmd3.CreateParameter("CardType",  adVarChar, adParamInput,20)			
			.Parameters.Append objCmd3.CreateParameter("CDMCProcessCSFileToOutput", adVarChar, adParamOutput, 200)
			
			.Parameters("UserID") = Session("UserID")	
			'Set the EmployeeID passed in to Zero(0) so all Employees are processed (otherwise only the EmployeeID will be processed)
			.Parameters("EmployeeID") ="0"
			.Parameters("CardType") = "DPC"  'Does both DPC and DTC even though DTC is passed in.
		
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd3.Execute 
		
		intRecord = objCmd3.Parameters.Item("CDMCProcessCSFileToOutput")

		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCDMCProcessCSFileToNAB"
			.CommandTimeout = 300

			.Parameters.Append objCmd.CreateParameter("UserID", adInteger)
			.Parameters.Append objCmd.CreateParameter("EmployeeID",  adVarChar, adParamInput,20)
			.Parameters.Append objCmd.CreateParameter("CardType",  adVarChar, adParamInput,20)			
			.Parameters.Append objCmd.CreateParameter("CDMCProcessCSFileToOutput", adVarChar, adParamOutput, 200)
			
			.Parameters("UserID") = Session("UserID")	
			'Set the EmployeeID passed in to Zero(0) so all Employees are processed (otherwise only the EmployeeID will be processed)
			.Parameters("EmployeeID") ="0"
			.Parameters("CardType") = "DTC"  'Does both DPC and DTC even though DTC is passed in.
		
			
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute 
		
		intRecord = objCmd.Parameters.Item("CDMCProcessCSFileToOutput")
		
		
		
		


		With objCmd1

			.CommandType = 4
			.CommandText = "spCAPSSendCardExpiryEmails"
			.CommandTimeout = 300

			.Parameters.Append objCmd1.CreateParameter("UpdatedBy", adInteger)
			.Parameters.Append objCmd1.CreateParameter("SysMessage", adVarChar, adParamOutput, 200)
			
			.Parameters("UpdatedBy") = Session("UserID")	
						
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd1.Execute 

		Dim strCardExpiryEmailMsg
		strCardExpiryEmailMsg = objCmd1.Parameters.Item("SysMessage")
	
	
	If intRecord = "0" or strCardExpiryEmailMsg = "ERROR" Then
		If intRecord = "0" Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No Records Processed in CDMC Process against Diners DTC card list files. Please notify System Administrators.</div>"
		End If

		If strCardExpiryEmailMsg = "ERROR" Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error in Card Expiry Email process. Please notify System Administrators.</div>"
		End If
	Else
		If IsNumeric(intRecord) Then intRecord = intRecord -1
		Response.Write "<div class=""alert alert-success"" role=""alert""> " & intRecord & " CDMC File records Processed and added to the DTC CS To Diners file</div>"
	End If
	
	If intRecordDPC = "0" or strCardExpiryEmailMsg = "ERROR" Then
		If intRecordDPC = "0" Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">No Records Processed in CDMC Process against Diners DPC card list files. Please notify System Administrators.</div>"
		End If

		If strCardExpiryEmailMsg = "ERROR" Then
			Response.Write "<div class=""alert alert-danger"" role=""alert"">Error in Card Expiry Email process. Please notify System Administrators.</div>"
		End If
	Else
		If IsNumeric(intRecordDPC) Then intRecordDPC = intRecordDPC -1
		Response.Write "<div class=""alert alert-success"" role=""alert""> " & intRecordDPC & " CDMC File records Processed and added to the DPC CS To Diners file</div>"
	End If
	
End Function

Public Function MoveFile(strFileFrom,strFileTo)

Dim objFSO
Dim strExtension

	Set objFSO = CreateObject("Scripting.FileSystemObject")    

'Response.Write "<BR>" & strFileFrom	
'Response.Write "<BR>" & strFileTo
Response.Write "<div class=""alert alert-success"" role=""alert"">CDMC File Moved From: " & strFileFrom & " to: " & strFileTo & "</div>"
    
	If objFSO.FileExists(strFileFrom) Then	

                    If objFSO.FileExists(strFileTo) Then
					
						'Delete File if it exists
						objFSO.DeleteFile strFileTo
						objFSO.MoveFile strFileFrom, strFileTo 
						
						MoveFile = "OK"
                        
					Else
                        
						'Move the file to the Loaded folder
                        objFSO.MoveFile strFileFrom, strFileTo 

						MoveFile = "OK"						
                                                                               
                    End If
					
	Else
	
		MoveFile = "File does not exist."
		
	End If
					
    Set objFSO = Nothing                        

End Function


Public Function SynchCDMC()
''New for Portal and changes to CDMC table to reload the CDMC Load table to the CDMC table
'Function to truncate tblCAPSCDMC and insert all records from tblCAPSCDMCLoad

Dim intRecord
	'Process History
			
		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCDMCSynch"
			.CommandTimeout = 300

			.Parameters.Append objCmd.CreateParameter("UserID", adInteger)
			
			.Parameters.Append objCmd.CreateParameter("CDMCOutput", adInteger, adParamOutput)
			
			.Parameters("UserID") = Session("UserID")	
						
			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute 

		'Return the result of the Save Function.
		intRecord = objCmd.Parameters.Item("CDMCOutput")

	''Respond with a message of success or failure
	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">No Records synchronised between CDMC and CDMCLoad. Please notify System Administrators.</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert""> " & intRecord & " Records synchronised between CDMC and CDMCLoad</div>"
	End If
	
	SynchCDMC = intRecord
	
End Function



''''New function to get tblCAPSCDMCLoad records
Public Function GetCMDCLoad()

Dim objRSC

Set objRSC= Server.CreateObject("ADODB.Recordset")

	objRSC.Open "SELECT Count(CDMCID) AS Recs FROM tblCAPSCDMCLoad WITH(NOLOCK) ",objCon

		If objRSC.EOF Then
			GetCMDCLoad = 0
		Else
			GetCMDCLoad = objRSC("Recs")
		End If

	objRSC.Close

	If IsNumeric(GetCMDCLoad) Then GetCMDCLoad = FormatNumber(GetCMDCLoad,0)

Set objRSC = Nothing

End Function


Set objRS = Nothing

 %>


