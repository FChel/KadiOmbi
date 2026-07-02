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


'If the load decode has been clicked then call the procedure to load from CMS/ProMaster
If Request.QueryString("Action")="RunNAB" Then
	Call RunNAB()
End If

If Request.QueryString("Action")="RunCDMC" Then
	Call RunCDMC()
End If

If Request.QueryString("Action")="RunCAPS" Then
	Call RunCAPS()
End If

If Not IsEmpty(Request.QueryString("FileDate")) Then

	dteBatchDate = Request.QueryString("FileDate")
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
	self.location="DCCPSynch.asp?Action=SaveFileLocal"
}


function RunNAB()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="DCCPSynch.asp?Action=RunNAB"
}

function RunCDMC()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="DCCPSynch.asp?Action=RunCDMC"
}

function RunCAPS()
{
	document.getElementById('Progress').style.display = "inline";
	self.location="DCCPSynch.asp?Action=RunCAPS"
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


</script>
<script src="../js/jquery.js"></script>

<body>
<main class="main py-3">
      <div class="container">
<form action="DCCPSynch.asp?Action=Save&chkDelete="+frm.chkDelete.value  method="POST" enctype="multipart/form-data" id="frm" name="frm">

<div class="content-wrapper">
    <div class="container-fluid">
	 
	<div class="row" id="basic-table">
  <div class="col-5">
    <div class="card">
     <div class="card-header">
          <h4 class="card-title">Defence Credit Card Portal (DCCP) Data Synchronisation</h4>
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
	<button type="button" class="btn btn-primary btn-xs" onclick="RunCAPS();" title="(spCAPSDCCPUpdateCAPS) Inserts all new Cards received in CAPS into the DCCP tbPORTALCards"><i class="fa fa-upload"></i> 1. CAPS New Cards to Portal</button>
  </div>

  <div class="col-4 text-right">
	<button type="button" class="btn btn-primary btn-xs" onclick="RunNAB();" title="(spCAPSDCCPUpdateNAB) Updates all existing cards in DCCP cards from CS File"><i class="fa fa-upload"></i> 2. Update Portal From CAPS</button>
  </div>
    <div class="col-4 text-right">
	<button type="button" class="btn btn-primary btn-xs" onclick="RunCDMC();" title="(spCAPSDCCPUpdateCDMC) Runs Updates in DCCP from CAPS updates"><i class="fa fa-upload"></i> 3. Update Portal Data in DCCP</button>
  </div>
 
</div>
</div>

<div class="col-lg-12 col-md-12">
<p class="text-left">
<span id="Progress" style="display:none"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>
<br>

<font color="red" size="2"><b>NOTE: </Font><font color="black" size="2">Each button above will update the related detail from CAPS to the DCCP Portal

           <!-- <BR>* This may take a few minutes and is best run outside normal hours, as best as possible-->
</B></Font>

</p>
</div>

<div class="row col-12">

 <div class="col-4 text-right">
	
  </div>
<!--   <div class="col-4 text-right">
	<button type="button" class="btn btn-outline-primary btn-xs" onclick="UploadLocalAccountsSQL();"><i class="fa fa-upload"></i> Update Accounts (SQL)</button>
  </div>
-->
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
		<%'DisplaySummary()%>	
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
		
	objRS.Open "SELECT TOP 20 * FROM tblCAPSFileLoad WITH(NOLOCK) order by fileloadid desc",objCon,0,1
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">There is no File Load Data</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Okay to Update Data</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 table-compact"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""16"" style=""text-align:left"">Sample of recent Files Loaded to CAPS</th></tr>" & _
		        "<tr><th style=""text-align:center"">File ID</th>" & _
				"<th style=""text-align:center"">File Type</th>" & _
				"<th style=""text-align:center"">File Name</th>" & _	
		        "<th style=""text-align:center"">Record Count</th>" & _
	 	        "<th style=""text-align:center"">File Date Time</th>" & _	
	 	        "<th style=""text-align:center"">Status</th><th style=""text-align:center"">DateLoaded</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR><TD>" & objRS("FileLoadID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("FileType") & "</TD><TD style=""text-align:center"">" & objRS("FileName") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("RecordCount") & "</TD><TD style=""text-align:center"">" & objRS("FileDateTime") & "</TD><TD style=""text-align:center"">" & objRS("Status") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("DateLoaded") & "</TD></TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close

        Response.Write "</table>"
		
End Sub

Sub DisplaySummary()

Dim ObjCmd1
Dim objRS12 
Dim x

Set ObjCmd1 = Server.CreateObject("ADODB.Command")
Set objRS12 = Server.CreateObject("ADODB.Recordset")

objCmd1.ActiveConnection = objCon 

On Error Resume Next

	objCmd1.CommandText="spCAPSDCCPUpdate1"
	objCmd1.CommandType = 4

	objCmd1.Parameters.Append objCmd1.CreateParameter("@ID",3,1,,1)
	objCmd1.Parameters.Append objCmd1.CreateParameter("@StringInput",3,1,,1)
	objCmd1.Parameters.Append objCmd1.CreateParameter("@IDOutput",3,2,,1)
	
	Set objRS12 = objCmd1.Execute

	If Not objRS12.Eof Then
		Response.Write "<table><tr>"

		For x = 0 to objRS12.Fields.Count-1
			Response.Write "<td>" & objRS12.Fields(x).Name & "</td>"
		Next
		
		Response.Write "</tr>"
	End If


	Do Until objRS12.Eof
		Response.Write "<tr>"

		For x = 0 to objRS12.Fields.Count-1
			Response.Write "<td>" & objRS12.Fields(x).Value & "</td>"
		Next
		
		Response.Write "</tr>"
	objRS12.Movenext

	Loop

	objRS12.Close

        Response.Write "</table>"

If Err.Number <> 0 Then
	Response.Write  Err.Number & " Source: " & Err.Source & " Desc: " &  Err.Description
    Err.Clear
End If

On Error goto 0

Set objRS12 = Nothing
Set ObjCmd1 = Nothing
	
End Sub



Sub RunNAB()

Dim ObjCmd1
Dim objRS12 
Dim x
Dim intRecord 

Set ObjCmd1 = Server.CreateObject("ADODB.Command")
Set objRS12 = Server.CreateObject("ADODB.Recordset")

objCmd1.ActiveConnection = objCon 

'On Error Resume Next

	objCmd1.CommandText="spCAPSDCCPUpdateNAB"
	objCmd1.CommandType = 4

	objCmd1.Parameters.Append objCmd1.CreateParameter("@ID",3,1,,1)
	objCmd1.Parameters.Append objCmd1.CreateParameter("@StringInput",3,1,,1)
	objCmd1.Parameters.Append objCmd1.CreateParameter("@IDOutput",3,2,,1)
	
	Set objRS12 = objCmd1.Execute

	'Return the result of the Save Function.
	intRecord = objCmd1.Parameters.Item("@IDOutput") 

	If intRecord  = "" or IsNull(intRecord) Then intRecord = 0

'	If Not objRS12.Eof Then
'		Response.Write "<table><tr>"
'
'		For x = 0 to objRS12.Fields.Count-1
'			Response.Write "<td>" & objRS12.Fields(x).Name & "</td>"
'		Next
'		
'		Response.Write "</tr>"
'	End If
'
'
'	Do Until objRS12.Eof
'		Response.Write "<tr>"
'
'		For x = 0 to objRS12.Fields.Count-1
'			Response.Write "<td>" & objRS12.Fields(x).Value & "</td>"
'		Next
'		
'		Response.Write "</tr>"
'	objRS12.Movenext
'
'	Loop
'
'	objRS12.Close
'
 '       Response.Write "</table>"


Response.Write "<div class=""alert alert-success"" role=""alert""><i class=""fa fa-check""></i> " & intRecord & " - NAB Updates to the DCCP Portal. (tblPortalCards updated with changes from CS file)</div>"



'If Err.Number <> 0 Then
'	Response.Write  Err.Number & " Source: " & Err.Source & " Desc: " &  Err.Description
'    Err.Clear
'End If

'On Error goto 0

Set objRS12 = Nothing
Set ObjCmd1 = Nothing
	
End Sub


Sub RunCDMC()

Dim ObjCmd1
Dim objRS12 
Dim x
Dim intRecord 

Set ObjCmd1 = Server.CreateObject("ADODB.Command")
Set objRS12 = Server.CreateObject("ADODB.Recordset")

objCmd1.ActiveConnection = objCon 

On Error Resume Next

	objCmd1.CommandText="spCAPSDCCPUpdateCDMC"
	objCmd1.CommandType = 4

	objCmd1.Parameters.Append objCmd1.CreateParameter("@ID",3,1,,1)
	objCmd1.Parameters.Append objCmd1.CreateParameter("@StringInput",3,1,,1)
	objCmd1.Parameters.Append objCmd1.CreateParameter("@IDOutput",3,2,,1)
	
	Set objRS12 = objCmd1.Execute

	'Return the result of the Save Function.
	intRecord = objCmd1.Parameters.Item("@IDOutput") 

	If intRecord  = "" or IsNull(intRecord) Then intRecord = 0

'	If Not objRS12.Eof Then
'		Response.Write "<table><tr>"
'
'		For x = 0 to objRS12.Fields.Count-1
'			Response.Write "<td>" & objRS12.Fields(x).Name & "</td>"
'		Next
'		
'		Response.Write "</tr>"
'	End If
'
'
'	Do Until objRS12.Eof
''		Response.Write "<tr>"
'
'		For x = 0 to objRS12.Fields.Count-1
'			Response.Write "<td>" & objRS12.Fields(x).Value & "</td>"
'		Next
'		
'		Response.Write "</tr>"
'	objRS12.Movenext
'
'	Loop
'
'	objRS12.Close
'
 '       Response.Write "</table>"

Response.Write "<div class=""alert alert-success"" role=""alert""><i class=""fa fa-check""></i> " & intRecord & " - CDMC Updates to the DCCP Portal.</div>"



'If Err.Number <> 0 Then
'	Response.Write  Err.Number & " Source: " & Err.Source & " Desc: " &  Err.Description
'    Err.Clear
'End If

'On Error goto 0

Set objRS12 = Nothing
Set ObjCmd1 = Nothing
	
End Sub


Sub RunCAPS()

Dim ObjCmd1
Dim objRS12 
Dim x
Dim intRecord 

Set ObjCmd1 = Server.CreateObject("ADODB.Command")
'Set objRS12 = Server.CreateObject("ADODB.Recordset")

objCmd1.ActiveConnection = objCon 

'On Error Resume Next

	objCmd1.CommandText="spCAPSDCCPUpdateCAPS"
	objCmd1.CommandType = 4

	objCmd1.Parameters.Append objCmd1.CreateParameter("@ID",3,1,,1)
	objCmd1.Parameters.Append objCmd1.CreateParameter("@StringInput",3,1,,1)
	objCmd1.Parameters.Append objCmd1.CreateParameter("@IDOutput",3,2,,1)
	
	Set objRS12 = objCmd1.Execute

	'Return the result of the Save Function.
	intRecord = objCmd1.Parameters.Item("@IDOutput") 

	If intRecord  = "" or IsNull(intRecord) Then intRecord = 0

'
'	If Not objRS12.Eof Then
'		Response.Write "<table><tr>"
'
'		For x = 0 to objRS12.Fields.Count-1
'			Response.Write "<td>" & objRS12.Fields(x).Name & "</td>"
'		Next
'		
'		Response.Write "</tr>"
'	End If
'
'
'	Do Until objRS12.Eof
'		Response.Write "<tr>"
'
'		For x = 0 to objRS12.Fields.Count-1
'			Response.Write "<td>" & objRS12.Fields(x).Value & "</td>"
'		Next
''		
'		Response.Write "</tr>"
'	objRS12.Movenext
'
'	Loop
'
	'objRS12.Close
'
'        Response.Write "</table>"


Response.Write "<div class=""alert alert-success"" role=""alert""><i class=""fa fa-check""></i> " & intRecord & " - CAPS Updates to the DCCP Portal. (Inserts to tblPORTALCards from tblCapsCard)</div>"

'If Err.Number <> 0 Then
'	Response.Write  Err.Number & " Source: " & Err.Source & " Desc: " &  Err.Description
 '   Err.Clear
'End If

'On Error goto 0

Set objRS12 = Nothing
Set ObjCmd1 = Nothing
	
End Sub



Set objRS = Nothing
Set objRS1 = Nothing
Set objCon = Nothing

 %>


