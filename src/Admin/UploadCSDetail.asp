<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #include file="upload.asp" -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #include file="../CC/CAPSFunctions.asp" -->
<%
'Description:	ANZ Upload Detail screen with search facility
'Author:		Michael Giacomin
'Date:			May 2020

    'If the session has expired then send the user back to the Default/login page
	If IsEmpty(Session("UserID")) Then Response.Redirect("Timeout.asp")
	
'Instantiate Common Page Variables.
Dim objCon
Dim objCmd
Dim objRS
Dim arrMonthName(12)
Dim strEID
Dim strTop
Dim strFileSeqNum
Dim strCardNum

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

If request.QueryString("Action")="Save" Then

	'Call StartLoad()
End If

If Not IsEmpty(Request.QueryString("EIDNo")) Then

	strEID = Request.QueryString("EIDNo")
End If

If Not IsEmpty(Request.QueryString("FileSeqNum")) Then

	strFileSeqNum = cstr(Request.QueryString("FileSeqNum"))
End If

If Not IsEmpty(Request.QueryString("CardNum")) Then

	strCardNum = cstr(Request.QueryString("CardNum"))
End If

If Not IsEmpty(Request.QueryString("Top")) Then

	strTop = Request.QueryString("Top")
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
        if(getFileExt(document.getElementById('FILE1').value)==".xls")
        {
            if(window.confirm('This will overwrite any existing CS From Diners file data! \n \n Continue?')==true)
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

$( document ).ready(function() {
 
  	
$('#DateClick2').on('input', ':text', function(){ alert('asas'); });


});

jQuery(document).ready(function($) {
    $(".clickable-row").click(function() {
        window.location = $(this).data("href");
    });
});

function loadDoc() {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("BatchNos").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetBatch.asp?FileSeqNum=" + frm.FileSeqNumS.value + "", true);
  xhttp.send();
}

function loadDocE() {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
  //alert(this.responseText);
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("EmployeeIDST").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetEmployees.asp?EmpID=" + frm.EmpIDS.value + "&FName=" + frm.FNamms.value + "&LName=" + frm.LNamms.value + "", true);
  xhttp.send();
}

function SelectFileSeqNum(varFileSeqNum) {

	varFileSeqNum.toString();
	
	if(varFileSeqNum==undefined) {	
	}
	{
	self.location = "UploadCSDetail.asp?Action=Search&FileSeqNum=" + varFileSeqNum + "&EIDNo=" + document.getElementById('EmpIDAnchor').text;
	}
}
function SelectCardNum(varCardNum) {

	varCardNum.toString();
	
	if(varCardNum==undefined) {	
	}
	{
	self.location = "UploadCSDetail.asp?Action=Search&CardNum=" + varCardNum + "&EIDNo=" + document.getElementById('EmpIDAnchor').text + "&FileSeqNum=" + varFileSeqNum;
	}
}
function SelectEmp(varEmpID) {

	varEmpID.toString();
	
	if(varEmpID==undefined) {	
	}
	{
	self.location = "UploadCSDetail.asp?Action=Search&EIDNo=" + varEmpID + "&FileSeqNum=" + document.getElementById('FileAnchor').text;
	}
}

</script>

<style>

    table.newd th, table.newd td{

        padding: 4px; 

    }

</style>
</head>
<body>

<form action="UploadCSDetail.asp?Action=Save&chkDelete="  method="POST" enctype="multipart/form-data" id="frm" name="frm">
	<!--scrolling File Seq Num Modal -->
            <div class="modal fade" id="exampleModalScrollable" tabindex="-1" role="dialog" aria-labelledby="exampleModalScrollableTitle" aria-hidden="true">
              <div class="modal-dialog modal-dialog-scrollable" role="document">
                <div class="modal-content">
                  <div class="modal-header">
                    <h5 class="modal-title" id="exampleModalScrollableTitle">CS From Diners File No.</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                      <i class="bx bx-x"></i>
                    </button>
                  </div>
                  <div class="modal-body">
                 <!-- -->
				 <label>Search for File No --</label><br>
						
						<label for="FileSeqNumS">File No (FileSeqNum)</label>
						<input type="text" name="FileSeqNumS" id="FileSeqNumS" class="form-control input-md" onKeyUp="loadDoc();">
											   
				  </div>
				  <div id="BatchNos">
				  
				  </div>
                  
                  <div class="modal-footer">
					<button type="button" class="btn btn-primary" onClick="ClearFile();" ><i class="fa fa-check"></i> Clear</button>
                    <button type="button" class="btn btn-outline-secondary" data-dismiss="modal">
                      <i class="bx bx-x d-block d-sm-none"></i>
					  <span class="d-none d-sm-block"><i class="fa fa-times"></i> Close</span>
                    </button>
                   
                  </div>
                </div>
			</div>
		</div>

<!--scrolling Employee Modal -->
            <div class="modal fade" id="EmployeeModal" tabindex="-1" role="dialog" aria-labelledby="exampleModalScrollableTitle" aria-hidden="true">
              <div class="modal-dialog modal-dialog-scrollable" role="document">
                <div class="modal-content">
                  <div class="modal-header">
                    <h5 class="modal-title" id="exampleModalScrollableTitle">CS Employee ID</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                      <i class="bx bx-x"></i>
                    </button>
                  </div>
                  <div class="modal-body">
					 <!-- -->
					 <label>Search for Employee ID</label><br>
						<label for="FNamms">First Name:</label>
						<input type="text" name="FNamms" id="FNamms" class="form-control input-md" onKeyUp="loadDocE();">
						<label for="LNamms">Last Name:</label>
						<input type="text" name="LNamms" id="LNamms" class="form-control input-md" onKeyUp="loadDocE();">
						<label for="email">Employee ID:</label>
						<input type="text" name="EmpIDS" id="EmpIDS" class="form-control input-md" onKeyUp="loadDocE();">

					  <div id="EmployeeIDST">
					  
					  </div>
                  </div>
                  <div class="modal-footer">
					<button type="button" class="btn btn-primary" onClick="ClearEmp();" ><i class="fa fa-check"></i> Clear</button>
                    <button type="button" class="btn btn-outline-secondary" data-dismiss="modal">
                      <i class="bx bx-x d-block d-sm-none"></i>
                      <span class="d-none d-sm-block"><i class="fa fa-times"></i> Close</span>
                    </button>
                   
                  </div>
                </div>
			</div>
		</div>
<!--scrolling Card Number Modal -->		
		<div class="modal fade" id="CardNoModal" tabindex="-1" role="dialog" aria-labelledby="exampleModalScrollableTitle" aria-hidden="true">
              <div class="modal-dialog modal-dialog-scrollable" role="document">
                <div class="modal-content">
                  <div class="modal-header">
                    <h5 class="modal-title" id="exampleModalScrollableTitle">CS From Diners File No.</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                      <i class="bx bx-x"></i>
                    </button>
                  </div>
                  <div class="modal-body">
                 <!-- -->
				 <label>Search for Card No --</label><br>
						
						<label for="CardNumS">Card Number</label>
						<input type="text" name="CardNumS" id="CardNumS" class="form-control input-md" onKeyUp="loadDoc();">
											   
				  </div>
				  <div id="CardNos">
				  
				  </div>
                  
                  <div class="modal-footer">
					<button type="button" class="btn btn-primary" onClick="ClearCard();" ><i class="fa fa-check"></i> Clear</button>
                    <button type="button" class="btn btn-outline-secondary" data-dismiss="modal">
                      <i class="bx bx-x d-block d-sm-none"></i>
                      <span class="d-none d-sm-block"><i class="fa fa-times"></i> Close</span>
                    </button>
                   
                  </div>
                </div>
			</div>
		</div>
		
<div class="content-wrapper">
    <div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-10">
				<h3 class="text-left">CS Upload Details <i class="fa help-tooltip fa-question-circle" data-toggle="tooltip" title="All records from the CS file selected which can be filtered and exported."></i></h3>
			</div>
			
			
		</div>

      </section>
	  
	  
	 <section>
	 
	 
	<div class="row" id="basic-table">
	 <div class="col-2">
    <div class="card">
     
      <div class="card-content">
        <div class="card-body">
		
		
		
 <button type="button" class="btn btn-secondary btn-xs mr-1 mb-1" onclick="self.location='UploadCS2.asp'"><i class="fa fa-arrow-left"></i> Back To CS Summary</button>
 
  </div>
	  </div>
    </div>
   </div>
   
  <!-- Example DataTables Card-->
  
  <div class="col-8">
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

<!-- #Include file=../CC/CAPSFooter.asp -->
</body>
</html>

<%
Sub DisplayTableDetails()

Dim strWhere

If strEID = "" or ISNull(strEID) Then
Else
	strWhere = " AND EIDNo = '" & strEID & "'"
End If

If strFileSeqNum = "" OR ISNull(strFileSeqNum) OR strFileSeqNum = "All" Then
Else
	strWhere = strWhere & " AND FileSeqNum = '" & strFileSeqNum & "'"
End If


If strTop = "" or ISNull(strTop) Then
	strTop = " TOP 50 "
Else
	strTop = " TOP " & strTop & ""
End If

'If Not IsEmpty(Request.QueryString("BatchNo")) Then
'	strWhere = "WHERE BatchNo = " & Request.QueryString("BatchNo") & ""
	
'	If Request.QueryString("BatchNo") = 0 Then strWhere = strWhere & " OR BatchNo IS NULL"
'Else
'	strWhere = ""
'End If

If Len(strWhere) > 4 Then
	strWhere = "WHERE " & Right(strWhere,Len(strWhere)-4)
End If

		objRS.Open "SELECT " & strTop & " * FROM qryCAPSCSFromDiners WITH(NOLOCK) "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">No CS From Diners data for " & strWhere & "</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Select Again</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""12"" style=""text-align:left"">CS From Diners data loaded for " & strWhere & "</th>" & _
				"<th colspan=""2""><button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""window.open('UploadCS2.asp')""><i class=""fa fa-table""></i> Export</button></th></tr>" & _
                "<tr><td colspan=""14"" style=""text-align:left; color:red;font-size:20px""><B>" & strWhere & "</B></td></tr><tr>" & _
		        "<th Style=""width:20px;"">ID</th>" & _
				"<th>File Date Time</th>" & _
				"<th>File Seq Num</th><th>EID No</th>" & _	
		        "<th>Card No</th>" & _
	 	        "<th>Card Expiry</th>" & _	
	 	        "<th>Card Status</th>" & _
		        "<th>Name On Card</th>" & _
                "<th>Address 1</th><th>Address 2</th><th>Address 3</th>" & _
                "<th>Suburb</th><th>State</th><th style=""background-color:white; color:black;"">Status</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR class='clickable-row' data-href='DisplayDataset.asp?tbl=tblCAPSCSFromDiners&W=WHERE CSFromDinersID=" & objRS("CSFromDinersID") & "' data-target='_blank'><TD>" & objRS("CSFromDinersID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("FileDateTime") & "</TD><TD style=""text-align:center"">" & objRS("FileSeqNum") & "</TD><TD style=""text-align:center"">" & objRS("EIDNo") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & MaskCard(objRS("CardNo")) & "</TD><TD style=""text-align:center"">" & objRS("CardExpiryDate") & "</TD><TD style=""text-align:center"">" & objRS("CardStatus") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("NameOnCard") & "</TD><TD style=""text-align:center"">" & objRS("Address1") & "</TD><TD style=""text-align:center"">" & objRS("Address2") & "</TD>" & _
								"<TD style=""text-align:center"">" & objRS("Address3") & "</TD><TD style=""text-align:center"">" & objRS("Suburb") & "</TD><TD style=""text-align:center"">" & objRS("State") & "</TD>" & _
			                    "<TD style=""text-align:center; background-color:white; color:black;"">" & objRS("Status") & "</TD></TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close

        Response.Write "</table>"
		
End Sub


Sub DisplaySummary()

Dim strFileSeqNumSearch
Dim strEmpIDSearch
Dim strCardNoSearch
Dim dteBatchDate
Dim dteBatchDateFormat

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)

	'Write the Summary table headers
	 Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
	"<tr><th colspan=""8"" style=""text-align:left"">CS From Diners Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" />)</th></tr>" & _
	"<th>File Seq Num</th>" & _
	"<th>EID</th>" & _
	"<th>Card No</th></tr>" 

	If IsNull(strFileSeqNum) or strFileSeqNum = "" Then
		strFileSeqNumSearch = "All"
	Else
		strFileSeqNumSearch = strFileSeqNum
	End If
	
	If IsNull(strEID) or strEID = "" Then
		strEmpIDSearch = "All"
	Else
		strEmpIDSearch = strEID
	End If
	
	If IsNull(strFileSeqNum) or strFileSeqNum = "" Then
		strCardNoSearch = "All"
	Else
		strCardNoSearch = strCardNum
	End If
	
	'Write the Summary table detail
	Response.Write "<TR><TD><a href=""#"" data-toggle=""modal"" data-target=""#exampleModalScrollable"" id=""FileAnchor"">" & strFileSeqNumSearch & "</A></B></TD>" & _
				"<TD style=""text-align:center""><a href=""#"" data-toggle=""modal"" data-target=""#EmployeeModal"" id=""EmpIDAnchor"">" & strEmpIDSearch & "</a></TD>" & _
				"<TD style=""text-align:center""><a href=""#"" data-toggle=""modal"" data-target=""#CardNoModal"" id=""CardNumAnchor"">" & strCardNoSearch & "</a></TD>" & _
				"</TR>"
	
        
	Response.Write "</table>"
		
End Sub


Set objRS = Nothing
Set objCon = Nothing

 %>


