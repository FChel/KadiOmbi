<%@ Language=VBScript %>
<!-- #include file="../upload.asp" -->
<!-- #Include file=../../ADOVBS.inc -->
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
Dim arrMonthName(12)
Dim strEID
Dim strTop

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

If Not IsEmpty(Request.QueryString("Top")) Then

	strTop = Request.QueryString("Top")
End If


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
    <!--<link rel="stylesheet" type="text/css" href="../../Frest/app-assets/css/core/menu/menu-types/horizontal-menu.min.css">-->
    <!-- END: Page CSS-->


	<!-- BEGIN: Page JS-->
    <script src="../../Frest/app-assets/js/scripts/modal/components-modal.min.js"></script>
    <!-- END: Page JS-->


  <!-- BEGIN: Custom CSS-->
    <!--<link rel="stylesheet" type="text/css" href="../../Frest/assets/css/style.css">-->
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
  xhttp.open("GET", "../../CC/AJAX/GetBatch.asp?FileSeqNum=" + frm.FileSeqNumS.value + "", true);
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
  xhttp.open("GET", "../../CC/AJAX/GetEmployees.asp?EmpID=" + frm.EmpIDS.value + "&FName=" + frm.FNamms.value + "&LName=" + frm.LNamms.value + "", true);
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
                    <h5 class="modal-title" id="exampleModalScrollableTitle">CS From Diners Batch No.</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                      <i class="bx bx-x"></i>
                    </button>
                  </div>
                  <div class="modal-body">
                 <!-- -->
				 <label>Search for Batch No --</label><br>
						
						<label for="FileSeqNumS">Batch No (FileSeqNum)</label>
						<input type="text" name="FileSeqNumS" id="FileSeqNumS" class="form-control input-md" onKeyUp="loadDoc();">
											   
				  </div>
				  <div id="BatchNos">
				  
				  </div>
                  
                  <div class="modal-footer">
                    <button type="button" class="btn btn-light-secondary" data-dismiss="modal">
                      <i class="bx bx-x d-block d-sm-none"></i>
                      <span class="d-none d-sm-block">Close</span>
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
                    <h5 class="modal-title" id="exampleModalScrollableTitle">CS From Diners Employee ID</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                      <i class="bx bx-x"></i>
                    </button>
                  </div>
                  <div class="modal-body">
					 <!-- -->
					 <label>Search for Employee ID/label><br>
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
                    <button type="button" class="btn btn-light-secondary" data-dismiss="modal">
                      <i class="bx bx-x d-block d-sm-none"></i>
                      <span class="d-none d-sm-block">Close</span>
                    </button>
                   
                  </div>
                </div>
			</div>
		</div>
		
<div class="content-wrapper">
    <div class="container-fluid">
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

<!-- BEGIN: Vendor JS-->
    <script src="../../Frest/app-assets/vendors/js/vendors.min.js"></script>
   
    <!-- BEGIN Vendor JS-->

    
</body>
</html>

<%
Sub DisplayTableDetails()

Dim strWhere

If strEID = "" or ISNull(strEID) Then
Else
	strWhere = " AND EIDNo = '" & strEID & "'"
End If

If strFileSeqNum = "" or ISNull(strFileSeqNum) Then
Else
	strWhere = strWhere & " AND FileSeqNum = '" & strFileSeqNum & "'"
End If


If strTop = "" or ISNull(strTop) Then
	strTop = " TOP 50 "
Else
	strTop = " TOP = " & strTop & ""
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

		objRS.Open "SELECT " & strTop & " * FROM qryCAPSCSFromDinersReco WITH(NOLOCK) "  & strWhere,objCon
		
		    If objRS.eof Then
		        Response.Write"<table Class=""table"" WIDTH=""100%"" BORDER=""1"" CELLSPACING=""1"" CELLPADDING=""1"">" & _
                "<tr><th colspan=7 style=""text-align:left"">No CS From Diners data for " & strWhere & "</B></th></tr>" & _
		        "<tr><td colspan=7 style=""text-align:left"">Select Again</td></tr>"
		    Else
		    
		         Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
                "<tr><th colspan=""15"" style=""text-align:left"">CS From Diners data loaded for " & strWhere & "</th>" & _
				"<th colspan=""2""><button type=""button"" class=""btn btn-secondary btn-xs"" onclick=""window.open('UploadCS2.asp')""><i class=""fa fa-table""></i> Export</button></th></tr>" & _
                "<tr><td colspan=""17"" style=""text-align:left; color:red;font-size:20px""><B>" & strWhere & "</B></td></tr><tr>" & _
		        "<th Style=""width:20px;"">ID</th>" & _
				"<th>File Date Time</th>" & _
				"<th>File Seq Num</th><th>EID No</th>" & _	
		        "<th>Card No</th>" & _
	 	        "<th>Card Expiry</th>" & _	
	 	        "<th>Card Status</th>" & _
		        "<th>Title</th><th>Given Names</th>" & _
                "<th>Surname</th><th>Name On Card</th>" & _
                "<th>Address 1</th><th>Address 2</th><th>Address 3</th>" & _
                "<th>Suburb</th><th>State</th><th style=""background-color:white; color:black;"">CS To Diners</th></tr>"
		        
		    End If
		    
		    Do until objRS.eof
		        'If objRS("GLCode") <> 0 Then
			    Response.Write "<TR class='clickable-row' data-href='DisplayDataset.asp?tbl=tblCAPSCSFromDiners&W=WHERE CSFromDinersID=" & objRS("CSFromDinersID") & "' data-target='_blank'><TD>" & objRS("CSFromDinersID") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("FileDateTime") & "</TD><TD style=""text-align:center"">" & objRS("FileSeqNum") & "</TD><TD style=""text-align:center"">" & objRS("EIDNo") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("CardNo") & "</TD><TD style=""text-align:center"">" & objRS("CardExpiryDate") & "</TD><TD style=""text-align:center"">" & objRS("CardStatus") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("Title") & "</TD><TD style=""text-align:center"">" & objRS("GivenNames") & "</TD><TD style=""text-align:center"">" & objRS("Surname") & "</TD>" & _
			                    "<TD style=""text-align:center"">" & objRS("NameOnCard") & "</TD><TD style=""text-align:center"">" & objRS("Address1") & "</TD><TD style=""text-align:center"">" & objRS("Address2") & "</TD>" & _
								"<TD style=""text-align:center"">" & objRS("Address3") & "</TD><TD style=""text-align:center"">" & objRS("Suburb") & "</TD><TD style=""text-align:center"">" & objRS("State") & "</TD>" & _
			                    "<TD style=""text-align:center; background-color:white; color:black;"">" & objRS("CSToDinersID") & "</TD></TR>"
    			'End If
    			
			    objRS.movenext
		    Loop
    			
	    objRS.Close

        Response.Write "</table>"
		
End Sub


Sub DisplaySummary()

Dim strFileSeqNumSearch
Dim strEmpIDSearch

	If IsNull(dteBatchDate) or dteBatchDate = "" Then dteBatchDate = DateAdd("d", -200, Now())
	
	dteBatchDate = Day(dteBatchDate) & "-" & MonthName(Month(dteBatchDate)) & "-" & Year(dteBatchDate)
	dteBatchDateFormat = Year(dteBatchDate) & "-" & Month(dteBatchDate) & "-" & Day(dteBatchDate)

	'Write the Summary table headers
	 Response.Write"<table Class=""table table-bordered mb-0 newd"" cellspacing=""0"" cellpadding=""0"">" & _
	"<tr><th colspan=""8"" style=""text-align:left"">CS From Diners Summary (Since <input class=""DateClick2"" type=""date"" style=""color:red;"" value=""" & dteBatchDateFormat & """ id=""CSDate"" />)</th></tr>" & _
	"<th>File Seq Num</th>" & _
	"<th>EID</th>" & _
	"<th>Total Cards</th><th>Total Employees</th>" & _	
	"<th>DTC</th>" & _
	"<th>CMC</th>" & _	
	"<th>Status</th>" & _
	"<th>Date Loaded</th></tr>" 

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
	
	'Write the Summary table detail
	Response.Write "<TR><TD><a href=""#"" data-toggle=""modal"" data-target=""#exampleModalScrollable"" id=""FileAnchor"">" & strFileSeqNumSearch & "</A></B></TD>" & _
				"<TD style=""text-align:center""><a href=""#"" data-toggle=""modal"" data-target=""#EmployeeModal"" id=""EmpIDAnchor"">" & strEmpIDSearch & "</a></TD><TD style=""text-align:center"">111</TD><TD style=""text-align:center"">111</TD>" & _
				"<TD style=""text-align:center"">1111</TD><TD style=""text-align:center"">111</TD><TD style=""text-align:center"">111</TD>" & _
				"<TD style=""text-align:center"">111</TD></TR>"
	
        
	Response.Write "</table>"
		
End Sub


Set objRS = Nothing
Set objCon = Nothing

 %>


