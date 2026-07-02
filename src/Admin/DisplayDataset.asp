<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #include file="upload.asp" -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #include file="../CC/CAPSFunctions.asp" -->
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

Dim strEID
Dim strTop
Dim strTable
Dim strWhere

'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
Set ObjCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")

If Not IsEmpty(Request.QueryString("EIDNo")) Then

	strEID = Request.QueryString("EIDNo")
End If

If Not IsEmpty(Request.QueryString("tbl")) Then

	strTable = cstr(Request.QueryString("tbl"))
End If

If Not IsEmpty(Request.QueryString("W")) Then

	strWhere = cstr(Request.QueryString("W"))
End If

'response.write "w=" & strWhere

 %>

	
<script language=javascript>

</script>

<style>

    table.newd th, table.newd td{

        padding: 4px; 

    }

</style>
</head>
<body>

<form action="UploadCSDetail.asp?Action=Save&chkDelete="  method="POST" enctype="multipart/form-data" id="frm" name="frm">
	
		
<div class="content-wrapper">
    <div class="container-fluid">
	
	<section class="breadcrumbs py-2">
		<div class="row">
			<div class="col-md-10">
				<h3 class="text-left">CS To Diners Data View and Export <i class="fa help-tooltip fa-question-circle" data-toggle="tooltip" title="All records from the file selected which can be filtered and exported."></i></h3>
			</div>
			
			
		</div>

      </section>
	  
	  
	<div class="row" id="basic-table">
	 <div class="col-2">
    <div class="card">
     
      <div class="card-content">
        <div class="card-body">
		
		
		
 <button type="button" class="btn btn-secondary btn-xs mr-1 mb-1" onclick="javascript:history.go(-1)"><i class="fa fa-arrow-left"></i> Back</button>
 <button type="button" class="btn btn-secondary btn-xs mr-1 mb-1" onclick="window.open('../CC/ExcelExport.asp?<%="tbl=" & strTable & "&W=" & strWhere%>')"><i class="fa fa-table"></i> Export To Excel</button>
 
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

Dim fldField
Dim strHeader
Dim strDetail
Dim x
Dim strSQL

	'If Session("EmployeeID") = 0 Then
		'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	'	strSQL = "SELECT * FROM tblCSToDiners " & strWhere
	'Else
		strSQL = "SELECT * FROM " & strTable & " WITH(NOLOCK) " & strWhere
		'strSQL = "SELECT * FROM qryApplications WHERE EmployeeID = " & Session("EmployeeID") & ""
	'End If

'response.write "SQ=" & strSQL

	Response.Write"<table Class=""table table-bordered table-hover mb-0 newd"" cellspacing=""0"" cellpadding=""0"">"

	objRS.Open strSQL,objCon
		
		
		If NOT objRS.Eof Then
		
			For each fldField in objRS.fields
			
				strHeader = strHeader & "<th style=""font-size:12px;"">" & fldField.name & "</th>"
			
			next
		
		strHeader = strHeader & "</tr>"
			
		
		End If
		
		Do until objRS.EOF
			
			strDetail = strDetail & "<tr>"
			
			For each fldField in objRS.fields
			
				strDetail = strDetail & "<td style=""font-size:12px;"">" & fldField.value & "</td>"
			
				x = x + 1
			next
			
			strDetail = strDetail & "</tr>"
			
			objRS.movenext
		Loop
		
		
						
objRS.Close

        strDetail = strDetail & "</table>"
		
		'Go back and add the first row in after counting the number of fields (columns)
		strHeader = "<tr><th colspan=""" & x & """>CS To Diners " & strWhere & "</th></tr>" & strHeader
				
		
		'Write the header and Detail to the screen
		Response.Write strHeader & strDetail
		
End Sub




Set objRS = Nothing
Set objCon = Nothing

 %>


