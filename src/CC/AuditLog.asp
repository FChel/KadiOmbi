<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=CAPSFunctions.asp -->
<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")

'Description:	Detail from the Audit Log table (tblCAPSAuditLog) which details all changes to cards
'Author:		Michael Giacomin
'Date:			July 2020

	Response.Expires = -1500	

Dim objCon
Dim objRS

Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL

Dim intEmployeeID

Dim lngApplicationID
Dim strEmployeeID
Dim strFirstName
Dim strLastName
Dim lngCDMCID
Dim strCardSearch

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
 	
    objCon.Open Session("DBConnection")	

	If IsNull(Session("CardID")) OR Session("CardID") = "" Then Session("CardID")= 0

	If Not IsEmpty(Request.QueryString("SearchAll")) Then
		Session("SearchAll") = Request.QueryString("SearchAll")
	End If

	If Not IsEmpty(Request.QueryString("ApplicationID")) Then
		Session("ApplicationID") = Request.QueryString("ApplicationID")
		Call LoadApplication(Session("ApplicationID"))
	End If

	If Not IsEmpty(Request.QueryString("CardID")) Then
		Session("CardID") = Request.QueryString("CardID")
	End If
	
	If Not IsEmpty(Request.QueryString("EmployeeID")) Then
		Session("EmployeeID") = Request.QueryString("EmployeeID")
	End If

	If Not IsEmpty(Request.QueryString("EmployeeSearchID")) Then
		Session("EmployeeSearchID") = Request.QueryString("EmployeeSearchID")
	End If
  
	If Not IsEmpty(Request.QueryString("Action"))  Then 
		If Request.QueryString("Action") = "Search" Then
			 Session("EmployeeSearchID") = Request.QueryString("EmployeeID")
		End If
	End If

	'If the is a global Card Id selected then get the Card Number (global function in include file) and allow them to remove it (button)
	If Not IsNull(Session("CardID")) Then
		If Session("CardID") <> 0 Then
			strCardSearch = " and Card " & MaskCard(GetCardNo(Session("CardID"))) & " <button type=""button"" class=""btn btn-outline-secondary btn-xs"" onclick=""self.location='AuditLog.asp?CardID=0'""; title=""Click to remove the Card Number from Search""><i class=""fa fa-times""></i> </button>"
		End If
	End If
	
	Call LoadEmployee()
  
%>

<script LANGUAGE="javascript">

f

function SaveData(){
	var varSubmit = true
		frm.msgbox.value='Saving.......';
		frm.submit();
	//}
}

function CloseScreen() {

    //if(top.Header1.Header2.document.form1.SaveStatus.value=='S')
    //{var x=window.confirm("Changes have been made, do you wish to save these changes?")
    //    if (x){
    //        SaveData();
    //        self.location='index.asp';
    //    }
    //    else
    //        self.location='index.asp';}
    //    else
    { self.location = 'HomeCC.asp'; }
}

setTimeout( 'ShowTimeoutWarning();', 1080000 );

function ShowTimeoutWarning () {     
    window.alert( "********** Warning! **********' \n \n 'You will be automatically logged out in 2 minutes unless you change screens, Close or Save!" ); 
}


function AppSearch() {
   
    self.location = "AuditLog.asp?Action=Search&EmployeeID2=" + frm.EmployeeIDSearch.value;
      
}

function SelectApp(varAppID) {
	if(varAppID==undefined) {
		alert(varAppID);
	}
	{
	self.location = "AuditLog.asp?Action=Search&ApplicationID=" + varAppID;
	}
}

function SelectEmp(varEmpID) {
	if(varEmpID==undefined) {
		alert(varEmpID);
	}
	{
	//self.location = "AuditLog.asp?Action=Search&EmployeeID=" + varEmpID;
	self.location = "AuditLog.asp?Action=Search&EmployeeID=" + varEmpID + "&FileSeqNum=1"
	//self.location = "AuditLog.asp?Action=Search&EmployeeID=" + varEmpID + "&FileSeqNum=" + document.getElementById('FileAnchor').text;
	}
}


function loadDoc() {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("EmpSear").innerHTML = this.responseText;
    }
  };
  //xhttp.open("GET", "GetEmployees.asp?EmpID=" + frm.EmpIDS.value + "&FName=" + frm.FirstName.value + "&LName=" + frm.LNamms.value + "", true);
  xhttp.open("GET", "AJAX/GetEmployees.asp?EmpID=" + frm.EmpIDS.value + "&FName=" + frm.FirstName.value + "&LName=" + frm.LNamms.value + "", true);
  xhttp.send();
}

function loadDetail(varID) {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("AuditLogDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetAuditLog.asp?AuditLogID=" + varID + "", true);
  xhttp.send();
}

</script>
	
<body >

<form action="AuditLog.asp?Action=Save" method="POST" id="frm" name="frm">

 <!-- Modal Compare -->
<div class="modal fade" id="AuditLogMod" tabindex="-1" role="dialog" aria-labelledby="compareModalLabel" aria-hidden="true">
     <div class="modal-dialog modal-large modal-dialog-centered modal-dialog-scrollable">
         <div class="modal-content">
              <div class="modal-header">
                <h5 class="modal-title" id="compareModalLabel">
                  Audit Log Detail
                </h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                  <span aria-hidden="true">&times;</span>
                </button>
              </div>
			<div class="modal-body" id="AuditLogDetail">
               
				  
                
            </div>

			<div class="modal-footer">
				<button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
			</div>
		</div>
	</div>
 </div>
<!-- Modal -->
<div class="modal fade" id="exampleModalCenter" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle"> Search for an Employee</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>Enter details and press Go.</label><br>
            <label for="FirstName">First Name:</label>
            <input type="text" name="FirstName" id="FirstName" class="form-control input-md">
			<label for="LNamms">Last Name:</label>
            <input type="text" name="LNamms" id="LNamms" class="form-control input-md">
			<label for="EmpIDS">Employee ID:</label>
            <input type="email" name="EmpIDS" id="EmpIDS" class="form-control input-md">
           
      </div>
	  <div id="EmpSear">
	  
	  </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
        <button type="button" class="btn btn-primary" onClick="loadDoc()">Go</button>
      </div>
	  <table><tr><td></td><td></td><td></td></tr></table>
    </div>
  </div>
</div>
</form>

    <main class="main py-5">
      <div class="container">
        <div class="row">
          <div class="col-md-9">
            <h4 class="py-2">Card Audit Log History: <span style="font-size:18px; color:grey; font-style:italic; font-weight:lighter;"><% Response.Write Session("EmployeeSearchID") & " - " & strFirstName & " " & strLastName & " " & strCardSearch%></span></h4>
          </div>
          <div class="col-md-3 text-right">
			<button type="button" class="btn btn-primary" data-toggle="modal" data-target="#exampleModalCenter"><i class="fa fa-search"></i> Search for Employee</button>
          
          </div>
        </div>
        <div class="row">
          <div class="col-12">
            <div id="table-tabs" class="as-tabs">
				
				 <h5 class="py-2">Audit Log</h5>
                  <table class="table table-compact">
                    <thead>
                      <tr>
                        <th scope="col" style="font-size:14px;">Change Date</th>
                        <th scope="col" style="font-size:14px;">Type</th>
                        <th scope="col" style="font-size:14px;">Sub Type</th>
                        <th scope="col" style="font-size:14px;">EID</th>
						<th scope="col" style="font-size:14px;">Card Type</th>
						<th scope="col" style="font-size:14px;">Card Number</th>
                        <!--<th scope="col" style="font-size:14px;">Actioned By</th>-->
                        <th scope="col" style="font-size:14px;">Value Before</th>
                        <th scope="col" style="font-size:14px;">Value After</th>
                        <th scope="col" style="font-size:14px;">Source File</th>
						<th scope="col" style="font-size:14px;">Change Details</th>
                        <!--<th scope="col">Card ID</th>
                        <th scope="col">Application ID</th>
                        <th scope="col">CS From Diners ID</th>
                        <th scope="col">CS to Diners ID</th>-->
						<th scope="col" style="font-size:14px;">Updated By</th>
                      </tr>
                    </thead>
                    <tbody>
					<% Call DisplayAuditLog()%>
                      
                    </tbody>
                  </table>
				  
          </div>
        </div>
      </div>
    </main>
	
	<!-- #Include file=CAPSFooter.asp -->
	
</body>
</html>
<%

Public Sub DisplayAuditLog()

Dim intRecord, intRows
Dim strCardNo
Dim strSourceFile
Dim dteChangeDate
Dim strAuditLogID
Dim strUpdateBy

If Session("CardID") = 0 Then
	strSQL = "SELECT * FROM qryCAPSAuditLog WITH(NOLOCK) WHERE EID = '" & Session("EmployeeID") & "' ORDER BY [ChangeDate] DESC"
Else
	strSQL = "SELECT * FROM qryCAPSAuditLog WITH(NOLOCK) WHERE CardID = '" & Session("CardID") & "' ORDER BY [ChangeDate] DESC"
End If


	objRS.Open strSQL,objCon

		Do Until objRS.Eof
		
			intRecord = intRecord + 1

			If IsNumeric(objRS("CardNumber")) Then
				strCardNo = MaskCard(objRS("CardNumber"))
			Else
				If IsNull(objRS("CardNumber")) Then
					strCardNo = 0
				Else
					strCardNo = objRS("CardNumber")
				End If
			End If
			
			'Format the Source File string and shorten if too long
			If IsNull(objRS("SourceFile")) Then
				strSourceFile = ""
			Else
				If Len(objRS("SourceFile")) > 20 Then
					strSourceFile = Right(objRS("SourceFile"),20)
				Else
					strSourceFile = objRS("SourceFile")
				End If
			End If
			'Format the Change Date and remove time
			If IsNull(objRS("ChangeDate")) Then
				dteChangeDate = ""
			Else
				If IsDate(objRS("ChangeDate")) Then
					dteChangeDate = FormatDateTime(objRS("ChangeDate"),vbShortDate)
				Else
					dteChangeDate = objRS("ChangeDate")
				End If
			End If
			
			If IsNull(objRS("AuditLogID")) Then
				strAuditLogID = ""
				Else
				strAuditLogID = objRS("AuditLogID")
			End If
			
			If IsNull(objRS("UpdatedByName")) Then
				strUpdateBy = "AUTO"
			Else
				strUpdateBy = objRS("UpdatedByName")
			End If
			
			'If intRecord < 4 then
			
			'	GetAuditLog = GetAuditLog & "<td>" & objRS("ChangeDate") & "</td><td>" & objRS("ValueAfter") & "</td><td></td>"
			'End If
		'	Response.Write "<tr><td><a data-toggle=""modal"" data-target=""#AuditLogMod"" HREF=""#"" onClick=""loadDetail(" & objRS("AuditLogID") & ")"">" & objRS("AuditLogID") & "</a></td><td style=""font-size:12px;"">" & dteChangeDate & "</td><td style=""font-size:12px; font-weight:bold;"">" & objRS("Type") & "</td><td style=""font-size:12px;"">" & objRS("SubType") & "</td><td style=""font-size:12px;"">" & objRS("EID") & "</td>" & _
		'			"<td style=""font-size:12px;"">" & objRS("CardType") & "</td><td style=""font-size:12px;"">" & MaskCard(strCardNo) & "</td><td style=""font-size:12px;"">" & objRS("ActionedByName") & "</td><td style=""font-size:12px;"">" & objRS("ValueBefore") & "</td><td style=""font-size:12px;"">" & objRS("ValueAfter") & "</td>" & _
		'			"<td style=""font-size:12px;"">" & strSourceFile & "</td><td style=""font-size:12px;"">" & objRS("ChangeDetails") & "</td><td>" & objRS("UpdatedByName") & "</td></tr>"
			'<td style=""font-size:12px;"">" & objRS("ActionedByName") & "</td>
			
			Response.Write "<tr><td><a data-toggle=""modal"" data-target=""#AuditLogMod"" HREF=""#"" onClick=""loadDetail(" & objRS("AuditLogID") & ")"" style=""font-size:12px;"">" & dteChangeDate & "</a></td>" & _
			"<td style=""font-size:12px; font-weight:bold;"">" & objRS("Type") & "</td><td style=""font-size:12px;"">" & objRS("SubType") & "</td><td style=""font-size:12px;"">" & objRS("EID") & "</td>" & _
			"<td style=""font-size:12px;"">" & objRS("CardType") & "</td><td style=""font-size:12px;"">" & strCardNo & "</td><td style=""font-size:12px;"">" & objRS("ValueBefore") & "</td><td style=""font-size:12px;"">" & objRS("ValueAfter") & "</td>" & _
			"<td style=""font-size:12px;"">" & strSourceFile & "</td><td style=""font-size:12px;"">" & objRS("ChangeDetails") & "</td><td>" & strUpdateBy & "</td></tr>"
					
			objRS.Movenext
		Loop

	objRS.Close
	
	'For intRows = 1 to 9 - intRecord
	
	'	GetAuditLog = GetAuditLog & "<td></td>"
		
	'Next

	'If Len(GetAuditLog) > 30  Then GetAuditLog = left(GetAuditLog,Len(GetAuditLog)-9)

End Sub


Public Sub LoadEmployee()
'Description:	Loads Employee details into page variables

	objRS.Open "SELECT * FROM qryCAPSCDMCHistorySearch WHERE EmployeeID = '" & Session("EmployeeSearchID") & "'",objCon
		  
		If Not objRS.EOF Then
		   
			'strEmployeeID = objRS("EmployeeID")
			strFirstName = objRS("FirstName")
			strLastName  = objRS("Surname")
			
	   End If

	objRS.Close

End Sub


Public Sub LoadApplication(lngAppID)
'Description:	Loads Employee ID for the Application selected/passed in, to use in search

	objRS.Open "SELECT [EmployeeID] FROM tblCAPSApplication WITH(NOLOCK) WHERE ApplicationID = " & lngAppID & "",objCon
		  
		If Not objRS.EOF Then
		   
			Session("EmployeeSearchID") = objRS("EmployeeID")
			
	   End If

	objRS.Close

End Sub

Set objRS = Nothing
Set objCon = Nothing
%>
