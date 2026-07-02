
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=ADOVBS.inc -->
<!-- #Include file=CAPSFunctions.asp -->
<%

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp?State=Expired")


'Description:	Create and view applications
'Author:		MG
'Date:			January 2020

	Response.Expires = -1500	

Dim objCon
Dim objCmd
Dim objRS
Dim objRS1
Dim strSelected
Dim x 
Dim strMessage
Dim strColour
Dim dteBatchDate
Dim intTotalTasks
Dim intTasksComplete
Dim intMorningTasks
Dim intAfternoonTasks
Dim intMorningTasksComplete
Dim intAfternoonTasksComplete
Dim strGlobalMessage

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")
Set objRS1 = Server.CreateObject("ADODB.Recordset")


'Open database connection
objCon.Open Session("DBConnection")

If IsNull(Session("BatchDate")) Or Session("BatchDate") = "" Then Session("BatchDate") = now()

If Not IsEmpty(Request.QueryString("FileDate")) Then

	Session("BatchDate") = Request.QueryString("FileDate")
End If

'Update the Global message
If Not IsEmpty(Request.QueryString("Action")) Then

	If Request.QueryString("Action") = "SaveGlobalMessageClear" Then
		Call SaveGlobalMessage("")
	Else
		Call SaveGlobalMessage(Request.QueryString("GlobalMessage"))
	End If
	
End If

'Call the procedure to load the count of tasks and tasks completed for display
'Call LoadTasksCount()

If IsNull(Application("GlobalMessage")) or Application("GlobalMessage") = "" Then
	strGlobalMessage = "<i>None</i>"
Else
	strGlobalMessage = Application("GlobalMessage")
End If

'Call CheckHost()

%>
<script>
function OpenSs(cb) {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	//var e = document.getElementById("CardTypeSelect");
	//var result = e.options[e.selectedIndex].value;
	
	//document.getElementById('CardType').value=result;
	alert('asa');
}

function DatePickChange() {
	self.location="HomeAdmin.asp?FileDate=" + document.getElementById("AdminDate").value;
}

function GlobalMessageChange() {
	self.location="HomeAdmin.asp?Action=SaveGlobalMessage&GlobalMessage=" + document.getElementById("GlobalMessage").value;
}

function loadDoc() {
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("UsersLoggedDetail").innerHTML = this.responseText;
    }
  };
  xhttp.open("GET", "AJAX/GetUsersLoggedIn.asp", true);
  xhttp.send();
  
}

</script>

<form action="HomeAdmin.asp?Action=Save" method="POST" id="frm" name="frm" class="inline">
<!-- Modal -->
<div class="modal fade" id="MessUse" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">CAPS Contact</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>Global Message to ALL Users</label><br>

			<label for="UserLoggedIn">Current Message:</label>
            <input type="text" name="GlobalMessage" id="GlobalMessage" class="form-control input-md" value="<%=strGlobalMessage%>">
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
		<button type="button" class="btn btn-primary" onClick='window.location="HomeAdmin.asp?Action=SaveGlobalMessageClear"'><i class="fa fa-times"></i> Clear Message</button>
        <button type="button" class="btn btn-primary" onClick="GlobalMessageChange();"><i class="fa fa-check"></i> Save</button>
      </div>
      </div>
    </div>
  </div>
</div>

<!-- Modal -->
<div class="modal fade" id="ModAppUsers" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">CAPS Users Logged In</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>CAPS Users Logged In:</label><br>

            <%=Application("NamedUsers")%>
			
      </div>
	  <div class="modal-body" id="UsersLoggedDetail" name="UsersLoggedDetail">
		</div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
      </div>
      </div>
    </div>
  </div>
  
</form>

    <main class="main py-3">
      <div class="container">
		<h2 class="text-left py-3">Admin Dashboard</h2>
	  
        <div class="row">
          <div class="col-md-6">
            <div class="panel panel-shadow mb-12">
              <div class="panel-header col-md-12">
                <h4>Admin Tasks</h4>
                <span class="panel-subheader">Admin Tasks for <% Call LoadDatePicker()%></span>
				<!--<div class="col-md-12 text-right my-auto"><p>Complete: <strong><%=intTasksComplete & " of " & intTotalTasks%></strong></p></div>-->
              </div>
              <div class="panel-content">
                	<% Call LoadTasksCount()%>		
			
			 <div class="my-information">
              <ul class="nav nav-tabs" id="myFiTab" role="tablist">
                <li class="nav-item" role="presentation">
                  <a class="nav-link active" id="my-cards-tab" data-toggle="tab" href="#Morning" role="tab" aria-controls="Morning" aria-selected="true">Opening Tasks</a>
                </li>
                <li class="nav-item" role="presentation">
                  <a class="nav-link" id="my-applications-tab" data-toggle="tab" href="#Afternoon" role="tab" aria-controls="Afternoon aria-selected="false">Closing Tasks</a>
                </li>
              </ul>
              <div class="tab-content overflow-auto" id="myFiTabContent" style="height:900px;">
                <div class="tab-pane fade show active" id="Morning" role="tabpanel" aria-labelledby="my-cards-tab">
				<%
					
					Call LoadTaskGroups()
				
				%>
                  
                
                </div>
              </div>
            </div>
		
		 </div>
            </div>
			
          </div>
          <div class="col-md-6 sidebar">
		  

              <div class="panel panel-shadow panel-validation mb-3">
				  <div class="panel-header">
					<h4>Recent Admin Tasks</h4>
					<span class="panel-subheader">Recently Run Tasks</span>
				  </div>
				  <div class="panel-content">
				  
				  <%Call LoadRecentTasks() %>
				</div>
			  </div>
            
			
            <div class="panel panel-shadow mb-3">
              <div class="panel-header">
                <h4>Summary</h4>
                <span class="panel-subheader">Admin Summary</span>
              </div>
			  <div class="panel-content row">
              <div class="col-md-6 text-left my-auto">
                <a href="#" class="block-link" data-toggle="modal" data-target="#ModAppUsers" onClick="loadDoc();">
                  <i class="fa fa-user"></i>
                  <span class="content">Users Logged In <p style="font-size:24px; font-weight:bold; color:black; text-align:right"><%=Application("users")%></p></span>
				 
                </a>
				</div>
				<div class="col-md-6 text-right my-auto">
                <a href="#" class="block-link" data-toggle="modal" data-target="#MessUse">
                  <i class="fa fa-bullhorn"></i>
                  <span class="content">Message All Users <p style="font-size:14px; font-weight:bold; color:black;"><%=Left(strGlobalMessage,20)%></p></span>
                </a></div></div>
				
				<!--<iframe width="600px;" height="200px;" src="../vendor/chart.js/BarChart2.html" scrollbar="no" border="0px;"></iframe>-->
				
              </div> 
			  
			   <div class="panel panel-shadow mb-3">
              
			  <div class="panel-content row">
				<iframe width="600px;" height="350px;" src="../vendor/chart.js/UserTypesBar.asp" scrolling="no" style="border:none;"></iframe>
				<!--<iframe width="200px;" height="200px;" src="../vendor/chart.js/PieChart1.html" scrolling="no" style="border:none;"></iframe>-->
			 </div>
				
              </div> 
			  
			  
            </div>
          </div>
        </div>
		
		
		
      </div>
    </main>

	<!--<script src="js/jquery.js"></script>
    <script src="bootstrap/js/bootstrap.bundle.min.js"></script>-->
    <script>
      jQuery(function ($) {
        $('[data-toggle="popover"]').popover();
      });
    </script>
	 <!-- Chart.js -->
    <script src="../js/plugins/Chart.min.js"></script>
	
<!-- #Include file=CAPSFooter.asp -->
  </body>
</html>

<%



Public Sub LoadTaskGroups()

Dim y, x, intTotal
Dim strLink
Dim strTaskType
Dim strAction
Dim strCreditLimit
Dim dteExiryDate
Dim strTaskTime
Dim strTaskName
Dim strTaskDescription
Dim strStatus
Dim arrTasks(2,100)
Dim strTaskStatus
Dim strAdminTaskDate
Dim lngAdminTaskID
Dim strSQL

'First get the list of tasks as this is fixed, then run through the tasks which have been completed for the day
'This is to avoid having a complex join, calendar table or multiple recordsets
strSQL = "SELECT * FROM tblCAPSTaskLog WITH(NOLOCK) WHERE ([AdminTaskDate]>= '" & Session("BatchDate") & "' AND [AdminTaskDate] <= DateAdd(day, 1, '" & Session("BatchDate") & "')) ORDER BY [AdminTaskID]"

	objRS.Open strSQL,objCon,0,1
	
		'Get the Task list for the current day		
		Do Until objRS.EOF
		
			y = y + 1
			
				arrTasks(0,y) = objRS("AdminTaskID")
				arrTasks(1,y) = objRS("TaskStatus")
				arrTasks(2,y) = objRS("AdminTaskDate")
			
		objRS.Movenext
		
		Loop

	objRS.Close

	'Set the Total Taks Complete for the date selected
	intTotalTasks = y
	
'Open the master recordset
'If strTaskTime = "" OR ISNull(strTaskTime) Then
'	strSQL = "SELECT * FROM tblCAPSAdminTasks WITH(NOLOCK) WHERE TaskTime = '" & strTaskTime & "'"
'Else
	'strSQL = "SELECT * FROM qryCAPSAdminTaskGroups WITH(NOLOCK) WHERE ([AdminTaskDate]>= '" & Session("BatchDate") & "' AND [AdminTaskDate] <= DateAdd(day, 1, '" & Session("BatchDate") & "')) ORDER BY TaskGroup"
	
	'Response.write strSQL
	strSQL = "SELECT DISTINCT TaskGroup,TaskTime,TaskGroupOrder FROM tblCAPSAdminTasks WHERE Active = 'Y' AND TaskGroup Is Not Null ORDER BY TaskGroupOrder"
	
'End If

''''''''Possibly should be below and loop through Task status to make sure all are complete
'SELECT DISTINCT TaskGroup,TaskTime,TaskGroupOrder,TaskStatus,MAX(tblCAPSTaskLog.Dateupdated) FROM tblCAPSAdminTasks 
'Inner Join tblCAPSTaskLog on tblCAPSAdminTasks.AdminTaskID= tblCAPSTaskLog.AdminTaskID
'WHERE TaskGroup Is Not Null
'Group By TaskGroup,TaskTime,TaskGroupOrder,TaskStatus
'ORDER BY TaskGroupOrder

'''End----- May have to add the date to the above

'Response.write strSQL

	objRS.Open strSQL,objCon,0,1
	
	If Not ObjRS.EOF Then
		If objRS("TaskTime") = "" OR IsNull(objRS("TaskTime")) Then
			strTaskTime = ""
		Else
			strTaskTime = objRS("TaskTime")
		End If
	End If
	
	Do Until objRS.EOF
		
		'Make sure that there is not a null AdminTaskID
		If IsNull(objRS("TaskGroup")) Then 
			lngAdminTaskID = 0 
		Else
			lngAdminTaskID = objRS("TaskGroup")
		End If
		
		If objRS("TaskTime") <> strTaskTime Then
			Response.Write "</div><div class=""tab-pane fade"" id=""" & objRS("TaskTime") & """ role=""tabpanel"" aria-labelledby=""" & objRS("TaskTime") & """>"
		End If
		
		If objRS("TaskTime") = "" OR IsNull(objRS("TaskTime")) Then
			strTaskTime = ""
		Else
			strTaskTime = objRS("TaskTime")
		End If		
		
		strTaskStatus = Get_Group_Status(objRS("TaskGroup"))
		
		If strTaskStatus = "Complete" Then
			strStatus = "<span class=""badge badge-pill badge-success"">Complete</span>"
		ElseIf strTaskStatus = "Error" Then
			strStatus = "<span class=""badge badge-pill badge-danger"">Error</span>"
		ElseIf strTaskStatus = "Started" Then
			strStatus = "<span class=""badge badge-pill badge-warning"">Started</span>"
		Else
			strStatus = "<span class=""badge badge-pill badge-danger"">Not Run</span>"
		End If
		
		'Response.Write Get_Group_Status(objRS("TaskGroup"))
		
		
		'Response.Write "<DIV>" & objRS("TaskGroup") & "</DIV>"
		Response.Write "<a href=""AdminTasks.asp?TaskGroup=" & objRS("TaskGroup") & "&TaskID=" & objRS("TaskGroup") & "#TaskSection" & objRS("TaskTime") & """ class=""section-link ""><div class=""status"">" & objRS("TaskTime") & "</div>" & _
                "<div class=""col-md-6 content""><span class=""title"" title=""" & strTaskDescription & """>" & objRS("TaskGroup") & ". " & strTaskName & "</span></div>" & _
				"<div class=""col-md-4 text-right my-auto"" title=""" & strAdminTaskDate & """><p>Status: <strong>" & strStatus & "</strong></p></div>" & _
				"</a>"
				
				
				
		'Response.Write "<div class=""panel panel-shadow mb-3""><div class=""panel-content row"">" & _
		'	"<div class=""col-md-1 text-center my-auto"">" & strLink & " " & strTaskType & "</div>" & _
		'	"<div class=""col-md-4""><h6>" & strTaskName & " " & strAction & "</h6></div>"  &_
		'	"<div class=""col-md-6 text-right my-auto"">" & strStatus & " " & strLink & "View Details <i class=""fa fa-arrow-right""></i></a></div></div></div>"
		
		intTotal = intTotal + 1

		
		objRS.Movenext
	Loop
	objRS.Close
	

End Sub


Public Sub LoadRecentTasks()

Dim y
Dim strLink
Dim strTaskType
Dim strAction
Dim strCreditLimit
Dim dteExiryDate
Dim strTaskTime
Dim strTaskName
Dim strTaskDescription
Dim dteDateRun
Dim strDateRunFormat
Dim strStatus
Dim strSQL

'Open the master recordset
'If strTaskTime = "" OR ISNull(strTaskTime) Then
'	strSQL = "SELECT * FROM tblCAPSAdminTasks WITH(NOLOCK) WHERE TaskTime = '" & strTaskTime & "'"
'Else
	strSQL = "SELECT top 4 * FROM qryCAPSAdminTaskLogSummary WITH(NOLOCK) WHERE Active = 'Y' AND TaskStatus = 'Complete' ORDER BY [TaskLogDateUpdated] DESC"
'End If

	objRS.Open strSQL,objCon,0,1
	
	'Get the TaskTime to write the headings for each tab
	If Not ObjRS.EOF Then
		If objRS("TaskTime") = "" OR IsNull(objRS("TaskTime")) Then
			strTaskTime = ""
		Else
			strTaskTime = objRS("TaskTime")
		End If
	End If
	
	Do Until objRS.EOF
		
		If objRS("TaskTime") <> strTaskTime Then
			'Response.Write "</div><div class=""tab-pane fade"" id=""" & objRS("TaskTime") & """ role=""tabpanel"" aria-labelledby=""" & objRS("TaskTime") & """>"
		End If
		
		If objRS("TaskTime") = "" OR IsNull(objRS("TaskTime")) Then
			strTaskTime = ""
		Else
			strTaskTime = objRS("TaskTime")
		End If
		
		If objRS("TaskName") = "" OR IsNull(objRS("TaskName")) Then
			strTaskName = ""
		Else
			strTaskName = objRS("TaskName")
		End If
		
		If objRS("TaskType") = "" OR IsNull(objRS("TaskType")) Then
			strTaskType = ""
		Else
			strTaskType = objRS("TaskType")
		End If
		
		If objRS("TaskDescription") = "" OR IsNull(objRS("TaskDescription")) Then
			strTaskDescription = ""
		Else
			strTaskDescription = Left(objRS("TaskDescription"),50)
		End If
		
		If Trim(strTaskType) = "Diners" Then
			strTaskType = "<img src=""../images/diners2.png"" Height=""40px"" Width=""60px"" title=""" & strTaskDescription & """> "
		ElseIf strTaskType = "ANZ" Then
			strTaskType = "<img src=""../images/ANZ.png"" Height=""40px"" Width=""60px"" Title=""" & strTaskDescription & """> " '& strCardType
		ElseIf strTaskType = "Defence" Then
			strTaskType = "<img src=""../images/logo_coa.png"" Height=""40px"" Width=""70px"" Title=""" & strTaskDescription & """> " '& strCardType
		ElseIf strTaskType = "CMS" Then
			strTaskType = "<img src=""../images/CMS2.png"" Height=""20px"" Width=""70px"" Title=""" & strTaskDescription & """> " '& strCardType
		Else
			strTaskType = "<img src=""../images/logo_coa.png"" Height=""40px"" Width=""60px"" Title=""" & strTaskDescription & """> "
		End If
		
		If IsNull(objRS("TaskLogDateUpdated")) Then
			dteDateRun = ""
		Else
			dteDateRun = objRS("TaskLogDateUpdated")
			If dteDateRun < now() -10 then
				strDateRunFormat = "color:red; font-weight:bold;"
			Else
				strDateRunFormat = "color:grey; font-size:13px;"
			End If
		End If
		
		If objRS("TaskStatus") = "Complete" Then
			strStatus = "<span class=""badge badge-pill badge-success"">Complete</span>"
		ElseIf objRS("TaskStatus") = "Error" Then
			strStatus = "<span class=""badge badge-pill badge-danger"">Error</span>"
		Else
			strStatus = "<span class=""badge badge-pill badge-danger"">Not Run</span>"
		End If
	
		
		'Response.Write "<div class=""panel panel-shadow mb-3""><div class=""panel-content row"">" & _
		'	"<div class=""col-md-1 text-center my-auto"">" & strLink & "" & strTaskType & "</div>" & _
		'	"<div class=""col-md-4""><h6>" & strTaskName & " " & strAction & "</h6></div>"  &_
		'	"<div class=""col-md-3 text-right my-auto"">" & strLink & "View Details <i class=""fa fa-arrow-right""></i></a></div></div></div>"
		
		Response.Write "<a href=""AdminTasks.asp?TaskID=" & objRS("AdminTaskID") & """ class=""section-link ""><div class=""status"">" & strTaskType & "</div>" & _
                "<div class=""col-md-6 content""><span class=""title"">" & strTaskName & "</span><span class=""description"">" & strTaskDescription & "</span></div>" & _
				"<div class=""col-md-4 text-right my-auto""><p>Status: <strong>" & strStatus & "</strong></p><p>Date Complete: <span style=""" & strDateRunFormat & " font-size:14px;"">" & dteDateRun & "</span></p></div>" & _
				"</a>"
		
		'Response.Write "<a href=""ApplicationsSubmit.asp?ApplicationChecks=Name"" class=""section-link ""><div class=""status"">" & strTaskType & "</div>" & _
        '        "<div class=""content""><span class=""title"">" & strTaskName & "</span><span class=""description"">" & strTaskDescription & "</span></div></a>"
				
		
		
		objRS.Movenext
	Loop
	objRS.Close
	

End Sub

Public Sub LoadDatePicker()

Dim dteBatchDateFormat

	If IsNull(Session("BatchDate")) or Session("BatchDate") = "" Then Session("BatchDate") = DateAdd("d", -1, Now())
	
	Session("BatchDate") = Day(Session("BatchDate")) & "-" & MonthName(Month(Session("BatchDate"))) & "-" & Year(Session("BatchDate"))
	dteBatchDateFormat = Year(Session("BatchDate")) & "-" & Month(Session("BatchDate")) & "-" & Day(Session("BatchDate"))
	
	dteBatchDateFormat = Year(Session("BatchDate")) & "-" & Right("0" & Month(Session("BatchDate")), 2) & "-" & Right("0" & Day(Session("BatchDate")), 2)

	'objRS.Open "SELECT TOP 6 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE DateLoaded > '" & Session("BatchDate") & "' AND FileType = 'ROMANCostCentres' ORDER BY FileSeqNum DESC",objCon
	
	Response.Write "<input class=""DateClick2"" type=""date"" style=""color:blue;"" value=""" & dteBatchDateFormat & """ id=""AdminDate"" name=""AdminDate"" onChange=""DatePickChange();"" />"

End Sub


Public Sub LoadTasksCount()

Dim strMorningColour
Dim strAfternoonColour

Dim dblResult
Dim strSQL

'Response.Write Session("BatchDate")

'First get the list of tasks as this is fixed, then run through the tasks which have been completed for the day
'This is to avoid having a complex join, calendar table or multiple recordsets
'strSQL = "SELECT [TaskLogID] FROM tblCAPSTaskLog WITH(NOLOCK) WHERE ([AdminTaskDate]>= '" & Session("BatchDate") & "' AND [AdminTaskDate] <= DateAdd(day, 1, '" & Session("BatchDate") & "')) ORDER BY [AdminTaskID]"
'strSQL = "SELECT [TaskLogID] FROM tblCAPSTaskLog WITH(NOLOCK) WHERE ([AdminTaskDate]>= convert(varchar,'" & now() & "',103) AND [AdminTaskDate] <= convert(varchar,DateAdd(day, 1, '" & now() & "'),103)) ORDER BY [AdminTaskID]"
strSQL = "SELECT * FROM qryCAPSTaskLogAdminTasks WITH(NOLOCK) WHERE ([AdminTaskDate]>= '" & Session("BatchDate") & "' AND [AdminTaskDate] <= DateAdd(day, 1, '" & Session("BatchDate") & "')) ORDER BY [AdminTaskID]"

intTasksComplete = 0
intMorningTasks = 0
intAfternoonTasks = 0
intMorningTasksComplete = 0
intAfternoonTasksComplete = 0

	objRS.Open strSQL,objCon,0,1
	
		'Get the Task list for the current day		
		Do Until objRS.EOF
		
			intTasksComplete = intTasksComplete + 1
			
			If objRS("TaskTime") = "Morning" Then
				intMorningTasksComplete = intMorningTasksComplete + 1
			End If
			
			If objRS("TaskTime") = "Afternoon" Then
				intAfternoonTasksComplete = intAfternoonTasksComplete + 1
			End If
			
		objRS.Movenext
		
		Loop

	objRS.Close

strSQL = "SELECT [AdminTaskID],[TaskTime] FROM qryCAPSAdminTaskLog WITH(NOLOCK) WHERE Active = 'Y' ORDER BY [SortOrder]"

objRS.Open strSQL,objCon,0,1
	
		'Get the Task list for the current day		
		Do Until objRS.EOF
		
			intTotalTasks = intTotalTasks + 1
			
			If objRS("TaskTime") = "Morning" Then
				intMorningTasks = intMorningTasks + 1
			End If
			
			If objRS("TaskTime") = "Afternoon" Then
				intAfternoonTasks = intAfternoonTasks + 1
			End If
			
		objRS.Movenext
		
		Loop

	objRS.Close
	
	strAfternoonColour = "red"
	strMorningColour = "red"
	
	If IsNull(intMorningTasks) or intMorningTasks = 0 Then
		strMorningColour = ""
	Else
		If IsNull(intMorningTasksComplete) or intMorningTasksComplete = 0 Then
		Else
			dblResult = intMorningTasksComplete/intMorningTasks
			Select Case intMorningTasks/intMorningTasksComplete
				Case 0.8
					strMorningColour = "green"
				Case 0.5
					strMorningColour = "yellow"
				Case 0.2
					strMorningColour = "red"
				Case Else
					strMorningColour = "red"
			End Select
		End If
	End If
	
	If IsNull(intAfternoonTasks) or intAfternoonTasks = 0 Then
		strAfternoonColour = ""
	Else
		If IsNull(intAfternoonTasksComplete) or intAfternoonTasksComplete = 0 Then
		Else
			dblResult = intAfternoonTasksComplete/intAfternoonTasks
			Select Case true
				Case dblResult > 0.9
					strAfternoonColour = "green"
				Case dblResult > 0.5
					strAfternoonColour = "yellow"
				Case dblResult > 0.2
					strAfternoonColour = "red"
				Case Else
					strAfternoonColour = "red"
			End Select
		End If
	End If
	
	Response.Write "<div class=""row number-container""><div class=""col-6"">" & _
		"<span class=""number number-lg color-" & strMorningColour & """>" & intMorningTasksComplete & "/" & intMorningTasks & "</span>" & _
		"<small>Opening Tasks Complete</small></div>" & _
		"<div class=""col-6""><span class=""number number-lg color-" & strAfternoonColour & """>" & intAfternoonTasksComplete & "/" & intAfternoonTasks & "</span>" & _
		"<small>Closing Tasks Complete</small></div></div>"
	
End Sub

Sub SaveGlobalMessage(strGlobalMessageSave)
'Procedure to update the application variable for the global message

	Application.Lock     
	Application("GlobalMessage") = strGlobalMessageSave
	Application.Unlock

End Sub

Public Sub CheckHost()
Dim strHost

	strHost = "vbmrsn05"
	
	If Ping(strHost) = True Then
		Response.Write strHost & " contacted"
	Else
		Response.Write strHost & " NOT contacted"
	End If

End Sub
	
Public Function Ping(strHost)
Dim objPing
Dim objRetStatus

	Set objPing = Server.CreateObject("winmgmts:{impersonationlevel=impersonate}").ExecQuery ("SELECT * FROM Win32_PingStatus WHERE Address = '" & strHost & "'")
	
	For Each objRetStatus in objPing
		If IsNull(objRetStatus.StatusCode) OR objRetStatus.StatusCode <> 0 Then
			Ping = False
			'Response.Write "Status Code is " & objRetStatus.StatusCode
		Else
			Ping = True
			'Response.Write "Time (ms) = " & vbTab & objRetStatus.BufferSize
			'Response.Write "Time (ms) = " & vbTab & objRetStatus.ResponseTime
		End If
	
	Next

End Function

Private Sub CheckWebsiteUp()
Dim strWebsite

	strWebsite = "vbmrsn05"

	If IsWebsiteUp( strWebsite ) Then
	   Response.Write "Web site " & strWebsite & " is up and running!"
	Else
		Response.Write "Web site " & strWebsite & " is down!!!"
	End If
End Sub

Function IsWebsiteUp( myWebsite )
' This function checks if a website is running by sending an HTTP request.
' If the website is up, the function returns True, otherwise it returns False.
' Argument: myWebsite [string] in "www.domain.tld" format, without the
' "http://" prefix.
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com
'
' The X-HTTP component is available at:
' http://www.xstandard.com/page.asp?p=C8AACBA3-702F-4BF0-894A-B6679AA949E6
' For more information on available functionality read:
' http://www.xstandard.com/printer-friendly.asp?id=32ADACB9-6093-452A-9464-9269867AB16E
    Dim objHTTP

    Set objHTTP = CreateObject( "XStandard.HTTP" )

    objHTTP.AddRequestHeader "User-Agent", _
                             "Mozilla/4.0 (compatible; MyApp 1.0; Windows NT 5.1)"
    objHTTP.Get "http://" & myWebsite

    If objHTTP.ResponseCode = 200 Then
        IsWebsiteUp = True
    Else
        IsWebsiteUp = False
    End If

    Set objHTTP = Nothing
End Function

Function Get_Group_Status(strTaskGroup)

Dim intGroupTasks
Dim intTasksCompleted

	objRS1.Open "SELECT Count(AdminTaskID) FROM tblCAPSAdminTasks WHERE ACTIVE = 'Y' AND TaskGroup = '" & strTaskGroup & "'",objCon
	
		If not objRS1.EOF Then
			intGroupTasks = objRS1(0)
		Else
			intGroupTasks = 0
		End If
	
	objRS1.Close
	
	objRS1.Open "SELECT Count(AdminTaskID) FROM qryCAPSTaskLogAdminTasks WHERE [AdminTaskDate] = '" & Session("BatchDate") & "' AND TaskGroup = '" & strTaskGroup & "' AND TaskStatus = 'Complete'",objCon

		If not objRS1.EOF Then
			intTasksCompleted = objRS1(0)
		Else
			intTasksCompleted = 0
		End If
	
	objRS1.Close
	

	If intGroupTasks = intTasksCompleted Then
		Get_Group_Status = "Complete"
	ElseIf intGroupTasks <> intTasksCompleted AND intTasksCompleted <> 0 Then
		Get_Group_Status = "Started"
	End If	

	
	objRS1.Open "SELECT Count(AdminTaskID) FROM qryCAPSTaskLogAdminTasks WHERE [AdminTaskDate] = '" & Session("BatchDate") & "' AND TaskGroup = '" & strTaskGroup & "' AND TaskStatus = 'Error'",objCon

		If not objRS1.EOF and objRS1(0) > 0 Then
			Get_Group_Status = "Error"	
		End If
	
	objRS1.Close


End Function


Set objRS = Nothing
Set objCon = Nothing

%>