
<!-- #Include file=CAPSHeader.asp -->
<!-- #Include file=ADOVBS.inc -->
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
Dim strViewButton
Dim dteBatchDate
Dim strTaskGroup

'Set Database objects

Set objCon = Server.CreateObject("ADODB.Connection")
Set objCmd = Server.CreateObject("ADODB.Command")
Set objRS = Server.CreateObject("ADODB.Recordset")

'Open database connection
objCon.Open Session("DBConnection")

'Collect the action results to save details
If Not IsEmpty(Request.QueryString("Task")) Then
	'Response.Write Request.QueryString("TaskID") & "," & Request.QueryString("Task") & "," & Request.QueryString("TaskLogID")
	Call SaveTaskLog(Request.QueryString("TaskID"),Request.QueryString("Task"),Request.QueryString("TaskLogID"))
End If

If Not IsEmpty(Request.QueryString("ViewButton")) Then
	strViewButton = Request.QueryString("ViewButton")
	Session("ViewButton") = strViewButton
End If

If IsNull(Session("BatchDate")) Or Session("BatchDate") = "" Then Session("BatchDate") = now()

If Not IsEmpty(Request.QueryString("FileDate")) Then
	Session("BatchDate") = Request.QueryString("FileDate")
End If

If Not IsEmpty(Request.QueryString("TaskTime")) Then
	Session("TaskTime") = Request.QueryString("TaskTime")
End If

If Not IsEmpty(Request.QueryString("TaskGroup")) Then
	Session("TaskGroup") = Request.QueryString("TaskGroup")
End If

%>
<script>
function OpenSs(cb) {

	//alert("asas");
	//var e = document.getElementById(this.cb);
	//var result = e.options[e.selectedIndex].value;
	
	//document.getElementById('ContinueMod').value=result;
	document.getElementById('ModApp').showModal();
}


function DatePickChange() {
	self.location="AdminTasks.asp?FileDate=" + document.getElementById("AdminDate").value;
}

</script>
<!-- Modal -->
<div class="modal fade" id="ModApp" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
  <div class="modal-dialog modal-dialog-centered" role="document">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLongTitle">CAPS Admin Tasks</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>Continue With Process</label><br>
            <input type="text" name="ContinueMod" id="ContinueMod" class="form-control input-md" value="">
			<input type="text" name="ContinueModTask" id="ContinueModTask" class="form-control input-md" value="">
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Close</button>
      </div>
    </div>
  </div>
</div>


    <main class="main py-3">
      <div class="container">
		<h3 class="text-left">Admin Dashboard</h3>
	  
        <div class="row">
          <div class="col-md-12">
            <div class="panel panel-shadow mb-3">
              <div class="panel-header">
			  <div class="panel-content row" Style="padding:0px;">
			  <div class="col-md-6">
                <h4>Admin Tasks</h4>
                <span class="panel-subheader">Admin Tasks for <% Call LoadDatePicker()%></span></div>
					<div class="col-md-4">
						<% Call LoadButtons() %>
					</div>
				</div>
				</div>
              </div>
              <div class="panel-content">
                
			 <div class="my-information">
              <ul class="nav nav-tabs" id="myFiTab" role="tablist">
                <li class="nav-item" role="presentation">
                  
				<%
				Call LoadTasks()
				
				%>
                  
                
                </div>
              </div>
            </div>
		
		 </div>
            </div>
			
          </div>
          </div>
        </div>
      </div>
    </main>

	
    <script>
     $(".ActionButton1").on("show.bs.dropdown", function(event){
	  var x = $(event.relatedTarget).text(); // Get the text of the element
	  alert(x);
	});
    </script>
	
	
<!-- #Include file=CAPSFooter.asp -->
  </body>
</html>

<%

Public Sub LoadTasks()

Dim y
Dim strLink
Dim strTaskType
Dim strAction
Dim strCreditLimit
Dim dteExiryDate
Dim strTaskTime
Dim strTaskName
Dim strTaskDescription
Dim strStatus
Dim dteDateRun
Dim strDateRunFormat
Dim strRunBy
Dim strImports
Dim strActionOptions
Dim lngTaskLogID
Dim strWhere
Dim arrTasks(6,50)
Dim strTaskStatus
Dim strAdminTaskDate
Dim arrActive(1)
Dim strRecordCount
Dim strSQL
Dim intTotalTasks

'First get the list of tasks as this is fixed, then run through the tasks which have been completed for the day
'This is to avoid having a complex join, calendar table or multiple recordsets
strSQL = "SELECT * FROM qryCAPSTaskLog WITH(NOLOCK) WHERE [AdminTaskDate] = '" & Session("BatchDate") & "' ORDER BY [AdminTaskID]"
'response.write strSQL
	objRS.Open strSQL,objCon,0,1
	
		'Get the Task list for the current day		
		Do Until objRS.EOF
		
			y = y + 1
			
			arrTasks(0,y) = objRS("AdminTaskID")
			arrTasks(1,y) = objRS("TaskStatus")
			arrTasks(2,y) = objRS("AdminTaskDate")
			arrTasks(3,y) = objRS("FName") & " " & objRS("LName")
			arrTasks(5,y) = objRS("RecordCount")
			arrTasks(6,y) = objRS("TaskLogID")
			
			'Determine if the Task Log Item should be displayed (further down in the Full Task list write section) based on the View Button clicked and the Task Status for the selected day
			If strViewButton = "All" Then
				strWhere = ""
				arrTasks(4,y) = "Y"
			ElseIf strViewButton = "Complete" Then
				
				If arrTasks(1,y) = "Complete" Then
					arrTasks(4,y) = "Y"
				End If
			ElseIf strViewButton = "Incomplete" Then
				If arrTasks(1,y) = "Not Complete" OR IsNull(arrTasks(1,y)) Then
					arrTasks(4,y) = "Y"
				End If
			End If

		objRS.Movenext
		
		Loop

	objRS.Close

	'Set the Total Task Complete for the date selected
	intTotalTasks = y

'Get the currently selected Time Period to load that tab of Task Data
If IsNull(Session("TaskTime")) or Session("TaskTime") = "" Then

Else
	strWhere = strWhere & " AND [TaskTime] = '" & Session("TaskTime") & "'"
End If

'Determine which tab is displayed as active
If Session("TaskTime") = "Morning" Then
	arrActive(0) = "active"
	arrActive(1) = ""
Else
	arrActive(0) = ""
	arrActive(1) = "active"
End If
'aria-selected=""true""

'Write the Start of the tabs (headers and active)
Response.Write "<a class=""nav-link " & arrActive(0) & """ id=""my-cards-tab"" href=""AdminTasks.asp?TaskTime=Morning"">Opening Tasks</a></li>" & _
		"<li class=""nav-item"" role=""presentation""><a class=""nav-link " & arrActive(1) & """ id=""my-applications-tab"" href=""AdminTasks.asp?TaskTime=Afternoon"" >Closing Tasks</a></li>" & _
		"</ul><div class=""tab-content"" id=""myFiTabContent""><div class=""tab-pane fade show active"" id=""Morning"" role=""tabpanel"" aria-labelledby=""my-cards-tab"">"
				
'Open the master recordset
'If strTaskTime = "" OR ISNull(strTaskTime) Then
'	strSQL = "SELECT * FROM tblCAPSAdminTasks WITH(NOLOCK) WHERE TaskTime = '" & strTaskTime & "'"
'Else
	strSQL = "SELECT * FROM qryCAPSAdminTaskLog WITH(NOLOCK) WHERE TaskGroup = '" & Session("TaskGroup") & "' AND Active = 'Y' " & strWhere & " ORDER BY [SortOrder]"
'End If

'response.write strSQL

	objRS.Open strSQL,objCon,0,1
	
	'Get the TaskTime to write the headings for each tab
	If Not ObjRS.EOF Then
		If objRS("TaskTime") = "" OR IsNull(objRS("TaskTime")) Then
			strTaskTime = ""
		Else
			strTaskTime = objRS("TaskTime")
		End If
	End If
	
	y = 0
	
	Do Until objRS.EOF
		
		y = y + 1
		
		If arrTasks(4,y) <> "Y" Then
		
		'First Loop through the array from the table tblCAPSTaskLog to to get any records which identify the task has been actioned for the currently selected date
		'And assign them to a local variable
		strTaskStatus = ""
		strAdminTaskDate = ""
		strRunBy = ""
		lngTaskLogID = 0
		
		For x = 1 to intTotalTasks
			If objRS("AdminTaskID") = arrTasks(0,x) Then
				lngTaskLogID = arrTasks(6,x) 
				strTaskStatus = arrTasks(1,x)
				strAdminTaskDate = arrTasks(2,x)
				strRunBy = arrTasks(3,x)
				strRecordCount = arrTasks(5,x)
			End If
			
			'response.write "<br>arrTasks(0,x)=" & arrTasks(0,x) & " y=" & y
			
		Next
		
		'If objRS("TaskTime") <> strTaskTime Then
		'	Response.Write "</div><div class=""tab-pane fade"" id=""" & objRS("TaskTime") & """ role=""tabpanel"" aria-labelledby=""" & objRS("TaskTime") & """>"
		'End If
		
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
			'strTaskDescription = Left(objRS("TaskDescription"),50)
			strTaskDescription = objRS("TaskDescription")
		End If
		
		If Trim(strTaskType) = "Diners" Then
			strTaskType = "<img src=""../images/diners2.png"" title=""" & strTaskDescription & """> "
		ElseIf strTaskType = "ANZ" Then
			strTaskType = "<img src=""../images/ANZ.png"" Title=""" & strTaskDescription & """> " '& strCardType
		ElseIf strTaskType = "Defence" Then
			strTaskType = "<img src=""../images/high-limit.png"" Title=""" & strTaskDescription & """> " '& strCardType
		ElseIf strTaskType = "CMS" Then
			strTaskType = "<img src=""../images/CMS.png"" Title=""" & strTaskDescription & """> " '& strCardType
		Else
			strTaskType = "<img src=""../images/high-limit.png"" Title=""" & strTaskDescription & """> "
		End If
		
		If strTaskStatus = "Complete" Then
			strTaskStatus = "<span class=""badge badge-pill badge-success"">Complete</span>"
		ElseIf strTaskStatus = "Error" Then
			strTaskStatus = "<span class=""badge badge-pill badge-danger"">Error</span>"
		Else
			strTaskStatus = "<span class=""badge badge-pill badge-warning"">Not Run</span>"
		End If
		
		If IsNull(strAdminTaskDate) or strAdminTaskDate = "" Then
			dteDateRun = "Not Run"
			strDateRunFormat = "font-style:italic; color:grey;"
		Else
			dteDateRun = strAdminTaskDate
			If dteDateRun < now() -10 then
				strDateRunFormat = "color:red; font-weight:bold;"
			Else
				strDateRunFormat = "color:black; font-weight:bold;"
			End If
		End If
		
		'If objRS("UserName") = "" OR IsNull(objRS("UserName")) Then
		'	strRunBy = ""
		'Else
		'	strRunBy = objRS("UserName")
		'End If
		
		'If objRS("RecordCount") = "" OR IsNull(objRS("RecordCount")) Then
		'	strImports = ""
		'Else
		'	strImports = objRS("RecordCount")
		'End If
		
		'If IsNull(objRS("TaskLogID")) or objRS("TaskLogID") = "" Then
		'	lngTaskLogID = 0
		'Else
		'	lngTaskLogID = objRS("TaskLogID")
		'End If
		
		strActionOptions = "<a class=""dropdown-item"" href=""AdminTasks.asp?Task=Complete&TaskID=" & objRS("AdminTaskID") & "&TaskLogID=" & lngTaskLogID & "#TaskSection" & objRS("SortOrder") - 1& """>Task Complete</a>" & _
				  "<a class=""dropdown-item"" href=""AdminTasks.asp?Task=NotComplete&TaskID=" & objRS("AdminTaskID") & "&TaskLogID=" & lngTaskLogID & "#TaskSection" & objRS("SortOrder") - 1& """>Not Complete</a>" & _
				  "<a class=""dropdown-item"" href=""AdminTasks.asp?Task=Error&TaskID=" & objRS("AdminTaskID") & "&TaskLogID=" & lngTaskLogID & "#TaskSection" & objRS("SortOrder") - 1 & """>Error</a>" & _
				  "<a class=""dropdown-item"" href=""AdminTasks.asp?Task=Reprocess&TaskID=" & objRS("AdminTaskID") & "&TaskLogID=" & lngTaskLogID & "#TaskSection" & objRS("SortOrder") - 1 & """>Re-Process</a>"
		
		'If there is a URL then add that as a seleciton for the Action Item
		If objRS("URL") = "" Or IsNull(objRS("URL")) Then
		Else
			strActionOptions = strActionOptions & "<a class=""dropdown-item"" href=""" & objRS("URL") & """>Go To Task...</a>"
		End If
		
		strAction = " <div class=""dropdown""><button class=""btn btn-outline-primary dropdown-toggle"" type=""button"" id=""ActionButton" & y  & """ data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"" href=""#"">" & _
					"Action</button><div class=""dropdown-menu"" aria-labelledby=""ActionButton" & y  & """>" & _
					strActionOptions & _
					"</div></div>"
		
		'strAction = " <div class=""dropdown""><button class=""btn btn-outline-primary dropdown-toggle"" type=""button"" id=""ActionButton" & y  & """ onClick=""OpenSs();"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false"">" & _
		'			"Action</button><div class=""dropdown-menu"" aria-labelledby=""ActionButton" & y  & """>" & _
		'			strActionOptions & _
		'			"</div></div>"
		
		Response.Write "<div class=""panel panel-shadow mb-3""><div class=""panel-content row"" id=""TaskSection" & objRS("SortOrder") & """>" & _
			"<div class=""col-md-1 text-center my-auto"">" & strLink & "" & strTaskType & " </div>" & _
			"<div class=""col-md-3""><h6>" & objRS("SortOrder") & ". " & strTaskName & "</h6> " & strTaskStatus & "</div>"  &_
			"<div class=""col-md-2""><span class=""title"">Date Run</span><p><span class=""description"" style=""" & strDateRunFormat & """>" & dteDateRun & "</span></p></div>"  &_
			"<div class=""col-md-2""><span class=""title"">" & objRS("TaskType") & "</span><p><span class=""description""><strong>" & strRecordCount & "</strong></span></p></div>"  &_
			"<div class=""col-md-2""><span class=""title"">Run By</span><p><span class=""description""><strong>" & strRunBy & "</strong></span></p></div>"  &_
			"<div class=""col-md-2 text-right my-auto"">" & strAction & " <i class=""fa fa-arrow-cog""></i></a></div></div></div>"
		

		'Response.Write "<div class=""panel panel-shadow mb-3""><div class=""panel-content row"">" & _
			'"<div class=""col-md-1 text-center my-auto"">" & strLink & "" & strTaskType & "</div>" & _
			'"<div class=""col-md-4""><h6>" & strTaskName & " " & strAction & _
			'"</h6><p>Number: <strong>" & strCardNo & "</strong></p></div>" & _
			'"<div class=""col-md-4""><p>Limit: <strong>" & strCreditLimit & "</strong></p></div>" & _
			'"<div class=""col-md-3 text-right my-auto"">" & strLink & "View Details <i class=""fa fa-arrow-right""></i></a></div></div></div>"		
		
		End If
		
		objRS.Movenext
	Loop
	
	objRS.Close
	
	Response.Write "</div>"
	
End Sub

Public Sub LoadDatePicker()

Dim dteBatchDateFormat

	If IsNull(Session("BatchDate")) or Session("BatchDate") = "" Then Session("BatchDate") = DateAdd("d", -1, Now())
	
	dteBatchDateFormat = Day(Session("BatchDate")) & "-" & MonthName(Month(Session("BatchDate"))) & "-" & Year(Session("BatchDate"))
	dteBatchDateFormat = Year(Session("BatchDate")) & "-" & Month(Session("BatchDate")) & "-" & Day(Session("BatchDate"))
	
	dteBatchDateFormat = Year(Session("BatchDate")) & "-" & Right("0" & Month(Session("BatchDate")), 2) & "-" & Right("0" & Day(Session("BatchDate")), 2)

	'objRS.Open "SELECT TOP 6 * FROM tblCAPSFileLoad WITH(NOLOCK) WHERE DateLoaded > '" & Session("BatchDate") & "' AND FileType = 'ROMANCostCentres' ORDER BY FileSeqNum DESC",objCon
	
	Response.Write "<input class=""DateClick2"" type=""date"" style=""color:blue;"" value=""" & dteBatchDateFormat & """ id=""AdminDate"" name=""AdminDate"" onChange=""DatePickChange();"" />"

End Sub

Public Sub LoadButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(3)

If Session("ViewButton") = "Complete" Then
	arrButton(2) = "active"
ElseIf strViewButton = "Incomplete" Then
	arrButton(3) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

	'Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				'"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='AdminTasks.asp?ViewButton=All';"">View All</button>" & _
			'	"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='AdminTasks.asp?ViewButton=Complete';"">Complete</button>" & _
			'	"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='AdminTasks.asp?ViewButton=Incomplete';"">Incomplete</button>" & _
			'	"</div>"

End Sub

Public Sub SaveTaskLog(lngAdminTaskID,strStatus,lngTaskLogID)

Dim intRecord
Dim strMessageIcon
Dim strMessageColour

  	With objCmd

		.CommandType = 4
		.CommandText = "spCAPSTaskLogSave"

		.Parameters.Append objCmd.CreateParameter("TaskLogID", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("AdminTaskID", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("FileLoadID", adInteger, adParamInput)
		.Parameters.Append objCmd.CreateParameter("AdminTaskDate", adDate, adParamInput)
		.Parameters.Append objCmd.CreateParameter("TaskStatus", adVarChar, adParamInput, 20)
		.Parameters.Append objCmd.CreateParameter("Notes", adVarChar, adParamInput, 500)
		.Parameters.Append objCmd.CreateParameter("UpdatedBy", adDouble, adParamInput) 
		.Parameters.Append objCmd.CreateParameter("TaskLogIDOutput", adInteger, adParamOutput)
		
		.Parameters("TaskLogID") = lngTaskLogID
		.Parameters("AdminTaskID") = lngAdminTaskID
		.Parameters("FileLoadID") = 0
		.Parameters("AdminTaskDate") = Session("BatchDate")'Now()
		.Parameters("TaskStatus") = strStatus
		.Parameters("Notes") = ""
		.Parameters("UpdatedBy") = Session("UserID")
		.ActiveConnection = objCon
		 
	End With

   objCmd.Execute        
  
	'Return the result of the Save Function.
	intRecord = objCmd.Parameters.Item("TaskLogIDOutput") 
 
	strMessageIcon = "&nbsp;&nbsp;<img src=""../images/check-icon.png"" /> Task " & intRecord & " updated!"

	strMessageColour = "Black"
		
End Sub


Set objRS = Nothing
Set objCon = Nothing

%>