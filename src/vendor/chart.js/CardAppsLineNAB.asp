<%


Dim objCon
Dim objRS
Dim strSQL
Dim strData
Dim strWHERE
Dim arrData(5)
Dim strMonth
Dim strYear
Dim arrDay(31)
Dim strWeekdayName
Dim intDays
Dim strDays(7)
Dim strDayLabels
Dim strMonthDay
Dim arrDay2(5,31)
Dim strDayName1
Dim strFirstDayName
Dim strDayOne
Dim strVBStartDay
Dim intStartDayFrom

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")	

'Sub LoadDetails()

If IsNull(Session("ReportYear")) OR Session("ReportYear") = "" Then
	strYear = "2020"
Else
	strYear = Session("ReportYear")
End If

If IsNull(Session("ReportMonth")) OR Session("ReportMonth") = "" Then
	strMonth = PadDigits(Month(now()),2)'"11"
	Session("ReportMonth") = Left(MonthName(Month(now())),3)
Else
	Select Case(Session("ReportMonth"))
		Case "Jan"
			strMonth = "01"
		Case "Feb"
			strMonth = "02"
		Case "Mar"
			strMonth = "03"
		Case"Apr"
			strMonth = "04"
		Case "May"
			strMonth = "05"
		Case "Jun"
			strMonth = "06"
		Case "Jul"
			strMonth = "07"
		Case "Aug"
			strMonth = "08"
		Case "Sep"
			strMonth = "09"
		Case "Oct"
			strMonth = "10"
		Case "Nov"
			strMonth = "11"
		Case "Dec"
			strMonth = "12"
		Case Else
			strMonth = "01"
	End Select
	'strMonth = PadDigits(Session("ReportMonth"),2)
End If


strDays(1) = "Mon"
strDays(2) = "Tue"
strDays(3) = "Wed"
strDays(4) = "Thu"
strDays(5) = "Fri"
strDays(6) = "Sat"
strDays(7) = "Sun"

'Get the First day NAME of the month
strFirstDayName = dateserial(year(date()),strMonth,01)
'response.write weekday(strFirstDayName, vbMonday) & " " & strFirstDayName
strDayOne = weekday(strFirstDayName, vbMonday)

strMonthFirstDayName = Left(FormatDateTime(strFirstDayName,1),3)

intDays = 0
intStartDayFrom = 0


'If the first day of the selected month is Saturday then add 2 days to the date and increment the date to the Monday (+2 days from Sat)
If strMonthFirstDayName = "Sat" Then 
	strMonthFirstDayName = "Mon"
	strFirstDayName = DateAdd("d",2,strFirstDayName)
	intStartDayFrom = 2
End If

'If the first day of the selected month is Sunday then add 1 day to the date and increment the date to the Monday (+1 days from Sat)
If strMonthFirstDayName = "Sun" Then 
	strMonthFirstDayName = "Mon"
	strFirstDayName = DateAdd("d",1,strFirstDayName)
	intStartDayFrom = 1
End If

'Increment the weekday name depending on when the first day of the month is
If strMonthFirstDayName = "Tue" Then intDays = 1
If strMonthFirstDayName = "Wed" Then intDays = 2
If strMonthFirstDayName = "Thu" Then intDays = 3
If strMonthFirstDayName = "Fri" Then intDays = 4


'Get the month days depending on whether the month is in the past or the current month
'Do this by adding 1 to the month selected then removing one day from the date to give the last day of the previous (current) month

	'response.write "strMonth=" & strMonth & " now=" &  Month(Now()) & " "
If cstr(strMonth) = cstr(Month(Now())) Then
	'strMonthDay = GetDayFromNow(Now(),vbLongDate )
	strMonthDay = GetDayFromNow(now())
	strMonthDay = Trim(Left(strMonthDay,2))
	'response.write strMonthDay & " - "
Else
	strMonthDay = Day(DateAdd("d",-1,"01/" & strMonth+1 & "/" & Session("ReportYear")))
	'response.write Day(DateAdd("d",-1,"01/" & strMonth+1 & "/" & Session("ReportYear"))) & "||"
End If
'response.write strMonthDay & " - " 
'response.write GetDayFromNow(now())
'response.write GetDayFromNow(now())

'Create the Week days
'For x = strDayOne to strMonthDay'Day(Now())
For x = 1 + intStartDayFrom to strMonthDay'Day(Now())

	intDays = intDays + 1

	If intDays=8 Then
		intDays = 1
	End If
		
		arrDay(x) = strDays(intDays)
	
		strDayLabels = strDayLabels & "'" & arrDay(x) & " " & x & "',"
	
Next 

strDayLabels = Left(strDayLabels,Len(strDayLabels)-1)
'response.write strDayLabels
strWhere = strYear & strMonth
'strWHERE = "112020"
'response.write strWhere

strSQL = "SELECT COUNT([EmployeeID]) As Apps,CardType,[ApplicationTypeName],[CardTypeSub],CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) As MonthYear " & _
	"FROM tblCAPSApplication WITH(NOLOCK) " & _
	"WHERE CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) = '" & strWHERE & "' " & _
	"GROUP BY [CardType],[ApplicationTypeName],[CardTypeSub],CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6))"

strSQL = "SELECT COUNT([EmployeeID]) As Apps,CardType,[ApplicationTypeName],[CardTypeSub],CAST(CAST(YEAR(DateSubmitted) as varchar(4)) As Varchar(8)) + '' + CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(RIGHT('00'+CAST(Day(DateSubmitted) as varchar(2)),2) as varchar(2)) As YearMonthDay, " & _
	"CAST(Day(DateSubmitted) as varchar(2)) As DayName1 " & _
	"FROM tblCAPSApplication WITH(NOLOCK)  " & _
	"WHERE CAST(CAST(YEAR(DateSubmitted) as varchar(4)) As Varchar(8)) + '' + CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) = '" & strWHERE & "' " & _
	"GROUP BY [CardType],[ApplicationTypeName],[CardTypeSub],CAST(CAST(YEAR(DateSubmitted) as varchar(4)) As Varchar(8)) + '' + CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(RIGHT('00'+CAST(Day(DateSubmitted) as varchar(2)),2) as varchar(2)), " & _
		"CAST(Day(DateSubmitted) as varchar(2)) " & _
	"ORDER BY CAST(CAST(YEAR(DateSubmitted) as varchar(4)) As Varchar(8)) + '' + CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(RIGHT('00'+CAST(Day(DateSubmitted) as varchar(2)),2) as varchar(2)) "

'response.write strSQL
   'Description:	Loads Position details into page if applicable.
	objRS.Open strSQL,objCon

		'arrData(1) = 0
		'arrData(2) = 0
		'arrData(3) = 0
		'arrData(4) = 0
		
		If Not objRS.EOF Then
			
			Do Until objRS.EOF
				strDayName1 = objRS("DayName1")
				
				If objRS("ApplicationTypeName") = "DTC NAB CiH" Then
					'arrData(1) =  arrData(1) + objRS("Apps")
					arrDay2(1,strDayName1) = objRS("Apps")'arrDay2(1,strDayName1) & "," & objRS("Apps")
					'response.write ", " & strDayName1 & " - " & objRS("Apps")
				End If
				
				If objRS("ApplicationTypeName") = "DTC NAB Lodge" Then
					'arrData(2) =  arrData(2) + objRS("Apps")
					arrDay2(2,strDayName1) = objRS("Apps")'arrDay2(2,strDayName1) & "," & objRS("Apps")
				End If
				
				If objRS("ApplicationTypeName") = "DTC NAB Dual" Then
					'arrData(2) =  arrData(2) + objRS("Apps")
					arrDay2(3,strDayName1) = objRS("Apps")'arrDay2(2,strDayName1) & "," & objRS("Apps")
				End If
				
				If objRS("CardTypeSub") = "NAB DPC" Then
					'arrData(3) =  arrData(3) + objRS("Apps")
					arrDay2(4,strDayName1) = objRS("Apps")'arrDay2(3,strDayName1) & "," & objRS("Apps")
				End If
				
				If Right(objRS("ApplicationTypeName"),12) = "Limit Change" Then
					arrDay2(5,strDayName1) = objRS("Apps")
					'arrData(5) =  arrData(5) + objRS("Apps")
				End If
				
				'If objRS("CardTypeSub") = "CTS" Then
				'	'arrData(4) =  arrData(4) + objRS("Apps")
				'	arrDay2(5,strDayName1) = objRS("Apps")'arrDay2(4,strDayName1) & "," & objRS("Apps")
				'End If
				
				'strData = strData & "," & objRS("Apps")
			
				objRS.Movenext
			Loop
		Else
			'strData = ","
	   End If
	   
	objRS.Close
	
	
	For x = 1 + intStartDayFrom to strMonthDay
		
		If arrDay2(1,x) = "" Or IsNull(arrDay2(1,x)) Then 
			arrDay2(1,x) = 0
		Else
			'arrDay2(1,x) = x
		End If
		
		If arrDay2(2,x) = "" Or IsNull(arrDay2(2,x)) Then 
			arrDay2(2,x) = 0
		Else
			'arrDay2(2,x) = x
		End If
		
		If arrDay2(3,x) = "" Or IsNull(arrDay2(3,x)) Then 
			arrDay2(3,x) = 0
		Else
			'arrDay2(3,x) = x
		End If
		
		If arrDay2(4,x) = "" Or IsNull(arrDay2(4,x)) Then 
			arrDay2(4,x) = 0
		Else
			'arrDay2(4,x) = x
		End If
		
		If arrDay2(5,x) = "" Or IsNull(arrDay2(5,x)) Then 
			arrDay2(5,x) = 0
		Else
			'arrDay2(4,x) = x
		End If
		
		'Get the Weekday Number (1 = Monday)
		strWeekdayName = Weekday(x, vbMonday)
		
		'strWeekdayName = Weekday(1+x, vbMonday)
		'strWeekdayName = Weekday(Now()+x, vbMonday)
		
		'If the weekday number is less than Saturday (6) then get the name
		'If strWeekdayName < 6 Then

			arrData(1) = arrData(1) & "," & arrDay2(1,x)
			arrData(2) = arrData(2) & "," & arrDay2(2,x)
			arrData(3) = arrData(3) & "," & arrDay2(3,x)
			arrData(4) = arrData(4) & "," & arrDay2(4,x)
			arrData(5) = arrData(5) & "," & arrDay2(5,x)
		'End If
	Next
	
	If len(arrData(1)) > 1 Then arrData(1) = Right(arrData(1),Len(arrData(1)) -1)
	If len(arrData(2)) > 1 Then arrData(2) = Right(arrData(2),Len(arrData(2)) -1)
	If len(arrData(3)) > 1 Then arrData(3) = Right(arrData(3),Len(arrData(3)) -1)
	If len(arrData(4)) > 1 Then arrData(4) = Right(arrData(4),Len(arrData(4)) -1)
	If len(arrData(5)) > 1 Then arrData(5) = Right(arrData(5),Len(arrData(5)) -1)
	'strData = arrData(1) & "," & arrData(2) & "," & arrData(3) & "," & arrData(4)
	'strData = Right(strData,Len(strData)-1)
	
	'response.write arrData(1)
	
'End Sub

Function PadDigits(val, digits)
  PadDigits = Right(String(digits,"0") & val, digits)
End Function


Function GetDayFromNow(str)
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
	
	'MediumDate = aDay & "-" & aMonth & "-" & aYear
	GetDayFromNow = aDay & "-" & aMonth 
End Function


%>

<!doctype html>
<html>

<head>
	<title>CAPS Card Apps Line Chart</title>
	<script src="Chart.min.js"></script>
	<script src="utils.js"></script>
	<style>
	canvas {
		-moz-user-select: none;
		-webkit-user-select: none;
		-ms-user-select: none;
	}
	</style>
</head>

<body>
	<div id="container" style="width: 100%;">
		<canvas id="line-chart" width="800" height="450"></canvas>
	</div>
	
	<script>
		new Chart(document.getElementById("line-chart"), {
  type: 'line',
  data: {
    labels: [<%=strDayLabels%>],
    datasets: [{ 
        data: [<%=arrData(1)%>],
        label: "DTC CiH",
        borderColor: "#3e95cd",
        fill: false
      }, { 
        data: [<%=arrData(2)%>],
        label: "Lodge",
        borderColor: "#8000FF",
        fill: false
      }, { 
        data: [<%=arrData(3)%>],
        label: "Dual",
        borderColor: "#31B404",
        fill: false
      },{ 
        data: [<%=arrData(4)%>],
        label: "DPC",
        borderColor: "#FF4000",
        fill: false
      }, { 
        data: [<%=arrData(5)%>],
        label: "Credit Limit",
        borderColor: "#fdb900",
        fill: false
      }, 
    ]
  },
  options: {
    title: {
      display: true,
      text: 'Card Applications for <%=Session("ReportMonth")%>'
    }
  }
});
	</script>
</body>

</html>
