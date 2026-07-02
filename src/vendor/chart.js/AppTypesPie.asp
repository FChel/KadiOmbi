<%


Dim objCon
Dim objRS
Dim strSQL
Dim strData
Dim strWHERE
Dim arrData(4)
Dim strMonth
Dim strYear
Dim arrPercent(4)

Set objCon = Server.CreateObject("ADODB.Connection")
Set objRS = Server.CreateObject("ADODB.Recordset")

objCon.Open Session("DBConnection")	


If IsNull(Session("ReportYear")) OR Session("ReportYear") = "" Then
	strYear = "2020"
Else
	strYear = Session("ReportYear")
End If

If IsNull(Session("ReportMonth")) OR Session("ReportMonth") = "" Then
	strMonth = "11"
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

strWhere = strMonth & strYear
'strWHERE = "112020"

'Sub LoadDetails()

'strWHERE = "112020"

strSQL = "SELECT COUNT([EmployeeID]) As Apps,ApplicationType,CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) As MonthYear " & _
	"FROM tblCAPSApplication WITH(NOLOCK) " & _
	"WHERE CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)) = '" & strWHERE & "' " & _
	"GROUP BY CAST(CAST(RIGHT('00'+CAST(Month(DateSubmitted) as varchar(2)),2) as varchar(2)) + '' + CAST(YEAR(DateSubmitted) as varchar(20)) AS Char(6)),ApplicationType"

   'Description:	Loads Position details into page if applicable.
	objRS.Open strSQL,objCon

		arrData(1) = 0
		arrData(2) = 0
		arrData(3) = 0

		If Not objRS.EOF Then
			
			Do Until objRS.EOF
				
				If objRS("ApplicationType") = "AE602 XML" Then
					arrData(1) =  arrData(1) + objRS("Apps")
				End If
				
				If objRS("ApplicationType") = "AE602 SC" Then
					arrData(2) =  arrData(2) + objRS("Apps")
				End If
				
				If objRS("ApplicationType") = "Portal" Then
					arrData(3) =  arrData(3) + objRS("Apps")
				End If

				strData = strData & "," & objRS("Apps")
			
				objRS.Movenext
			Loop
		Else
			strData = ","
	   End If
	   
	objRS.Close
	
	strData = arrData(1) & "," & arrData(2) & "," & arrData(3) & "," & arrData(4)
	'strData = Right(strData,Len(strData)-1)
	
	'response.write strData
	
	If IsNull(arrData(1)) Then arrData(1) = 0
	If IsNull(arrData(2)) Then arrData(2) = 0
	If IsNull(arrData(3)) Then arrData(3) = 0
	If IsNull(arrData(4)) Then arrData(4) = 0
	
	'Get the percentage of each of the pie portions (there are currently only 2) for siplay
	arrPercent(1) = CInt(arrData(1)) + CInt(arrData(2))	+ CInt(arrData(3)) + Cint(arrData(4))
	arrPercent(2) = arrPercent(1)
	arrPercent(3) = arrPercent(1)

	If CInt(arrPercent(1))=0 Then
		arrPercent(1) = 0
	Else
		arrPercent(1) = (CInt(arrData(1))/arrPercent(1) ) * 100
	End If
	
	If CInt(arrPercent(2))=0 Then
		arrPercent(2) = 0
	Else
		arrPercent(2) = (CInt(arrData(2))/arrPercent(2) ) * 100
	End If
	
	If CInt(arrPercent(3))=0 Then
		arrPercent(3) = 0
	Else
		arrPercent(3) = (CInt(arrData(3))/arrPercent(3) ) * 100
	End If

	'Round the Percentage to 2 decomals
	arrPercent(1) = Round(arrPercent(1),2)
	arrPercent(2) = Round(arrPercent(2),2)
	arrPercent(3) = Round(arrPercent(3),2)

'End Sub


%>
<!doctype html>
<html>

<head>
	<title>Application Types Chart</title>
	<script src="Chart.min.js"></script>
	<script src="utils.js"></script>
	<style>
	canvas {
		-moz-user-select: none;
		-webkit-user-select: none;
		-ms-user-select: none;
	}
	</style>
	
<script>


</script>
</head>

<body>

	<div id="canvas-holder" style="width:100%; height:100%;">
		<canvas id="chart-area" style="width:100%; height:100%; align-content: left;"></canvas>
	</div>
	
	<script>
		var randomScalingFactor = function() {
			return Math.round(Math.random() * 100);
		};

		var config = {
			type: 'pie',
			data: {
				datasets: [{
					data: [
						<%=strData%>
					],
					backgroundColor: [
						'#009900','#ff9900','#4287f5',
						
						
					]
				}], 
			labels: [
        'AE602 XML (' + <%=arrPercent(1)%> + '%)',
        'Service Connect (' + <%=arrPercent(2)%> + '%)',
	'Portal (' + <%=arrPercent(3)%> + '%)'
    ]
				
			},
			options: {
				responsive: false,

				title: {
					display: true,
					text: 'Application Types'
				},
				animation: {
					animateScale: true,
					animateRotate: true
				},
				pointLabel: {
					display: false,
					label: 'Applications'
				},
				legend: { position: 'right' },
				
				 plugins: {
            datalabels: {
                formatter: (value, ctx) => {
                
                  let sum = 0;
                  let dataArr = ctx.chart.data.datasets[0].data;
                  dataArr.map(data => {
                      sum += data;
                  });
                  let percentage = (value*100 / sum).toFixed(2)+"%";
                  return percentage;

              
                },
                color: '#fff',
                     },
					 }
			}
		};

		window.onload = function() {
			var ctx = document.getElementById('chart-area').getContext('2d');
			window.myDoughnut = new Chart(ctx, config);
		};

		document.getElementById('randomizeData').addEventListener('click', function() {
			config.data.datasets.forEach(function(dataset) {
				dataset.data = dataset.data.map(function() {
					return randomScalingFactor();
				});
			});

			window.myDoughnut.update();
		});

		var colorNames = Object.keys(window.chartColors);
		document.getElementById('addDataset').addEventListener('click', function() {
			var newDataset = {
				backgroundColor: [],
				data: [],
				label: 'New dataset ' + config.data.datasets.length,
			};

			for (var index = 0; index < config.data.labels.length; ++index) {
				newDataset.data.push(randomScalingFactor());

				var colorName = colorNames[index % colorNames.length];
				var newColor = window.chartColors[colorName];
				newDataset.backgroundColor.push(newColor);
			}

			config.data.datasets.push(newDataset);
			window.myDoughnut.update();
		});

		document.getElementById('addData').addEventListener('click', function() {
			if (config.data.datasets.length > 0) {
				config.data.labels.push('data #' + config.data.labels.length);

				var colorName = colorNames[config.data.datasets[0].data.length % colorNames.length];
				var newColor = window.chartColors[colorName];

				config.data.datasets.forEach(function(dataset) {
					dataset.data.push(randomScalingFactor());
					dataset.backgroundColor.push(newColor);
				});

				window.myDoughnut.update();
			}
		});

		document.getElementById('removeDataset').addEventListener('click', function() {
			config.data.datasets.splice(0, 1);
			window.myDoughnut.update();
		});

		document.getElementById('removeData').addEventListener('click', function() {
			config.data.labels.splice(-1, 1); // remove the label first

			config.data.datasets.forEach(function(dataset) {
				dataset.data.pop();
				dataset.backgroundColor.pop();
			});

			window.myDoughnut.update();
		});

		document.getElementById('changeCircleSize').addEventListener('click', function() {
			if (window.myDoughnut.options.circumference === Math.PI) {
				window.myDoughnut.options.circumference = 2 * Math.PI;
				window.myDoughnut.options.rotation = -Math.PI / 2;
			} else {
				window.myDoughnut.options.circumference = Math.PI;
				window.myDoughnut.options.rotation = -Math.PI;
			}

			window.myDoughnut.update();
		});
		
	
	
	</script>
</body>

</html>