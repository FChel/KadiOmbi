
<!-- #Include file=../CC/CAPSHeader.asp -->
<!-- #Include file=../ADOVBS.inc -->
<!-- #Include file=../CC/CAPSFunctions.asp -->
<%

If IsEmpty(Session("UserID")) Then Response.Redirect("../Timeout.asp")

Response.ContentType = "text/html" 
Response.CodePage = 65001
Response.CharSet = "UTF-8"

'Description:	Training Course detail report/screen for training requirement
'Author:		Michael Giacomin
'Date:			January 2021

	Response.Expires = -1500	

Dim objCon
Dim objRS
Dim objRS1
Dim objRS2
Dim objCmd

Dim x
Dim strMessage
Dim strSelected
Dim strMessageIcon
Dim strMessageColour
Dim strSQL

Dim strWhere
Dim strSearchInput

    Set objCon = Server.CreateObject("ADODB.Connection")
    Set objRS = Server.CreateObject("ADODB.Recordset")
    Set objRS1 = Server.CreateObject("ADODB.Recordset")
    Set objCmd = Server.CreateObject("ADODB.Command")
	
    objCon.Open Session("DBConnection")	

    'Session("CurrentPage") = "Admin/EmailError.asp"

	If Not IsEmpty(Request.QueryString("PageCombo")) Then
		Session("PageCombo") = Request.QueryString("PageCombo")
	End If
	
	If Not IsEmpty(Request.QueryString("ViewButton")) Then
		Session("CATViewButton") = Request.QueryString("ViewButton")
	End If
	
	'Get the Group Name Selected
	If Not IsEmpty(Request.QueryString("GroupName")) Then
		Session("GroupNameTraining") = Request.QueryString("GroupName")
	End If
	
	'Set the default Card type to DTC
	If IsEmpty(Session("ViewExtraButton")) Then Session("ViewExtraButton") = "DTC"
	
	If IsEmpty(Session("CATViewButton")) Then Session("CATViewButton") = "NotExported"
	
	'Execute Action
	If Request.QueryString("Action") = "Delete" Then   
		
		Call DeleteData(Request.QueryString("EmployeeID"))
	End If
	
	'Get the Card Account Transfer Account Type Selected
	If Not IsEmpty(Request.QueryString("StatusViewButton")) Then
		Session("StatusViewButton") = Request.QueryString("StatusViewButton")
	End If
	
	'Normal transfer for selected AHTo --Called/String passed from the Modal
	If Request.QueryString("Action") = "Transfer" Then   
		Call TransferAccount(Request.QueryString("TransferComments"))		
	End If
	
	'Return transfer for Database sourced AHTo --Called/String passed from the Modal
	If Request.QueryString("Action") = "TransferReturn" Then   
		Call TransferAccountReturn(Request.QueryString("TransferComments"))		
	End If
	
	'Call the Delete (change Exported) value if the delete button is clicked for a record
	If Request.QueryString("Action") = "RemoveCAT" Then   
		Call DeleteTransfer(Request.QueryString("Type"),Request.QueryString("CardAccountTransferID"))		
	End If
	
	'Clear the selected AH To
	If Request.QueryString("Action") = "ClearAHTo" Then   
		Session("CATEmployeeID") = ""
	End If
	
	If Request.QueryString("Action") = "Search" Then   
		
		If IsNull(Request.QueryString("EmployeeID")) OR Request.QueryString("EmployeeID")= "" Then
			Session("CATEmployeeID") = ""
		Else
			Session("CATEmployeeID") = Request.QueryString("EmployeeID")
		End If
		
	End If

	'Clear the Selected Now selections
	If Request.QueryString("Action") = "ClearSelected" Then   
		Call ClearTrainingAction()
	End If
	
	'Get any search input value to use when clicking buttons so the Employee search is not lost
	If Not IsNull(Request.QueryString("SearchInput")) Then
		strSearchInput = Request.QueryString("SearchInput")
	End IF
	
	'Make sure there is a (0) zero CATEmployeEID 
	If IsNull(Session("CATEmployeeID")) OR Session("CATEmployeeID") = "" Then Session("CATEmployeeID") = 0
%>

<html>
<head>
<script LANGUAGE="javascript">

	setTimeout( 'ShowTimeoutWarning();', 1080000 );

function ShowTimeoutWarning () {     
    window.alert( "********** Warning! **********' \n \n 'You will be automatically logged out in 2 minutes unless you change screens, Close or Save!" ); 
}

function DeleteModalClose(cb) {
   
	document.getElementById("ModalDelete").style.display = "none";
        
}  

function SearchUsers() {
	
	var id = document.getElementById('SearchInput').value
	self.location = "CardAccountTransfer.asp?SearchInput=" + id;

}

function OpenExcelReport() {
	var strExcel = document.getElementById('WhereClause').value;
	
	window.open('../CC/ExcelExport.asp?' + strExcel + '')
	//window.open('../CC/ExcelExport.asp?tbl=qryCAPSTrainingReportGroupExport&W=' + strExcel + '')
	//window.open('../CC/ExcelExport.asp?tbl=qryCAPSTrainingReportGroupExport&W=' + strExcel + '&Top=100')
	//window.open('../CC/ExcelExport.asp?tbl=qryCAPSTrainingReportGroup&Top=100')

}

function loadDocE(cb) {

  var id = cb.getAttribute('data-id');
  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("GetUserDetail").innerHTML = this.responseText;
    }
  };

  xhttp.open("GET", "../CC/AJAX/GetUserDetail.asp?UserID=" + id, true);
  xhttp.send();
}

function LoadFiles() {

  //var id = cb.getAttribute('data-id');
  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ModalFileMessage").innerHTML = this.responseText;
    }
  };

  xhttp.open("GET", "../CC/AJAX/GetExcelFiles.asp", true);
  xhttp.send();
}

function loadDelete(cb) {


var id = cb.getAttribute('data-id');
var name = cb.getAttribute('data-EmployeeName');

	document.getElementById("ModalDeleteMessage").innerHTML = 'Deactivate User - ' + name + '?';
	document.getElementById("ModalDelete").style.display = "block";
	document.getElementById("EmployeeDeleteID").value = id;
	
}

function LoadReview() {
//console.log('assa');
//	const chks = document.querySelectorAll(".PrcChecks");
	
	//var chek = document.getElementById(
//	for (var i = 0; i < chks.length; i++) {
		//if (chks.checked==false) {
//	  console.log('chks: ', chks[i]);
	  //}
//	}
	
//	for (const PrcCheck of chks.values()) {
//		console.log('PrcCheck: ', PrcCheck.value());
//	}
//	var i=0;
//	var markedCheckbox = document.getElementsByName('Chk');  
 // for (var checkbox of markedCheckbox) {  
 //   if (checkbox.checked)  
 //     console.log(checkbox.value + ' ');  
//	  i++;
	  //console.log(i + ' ~');
 // }  
//	document.getElementById("ModalDeleteMessage").innerHTML = 'Employees Selected ' + i;

var inste = document.querySelectorAll('input[type="checkbox"]:checked').length;
//alert(inste);
document.getElementById("ModalReviewDetailTop").innerHTML = 'Employees Selected ' + inste;


//var id = cb.getAttribute('data-id');
  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ModalReviewDetail").innerHTML = this.responseText;
    }
  };

  //xhttp.open("GET", "../CC/AJAX/GetCardAccountTransfer.asp?EmployeeID=" + id, true);
  xhttp.open("GET", "../CC/AJAX/GetCardAccountTransfer.asp?EmployeeID=", true);
  xhttp.send();


}

function LoadReview2(cb) {


//var id = cb.getAttribute('data-id');
//var name = cb.getAttribute('data-EmployeeName');

	document.getElementById("ModalReviewDetail").innerHTML = 'Training Email Review?';
	//document.getElementById("ReviewModalDelete").style.display = "block";
	//document.getElementById("EmployeeDeleteID").value = id;
	
	//if (this.checked) {
	//	   $(':checkbox').each(function() {
	//	document.getElementById("ModalReviewDetail").innerHTML = 1;
	//	}
	//}
	
	var checked = 0;
 
        //Reference the Table.
        //var tblFruits = document.getElementById("tblFruits");
 
        //Reference all the CheckBoxes in Table.
        //var chks = tblFruits.getElementsByTagName("INPUT");
 
        //Loop and count the number of checked CheckBoxes.
        //for (var i = 0; i < chk.length; i++) {
		for (var i = 0; i < 10; i++) {
            if (document.getElementById("chk[i]").checked) {
                checked++;
            }
        }
 
        if (checked > 0) {
            alert(checked + " CheckBoxe(s) are checked.");
            return true;
        } else {
            alert("Please select CheckBoxe(s).");
            return false;
        }
}

function deleteEmployee(cb) {

	var id = document.getElementById("EmployeeDeleteID").value

	self.location = "CardAccountTransfer.asp?Action=Delete&EmployeeID=" + id;

}

function SaveConfirmTransfer() {

	var id = document.getElementById("TransferComments").value;
	var rtnType = document.getElementById("ReturnType").value;
	
	if(rtnType=='Return') {
		self.location = "CardAccountTransfer.asp?Action=TransferReturn&TransferComments=" + id;
	} else {
		self.location = "CardAccountTransfer.asp?Action=Transfer&TransferComments=" + id;
	}
}

function UpdateTraining(cb) {

	var id = document.getElementById("dteTrain").value
	//var EmailID = document.getElementById("SelEmail").selectedIndex;
	
	var e = document.getElementById("SelEmail");
	var EmailID = e.options[e.selectedIndex].value;
	
	self.location = "CardAccountTransfer.asp?Action=UpdateTraining&ExemptDate=" + id + '&EmailTemplateID='+EmailID;

}

function ChangeGroup(cb) {
	
	var e = document.getElementById("GroupSelect");
	var GroupID = e.options[e.selectedIndex].value;
	
	self.location = "CardAccountTransfer.asp?GroupName=" + GroupID;

}

function SaveCATAction(cb) {

  var id = cb.getAttribute('data-id');
  var Courseid = cb.getAttribute('data-CourseID');
  var varSelected = cb.checked
  
  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("TrainingActionSelected").innerHTML = this.responseText;
    }
  };

  //xhttp.open("GET", "../CC/AJAX/GetCardAccountTransfer.asp?EmployeeID=" + id, true);
  xhttp.open("GET", "../CC/AJAX/SaveCardAccountTransfer.asp?EmployeeID="+id+"&CourseID="+Courseid+"&Selected="+varSelected, true);
  xhttp.send();



}

function GetTrainingAction() {
//Run the same procedure to save the Training selected, but only load the details for the current userID
  //var id = cb.getAttribute('data-id');
  //var Courseid = cb.getAttribute('data-CourseID');
  //var varSelected = cb.checked
  
  
  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("TrainingActionSelected").innerHTML = this.responseText;
    }
  };

  //xhttp.open("GET", "../CC/AJAX/GetCardAccountTransfer.asp?EmployeeID=" + id, true);
  //xhttp.open("GET", "../CC/AJAX/SaveCardAccountTransfer.asp?EmployeeID="+id+"&CourseID="+Courseid+"&Selected="+varSelected, true);
  xhttp.open("GET", "../CC/AJAX/SaveCardAccountTransfer.asp", true);
  xhttp.send();

}

function ConfirmCATAction(varReturn) {

  //var id = cb.getAttribute('data-id');
  //var Courseid = cb.getAttribute('data-CourseID');
  //var varSelected = cb.checked

  if (varReturn=='Return') {
	varAHTo='Return each Card Selected below to Cardholder as Account Holder'
  } else {
	varAHTo = document.getElementById("AHTo").innerHTML
 };
  
  document.getElementById("ModalConfirmMessageTo").innerHTML = varAHTo;
  
  var xhttp = new XMLHttpRequest();
 
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("ModalConfirmMessage").innerHTML = this.responseText;
    }
  };

  //xhttp.open("GET", "../CC/AJAX/GetCardAccountTransfer.asp?EmployeeID=" + id, true);
  xhttp.open("GET", "../CC/AJAX/SaveCardAccountTransfer.asp?Confirm=1&Return="+varReturn, true);
  xhttp.send();

	//frm.TransferComments.focus();
	
	//setTimeout(document.getElementByName("TransferComments").focus();,3000)
	setTimeout(ComFocus,1000);
}

function ComFocus() {
	document.getElementById("TransferComments").focus();
}

function ChangePage() {

	//var id = cb.getAttribute('CardTypeSelect');
	//var id = document.getElementByName("CardTypeSelect").value;
	//alert(id);
	var e = document.getElementById("PageCombo");
	var result = e.options[e.selectedIndex].value;
	
	self.location = 'CardAccountTransfer.asp?PageCombo=' + result;
	//alert(result);
	//document.getElementById('CardType').value=result;
	
}


	
$(document).ready(function(){
	$('#SelectAllCheck').click(function() {
	   if (this.checked) {
		   $(':checkbox').each(function() {
			   this.checked = true;   
			   //Run the procedure to save the selected employee selection (on)
				SaveCATAction(this);
		   });
	   } else {
		  $(':checkbox').each(function() {
			   this.checked = false; 
			   //Run the procedure to save the selected employee selection (off)
				SaveCATAction(this);
		   });
	   } 
	});
	
	//Check to see which checkboxes are selected
	$('#ProcessSelected').click(function() {
	var i=0;
	//var boxes = $(":checkbox:checked");
	//console.log(boxes +' checked');
	//alert(boxes);
	if (this.checked){
	console.log($(this).val());
	i++;
	}
	console.log(i);
	});
	
	$(".clickable-row").click(function() {
		window.location = $(this).data("href");
	});

	
});

function SelectEmp(varEmpID) {
	if(varEmpID==undefined) {
		alert(varEmpID);
	}
	{
	//self.location = "EmployeeHistory.asp?Action=Search&EmployeeID=" + varEmpID;
	self.location = "CardAccountTransfer.asp?Action=Search&EmployeeID=" + varEmpID + "&Type=AHTo"
	//self.location = "EmployeeHistory.asp?Action=Search&EmployeeID=" + varEmpID + "&FileSeqNum=" + document.getElementById('FileAnchor').text;
	}
}

function loadDoc() {
var varEmpID
var varFName
var varLName

	document.getElementById('Progress').style.display = "inline";

	if(frm.EmpIDS.value==undefined) {
		varEmpID='';}
	else {
		varEmpID=frm.EmpIDS.value
	}
	if(frm.FirstName.value==undefined) {
		varFName='';}
	else {
		varFName=frm.FirstName.value
	}
	if(frm.LNamms.value==undefined) {
		varLName='';}
	else {
		varLName=frm.LNamms.value
	}
	
  var xhttp = new XMLHttpRequest();
  xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
     document.getElementById("EmpSear").innerHTML = this.responseText;
	 document.getElementById('Progress').style.display = "none";
    }
  };
  xhttp.open("GET", "../CC/AJAX/GetEmployees.asp?EmpID=" + varEmpID + "&FName=" + varFName + "&LName=" + varLName + "", true);
  //xhttp.open("GET", "../CC/AJAX/GetEmployees.asp?EmpID=" + frm.EmpIDS.value + "&FName=" + frm.FirstName.value + "&LName=" + frm.LNamms.value + "", true);
  //xhttp.open("GET", "../CC/AJAX/GetEmployees.asp?From=CAT&EmpID=" + frm.EmpIDS.value + "&FName=" + frm.FirstName.value + "&LName=" + frm.LNamms.value + "&NameOnCard=" + frm.NonCS.value + "&SearchType=" + varSearchType + "", true);
  xhttp.send();
}

</script>

</head>
<body onLoad="GetTrainingAction();">
<!--<form name="frm">-->
	<main class="main py-3">
		<div class="container">
	
	
	<form action="CardAccountTransfer.asp?Action=Save" method="POST" id="frm" name="frm">
	
	<!-- Select AH To Modal -->
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
        <label>Enter details the click Search...</label><br>
            <label for="FirstName">First Name:</label>
            <input type="text" name="FirstName" id="FirstName" class="form-control input-md">
			<label for="LNamms">Last Name:</label>
            <input type="text" name="LNamms" id="LNamms" class="form-control input-md">
			<label for="EmpIDS">Employee ID:</label>
            <input type="text" name="EmpIDS" id="EmpIDS" class="form-control input-md">
			<!--New Name on Card addition ---not being used as CTS Cards cannot be transferred to (no CMS Accounts)
			<label for="NonCS">Name On Card:</label>
            <input type="text" name="NonCS" id="NonCS" class="form-control input-md">
			END NAME on CARD Additin (not used)-->
      </div>
	  <div id="EmpSearProg">
		<span id="Progress" style="display:none; padding-left:20px; padding-bottom:20px;"><img src="../Images/progress.gif" />  &nbsp;&nbsp;&nbsp; <b>Loading...</b></span>
	  </div>
	  <div id="EmpSear">
		
	  </div>
      <div class="modal-footer">
		<!--<button type="button" class="btn btn-primary" onClick="loadDoc('NameOnCard')"><i class="fa fa-credit-card"></i> Search Name On Card</button>-->
        <button type="button" class="btn btn-primary" onClick="loadDoc()"><i class="fa fa-check"></i> Search</button>
		<button type="button" class="btn btn-secondary" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
      </div>
	  <table><tr><td></td><td></td><td></td></tr></table>
    </div>
  </div>
</div>
<!-- END Select AH To Modal -->

</form>

	<!-- Start Delete Modal -->

	<div class="modal fade" id="ModalDelete" tabindex="-1" role="dialog" aria-labelledby="ModalDeleteCenterTitle" aria-hidden="true">
		<div class="modal-dialog modal-dialog-centered" role="document">
		  	<div class="modal-content">
				<div class="modal-header">
			  		<h5 class="modal-title" id="ModalDeleteLongTitle" style="font-weight:bold;">Delete User</h5>
			  		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				</div>
				<div class="modal-body">			  
				  	<div class="row">
						<div class="col-md-12 mb-3">			
					 		<div class="alert alert-danger" role="alert" id="AlertDanger" style="display:block">
						  		<div id="ModalDeleteMessage"></div>
					  		</div>
						</div>
				  	</div>
				  	<div class="row">
					  	<div class="col-md-12 mb-3" style="text-align:right;">
						  	<input type="hidden" id="EmployeeDeleteID"></input>
						  	<button class="btn btn-primary btn-sm" onClick="deleteEmployee(this)"><i class="fa fa-check"></i> Yes</button>
							<button type="button" class="btn btn-secondary btn-sm" data-dismiss="modal" onClick="DeleteModalClose(this);"><i class="fa fa-times"></i> No</button>
					  	</div>
				  	</div>
			  	</div>
			</div>	
		</div>
		<div class="modal-footer"></div>
	</div>

	<!-- End Delete Modal -->
	
	<!-- Start Download Excel Modal -->
	<div class="modal fade" id="ModalExcel" tabindex="-1" role="dialog" aria-labelledby="ModalDeleteCenterTitle" aria-hidden="true">
		<div class="modal-dialog modal-dialog-centered" role="document">
		  	<div class="modal-content">
				<div class="modal-header">
			  		<h5 class="modal-title" id="ModalDeleteLongTitle" style="font-weight:bold;">Download Excel file</h5>
			  		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				</div>
				<div class="modal-body">			  
				  	<div class="row">
						<div class="col-md-12 mb-3">			
						  		<div id="ModalFileMessage"></div>
						</div>
				  	</div>
				  	<div class="row">
					  	<div class="col-md-12 mb-3" style="text-align:right;">
						  	<button type="button" class="btn btn-secondary btn-sm" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
					  	</div>
				  	</div>
			  	</div>
			</div>	
		</div>
		<div class="modal-footer"></div>
	</div>
	<!-- End Download Excel Modal -->

	<!-- Start Edit Modal -->
	<!--<form action="CardAccountTransfer.asp?Action=Save" method="POST" id="frm" name="frm" class="needs-validation" novalidate>-->
		<div class="modal fade" id="ReviewModal" tabindex="-1" role="dialog" aria-labelledby="ReviewModalTitle" aria-hidden="true">
			<div class="modal-dialog modal-dialog-centered" role="document">
				<div class="modal-content">
					<div class="modal-header">
						<h5 class="modal-title" id="emailModalLabel">Training Compliance Process <i class="fa fa-cogs"></i></h5><button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
					</div>
					<div class="modal-body">
						<div class="col-md-12">
							<table class="table table-bordered table-hover CAPS">	
								<div id="ModalReviewDetailTop"></div>
								<div id="ModalReviewDetail"></div>	
								
							</table>
						</div>	
						
						<div class="row">
							<div class="col-md-12 mb-3" style="text-align:right;">
								<input type="hidden" id="EmployeeDeleteID"></input>
								<!--<button class="btn btn-primary btn-sm" onClick="self.location='CardAccountTransfer.asp?Action=UpdateTraining'"><i class="fa fa-check"></i> Save</button>-->
								<button class="btn btn-primary btn-sm" onClick="UpdateTraining(this)"><i class="fa fa-check"></i> Save</button>
								<button type="button" class="btn btn-secondary btn-sm" data-dismiss="modal"><i class="fa fa-times"></i> Close</button>
							</div>
						</div>
					
					</div>
					<div class="modal-footer"></div>
				</div>
			</div>
		</div>
	<!--</form>-->
	<!-- End Edit Modal -->
	
	
	<!-- Start Confirm Transfer Modal -->

	<div class="modal fade" id="ModalConfirm" tabindex="-1" role="dialog" aria-labelledby="ModalConfirmTitle" aria-hidden="true">
		<div class="modal-dialog modal-dialog-centered" role="document">
		  	<div class="modal-content">
				<div class="modal-header">
			  		<h5 class="modal-title" id="ModalConfirmTitle" style="font-weight:bold;">Confirm Transfer - Add to Transfer File now?</h5>
			  		<button type="button" class="close" data-dismiss="modal" aria-label="Close"><span aria-hidden="true">&times;</span></button>
				</div>
				<div class="modal-body">			  
				  	<div class="row">
						<div class="col-md-12 mb-3">	
							<div class="alert alert-info" role="alert" id="AlertInfoConfirm" style="display:block">						
					 		<h6 class="modal-title" id="ModalConfirmTitleTo" style="font-weight:bold;">Account Transferring To</h6>
								<div id="ModalConfirmMessageTo"></div>
							</div>
						  		<div id="ModalConfirmMessage"></div>
					  		
						</div>
				  	</div>
				  	<div class="row">
					  	<div class="col-md-12 mb-3" style="text-align:right;">
						  	<input type="hidden" id="EmployeeDeleteID"></input>
						  	<button class="btn btn-primary btn-sm" onClick="SaveConfirmTransfer();"><i class="fa fa-check"></i> Yes</button>
							<button type="button" class="btn btn-secondary btn-sm" data-dismiss="modal"><i class="fa fa-times"></i> No</button>
					  	</div>
				  	</div>
			  	</div>
			</div>	
		</div>
		<div class="modal-footer"></div>
	</div>

	<!-- End Confirm Transfer Modal -->
	
	
<!-- End the first part of the Header Container -->

	
	
		  <div class="row mb-2">
			<div class="col-md-5">
			  <h3>Card Account Transfers</h3>
			</div>
			<div class="col-md-7 text-right">
				<%Call LoadExtraButtons()%>
			</div>
		  </div>
		  
		  <div class="row py-2">
            <div class="col-md-8">
              <%Call LoadViewButtons()%>
            </div>
			<div class="col-md-4">
				<div class="form-group has-search">
					<span class="fa fa-search form-control-feedback"></span>
				 <input type="text" class="form-control" type="search" id="SearchInput" name="SearchInput" title="Enter [EID] (OR multiple EIDs with comma) OR [U:] for CMSUser OR [FName LName] for cardholder name OR [CTS:] for Name On Card" placeholder="Enter EID (EIDs spaced by comma) or U: for CMSUser" onChange="SearchUsers();" value="<%=Request.QueryString("SearchInput")%>"/>
				 </div>
			</div>
          </div>
		  
		<div class="row py-2">
			<div class="col-md-3">
								
				<div class="panel panel-shadow mb-3">

					<div class="panel-header">
						<% Call LoadSelectAHNow()%>
							<!--Loaded from the AJAX page SaveCardAccountTransfer.asp-->
					  </div>
				</div>
				
					<div class="panel panel-shadow mb-3">

					<div class="panel-header">
						<% Call LoadSelectNow()%>
							<!--Loaded from the AJAX page SaveTraining.asp-->
					  </div>
				</div>
			</div>
			

            <div class="col-md-9">
				  <% 'Determine the Table/details to display based on the selection
					If Session("CATViewButton") = "NewTransfer" Then
						Call DisplayTableCards() 
					Else
						Call DisplayTable()
					End If
				  %>

			</div>
        </div>

	
</main>

</form>



<!-- #Include file=../cc/CAPSFooter.asp -->

</body>
</html>

<%

Public Sub DisplayTable()

'Displays the main table with User details and results of searches
'Dim y
Dim strStatus
Dim strSelected
Dim intRecordCount

Dim strSearch
Dim strSort
'Dim strWhere
Dim strRecordMessage
Dim lngStartingPage
Dim lngCurrentPage

Dim lngTotalPages
Dim lngTotalRecords
Dim arrPagecombo(6)
Dim strPageCombo
Dim strOrderType
Dim arrSort(10)
Dim strSortArrow
Dim arrNames
Dim strFNameSearch
Dim strLNameSearch
Dim strSearchDate
Dim strSearchDay
Dim strSearchMonth
Dim strSearchYear
Dim strDateFrom
Dim strDateTo
Dim bolSkip
Dim strActive
Dim strPages2
Dim strEmail
Dim strFName
Dim strLName
Dim strCampus
Dim strDateBy
Dim strCompletion
Dim strCardType
Dim strEmailShort
Dim strGroup

Dim strTop
Dim strSearch2
Dim strExempt
Dim strCardNo
Dim strRemoveCAT
Dim strGroupTitle

Dim strCurrentUser
Dim strToAH
Dim strUserName

Dim strExported
Dim strExportedText
Dim strExportedDelete
Dim strRemoveHeader
Dim strAcctRefHeader
Dim strAccountRef
Dim strFromAHEID
Dim strToAHEID
Dim strComment
Dim strCommentTitle
Dim intSearchTotal

strSearch = Replace(Request.QueryString("SearchInput"), "'", "''")

'strTop = " TOP 500 "

	'Build the TOP Statement
	If Session("PageCombo") = "" Or IsNull(Session("PageCombo")) Then
		Session("PageCombo") = 500
	End If
	
	If IsNull(strTop) or strTop = "" Then
		If Session("PageCombo") = "All" Then
			strTop = ""
		Else
			strTop = "TOP " & Session("PageCombo")
		End If
	End If
	
	If IsEmpty(Request.QueryString("SortType")) Then
		'strOrderType = "ASC"
	Else
		If Request.QueryString("SortType") = "ASC" Then
			strOrderType = "DESC"
			'Set the variable to be used in the sort order Fontawesone image
			strSortArrow = "-down"
		Else
			strOrderType = "ASC"
			'Set the variable to be used in the sort order Fontawesone image
			strSortArrow = "-up"
		End If
	End If
	
	'Get the Sort order if selected, otherwise sort by the CATID desc as default
	If IsEmpty(Request.QueryString("Sort")) Then
		strSort = " ORDER BY CardAccountTransferID DESC"
	Else		
		strSort = " ORDER BY " & Request.QueryString("Sort") & " " & strOrderType
	End If
	
	'If there is no sort then sort by the most recent submitted
	If IsNull(strSort) Or strSort = "" Then strSort = " ORDER BY RequestedDate DESC"
	
	If Session("CATViewButton") = "Exported" Then
		strWhere = " AND [Exported] ='Y'"
		strRemoveHeader = "Re-Add"
		strAcctRefHeader = ""
		
	ElseIf Session("CATViewButton") = "NotExported" Then
		strWhere = " AND [Exported] ='N' "
		strRemoveHeader = "Remove"
		strAcctRefHeader = "<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=AccountRefNo&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Account Ref No <i class=""fa fa-sort" & arrSort(3) & """></i></a></th>"
	Else
		'This catches ALL
		strWhere = " "
		strRemoveHeader = "Re-Add"
		strAcctRefHeader = "<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=AccountRefNo&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Account Ref No <i class=""fa fa-sort" & arrSort(3) & """></i></a></th>"
	End If
	
	
If strSearch = "" OR ISNull(strSearch) Then
		strSQL = "SELECT " & strTop & " * FROM qryCAPSCardAccountTransferUser WITH(NOLOCK) WHERE [CardAccountTransferID] >0 " & strWhere & strSort
			
Else
	'If the user has entered the date lookup then process this, otherwise perform the EmployeeID and Name searches
	If Left(strSearch,2) = "d:" Then
	
		strSearchDate = Right(strSearch,Len(strSearch)-2)
		
		If Len(strSearchDate) > 5 Then
			'strSearchDate = MediumDate(strSearchDate)
			'Check for the Day in the string, which could be one or two characters (numbers)
			If Mid(strSearchDate,2,1) = "/" OR Mid(strSearchDate,2,1) = "-" Then
				strSearchDay = "0" & Mid(strSearchDate,2,1)
				'Get the Month value
				If Mid(strSearchDate,4,1) = "/" OR Mid(strSearchDate,4,1) = "-" Then
					strSearchMonth = Mid(strSearchDate,3,1)
				Else
					strSearchMonth = Mid(strSearchDate,3,2)
				End If
			Else
				strSearchDay = Left(strSearchDate,2)
				'Get the Month value
				If Mid(strSearchDate,5,1) = "/" OR Mid(strSearchDate,5,1) = "-" Then
					strSearchMonth = Mid(strSearchDate,4,1)
				Else
					strSearchMonth = Mid(strSearchDate,4,2)
				End If
			End If
			
			'Get the year value
			'response.write "mid=" & Mid(strSearchDate,Len(strSearchDate)-3,1) & "</br>"
			If Mid(strSearchDate,Len(strSearchDate)-2,1) = "/" OR Mid(strSearchDate,Len(strSearchDate)-2,1) = "-" Then
				strSearchYear = "20" & Right(strSearchDate,2)
			Else
				strSearchYear = Right(strSearchDate,4)
			End If
			
			'Original Search date string before splitting to Day, Month and Year.
			'strSearchDate = Left(strSearchDate,2) & "-" & MonthName(Mid(strSearchDate,4,2)) & "-" & Right(strSearchDate,4)
			
			strSearchDate = strSearchDay & "-" & MonthName(strSearchMonth) & "-" & strSearchYear
		End If
		'response.write strSearchDate
		If IsDate(strSearchDate) Then 'DateAdd("d", -1, strSearchDate)
		
			'strDateFrom = DateAdd("d", -1, strSearchDate)
			strDateFrom = strSearchDate
			strDateTo = DateAdd("d", +1, strSearchDate)

			strSQL = "SELECT " & strTop & " * FROM qryCAPSCardAccountTransferUser WITH(NOLOCK) WHERE (Requested > '" & strDateFrom & "' AND Requested < '" & strDateTo & "')" & strWhere & strSort
			
		Else
			strSQL = "SELECT " & strTop & " * FROM qryCAPSCardAccountTransferUser WITH(NOLOCK) WHERE (Requested = '" & strSearchDate & "')" & strWhere & strSort
		End If
		
	Else
		'If Session("UserView") = "All" Then
		'If there is a number of EIDs entered separated by a comma then search for all employee IDs
		If Instr(1,strSearch,",")>0 Then
			arrNames = Split(strSearch,",")
			For x = 0 to UBound(arrNames)
				strSearch2=strSearch2 & "'" & arrNames(x) & "',"
			Next
			
			strSearch = left(strSearch2,Len(strSearch2)-1)
			strSQL = "SELECT " & strTop & " * FROM qryCAPSCardAccountTransferUser WITH(NOLOCK) WHERE FromAH IN (" & strSearch & ") " & strSort
			
		Else
			strSQL = "SELECT " & strTop & " * FROM qryCAPSCardAccountTransferUser WITH(NOLOCK) WHERE (NameOnCard Like '%" & strSearch & "%') " & strSort
			
		End If
		
	End If
	
End If

'Get the Sort Order for the sort order fontawesome image for the selected sort field
'For x = 1 to 8

	Select Case Request.QueryString("Sort")
	
		Case "FromAH"
			arrSort(1) = strSortArrow
		Case "ToAH"
			arrSort(2) = strSortArrow
		Case "AccountRefNo"
			arrSort(3) = strSortArrow
		Case "NameOnCard"
			arrSort(4) = strSortArrow
		Case "Comment"
			arrSort(5) = strSortArrow
		Case "Exported"
			arrSort(6) = strSortArrow
		Case "RequestedBy"
			arrSort(7) = strSortArrow
		Case "RequestedDate"
			arrSort(8) = strSortArrow
		Case "CardAccountTransferID"
			arrSort(8) = strSortArrow
		Case "UserName"
			arrSort(9) = strSortArrow
	End Select

'Next

'Build the message displayed at the bottom of the screen with the search details
'If Session("UserView") = "All" Then
	strRecordMessage = strSearch
'Else
'	strRecordMessage = Session("UserName") 
'End If

	'y = 0
	
'	objRS.Open "Select Count(*) AS [CountEIDs] FROM qryCAPSCardAccountTransferUser WITH(NOLOCK) WHERE [CardAccountTransferID]>0 " & strWhere,objCon,3,1
'		If not objRS.EOF Then
		
'			lngTotalRecords = objRS("CountEIDs")
'			If IsNumeric(lngTotalRecords) Then lngTotalRecords = FormatNumber(lngTotalRecords,0)
'		End If
		
'	objRS.Close
	lngTotalRecords=0
	'Get the total returned so as not to show more displaying than records returned
	If IsNull(strTop) or strTop = "" Then
		intSearchTotal = 0
	Else
		intSearchTotal = Right(strTop,Len(strTop)-4)
	End IF
	
	If clng(intSearchTotal) > clng(lngTotalRecords) Then
		intSearchTotal = lngTotalRecords
	Else
		intSearchTotal = Session("PageCombo")
	End IF
	
	
	'response.write strSQL
	objRS.Open strSQL,objCon,3,1
	
'Response.Write strSQL
	
	'Set the Page combos here so can be transferred to other pages together
	arrPagecombo(1) = "50"
	arrPagecombo(2) = "100"
	arrPagecombo(3) = "200"
	arrPagecombo(4) = "500"
	arrPagecombo(5) = "1000"
	arrPagecombo(6) = "All"
	
	'Build the Page Combo for TOP statement
	For x = 1 to 6
	
		If x = 6 OR IsNumeric(Session("PageCombo")) = False Then
			If Session("PageCombo") = arrPagecombo(x) Then
				strTop = lngTotalRecords
				strSelected = " SELECTED "
			End If
		Else
			If cint(Session("PageCombo")) = cint(arrPagecombo(x)) Then
				strSelected = " SELECTED "
				
				'If x = 6 Then strTop = lngTotalRecords
				
			Else
				strSelected = ""
			End If
		End If
		
		strPageCombo = strPageCombo & "<option " & strSelected & " value=""" & arrPagecombo(x) & """>" & arrPagecombo(x) & "</option>"
	Next
	
	strPageCombo = "<SELECT ID=""PageCombo"" Name=""PageCombo"" onChange=""ChangePage();"">" & strPageCombo & "</select>"
		
	'Write a message in the list if there are no Users
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Card Account Transfer record for " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	Else
		
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Displaying " & intSearchTotal & " of " & lngTotalRecords & " cardholders </span><span class=""panel-subheader"" style=""float:right;"">Number of records per page: " & strPageCombo  & "</span></div></div>"
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=CardAccountTransferID&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> CAT ID <i class=""fa fa-sort" & arrSort(9) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=FromAH&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> From AH <i class=""fa fa-sort" & arrSort(1) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=ToAH&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> To AH <i class=""fa fa-sort" & arrSort(2) & """></i></a></th>" & _
			strAcctRefHeader & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=NameOnCard&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Name On Card  <i class=""fa fa-sort" & arrSort(4) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=Comment&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Comment <i class=""fa fa-sort" & arrSort(5) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=Exported&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Exported <i class=""fa fa-sort" & arrSort(6) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=RequestedBy&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Requested By <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=RequestedDate&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """ >Requested Date <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px; text-align:center;""><a href=""CardAccountTransfer.asp?Sort=UserName&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """ >Current CMS User <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px; text-align:center;"">" & strRemoveHeader & "</th>" & _
			"</tr></thead><tbody class=""text-left"">"
				
	End If
    
	x = 0
	
intRecordCount = 0

'objRS.Open strSQL,objCon   
    
	
	Do until objRS.EOF	
		
		'Get the Current User and if the same as the To AH then colour
		If IsNull(objRS("UserName")) Then
			strUserName = ""
		Else
			strUserName = Trim(objRS("UserName"))
		End If
		
		If IsNull(objRS("ToAH")) Then
			strToAH = ""
		Else
			strToAH = Trim(objRS("ToAH"))
		End If
		
		If strUserName = strToAH Then
			strCurrentUser = "<span class=""badge badge-pill badge-success"">" & strUserName & "</span>"
		Else
			strCurrentUser = "<span class=""badge badge-pill badge-danger"">" & strUserName & "</span>"
		End If
		
		If IsNull(objRS("Exported")) Then
			strExported = ""
			strExportedDelete = "D"
		Else
			strExported = Trim(objRS("Exported"))
			
			If strExported = "N" Then
				strExported = "<span class=""badge badge-pill badge-info"">" & strExported & "</span>"
				strExportedText = "Title=""Not Yet Exported. Will be Exported in ProMaster exports process of the next Afternoon Run."" "
				strExportedDelete = "D"
				strAccountRef = "<td style=""font-size:11px; text-align:center;"">" & objRS("AccountRefNo") & "</td>"
			ElseIf strExported = "Y" Then
				strExported = "<span class=""badge badge-pill badge-success"">" & strExported & "</span>"
				strExportedText = "Title=""Exported. Request for export made on - " & objRS("RequestedDate") & "."" "
				strExportedDelete = ""
				If Session("CATViewButton") = "Exported" Then
					strAccountRef = ""
				Else
				strAccountRef = "<td></td>"
				End If
			'This will catch the Deleted (D) records
			Else
				strExported = "<span class=""badge badge-pill badge-danger"">" & strExported & "</span>"
				strExportedText = "Title=""Deleted before being Exported."" "
				strExportedDelete = "N"
				strAccountRef = "<td style=""font-size:11px; text-align:center;"">" & objRS("AccountRefNo") & "</td>"
			End If
		End If
		
		strRemoveCAT = ""
		
		'Set the Delete or add buttons based on the card account transfer exported field
		If Session("CATViewButton") = "Exported" Then
			'Only display the add if the record has been deleted
			If strExportedDelete = "N" Then
				'Add a pill badge for removal of the exempt date
				strRemoveCAT = "<a href=""CardAccountTransfer.asp?Action=RemoveCAT&CardAccountTransferID=" & objRS("CardAccountTransferID") & "&Type=" & strExportedDelete & """ title=""Re-add Card Account Transfer""><span class=""badge badge-pill-xs badge-success badge-xs"" style=""padding:10px;"">+</span></a>"
			End If
		ElseIf Session("CATViewButton") = "NotExported" Then
			'Add a pill badge for removal of the exempt date
			strRemoveCAT = "<a href=""CardAccountTransfer.asp?Action=RemoveCAT&CardAccountTransferID=" & objRS("CardAccountTransferID") & "&Type=" & strExportedDelete & """ title=""Remove Card Account Transfer""><span class=""badge badge-pill-xs badge-danger badge-xs"" style=""padding:10px;"">X</span></a>"
		Else
			'Only display the add if the record has been deleted
			If strExportedDelete = "N" Then
				'Add a pill badge for removal of the exempt date
				strRemoveCAT = "<a href=""CardAccountTransfer.asp?Action=RemoveCAT&CardAccountTransferID=" & objRS("CardAccountTransferID") & "&Type=" & strExportedDelete & """ title=""Re-add Card Account Transfer""><span class=""badge badge-pill-xs badge-success badge-xs"" style=""padding:10px;"">+</span></a>"
			End If
		End If
	
		'Create the Title (mouse over) for the Account Holder From and To EID details
		If IsNull(objRS("FromAHEmployeeID")) Then
			strFromAHEID = ""
		Else
			strFromAHEID = "Title=""" & Trim(objRS("FromAHEmployeeID")) & " - " & Trim(objRS("FromAHName")) & " """
		End If
		
		If IsNull(objRS("ToAHEmployeeID")) Then
			strToAHEID = ""
		Else
			strToAHEID = "Title=""" & Trim(objRS("ToAHEmployeeID")) & " - " & Trim(objRS("ToAHName")) & " """
		End If		
		
		'Get the comments and shorten to save space
		If IsNull(objRS("Comment")) Then
			strComment = ""
			strCommentTitle = ""
		Else
			If Len(objRS("Comment")) > 10 Then
				strComment = Left(objRS("Comment"),10)
				strCommentTitle = "Title=""" & objRS("Comment") & " """
			Else
				strComment = objRS("Comment")
				strCommentTitle = "Title=""" & objRS("Comment") & " """
			End If
		End If
		
		Response.Write "<tr><td style=""font-size:12px;"">" & objRS("CardAccountTransferID") & "</td>" & _
			"<td style=""font-size:12px;"" " & strFromAHEID & ">" & objRS("FromAH") & "</td>" & _
			"<td style=""font-size:12px;"" " & strToAHEID & ">" & objRS("ToAH") & "</td>" & _
			strAccountRef & _
			"<td style=""font-size:11px; text-align:center;"">" & objRS("NameOnCard") & "</td>" & _ 
			"<td style=""font-size:12px;"" " & strCommentTitle & ">" & strComment & "</td>" & _
			"<td style=""font-size:12px;"" " & strExportedText & " >" & strExported & "</td>" & _
			"<td style=""font-size:11px;"">" & objRS("RequestedBy") & "</td>" & _
			"<td style=""font-size:13px;"">" & objRS("RequestedDate") & "</td>" & _
			"<td style=""font-size:12px; text-align:center;"">" & strCurrentUser & "</td>" & _
			"<td style=""font-size:13px; text-align:center;"">" & strRemoveCAT & "</td></tr>"
			'"<td style=""font-size:13px; text-align:center;""><input class=""PrcChecks"" type=""checkbox"" id=""Chk" & intRecordCount & """ name=""Chk"" data-id=""" & objRS("CardAccountTransferID") & """ data-CourseID=""" & objRS("CardAccountTransferID") & """ onClick=""SaveTrainingActionw(this);""></td></tr>"
			
		intRecordCount = intRecordCount + 1
	
	objRS.movenext
	
	Loop	
	strWhere = "WHERE [CardAccountTransferID]>0 " & strWhere		
	Response.Write "<TR><TH colspan=""7"">Total <input type=""HIDDEN"" id=""WhereClause"" name=""WhereClause"" value=""tbl=qryCAPSCardAccountTransferUser&W=" & strWhere & "&Top=" & strTop & """ ></TH>" & _
			"<TH colspan=""2"" style=""text-align:center;"">" & intRecordCount & "</TH></TR></tbody></table></div></div>"
	
objRS.Close

End Sub


Public Sub DisplayTableCards()

'Displays the main table with User details and results of searches
'Dim y
Dim strStatus
Dim strSelected
Dim intRecordCount

Dim strSearch
Dim strSort

Dim lngTotalPages
Dim lngTotalRecords
Dim strOrderType
Dim arrSort(10)
Dim strSortArrow
Dim arrNames
Dim strFNameSearch
Dim strLNameSearch
Dim strTop
Dim strRecordMessage
Dim strCardType
Dim strCardNumber
Dim strSearchEID
Dim strSearch2
Dim strEmployeeName
Dim intSearchTotal

strSearch = Replace(Request.QueryString("SearchInput"), "'", "''")
strTop = "Top 500"

	'Determine the EmployeeID field to query based on whether account holder (from CMS) or Cardholder (from CAPS) is selected
	If Session("StatusViewButton") = "Cardholder" Then
		strSearchEID = "EmployeeID"
	Else
		strSearchEID = "employee_id"
	End If
	
	If IsEmpty(Request.QueryString("SortType")) Then
		'strOrderType = "ASC"
	Else
		If Request.QueryString("SortType") = "ASC" Then
			strOrderType = "DESC"
			'Set the variable to be used in the sort order Fontawesone image
			strSortArrow = "-down"
		Else
			strOrderType = "ASC"
			'Set the variable to be used in the sort order Fontawesone image
			strSortArrow = "-up"
		End If
	End If
	
	If IsEmpty(Request.QueryString("Sort")) Then
		strSort = ""
	Else		
		strSort = " ORDER BY " & Request.QueryString("Sort") & " " & strOrderType
	End If
	
	'If there is no sort then sort by the most recent submitted
	If IsNull(strSort) Or strSort = "" Then strSort = " ORDER BY Surname"
	
'	If Session("CATViewButton") = "Exported" Then
'		strWhere = " AND [Exported] ='Y'"
'	ElseIf Session("CATViewButton") = "NotExported" Then
'		strWhere = " AND [Exported] <>'Y' "
'	Else
'		'This catches ALL
'		strWhere = " "
'	End If
	
'The Default search'
strSQL = "SELECT * FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE (" & strSearchEID & " Like '%" & strSearch & "%') " & strSort
strWhere = " AND (" & strSearchEID & " Like '%" & strSearch & "%')"

If strSearch = "" OR ISNull(strSearch) Then
		strSQL = "SELECT " & strTop & " * FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE [CardID] >0 " & strWhere & strSort
			
Else
	If Len(strSearch)=7 Then
		If Left(strSearch,1) ="8" Then
			'Get the name for the EID being searched for (Function in CAPSFunctions.asp)
			strEmployeeName = GetEmployeeName(strSearch,"Y")
		End IF
	End If
	
	If Len(strSearch)>2 Then
		If UCASE(Left(strSearch,2)) = "U:" Then
			strSearch = Right(strSearch,Len(strSearch)-2)
			'response.write strSearch
			
			strSQL = "SELECT * FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE (user_name Like '%" & strSearch & "%') " & strSort
			strWhere = " AND (user_name Like '%" & strSearch & "%')"
		ElseIf UCASE(Left(strSearch,4)) = "CTS:" Then
			strSearch = Right(strSearch,Len(strSearch)-4)
			'response.write strSearch
			
			strSQL = "SELECT * FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE CardTypeSub='CTS' AND (NameOnCard Like '%" & strSearch & "%') " & strSort
			strWhere = " AND (NameOnCard Like '%" & strSearch & "%')"
		End If
		
		'strSQL = "SELECT * FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE (" & strSearchEID & " Like '%" & strSearch & "%') " & strSort
		'strSQL = "SELECT * FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strSort
	Else
		strSQL = "SELECT * FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strSort
		'strSQL = "SELECT * FROM tblCAPSCard WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strSort
		strWhere = " AND (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%')"
	End If
	
	'If Session("UserView") = "All" Then
	'If there is a number of EIDs entered separated by a comma then search for all employee IDs
	If Instr(1,strSearch,",")>0 Then
		arrNames = Split(strSearch,",")
		For x = 0 to UBound(arrNames)
			strSearch2=strSearch2 & "'" & arrNames(x) & "',"
		Next
		
		strSearch = left(strSearch2,Len(strSearch2)-1)
		strSQL = "SELECT * FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE " & strSearchEID & " IN (" & strSearch & ") " & strSort
		'strSQL = "SELECT * FROM tblCAPSCard WITH(NOLOCK) WHERE EmployeeID IN (" & strSearch & ") " & strSort
		strWhere = " AND " & strSearchEID & " IN (" & strSearch & ")"
	Else
		'If the user has entered a search term with a space then assume this is a first and last name so search on that only
			If Instr(1,strSearch," ")>0 Then
				arrNames = Split(strSearch," ")
				strFNameSearch = arrNames(0)
				strLNameSearch = arrNames(1)
			
				strSQL = "SELECT * FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE (FirstName Like '%" & strFNameSearch & "%' AND Surname Like '%" & strLNameSearch & "%') " & strSort
				strWhere = " AND (FirstName Like '%" & strFNameSearch & "%' AND Surname Like '%" & strLNameSearch & "%')"
			End If

		'strSQL = "SELECT * FROM tblCAPSCard WITH(NOLOCK) WHERE (EmployeeID Like '%" & strSearch & "%' OR FirstName Like '%" & strSearch & "%' OR Surname Like '%" & strSearch & "%') " & strSort
		
	End If
			
End If

'Get the Sort Order for the sort order fontawesome image for the selected sort field
'For x = 1 to 8

	Select Case Request.QueryString("Sort")
	
		Case "CardID"
			arrSort(1) = strSortArrow
		Case "CardType"
			arrSort(2) = strSortArrow
		Case "Status"
			arrSort(3) = strSortArrow
		Case "NameOnCard"
			arrSort(4) = strSortArrow
		Case "EmployeeID"
			arrSort(5) = strSortArrow
		Case "CMSUser"
			arrSort(6) = strSortArrow
		Case "Expiry"
			arrSort(7) = strSortArrow
		Case "DateIssued"
			arrSort(8) = strSortArrow

	End Select

'Next

'Build the message displayed at the bottom of the screen with the search details
'If Session("UserView") = "All" Then
	strRecordMessage = strSearch
'Else
'	strRecordMessage = Session("UserName") 
'End If

	'y = 0
	'Response.Write "Select Count(*) AS [CountEIDs] FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE [CardID]>0 " & strWhere
	objRS.Open "Select Count(*) AS [CountEIDs] FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE [CardID]>0 " & strWhere,objCon,3,1
		If not objRS.EOF Then
		
			lngTotalRecords = objRS("CountEIDs")
			If IsNumeric(lngTotalRecords) Then lngTotalRecords = FormatNumber(lngTotalRecords,0)
		End If
		
	objRS.Close
	
	'Get the total returned so as not to show more displaying than records returned
	If IsNull(Session("PageCombo")) or Session("PageCombo") = "" Then
		intSearchTotal = 0
	Else
		intSearchTotal = Session("PageCombo")'Right(Session("PageCombo"),Len(Session("PageCombo"))-4)
	End IF
	
	If clng(intSearchTotal) > clng(lngTotalRecords) Then
		intSearchTotal = lngTotalRecords
	Else
		intSearchTotal = Session("PageCombo")
	End IF
	
	'response.write strSQL
	objRS.Open strSQL,objCon,3,1
	
Response.Write strSQL
	
	'Write a message in the list if there are no Users
	If objRS.EOF Then
		Response.Write "<TR><TH colspan=""10"" Style=""text-align:center;"">No Card Account Transfer record for " & strRecordMessage & "</TH>" & _
				"<TH colspan=""3"" style=""text-align:center;""></TH></TR>"
	Else
		
		'Write the Table header with search results
		Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & _
                  "<span class=""panel-subheader"">Displaying " & intSearchTotal & " of " & lngTotalRecords & " cardholders </span><span style=""color:red; font-weight:bold; align:right; text-align:right; float:right;"">" & strEmployeeName & "</span></div></div>"
		
		Response.Write "<div class=""row""><div class=""col-12"">" & _
			"<table class=""table table-compact text-left""><thead><tr>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=CardID&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Card ID <i class=""fa fa-sort" & arrSort(1) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=CardType&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Card Type <i class=""fa fa-sort" & arrSort(2) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=Status&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Status <i class=""fa fa-sort" & arrSort(3) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=NameOnCard&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Name On Card <i class=""fa fa-sort" & arrSort(4) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=EmployeeID&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> EmployeeID  <i class=""fa fa-sort" & arrSort(5) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=CMSUser&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Orig AH <i class=""fa fa-sort" & arrSort(6) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=user_name&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Current AH <i class=""fa fa-sort" & arrSort(7) & """></i></a></th>" & _
			"<th scope=""col"" style=""font-size:13px;""><a href=""CardAccountTransfer.asp?Sort=AccountRefNo&Link=CD&SortType=" & strOrderType & "&SearchInput=" & strSearchInput & """> Account Ref <i class=""fa fa-sort" & arrSort(8) & """></i></a></th>" & _
			"<th scope=""col"" style=""text-align:center;"">Select <input type=""checkbox"" id=""SelectAllCheck"" name=""SelectAllCheck""></th>" & _
			"</tr></thead><tbody class=""text-left"">"
				
	End If
    
	x = 0
	
intRecordCount = 0

'objRS.Open strSQL,objCon   
    
	
	Do until objRS.EOF	
		
		If isNull(objRS("CardTypeSub")) Then
			strCardType = ""
		Else
			strCardType  = objRS("CardType") & " " & objRS("CardTypeSub")
		End If
		
		
		If IsNull(objRS("Status")) Then
				strStatus = ""
		Else
			'Determine the Status display based on the Card Type
			If strCardType = "DPC ANZ" Then
				If objRS("Status") = "" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
				ElseIf objRS("Status") = "L" OR objRS("Status") = "C" OR objRS("Status") = "S" Then
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				ElseIf objRS("Status") = "T" Then
					strStatus = "<span class=""badge badge-pill badge-warning"">Temporary Hold</span>"
				Else
					strStatus = ""
				End If
			Else
				If objRS("Status") = "00" Then
					strStatus = "<span class=""badge badge-pill badge-success"">Active</span>"
				ElseIf objRS("Status") = "01" OR objRS("Status") = "02" Then
					strStatus = "<span class=""badge badge-pill badge-danger"">Cancelled</span>"
				Else
					strStatus = ""
				End If
			End If
		End If
		
		If isNull(objRS("CardNumberShort")) Then
			strCardNumber = ""
		Else
			strCardNumber  = "title=""" & objRS("CardNumberShort") &  """"
		End If
		
		
			
		Response.Write "<tr><td style=""font-size:13px;"">" & objRS("CardID") & "</td>" & _
			"<td style=""font-size:12px;"" " & strCardNumber & ">" & objRS("CardType") & " " & objRS("CardTypeSub") & "</td>" & _
			"<td style=""font-size:13px;"">" & strStatus & " " & objRS("Status") & "</td>" & _
			"<td style=""font-size:12px;"">" & objRS("NameOnCard") & "</td>" & _
			"<td style=""font-size:13px; text-align:center;"">" & objRS("EmployeeID") & "</td>" & _ 
			"<td style=""font-size:13px;"">" & objRS("CMSUser") & "</td>" & _
			"<td style=""font-size:12px;"">" & objRS("user_name") & "</td>" & _
			"<td style=""font-size:12px;"">" & objRS("AccountRefNo") & "</td>" & _
			"<td style=""font-size:13px; text-align:center;""><input class=""PrcChecks"" type=""checkbox"" id=""Chk" & intRecordCount & """ name=""Chk"" data-id=""" & objRS("EmployeeID") & """ data-CourseID=""" & objRS("CardID") & """ onClick=""SaveCATAction(this);""></td></tr>"
			
		intRecordCount = intRecordCount + 1
	
	objRS.movenext
	
	Loop	
			
	Response.Write "<TR><TH colspan=""7"">Total <input type=""HIDDEN"" id=""WhereClause"" name=""WhereClause"" value=""" & strWhere & """ ></TH>" & _
			"<TH colspan=""2"" style=""text-align:center;"">" & intRecordCount & "</TH></TR></tbody></table></div></div>"
	
objRS.Close

End Sub


Sub TransferAccount(strComment)

Dim intRecord
Dim intEmailTemplateID

	If IsNull(strComment) OR strComment = "" Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: No Comment Entered. Please enter a COMMENT for the transfer.</div>"
		Exit Sub
	End If

	'Get the EmailTemplateId from the Query String
	If Request.QueryString("EmailTemplateID") = "" Or IsNull(Request.QueryString("EmailTemplateID")) Then
		intEmailTemplateID = 0
	Else
		intEmailTemplateID = Request.QueryString("EmailTemplateID")
	End If
	
  		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCardAccountTransferSaveAll"

			.Parameters.Append objCmd.CreateParameter("CardAccountTransferID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("ToAH", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("Comment", adVarChar, adParamInput, 100)
			.Parameters.Append objCmd.CreateParameter("RequestedBy", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("CardAccountTransferIDOutput", adInteger, adParamOutput)
			
			.Parameters("CardAccountTransferID") = 0
			.Parameters("ToAH") = Session("CATEmployeeID")
			.Parameters("Comment") = strComment
			.Parameters("RequestedBy") = Session("UserID")

			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		intRecord = objCmd.Parameters.Item("CardAccountTransferIDOutput")
	
	If intRecord = -1 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Card Transfers Not added! No Account Holder details found for: " & Session("CATEmployeeID")  & "</div>"
	ElseIf intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Card Transfers Not added! See System Admin.</div>"
	Else
		If IsNumeric(intRecord) Then intRecord = intRecord -1
		
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Card Account Transfers Added</div>"
	End If
		
End Sub


Sub TransferAccountReturn(strComment)

Dim intRecord
Dim intEmailTemplateID

	If IsNull(strComment) OR strComment = "" Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: No Comment Entered. Please enter a COMMENT for the transfer.</div>"
		Exit Sub
	End If

	'Get the EmailTemplateId from the Query String
	If Request.QueryString("EmailTemplateID") = "" Or IsNull(Request.QueryString("EmailTemplateID")) Then
		intEmailTemplateID = 0
	Else
		intEmailTemplateID = Request.QueryString("EmailTemplateID")
	End If
	
  		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSCardAccountTransferReturn"

			.Parameters.Append objCmd.CreateParameter("CardAccountTransferID", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("Comment", adVarChar, adParamInput, 100)
			.Parameters.Append objCmd.CreateParameter("RequestedBy", adVarChar, adParamInput, 50)
			.Parameters.Append objCmd.CreateParameter("CardAccountTransferIDOutput", adInteger, adParamOutput)
			
			.Parameters("CardAccountTransferID") = 0
			.Parameters("Comment") = strComment
			.Parameters("RequestedBy") = Session("UserID")

			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		intRecord = objCmd.Parameters.Item("CardAccountTransferIDOutput")
	
	If intRecord = -1 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Card Transfers Not added! No Account Holder details found for: " & Session("CATEmployeeID")  & "</div>"
	ElseIf intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error: Card Transfers Not added! See System Admin.</div>"
	Else
		If IsNumeric(intRecord) Then intRecord = intRecord -1
		
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Card Account Transfers Added</div>"
	End If
		
End Sub


Public Sub LoadViewButtons
'Load the the View Selector buttons depending on what has been clicked
Dim arrButton(4)
Dim strGroups
Dim strSelected

	
If Session("CATViewButton") = "Exported" Then
	arrButton(2) = "active"
ElseIf Session("CATViewButton") = "NotExported" Then
	arrButton(3) = "active"
ElseIf Session("CATViewButton") = "NewTransfer" Then
	arrButton(4) = "active"
Else
	'This catches ALL
	arrButton(1) = "active"
End If

	Response.Write "<div class=""form-row""><div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(1) & """ onClick=""self.location.href='CardAccountTransfer.asp?Link=CD&ViewButton=All';""><i class=""fa fa-folder""></i> View All</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(2) & """ onClick=""self.location.href='CardAccountTransfer.asp?Link=CD&ViewButton=Exported';""><i class=""fa fa-file-alt""></i> Exported</button>" & _
				"<button type=""button"" class=""btn btn-outline-primary " & arrButton(3) & """ onClick=""self.location.href='CardAccountTransfer.asp?Link=CD&ViewButton=NotExported';""><i class=""fa fa-flag""></i> Not Yet Exported</button>" & _
				"</div><button type=""button"" class=""btn btn-outline-primary " & arrButton(4) & """ onClick=""self.location.href='CardAccountTransfer.asp?Link=CD&ViewButton=NewTransfer';""><i class=""fa fa-file""></i> New Transfer</button>"'</div>"				

				
'Load the the Status Selector buttons depending on what has been clicked
Dim strStatusButton

'Get the Status View button depending on what has been selected
If Session("StatusViewButton") = "Cardholder" Then
	strStatusButton ="<button type=""button"" class=""btn btn-outline-info active"" onClick=""self.location.href='CardAccountTransfer.asp?Link=CD&StatusViewButton=AccountHolder&SearchInput=" & strSearchInput & "';"" title=""Click to view details as Cardholder""><i class=""fa fa-address-card""></i> As Cardholder</button>"
Else
	strStatusButton ="<button type=""button"" class=""btn btn-outline-secondary active"" onClick=""self.location.href='CardAccountTransfer.asp?Link=CD&StatusViewButton=Cardholder&SearchInput=" & strSearchInput & "';"" title=""Click to view details as Account Holder""><i class=""fa fa-address-book""></i> As Account Holder</button>"
End If

	'Response.Write 	"<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"">" & _
	Response.Write	"&nbsp;&nbsp;&nbsp;" & strStatusButton & _
				"</div>"
				
				
End Sub


Public Sub LoadExtraButtons
'Load the Excel Export and Card Type buttons depending on what has been clicked

Dim strDropDown

	
	strDropDown = "<div class=""dropdown""><button class=""btn btn-outline-success dropdown-toggle"" type=""button"" id=""dropdownMenuButton"" data-toggle=""dropdown"" aria-haspopup=""true"" aria-expanded=""false""><i class=""fa fa-file-excel""></i> Export To Excel</button>" & _
		"<div class=""dropdown-menu"" aria-labelledby=""dropdownMenuButton"">"
	
	strDropDown = strDropDown & "<a class=""dropdown-item"" target=""_new"" href=""#"" onClick=""OpenExcelReport();"">On Screen List (Small only)</a>"
	strDropDown = strDropDown & "<a class=""dropdown-item"" href=""Training.asp?Action=ExportExcelReport"">Filter Button Selected</a>"
	strDropDown = strDropDown & "</div></div>"

	Response.Write "<div class=""btn-group btn-selector"" role=""group"" aria-label=""Basic example"" style=""float:right;"">" & _
		strDropDown & _
		"<button type=""button"" class=""btn btn-outline-success""data-toggle=""modal"" data-target=""#ModalExcel"" onClick=""LoadFiles();""> View Files</button>" & _
		"</div>"
	
End Sub


Public Sub LoadSelectAHNow()
'Procedure to Load the Header for the Selected Now side panel.
Dim strCards
Dim strTitle
Dim strClear

	Response.Write "<h4><a href=""#"" data-toggle=""modal"" data-target=""#exampleModalCenter"">Selected AH To</a></h4><div id=""TrainingActionSelectedAH"">"
	
	'Allow the user to clear the selected batch if one is selected
	If IsNull(Session("CATEmployeeID")) Or Session("CATEmployeeID") = "0"  Then
		strClear = ""
		strClear = strClear & " <button type=""button"" class=""btn btn-outline-secondary btn-xs"" data-toggle=""modal"" data-target=""#ModalConfirm"" OnClick=""ConfirmCATAction('Return')"" title=""Click to Return the all selected Account Holders to each Cardholder as their own Account Holder""><i class=""fa fa-check""></i> Return To CH </button>"
	Else
		strClear = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onClick=""self.location.href='CardAccountTransfer.asp?Action=ClearAHTo';""><i class=""fa fa-times""></i> Clear AH To</button>"
	End If
	
	Response.Write "<div class=""panel panel-light mb-3""><div class=""panel-header""><h4></h4>" & strClear & "</div></div>"
                  '"<span class=""panel-subheader"" style=""font-weight:bold;"">Unactivated Batches</span>" & strClear & "</div></div>"
		
		
	Response.Write "<div class=""row""><div class=""col-12"" id=""AHTo"">" & _
		"<table class=""table table-compact table-hover text-left""><thead>" & _
		"<tr><th style=""font-weight:bold; font-size:12px;"">Name</th><th style=""font-weight:bold; font-size:12px;"">PMKeYS ID</th>" & _
		"<th style=""font-weight:bold; font-size:12px;"">CMS UserName</th></tr></thead><tbody>"
	
	
'Select all Groups from the Corporate Directory and load them into the select options
	'objRS.Open "SELECT * FROM qryCAPSCardAccountTransferSearch WITH(NOLOCK) WHERE Employee_ID='" & Session("CATEmployeeID") & "'" ,objCon
	
	If Session("CATEmployeeID") <> 0 Then
	objRS.Open "SELECT TOP 1 * FROM [tblCAPSProMasterUser] WHERE [employee_id] = '" & Session("CATEmployeeID") & "'" ,objCon
	
		Session("TrainingSelectedNow") = ""
		
		If Not objRS.EOF Then
			
			strCards = objRS("first_name") & " " & objRS("surname")
			'strTitle = " title=""" & objRS("CardNumberShort") & """ "
			strTitle = " title=""Date Created: " & objRS("create_date") & " - Admin Cntr: " & objRS("admin_ctr_name") & """ "
			
			
			Response.Write "<tr class=""clickable-row"" data-href='CardAccountTransfer.asp?EmployeeID=" & objRS("Employee_ID") & "' data-target='_blank' " & strTitle & "><td style=""font-size:12px;"">" & strCards & "</td><td style=""font-size:12px;"">" & objRS("Employee_ID") & "</td>" & _
				"<td style=""font-size:12px;"">" & objRS("User_Name") & "</td></TR>"
			
			'Response.Write "<tr class=""clickable-row"" data-href='CardAccountTransfer.asp?EmployeeID=" & objRS("EmployeeID") & "' data-target='_blank' " & strTitle & "><td style=""font-size:12px;"">" & strCards & "</td><td style=""font-size:12px;"">" & objRS("EmployeeID") & "</td>" & _
			'	"<td style=""font-size:12px;"">" & objRS("UserName") & "</td></TR>"
				
		Else
			Response.Write "<tr><td style=""font-size:12px;"">No Data Found for: " & Session("CATEmployeeID") & "</td></TR>"
		End If
		'objRS.Movenext
		'Loop

	objRS.Close
	
	End If
	
	Response.write "</tbody></table></div></div></div>"
	
	'If IsNull(Session("CATEmployeeID")) or Session("CATEmployeeID") = "" Then
	'	Response.write "</div>"
	'Else
	'	Response.Write Session("CATEmployeeID") & "</div>"
	'End If
	
End Sub

Public Sub LoadSelectNow()
'Procedure to Load the Header for the Selected Now side panel.

Dim strClear

''The Transfer Now button and records are loaded from the AJAX page (SaveCardAccountTransfer.asp) called when Select checkbox is clicked

'Select all Groups from the Corporate Directory and load them into the select options
	objRS.Open "SELECT [EmployeeID] FROM tblCAPSTrainingAction WITH(NOLOCK) WHERE [Selected]= 'Y' AND UpdatedBy='" & Session("UserID") & "'" ,objCon
		
		Session("TrainingSelectedNow") = ""
		
		Do Until objRS.EOF
			
			Session("TrainingSelectedNow") = Session("TrainingSelectedNow") & "," & objRS("EmployeeID")
			
		objRS.Movenext
		Loop

	objRS.Close
	
	'Remove the last comma if more than one EID is selected
	If Len(Session("TrainingSelectedNow"))>1 Then 
		Session("TrainingSelectedNow") = Right(Session("TrainingSelectedNow"),Len(Session("TrainingSelectedNow"))-1)
	End If
	
	'Allow the user to clear the selected Selected Now if at least one is selected ------ Moved to AJAX page SaveAccountTransfer.asp
	'If IsNull(Session("TrainingSelectedNow")) Or Session("TrainingSelectedNow") = ""  Then
	'	strClear = ""
	'Else
	'	strClear = "<button type=""button"" class=""btn btn-outline-secondary btn-xs"" onClick=""self.location.href='CardAccountTransfer.asp?BatchID=';""><i class=""fa fa-times""></i> Clear Selected</button>"
	'End If
	
	Response.Write "<h4><a href=""CardAccountTransfer.asp?SearchInput=" & Session("TrainingSelectedNow") & """>Selected Now </a></h4><div id=""TrainingActionSelected""></div>"

End Sub

Sub TransferAccount2(strExemptDate)

Dim intRecord
Dim intEmailTemplateID

response.write "Saved!"

Exit Sub

	If IsNull(Session("CourseID")) OR Session("CourseID") = "" Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error No Course ID for " & Session("CardType") & ". Please notify CAPS System Admin.</div>"
		Exit Sub
	End If
	
	'Get the EmailTemplateId from the Query String
	If Request.QueryString("EmailTemplateID") = "" Or IsNull(Request.QueryString("EmailTemplateID")) Then
		intEmailTemplateID = 0
	Else
		intEmailTemplateID = Request.QueryString("EmailTemplateID")
	End If
	
  		With objCmd

			.CommandType = 4
			.CommandText = "spCAPSTrainingActionExempt"

			.Parameters.Append objCmd.CreateParameter("CourseID", adVarChar, adParamInput, 10)
			.Parameters.Append objCmd.CreateParameter("BusinessArea", adVarChar, adParamInput, 20)
			.Parameters.Append objCmd.CreateParameter("ExemptDate", adDate, adParamInput)  
			.Parameters.Append objCmd.CreateParameter("EmailDetailID", adInteger, adParamInput)			
			.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
			.Parameters.Append objCmd.CreateParameter("TrainingActionIDOutput", adInteger, adParamOutput)
			
			.Parameters("CourseID") = Session("CourseID")
			.Parameters("BusinessArea") = Session("BusinessArea")
			If strExemptDate = "" Or IsNull(strExemptDate) Then
			Else
			.Parameters("ExemptDate") = strExemptDate
			End IF
			.Parameters("EmailDetailID") = intEmailTemplateID
			.Parameters("UpdatedBy") = Session("UserID")

			.ActiveConnection = objCon
			 
		End With
	   
		objCmd.Execute        
	  
		intRecord = objCmd.Parameters.Item("TrainingActionIDOutput")
		'Session("EmailErrorMsgID") = lngEmailErrorMsgID		
	
	If intRecord = 0 Then
		Response.Write "<div class=""alert alert-danger"" role=""alert"">Error Exempt Date not updated</div>"
	Else
		Response.Write "<div class=""alert alert-success"" role=""alert"">" & intRecord & " Employees Updated</div>"
	End If
		
End Sub


Public Sub DeleteTransfer(strType, lngCATID)

	strSQL = "UPDATE tblCAPSCardAccountTransfer SET Exported = '" & strType & "' WHERE CardAccountTransferID = " & lngCATID & ""
	
	objCon.Execute strSQL
	
	Response.Write "<div class=""alert alert-success"" role=""alert"">Card Account Transfer changed to " & strType & "!</div>"
	
End Sub


Sub ClearTrainingAction()
'Procedure to clear all Card Account Transfers selected on screen under SELECTED NOW
	
	With objCmd

	.CommandType = 4
		.CommandText = "spCAPSCardAccountTransferActionClear"

		.Parameters.Append objCmd.CreateParameter("BusinessArea", adVarChar, adParamInput, 20)
		.Parameters.Append objCmd.CreateParameter("Selected", adChar, adParamInput, 1)              
		.Parameters.Append objCmd.CreateParameter("UpdatedBy", adInteger, adParamInput)
		

		.Parameters("BusinessArea") = "CardAccountTransfer"'Session("BusinessArea")'strBusinessArea
		.Parameters("Selected") = "Y"'strSelected
		.Parameters("UpdatedBy") = Session("UserID")
		
		.ActiveConnection = objCon
		 
	End With
   
	objCmd.Execute 

End Sub


Set objRS = Nothing
Set objCon = Nothing
%>
