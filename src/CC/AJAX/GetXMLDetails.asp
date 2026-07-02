<!-- #Include file=../CAPSFunctions.asp -->
<html lang="en">
<%

Dim objCon
Dim objRS
Dim x
Dim strApplicationID
Dim strStatus
Dim strName
Dim strSubStatPos
Dim strNameCard
Dim strSubStatPosCard
Dim strEmployeeID
Dim strFamilyName
Dim strGivenNames
Dim intAppType
Dim strAppType
Dim strAppDOBAge

Dim dblCurrentLimit
Dim dblNewLimit
Dim dblNewTransactionLimit
Dim dteTransactionDateFrom
Dim dteTransactionDateTo

Dim strUpdatedDTCSignedApplicant
Dim strUpdatedDTCDualSignedApplicant
Dim strUpdatedCMCSignedApplicant
Dim strUpdatedDPCSignedApplicant
Dim strUpdatedDTCLCSignedApplicant
Dim strUpdatedDPCLCSignedApplicant
Dim strDPCandDTCSignedApplicant
Dim subDigitalSignature6_chkSES
Dim subDigitalSignature6_chkASFIN
Dim subDigitalSignature6_chkASFINSupport


'Open Database Connection
Set objCon = Server.CreateObject("ADODB.Connection")
    
    objCon.Open Session("DBConnection")

Set objRS = Server.CreateObject("ADODB.Recordset")

If Not IsEmpty(Request.QueryString("ApplicationID")) Then
	strApplicationID = " WHERE ApplicationID = " & Request.QueryString("ApplicationID") & ""
End If

If strApplicationID = "" or IsNull(strApplicationID) Then
	strApplicationID = " WHERE ApplicationID = " & Session("ApplicationID") & ""
End If

Public Sub ShowDetails()

	'Description:	Loads CDMC Details onto the page called from
	objRS.Open "SELECT * FROM tblCAPSXMLApplication WITH(NOLOCK) " & strApplicationID,objCon


		If Not objRS.EOF Then
			
			'Write the Header			
			Response.write "<div class=""row""><div class=""col-md-12"">" & _
				"<table class=""table table-compact"">"
				
			'Make the CS Title Given names and Surname into one field/variable (for space savings)
			If IsNull(objRS("subLimitChange_fldCDFamilyName")) Then
				strName = ""
			Else
				strName = Trim(Trim(objRS("subLimitChange_fldCDGivenNames")) & " " & Trim(objRS("subLimitChange_fldCDFamilyName")))
			End If
			
			'Make the CDMC Suburb, State, PostCode into one field/variable (for space savings)
			'If IsNull(objRS("OutSuburb")) Then
			'	strSubStatPosCard = ""
			'Else
			'	strSubStatPosCard = Trim(Trim(objRS("OutSuburb")) & " " & Trim(objRS("OutState")) & " " & Trim(objRS("OutPostCode")))
			'End If	
			'The variable for the DTC dual and DPC 
			strDPCandDTCSignedApplicant = ""	
		
			
			'Determine the field based on the applicaiton type
			If Not IsNull(objRS("grpIAmApplyingFor")) Then
			
				Select Case objRS("grpIAmApplyingFor")
				
					Case 1
						intAppType = 1
						strAppType = "<tr class=""updated""><td style=""font-weight:bold;"">Application Type  </td><td>DTC Only</td></tr>"
						strUpdatedDTCSignedApplicant = "class=""updated"""
					Case 2
						intAppType = 1
						strAppType = "<tr class=""updated""><td style=""font-weight:bold;"">Application Type  </td><td>DTC Lodge</td></tr>"
						strUpdatedDTCSignedApplicant = "class=""updated"""
					Case 3
						intAppType = 1
						strAppType = "<tr class=""updated""><td style=""font-weight:bold;"">Application Type  </td><td>DTC Dual</td></tr>"
						strUpdatedDTCDualSignedApplicant = "class=""updated"""
						strDPCandDTCSignedApplicant = "<tr><td style=""font-weight:bold;"" " & strUpdatedDTCDualSignedApplicant & ">DTC Dual Signed By Applicant  </td><td " & strUpdatedDTCDualSignedApplicant & ">" & Trim(objRS("subApply_subDigitalSignature1_fldSignFlag")) & "</td></tr>"
					Case 4
						intAppType = 1
						strAppType = "<tr class=""updated""><td style=""font-weight:bold;"">Application Type  </td><td>DPC NAB</td></tr>"
						strUpdatedDPCSignedApplicant = "class=""updated"""
						strDPCandDTCSignedApplicant = "<tr><td style=""font-weight:bold;"" " & strUpdatedDPCSignedApplicant & ">DPC Signed By Applicant  </td><td " & strUpdatedDPCSignedApplicant & ">" & Trim(objRS("subApply_subDigitalSignature1_fldSignFlag")) & "</td></tr>"
					Case 5
						intAppType = 2
						strAppType = "<tr class=""updated""><td style=""font-weight:bold;"">Application Type  </td><td>DTC Limit Change</td></tr>"
						strUpdatedDTCLCSignedApplicant = "class=""updated"""
					Case 6
						intAppType = 3
						strAppType = "<tr class=""updated""><td style=""font-weight:bold;"">Application Type  </td><td>DPC Limit Change</td></tr>"
						strUpdatedDPCLCSignedApplicant = "class=""updated"""
					Case 7
						intAppType = 1
						strAppType = "<tr class=""updated""><td style=""font-weight:bold;"">Application Type  </td><td>DTC NAB CiH</td></tr>"
						strUpdatedDTCSignedApplicant = "class=""updated"""
					Case 8
						intAppType = 2
						strAppType = "<tr class=""updated""><td style=""font-weight:bold;"">Application Type  </td><td>Lodge Limit Change</td></tr>"
						strUpdatedDTCLCSignedApplicant = "class=""updated"""
				End Select
			
			End If
			
			'Determine the field by the application Type (1 = DTC, DPC or CMC, 2 = Limit Change) 
			If intAppType = 1 then
				strEmployeeID = Trim(objRS("fldAppEmployeeID"))
				strFamilyName = Trim(objRS("fldAppFamilyName"))
				strGivenNames =  Trim(objRS("fldAppGivenNames"))

			Else
				strEmployeeID = Trim(objRS("subLimitChange_fldCDEmployeeID"))
				strFamilyName = Trim(objRS("subLimitChange_fldCDFamilyName"))
				strGivenNames =  Trim(objRS("subLimitChange_fldCDGivenNames"))

			End If
			
			If Not IsNull(objRS("subApply_dteAppAge")) Then
				If IsDate(objRS("subApply_dteAppAge")) Then
					strAppDOBAge = "&nbsp;&nbsp;&nbsp; <span class=""badge badge-pill badge-info"">" & DateDiff("yyyy",objRS("subApply_dteAppAge"),now()) & " Years of Age today</span>"
				Else
					strAppDOBAge = objRS("subApply_dteAppAge")
				End If
			Else
				strAppDOBAge = ""
			End If
			
			'Get the Current Limit and Format
			If Not IsNull(objRS("subLimitChange_numCurrentLimit")) Then
				If IsNumeric(objRS("subLimitChange_numCurrentLimit")) Then
					dblCurrentLimit = FormatCurrency(objRS("subLimitChange_numCurrentLimit"),0)
				End If
			End If
			
			'Get the New Limit and Format
			If Not IsNull(objRS("subLimitChange_numNewLimit")) Then
				If IsNumeric(objRS("subLimitChange_numNewLimit")) Then
					dblNewLimit = FormatCurrency(objRS("subLimitChange_numNewLimit"),0)
				End If
			End If
			
			'Get the New Transaction Limit and Format
			If Not IsNull(objRS("subLimitChange_subRequestedTransactionAmountDates_numLimitAmount")) Then
				If IsNumeric(objRS("subLimitChange_subRequestedTransactionAmountDates_numLimitAmount")) Then
					dblNewTransactionLimit = FormatCurrency(objRS("subLimitChange_subRequestedTransactionAmountDates_numLimitAmount"),0)
				End If
			End If
			
			'Get the Trasnaction Date From
			If Not IsNull(objRS("dteRequestedDateFrom")) Then
				If Len(objRS("dteRequestedDateFrom"))>11 Then
					dteTransactionDateFrom = Left(objRS("dteRequestedDateFrom"),11)
				End If
			End If
			
			'Get the Trasnaction Date To
			If Not IsNull(objRS("dteRequestedDateTo")) Then
				If Len(objRS("dteRequestedDateTo"))>11 Then
					dteTransactionDateTo = Right(objRS("dteRequestedDateTo"),11)
				End If
			End If
			
			'Get the SES Signature check box details Date To
			If Not IsNull(objRS("subDigitalSignature6_chkSES")) Then
				If objRS("subDigitalSignature6_chkSES")=1 Then
					subDigitalSignature6_chkSES = "<i class=""fa fa-check-square""></i> Checked"
				Else
					subDigitalSignature6_chkSES = "<i class=""fa fa-times""></i> Not Checked"
				End If
			End If
			
			'Get the ASFIN Signature check box details Date To
			If Not IsNull(objRS("subDigitalSignature6_chkASFIN")) Then
				If objRS("subDigitalSignature6_chkASFIN")=1 Then
					subDigitalSignature6_chkASFIN = "<i class=""fa fa-check-square""></i> Checked"
				Else
					subDigitalSignature6_chkASFIN = "<i class=""fa fa-times""></i> Not Checked"
				End If
			End If
			
			'Get the ASFIN Support Signature check box details Date To
			If Not IsNull(objRS("subDigitalSignature6_chkASFINSupport")) Then
				If objRS("subDigitalSignature6_chkASFINSupport")=1 Then
					subDigitalSignature6_chkASFINSupport = "<i class=""fa fa-check-square""></i> Checked"
				Else
					subDigitalSignature6_chkASFINSupport = "<i class=""fa fa-times""></i> Not Checked"
				End If
			End If
			
			Response.Write "<tr><td style=""font-weight:bold;"">XML ID</td><td>" & objRS("XMLApplicationID") & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">EID No</td><td>" & objRS("subLimitChange_fldCDEmployeeID") & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Applicant</td><td>" & strName & "</td></tr>" & _
				"<tr class=""updated""><td colspan=""2"">AE602 XML Application Details " & strApplicationID & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Application Version </td><td>" & Trim(objRS("AppVersion")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Applicant Age </td><td>" & Trim(objRS("subApply_dteAppAge")) & " " & strAppDOBAge & "</td></tr>" & _
				strAppType & _
				"<tr><td style=""font-weight:bold;"">EmployeeID  </td><td>" & strEmployeeID & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">FamilyName  </td><td>" & strFamilyName & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">GivenNames  </td><td>" & strGivenNames & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">EmailAddress  </td><td>" & Trim(objRS("subLimitChange_fldCDWorkEmailAddress")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">EmployeeID  </td><td>" & Trim(objRS("subLimitChange_fldCSEmployeeID")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">FamilyName  </td><td>" & Trim(objRS("subLimitChange_fldCSFamilyName")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">GivenNames  </td><td>" & Trim(objRS("subLimitChange_fldCSGivenNames")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">AppEmployeeID  </td><td>" & Trim(objRS("subLimitChange_fldDTCAppEmployeeID")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">AppFamilyName  </td><td>" & Trim(objRS("subLimitChange_fldDTCAppFamilyName")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">AppGivenNames  </td><td>" & Trim(objRS("subLimitChange_fldDTCAppGivenNames")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">JustificationForChange</td><td>" & Trim(objRS("subLimitChange_fldJustificationForChange")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">LastFourDigits  </td><td>" & Trim(objRS("subLimitChange_fldLastFourDigits")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">ChangesPermanent  </td><td>" & Trim(objRS("subLimitChange_grpChangesPermanent")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">CurrentLimit  </td><td>" & dblCurrentLimit & " <span style=""font-size:12px;"">Unformatted from application: (" & Trim(objRS("subLimitChange_numCurrentLimit")) & ")</span></td></tr>" & _
				"<tr><td style=""font-weight:bold;"">NewLimit  </td><td>" & dblNewLimit & " <span style=""font-size:12px;"">Unformatted from application: (" & Trim(objRS("subLimitChange_numNewLimit")) & ")</span></td></tr>" & _
				"<tr><td style=""font-weight:bold;"">HigherLimitTransactions  </td><td>" & Trim(objRS("subLimitChange_subCreditLimitJustification_ddlHigherLimitTransactions")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">NatureofTransaction  </td><td>" & Trim(objRS("subLimitChange_subCreditLimitJustification_ddlNatureofTransaction")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">OperationorExercise  </td><td>" & Trim(objRS("subLimitChange_subCreditLimitJustification_ddlOperationorExercise")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">ProposedCreditLimit  </td><td>" & Trim(objRS("subLimitChange_subCreditLimitJustification_ddlProposedCreditLimit")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"" " & strUpdatedDTCSignedApplicant & ">DTC Signed By Applicant  </td><td " & strUpdatedDTCSignedApplicant & ">" & Trim(objRS("subApply_subDigitalSignature1_fldSignFlag")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"" " & strUpdatedCMCSignedApplicant & ">CMC Signed By Applicant  </td><td  " & strUpdatedCMCSignedApplicant & ">" & Trim(objRS("subApply_subDigitalSignature3_fldSignFlag")) & "</td></tr>" & _
				strDPCandDTCSignedApplicant & _
				"<tr><td style=""font-weight:bold;"">DPC Signed By Supervisor  </td><td>" & Trim(objRS("subApply_subDigitalSignature2_fldSignFlag")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"" " & strUpdatedDTCLCSignedApplicant & ">DTC Limit Change Signed By Applicant  </td><td " & strUpdatedDTCLCSignedApplicant & ">" & Trim(objRS("subLimitChange_subDigitalSignature4_fldSignFlag")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">ValidDate  </td><td>" & Trim(objRS("subLimitChange_subDigitalSignature5_dteValidDate")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">CardHoldersName  </td><td>" & Trim(objRS("subLimitChange_subDigitalSignature5_fldCardHoldersName")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Account Holder CMS User ID  </td><td>" & Trim(objRS("subApply_fldAccHolderEmployeeID")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"" " & strUpdatedDPCLCSignedApplicant & ">DPC Limit Change Signed By Applicant  </td><td " & strUpdatedDPCLCSignedApplicant & ">" & Trim(objRS("subLimitChange_subDigitalSignature5_fldSignFlag")) & "</td></tr>" & _
				"<tr><td style=""font-weight:italic;"">Date_dteFrom (Old) </td><td>" & Trim(objRS("subLimitChange_subPeriodofChangeFromDateToDate_dteFrom")) & "</td></tr>" & _
				"<tr><td style=""font-weight:italic;"">ToDate_dteTo (old) </td><td>" & Trim(objRS("subLimitChange_subPeriodofChangeFromDateToDate_dteTo")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Start Date (New) Monthly Credit Limit </td><td>" & Trim(objRS("subLimitChange_subPeriodofChangeFromDateToDate_dteStart")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">End Date (New) Monthly Credit Limit </td><td>" & Trim(objRS("subLimitChange_subPeriodofChangeFromDateToDate_dteEnd")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Diners New Transaction Limit (Num Limit Amount) </td><td>" & dblNewTransactionLimit & " <span style=""font-size:12px;"">Unformatted from application: (" & Trim(objRS("subLimitChange_subRequestedTransactionAmountDates_numLimitAmount")) & ")</span></td></tr>" & _
			    "<tr><td style=""font-weight:bold;"">NAB New Transaction Limit (Num Limit Amount) </td><td>" & Trim(objRS("SubLimitChange_subRequestedTransactionAmountDates_ddlRequestedLimit")) & " <span style=""font-size:12px;"">Unformatted from application: (" & Trim(objRS("SubLimitChange_subRequestedTransactionAmountDates_ddlRequestedLimit")) & ")</span></td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Transaction Date From</td><td>" & dteTransactionDateFrom & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">Transaction Date To</td><td>" & dteTransactionDateTo & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">RequestedDateFrom</td><td>" & Trim(objRS("subLimitChange_subRequestedTransactionAmountDates_subRequestedPeriodFromDateToDate_dteRequestedDateFrom")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">RequestedDateTo</td><td>" & Trim(objRS("subLimitChange_subRequestedTransactionAmountDates_subRequestedPeriodFromDateToDate_dteRequestedDateTo")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">PurchaseOrder</td><td>" & Trim(objRS("subLimitChange_ddlRaisePurchaseOrder")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">CSValidDate</td><td>" & Trim(objRS("subLimitChange_dteCSValidDate")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">AppValidDate</td><td>" & Trim(objRS("subLimitChange_dteDTCAppValidDate")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">subSESApprover</td><td>" & Trim(objRS("subSESApprover")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">subDigitalSignature6_chkSES</td><td>" & subDigitalSignature6_chkSES & " <span style=""font-size:12px;"">Unformatted from application: (" & Trim(objRS("subDigitalSignature6_chkSES")) & ")</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">subDigitalSignature6_chkASFIN</td><td>" & subDigitalSignature6_chkASFIN & " <span style=""font-size:12px;"">Unformatted from application: (" & Trim(objRS("subDigitalSignature6_chkASFIN")) & ")</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">subDigitalSignature6_chkASFINSupport</td><td>" & subDigitalSignature6_chkASFINSupport & " <span style=""font-size:12px;"">Unformatted from application: (" & Trim(objRS("subDigitalSignature6_chkASFINSupport")) & ")</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">subDigitalSignature7_fldSAEmployeeID</td><td>" & Trim(objRS("subDigitalSignature7_fldSAEmployeeID")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">subDigitalSignature7_fldCMAFamilyName</td><td>" & Trim(objRS("subDigitalSignature7_fldCMAFamilyName")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">subDigitalSignature7_fldCMAGivenNames</td><td>" & Trim(objRS("subDigitalSignature7_fldCMAGivenNames")) & "</td></tr>" & _
				"<tr><td style=""font-weight:bold;"">subDigitalSignature7_fldSignFlag</td><td>" & Trim(objRS("subDigitalSignature7_fldSignFlag")) & "</td></tr>"
			
		Else
			Response.write "No XML Application Record for Application ID " & Request.QueryString("ApplicationID")
	   End If

	objRS.Close
	
	Response.write "</div></table>"
	
	
End Sub
%>
<head>
  
</head>

<body>
  <%
  Call ShowDetails()
  
  %>
</body>

</html>
<%

Set objRS = Nothing
Set objCon = Nothing
  
  %>