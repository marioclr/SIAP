var bIsNetscapeForCommonLibrary = (document.all) ? 0 : 1;
var bExpanded = false;
var oNewWindow = null;

var aMonthNames = new Array('', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December');
//var aMonthNames = new Array('', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre');

function BuildDateString(sYear, sMonth, sDay) {
//************************************************************
//Purpose: To concatenate the year, month and day values using
//         the given format.
//Inputs:  sYear, sMonth, sDay
//Outputs: The formated date
//************************************************************
	return '' + sYear + sMonth + sDay;
}

function ChangeDaysListGivenTheMonthAndYear(iMonth, iYear, oList){
//************************************************************
//Purpose: To update the days list given the month.
//Inputs:  iMonth
//Outputs: oList
//************************************************************
	var oOption = null;
	var iIncrement = 0;
	if (oList.options[0].text == '')
		iIncrement = 1;

	if ((iMonth == 4) || (iMonth == 6) || (iMonth == 9) || (iMonth == 11)) {
		oOption = oList.options[30 + iIncrement];
		if (oOption)
			oList.options[30 + iIncrement] = null;
	}
	else if (iMonth == 2) {
		oOption = oList.options[30 + iIncrement];
		if (oOption)
			oList.options[30 + iIncrement] = null;
		oOption = oList.options[29 + iIncrement];
		if (oOption)
			oList.options[29 + iIncrement] = null;
		oOption = oList.options[28 + iIncrement];
		if (oOption)
			oList.options[28 + iIncrement] = null;
	}

	if ((iYear%4 == 0) && (iMonth == 2)) {
		oOption = oList.options[28 + iIncrement];
		if (! oOption) {
			oOption = new Option('29', '29', false, false);
			if (oOption)
				oList.options[28 + iIncrement] = oOption;
		}
	}
	if (iMonth != 2) {
		oOption = oList.options[28 + iIncrement];
		if (! oOption) {
			oOption = new Option('29', '29', false, false);
			if (oOption)
				oList.options[28 + iIncrement] = oOption;
		}
		oOption = oList.options[29 + iIncrement];
		if (! oOption) {
			oOption = new Option('30', '30', false, false);
			if (oOption)
				oList.options[29 + iIncrement] = oOption;
		}
	}
	if ((iMonth == 1) || (iMonth == 3) || (iMonth == 5) || (iMonth == 7) || (iMonth == 8) || (iMonth == 10) || (iMonth == 12)) {
		oOption = oList.options[30 + iIncrement];
		if (! oOption) {
			oOption = new Option('31', '31', false, false);
			if (oOption)
				oList.options[30 + iIncrement] = oOption;
		}
	}
} // End of ChangeDaysListGivenTheMonthAndYear

function DoNothing() {
//************************************************************
//Purpose: To do nothing
//************************************************************
}

function ExpandCollapseHelpSection() {
//************************************************************
//Purpose: To expand and collapse the section with the tips and
//         to switch the arrow image.
//************************************************************
	if (! bIsNetscapeForCommonLibrary) {
		if (bExpanded) {
			document.images['ArrowForHelpSectionImg'].src = 'Images/RightArrowTeal.gif';
			document.images['ArrowForHelpSectionImg'].alt = 'Ocultar tips de ayuda';
			document.all('TipsSectionDiv').style.display = 'none';
			
		}
		else {
			document.images['ArrowForHelpSectionImg'].src = 'Images/DownArrowTeal.gif';
			document.images['ArrowForHelpSectionImg'].alt = 'Mostrar tips de ayuda';
			document.all('TipsSectionDiv').style.display = '';
		}
		bExpanded= !bExpanded;
	}
} // End of ExpandCollapseHelpSection

function JSFormatNumber(dNumber, iDecimals) {
//************************************************************
//Purpose: To format a number adding comas.
//Inputs:  dNumber
//************************************************************
	var asNumber = (dNumber + '.00').split('.');
	var sNewNumber = '';
	var lExtra = 0;
	var sFloat = '';

	if (asNumber.length > 1) {
		lExtra = parseInt((asNumber[1] + '0000000000000000000').substr(iDecimals, 1));
		if (lExtra >= 4)
			lExtra = 1;
		else
			lExtra = 0;
		sFloat = ((parseInt(('1' + asNumber[1] + '0000000000000000000').substr(0, iDecimals + 1)) + lExtra) + '0000000000000000000').substr(0, iDecimals + 1);
		if (sFloat.substr(0, 1) == '2')
			asNumber[0] = '' + (parseInt(asNumber[0]) + 1);
	}
	while (asNumber[0].length > 3) {
		sNewNumber = ',' + asNumber[0].substr(asNumber[0].length - 3, 3) + sNewNumber;
		asNumber[0] = asNumber[0].substr(0, asNumber[0].length - 3);
	}
	if (asNumber.length > 1) {
		sNewNumber += '.' + sFloat.substr(1);
	}
	return asNumber[0] + sNewNumber;
} // End of JSFormatNumber

function HideShowGroupsList(oList, sItemName, oItem) {
//************************************************************
//Purpose: To hide/show the groups list depending on the
//         user type selection.
//Inputs:  oList, sItemName, oItem
//************************************************************
	var oPopupItem;

	if (oList.value == 'Tutor') {
		if (bIsNetscapeForCommonLibrary){
			if (oItem)
				oItem.visibility = 'visible';
		}
		else {
			oPopupItem = document.all(sItemName)
			if (oPopupItem)
				oPopupItem.style.visibility = 'visible';
		}
	}
	else {
		if (bIsNetscapeForCommonLibrary){
			if (oItem)
				oItem.visibility = 'hidden';
		}
		else {
			oPopupItem = document.all(sItemName)
			if (oPopupItem)
				oPopupItem.style.visibility = 'hidden';
		}
	}
} //End of HideShowGroupsList

function MarkAllCheckboxes(bValue, oCheckBoxes) {
//************************************************************
//Purpose: To check the checkboxes using the boolean value
//Inputs:  bValue
//Outputs: oCheckBoxes
//************************************************************
	var i;
	oCheckBoxes.checked = bValue;
	for (var i=0; i<oCheckBoxes.length; i++)
		oCheckBoxes[i].checked = bValue;

} // End of MarkAllCheckboxes

function MarkCheckboxesBitwise(lValue, oCheckBoxes) {
//************************************************************
//Purpose: To check the checkboxes using the bitwise value
//Inputs:  lValue
//Outputs: oCheckBoxes
//************************************************************
	var i;
	lValue = parseFloat(lValue);
	for (var i=0; i<oCheckBoxes.length; i++)
		oCheckBoxes[i].checked = ((lValue & parseFloat(oCheckBoxes[i].value)) != 0);
} //End of MarkCheckboxesBitwise

function OpenCalendarWindow(sFormName, sDateCombo) {
//************************************************************
//Purpose: To open a new instance of the browser with a
//         calendar
//Inputs:  sFormName, sDateCombo
//************************************************************
		oNewWindow = window.open('BrowserMonth.asp?HideDesc=1&FormName=' + sFormName + '&DateCombo=' + sDateCombo, 'Calendar', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=165,height=110');
} // End of OpenCalendarWindow

function OpenFileInNewWindow(sURL_IE, sURL_NC, sTitle) {
//************************************************************
//Purpose: To open a file in a new instance of the browser
//Inputs:  sURL_IE, sURL_NC, sTitle
//************************************************************
	if (! sURL_NC)
		sURL_NC = sURL_IE;
	if (bIsNetscapeForCommonLibrary)
		oNewWindow = window.open(sURL_NC, sTitle, 'toolbar=no,location=no,directories=no,status=yes,menubar=yes,scrollbars=yes,resizable=yes,copyhistory=no,width=640,height=480');
	else
		oNewWindow = window.open(sURL_IE, sTitle, 'toolbar=no,location=no,directories=no,status=yes,menubar=yes,scrollbars=yes,resizable=yes,copyhistory=no,width=640,height=480');
} // End of OpenFileInNewWindow

function OpenNewWindow(sURL_IE, sURL_NC, sTitle, sWidth, sHeight, sScrolls, sStatusBar) {
//************************************************************
//Purpose: To open a new instance of the browser with the given
//         file as document
//Inputs:  sURL_IE, sURL_NC, sTitle, sWidth, sHeight, sScrolls, sStatusBar
//************************************************************
	if (! sURL_NC)
		sURL_NC = sURL_IE;
	if (bIsNetscapeForCommonLibrary)
		if (sURL_NC.search(/export\.asp/gi) > -1) {
			oNewWindow = window.open(sURL_NC, sTitle, 'toolbar=no,location=no,directories=no,status=no,menubar=yes,scrollbars=' + sScrolls + ',resizable=' + sScrolls + ',copyhistory=no,width=' + sWidth + ',height=' + sHeight);
		} else {
			oNewWindow = window.open(sURL_NC, sTitle, 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=' + sScrolls + ',resizable=' + sScrolls + ',copyhistory=no,width=' + sWidth + ',height=' + sHeight);
		}
	else
		if (sURL_IE.search(/export\.asp/gi) > -1) {
			oNewWindow = window.open(sURL_IE, sTitle, 'toolbar=no,location=no,directories=no,status=' + sStatusBar + ',menubar=yes,scrollbars=' + sScrolls + ',resizable=' + sScrolls + ',copyhistory=no,width=' + sWidth + ',height=' + sHeight);
		} else {
			oNewWindow = window.open(sURL_IE, sTitle, 'toolbar=no,location=no,directories=no,status=' + sStatusBar + ',menubar=no,scrollbars=' + sScrolls + ',resizable=' + sScrolls + ',copyhistory=no,width=' + sWidth + ',height=' + sHeight);
		}
		//oNewWindow = window.showModalDialog(sURL_IE, sTitle, 'toolbar=no,location=no,directories=no,status=' + sStatusBar + ',menubar=no,scrollbars=' + sScrolls + ',resizable=' + sScrolls + ',copyhistory=no,width=' + sWidth + ',height=' + sHeight);
} // End of OpenNewWindow

var iTimer = 0;
function SearchForRecord(oField, sAction, sTarget) {
//************************************************************
//Purpose: To send the URL to the 'Search Record' frame
//Inputs:  oField, sAction, sTarget
//************************************************************
	var sValue = '0000000000' + oField.value;
	sValue = sValue.substr(sValue.length - oField.size, oField.size);
	if (iTimer == 0)
		SearchRecord(sValue, sAction, 'SearchRecordIFrame', sTarget);
} // End of SearchForRecord

function SearchRecord(sRecordID, sAction, sFrameName, sTargetField) {
//************************************************************
//Purpose: To send the URL to the 'Search Record' frame
//Inputs:  sRecordID, sAction, sFrameName
//Outputs: sTargetField
//************************************************************
	var oIFrame = eval('document.' + sFrameName + '.location')
	if (sRecordID == '') {
		alert('Favor de introducir un valor para realizar la búsqueda.');
	} else {
		if (oIFrame)
			oIFrame.href = 'SearchRecord.asp?Action=' + sAction + '&RecordID=' + sRecordID + '&TargetField=' + sTargetField;
	}
} // End of SearchRecord

function SetDateCombos(sYear, sMonth, sDay, oYearCombo, oMonthCombo, oDayCombo) {
//************************************************************
//Purpose: To select the specified date in the Date Combos
//Inputs:  sYear, sMonth, sDay
//Outputs: oYearCombo, oMonthCombo, oDayCombo
//************************************************************
	var i;
	var oDate = new Date();

	if (oYearCombo) {
		if (sYear == '')
			sYear = oDate.getFullYear() + '';
		for (i=0; i<oYearCombo.length; i++)
			if (oYearCombo[i].value == sYear) {
				oYearCombo[i].selected = true;
				i = oYearCombo.length;
			}
			else
				oYearCombo[i].selected = false;
	}

	if (oMonthCombo) {
		if (sMonth == '')
			if (oDate.getMonth() < 9)
				sMonth = '0' + (oDate.getMonth() + 1);
			else
				sMonth = (oDate.getMonth() + 1) + '';
		for (i=0; i<oMonthCombo.length; i++)
			if (oMonthCombo[i].value == sMonth) {
				oMonthCombo[i].selected = true;
				i = oMonthCombo.length;
			}
			else
				oMonthCombo[i].selected = false;
	}

	if (oDayCombo) {
		if (sDay == '')
			if (oDate.getDate() < 10)
				sDay = '0' + oDate.getDate();
			else
				sDay = oDate.getDate() + '';
		for (i=0; i<oDayCombo.length; i++)
			if (oDayCombo[i].value == sDay) {
				oDayCombo[i].selected = true;
				i = oDayCombo.length;
			}
			else
				oDayCombo[i].selected = false;
	}

} // End of SetDateCombos

function SetHiddenValueForCheckBox(bChecked, oHiddenField) {
//************************************************************
//Purpose: To set the value for the specified hidden field
//         depending if the check box is checked
//Inputs:  bChecked
//Outputs: oHiddenField
//************************************************************
	if (oHiddenField) {
		if (bChecked)
			oHiddenField.value = 1;
		else
			oHiddenField.value = 0;
	}
} // End of SetHiddenValueForCheckBox

function ToogleIFrameContents(sURL, sTargetName, vImage) {
	var oIFrame = eval('document.' + sTargetName + 'IFrm');
	var oDiv = document.all[sTargetName + 'Div'];

	if (oIFrame) {
		if (sURL == oIFrame.location) {
			oIFrame.location = 'about:blank';
			oDiv.style.display = 'none';
		}
		else {
			oIFrame.location = sURL;
			oDiv.style.display = '';
		}
		ToogleImage(vImage, 'Images/BtnExpand.gif', 'Images/BtnCollapse.gif');
	}
}