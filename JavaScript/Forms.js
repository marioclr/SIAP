var bIsNetscapeForForms = (document.all) ? 0 : 1;

function CheckAllItemsFromCheckboxes(oCheckbox) {
//************************************************************
//Purpose: To check all the items from the checkbox collection 
//Inputs:  oCheckbox
//Outputs: oCheckbox
//************************************************************
	var i;

	for (i=0; i<oCheckbox.length; i++)
		oCheckbox[i].checked = true;
} // End CheckAllItemsFromCheckboxes

function CheckCheckBoxSelection(oCheckboxesList) {
//************************************************************
//Purpose: To check if a checkbox is checked before send
//         the information to process.
//Inputs:  oCheckboxesList
//************************************************************
	if (oCheckboxesList.checked)
		return true;
	for (var i=0; i<oCheckboxesList.length; i++)
		if (oCheckboxesList[i].checked)
			return true;

	alert('Favor de seleccionar al menos un registro.');
	return false;
} // End of CheckCheckBoxSelection

function CheckRadioSelection(oRadioButtonsList) {
//************************************************************
//Purpose: To check if a radio button is selected before send
//         the information to process.
//Inputs:  oRadioButtonsList
//************************************************************
	if (! oRadioButtonsList)
		return false;

	if (oRadioButtonsList.checked)
		return true;
	for (var i=0; i<oRadioButtonsList.length; i++)
		if (oRadioButtonsList[i].checked)
			return true;

	alert('Favor de seleccionar un registro.');
	return false;
} // End of CheckRadioSelection

function GetCheckBoxSelection(oCheckboxesList) {
//************************************************************
//Purpose: To check if a checkbox is checked before send
//         the information to process.
//Inputs:  oCheckboxesList
//************************************************************
	var sValues = '-760211';

	if (oCheckboxesList.checked)
		sValues = oCheckboxesList.value;
	for (var i=0; i<oCheckboxesList.length; i++)
		if (oCheckboxesList[i].checked)
			sValues += ';;;' + oCheckboxesList[i].value;

	return sValues;
} // End of GetCheckBoxSelection

function SendURLValuesToForm(sURL, oForm) {
//************************************************************
//Purpose: To set the element values of a form using an URL
//Inputs:  sURL, oForm
//Outputs: oForm
//************************************************************
	if (sURL.length > 0) {
		var aURLElements = sURL.split('&');
		var aURLElement;
		var sValue = '';
		var oField = null;
		var oFieldYear = null;
		var oFieldMonth = null;

		if (oForm) {
			for (var i=0; i<aURLElements.length; i++) {
				aURLElement = aURLElements[i].split('=');
				oField = eval('oForm.' + aURLElement[0]);
				sValue = aURLElement[1];
				if (oField) {
					if (oField[0]) {
						if (oField[0].id.search(/Rd/ig) != -1)
							SetRadioButtonsValue(sValue, oField);
						if (oField[0].id.search(/ChPm/ig) != -1)
							MarkCheckboxesBitwise(parseInt(sValue), oField);
						if (oField[0].id.search(/Chk/ig) != -1) {
							UncheckAllItemsFromCheckboxes(oField);
							for (var j=0; j<oField.length; j++) {
								var sRegExp = eval('/\,' + oField[j].value + '\,/gi');
								if ((',' + sValue + ',').search(sRegExp) != -1)
									oField[j].checked = true;
							}
						}
					}

					if (oField.id) {
						if (oField.id.search(/Cmb/ig) != -1) {
							if (oField.id.search(/DayCmb/ig) != -1) {
								oFieldYear = eval('oForm.' + aURLElement[0].replace(/DayCmb/ig, 'YearCmb'));
								oFieldMonth = eval('oForm.' + aURLElement[0].replace(/DayCmb/ig, 'MonthCmb'));
								if (oFieldYear && oFieldMonth)
									ChangeDaysListGivenTheMonthAndYear(oFieldMonth.value, oFieldYear.value, oField);
							}
							SelectItemByValue(sValue, false, oField);
						}
						if (oField.id.search(/Lst/ig) != -1) {
							var aValues = sValue.split(',');
							for (var j=0; j<aValues.length; j++)
								SelectListItemByValue(aValues[j], true, oField);
						}
						if (oField.id.search(/Hdn/ig) != -1)
							oField.value = sValue;
						if (oField.id.search(/Txt/ig) != -1)
							oField.value = sValue;
						if (oField.id.search(/Pwd/ig) != -1)
							oField.value = sValue;
						if (oField.id.search(/TxtArea/ig) != -1)
							oField.value = oField.value.replace(/<BR \/>/ig, '\n').replace(/%0D%0A/ig, '\n').replace(/%250D%250A/ig, '\n');
						if (oField.id.search(/Chk/ig) != -1)
							oField.checked = (oField.value == sValue);
					}
				}
			}
		}
	}
} // End of SendURLValuesToForm

function SetCheckboxesValue(sValue, oCheckboxes) {
//************************************************************
//Purpose: To check the box with the given value
//Inputs:  sValue
//Outputs: oCheckboxes
//************************************************************
	if (oCheckboxes) {
		if (oCheckboxes.value == sValue)
			oCheckboxes.checked = true;
		for (var i=0; i<oCheckboxes.length; i++) {
			if (oCheckboxes[i].value == sValue)
				oCheckboxes[i].checked = true;
		}
	}
} // End of SetCheckboxesValue

function SetHiddenValueForCheckBox(bChecked, oHiddenField) {
//************************************************************
//Purpose: To set the value of a hidden field using a boolean
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

function SetRadioButtonsValue(sValue, oRadioButtons) {
//************************************************************
//Purpose: To check the radio button with the given value
//Inputs:  sValue
//Outputs: oRadioButtons
//************************************************************
	if (oRadioButtons) {
		oRadioButtons.checked = (oRadioButtons.value == sValue);
		for (var i=0; i<oRadioButtons.length; i++) {
			oRadioButtons[i].checked = (oRadioButtons[i].value == sValue);
		}
	}
} // End of SetRadioButtonsValue

function UncheckCheckboxesValue(sValue, oCheckboxes) {
//************************************************************
//Purpose: To uncheck the box with the given value
//Inputs:  sValue
//Outputs: oCheckboxes
//************************************************************
	if (oCheckboxes) {
		if (oCheckboxes.value == sValue)
			oCheckboxes.checked = false;
		for (var i=0; i<oCheckboxes.length; i++) {
			if (oCheckboxes[i].value == sValue)
				oCheckboxes[i].checked = false;
		}
	}
} // End of UncheckCheckboxesValue

function UncheckAllItemsFromCheckboxes(oCheckbox) {
//************************************************************
//Purpose: To uncheck all the items from the checkbox collection 
//Inputs:  oCheckbox
//Outputs: oCheckbox
//************************************************************
	var i;

	for (i=0; i<oCheckbox.length; i++)
		oCheckbox[i].checked = false;
} // End UncheckAllItemsFromCheckboxes