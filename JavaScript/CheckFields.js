var N_NO_RANK_FLAG = 0;
var N_MINIMUM_ONLY_FLAG = 1;
var N_MAXIMUM_ONLY_FLAG = 2;
var N_BOTH_FLAG = 3;

var N_OPEN_FLAG = 0;			// Minimum and maximum are open
var N_MINIMUM_OPEN_FLAG = 1;	// Maximum is closed
var N_MAXIMUM_OPEN_FLAG = 2;	// Minimum is closed
var N_CLOSED_FLAG = 3;			// Minimum and maximum are closed

function CheckIntegerValue(oField, sFieldDescription, iFlags, iIntervalType, iMinimumValue, iMaximumValue) {
//************************************************************
//Purpose: To check if the given value is a valid integer and
//         if it's in the given rank.
//Inputs:  oField, sFieldDescription, iFlags, iIntervalType, iMinimumValue, iMaximumValue
//Outputs: A boolean
//************************************************************
	var sTemp;
	if (oField.value.length == 0) {
		alert('Favor de introducir ' + sFieldDescription + '.');
		oField.focus();
		return false;
	}
	else {
		sTemp = oField.value;
		if (isNaN(parseInt(sTemp))) {
			alert('Favor de introducir un valor numérico para ' + sFieldDescription + '.');
			oField.focus();
			return false;
		}
		else {
			sTemp = parseInt(sTemp);
			switch (iFlags) {
				case N_MINIMUM_ONLY_FLAG:
					if ((iIntervalType == N_CLOSED_FLAG) || (iIntervalType == N_MAXIMUM_OPEN_FLAG)){
						if (sTemp < iMinimumValue) {
							alert('Favor de introducir un valor mayor o igual a ' + iMinimumValue + ' para ' + sFieldDescription + '.');
							oField.focus();
							return false;
						}
					}
					else {
						if (sTemp <= iMinimumValue) {
							alert('Favor de introducir un valor mayor a ' + iMinimumValue + ' para ' + sFieldDescription + '.');
							oField.focus();
							return false;
						}
					}
					break;
				case N_MAXIMUM_ONLY_FLAG:
					if ((iIntervalType == N_CLOSED_FLAG) || (iIntervalType == N_MINIMUM_OPEN_FLAG)) {
						if (sTemp > iMaximumValue) {
							alert('Favor de introducir un valor menor o igual a ' + iMaximumValue + ' para ' + sFieldDescription + '.');
							oField.focus();
							return false;
						}
					}
					else {
						if (sTemp >= iMaximumValue) {
							alert('Favor de introducir un valor menor a ' + iMaximumValue + ' para ' + sFieldDescription + '.');
							oField.focus();
							return false;
						}
					}
					break;
				case N_BOTH_FLAG:
					switch (iIntervalType) {
						case N_OPEN_FLAG:
							if ((sTemp <= iMinimumValue) || (sTemp >= iMaximumValue)) {
								alert('Favor de introducir un valor entre ' + iMinimumValue + ' y ' + iMaximumValue + ' para ' + sFieldDescription + '.');
								oField.focus();
								return false;
							}
							break;
						case N_MINIMUM_OPEN_FLAG:
							if ((sTemp <= iMinimumValue) || (sTemp > iMaximumValue)) {
								alert('Favor de introducir un valor entre ' + iMinimumValue + ' y ' + iMaximumValue + ' para ' + sFieldDescription + '.');
								oField.focus();
								return false;
							}
							break;
						case N_MAXIMUM_OPEN_FLAG:
							if ((sTemp < iMinimumValue) || (sTemp >= iMaximumValue)) {
								alert('Favor de introducir un valor entre ' + iMinimumValue + ' y ' + iMaximumValue + ' para ' + sFieldDescription + '.');
								oField.focus();
								return false;
							}
							break;
						case N_CLOSED_FLAG:
							if ((sTemp < iMinimumValue) || (sTemp > iMaximumValue)) {
								alert('Favor de introducir un valor entre ' + iMinimumValue + ' y ' + iMaximumValue + ' para ' + sFieldDescription + '.');
								oField.focus();
								return false;
							}
							break;
					}
					break;
			}
		}
	}
	return true;
} // End of CheckIntegerValue

function CheckFloatValue(oField, sFieldDescription, iFlags, iIntervalType, iMinimumValue, iMaximumValue) {
//************************************************************
//Purpose: To check if the given value is a valid float and
//         if it's in the given rank.
//Inputs:  oField, sFieldDescription, iFlags, iIntervalType, iMinimumValue, iMaximumValue
//Outputs: A boolean
//************************************************************
	var sTemp;

	if (oField.value.length == 0) {
		alert('Favor de introducir ' + sFieldDescription + '.');
		oField.focus();
		return false;
	}
	else {
		sTemp = oField.value.replace(/,/g, '.');
		//sTemp = oField.value.replace(/\x2E/g, ',');
		if (! isNaN(parseFloat(oField.value)))
			oField.value = sTemp;

		if (isNaN(parseFloat(oField.value))) {
			alert('Favor de introducir un valor numérico para ' + sFieldDescription + '.');
			oField.focus();
			return false;
		}
		else {
			oField.value = parseFloat(oField.value);
			switch (iFlags) {
				case N_MINIMUM_ONLY_FLAG:
					if ((iIntervalType == N_CLOSED_FLAG) || (iIntervalType == N_MAXIMUM_OPEN_FLAG)){
						if (parseFloat(oField.value) < iMinimumValue) {
							alert('Favor de introducir un valor mayor o igual a ' + iMinimumValue + ' para ' + sFieldDescription + '.');
							oField.focus();
							return false;
						}
					}
					else {
						if (parseFloat(oField.value) <= iMinimumValue) {
							alert('Favor de introducir un valor mayor a ' + iMinimumValue + ' para ' + sFieldDescription + '.');
							oField.focus();
							return false;
						}
					}
					break;
				case N_MAXIMUM_ONLY_FLAG:
					if ((iIntervalType == N_CLOSED_FLAG) || (iIntervalType == N_MINIMUM_OPEN_FLAG)) {
						if (parseFloat(oField.value) > iMaximumValue) {
							alert('Favor de introducir un valor menor o igual a ' + iMaximumValue + ' para ' + sFieldDescription + '.');
							oField.focus();
							return false;
						}
					}
					else {
						if (parseFloat(oField.value) >= iMaximumValue) {
							alert('Favor de introducir un valor menor a ' + iMaximumValue + ' para ' + sFieldDescription + '.');
							oField.focus();
							return false;
						}
					}
					break;
				case N_BOTH_FLAG:
					switch (iIntervalType) {
						case N_OPEN_FLAG:
							if ((parseFloat(oField.value) <= iMinimumValue) || (parseFloat(oField.value) >= iMaximumValue)) {
								alert('Favor de introducir un valor entre ' + iMinimumValue + ' y ' + iMaximumValue + ' para ' + sFieldDescription + '.');
								oField.focus();
								return false;
							}
							break;
						case N_MINIMUM_OPEN_FLAG:
							if ((parseFloat(oField.value) <= iMinimumValue) || (parseFloat(oField.value) > iMaximumValue)) {
								alert('Favor de introducir un valor entre ' + iMinimumValue + ' y ' + iMaximumValue + ' para ' + sFieldDescription + '.');
								oField.focus();
								return false;
							}
							break;
						case N_MAXIMUM_OPEN_FLAG:
							if ((parseFloat(oField.value) < iMinimumValue) || (parseFloat(oField.value) >= iMaximumValue)) {
								alert('Favor de introducir un valor entre ' + iMinimumValue + ' y ' + iMaximumValue + ' para ' + sFieldDescription + '.');
								oField.focus();
								return false;
							}
							break;
						case N_CLOSED_FLAG:
							if ((parseFloat(oField.value) < iMinimumValue) || (parseFloat(oField.value) > iMaximumValue)) {
								alert('Favor de introducir un valor entre ' + iMinimumValue + ' y ' + iMaximumValue + ' para ' + sFieldDescription + '.');
								oField.focus();
								return false;
							}
							break;
					}
					break;
			}
		}
	}
	return true;
} // End of CheckFloatValue

var bCheckNewItemDone = false;

function CheckItemToChange(sFormName, sSourceFieldName, sIDTargetFieldName, sNameTargetFieldName) {
	var oSourceField = eval('document.' + sFormName + '.' + sSourceFieldName);
	var oIDTargetField = eval('document.' + sFormName + '.' + sIDTargetFieldName);
	var oNameTargetField = eval('document.' + sFormName + '.' + sNameTargetFieldName);
	var aFieldValue;

	if (oSourceField) {
		if (oSourceField.value != '') {
			aFieldValue = oSourceField.value.split(';;;');
			oIDTargetField.value = aFieldValue[0];
			oNameTargetField.value = aFieldValue[1];
			oSourceField.value = '';
			bCheckNewItemDone = true;
		}
		else {
			if (! bCheckNewItemDone)
				window.setTimeout('CheckItemToChange(\'' + sFormName + '\', \'' + sSourceFieldName + '\', \'' + sIDTargetFieldName + '\', \'' + sNameTargetFieldName + '\')', 500);
			bCheckNewItemDone = false;
		}
	}
} // End of CheckItemToChange

function CheckNewItemToAdd(sFormName, sSourceFieldName, sTargetFieldName) {
	var oSourceField = eval('document.' + sFormName + '.' + sSourceFieldName);
	var oTargetField = eval('document.' + sFormName + '.' + sTargetFieldName);
	var aFieldValue;

	if (oSourceField) {
		if (oSourceField.value != '') {
			if (oTargetField.id.search(/lst/gi) != -1) {
				aFieldValue = oSourceField.value.split(';;;');
				AddItemToList(aFieldValue[1], aFieldValue[0], null, oTargetField);
			}
			else {
				oTargetField.value = oSourceField.value + oTargetField.value;
			}
			oSourceField.value = '';
			bCheckNewItemDone = true;
		}
		else {
			if (! bCheckNewItemDone)
				window.setTimeout('CheckNewItemToAdd(\'' + sFormName + '\', \'' + sSourceFieldName + '\', \'' + sTargetFieldName + '\')', 500);
			bCheckNewItemDone = false;
		}
	}
} // End of CheckNewItemToAdd
