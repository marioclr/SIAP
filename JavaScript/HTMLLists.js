function AddItemToList(sText, sValue, aEmptyOption, oList) {
//************************************************************
//Purpose: To add a new option at the end of the list
//Inputs:  sText, sValue, aEmptyOption
//Outputs: oList
//************************************************************
	var oNewOption = null;

	if (oList.length == 1)
		if (aEmptyOption)
			if ((oList.options[0].text == aEmptyOption[0]) || (oList.options[0].value == aEmptyOption[1]))
				oList.options[0] = null;

	//if ((sText != '') && (sValue != ''))
		oNewOption = new Option(sText, sValue, false, false);
	if (oNewOption)
		oList.options[oList.length] = oNewOption;
} // End of AddItemToList

function ChangeSelectedItem(sText, sValue, oList) {
//************************************************************
//Purpose: To replace the selected list options with a new one
//Inputs:  sText, sValue
//Outputs: oList
//************************************************************
	var i;

	for (i=0; i<oList.length; i++)
		if (oList.options[i].selected) {
			oList.options[i] = new Option(sText, sValue, false, true);
			i = oList.length;
		}
} // End of ChangeSelectedItem

function ChangeSelectedItems(sText, sValue, oList) {
//************************************************************
//Purpose: To replace the selected list options with a new one
//Inputs:  sText, sValue
//Outputs: oList
//************************************************************
	var i;

	for (i=0; i<oList.length; i++)
		if (oList.options[i].selected)
			oList.options[i] = new Option(sText, sValue, false, true);
} // End of ChangeSelectedItems

function CountSelectedItems(oSourceList) {
//************************************************************
//Purpose: To get the number of the selected items
//Inputs:  oSourceList
//Outputs: A number with the selected items
//************************************************************
	var i;
	var iCounter = 0;

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList[i].selected)
			iCounter++;

	return iCounter;
} // End CountSelectedItems

function GetSelectedItems(oSourceList) {
//************************************************************
//Purpose: To get a list of the selected items
//Inputs:  oSourceList
//Outputs: A string with the indexes of the selected items
//************************************************************
	var i;
	var sItems = '';

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList[i].selected)
			sItems += i + ',';

	if (sItems != '')
		sItems = sItems.substr(0, (sItems.length - ','.length));
	return sItems;
} // End GetSelectedItems

function GetSelectedText(oSourceList) {
//************************************************************
//Purpose: To get the first selected text
//Inputs:  oSourceList
//Outputs: A string with the the selected text
//************************************************************
	var i;

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList[i].selected)
			return oSourceList[i].text;
} // End GetSelectedText

function GetSelectedTexts(oSourceList) {
//************************************************************
//Purpose: To get a list of the selected texts
//Inputs:  oSourceList
//Outputs: A string with the indexes of the selected texts
//************************************************************
	var i;
	var sTexts = '';

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList[i].selected)
			sTexts += oSourceList[i].text + ';;;';

	if (sTexts != '')
		sTexts = sTexts.substr(0, (sTexts.length - ';;;'.length));
	return sTexts;
} // End GetSelectedTexts

function GetSelectedValues(oSourceList) {
//************************************************************
//Purpose: To get a list of the selected values
//Inputs:  oSourceList
//Outputs: A string with the indexes of the selected values
//************************************************************
	var i;
	var sValues = '';

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList[i].selected)
			sValues += oSourceList[i].value + ';;;';

	if (sValues != '')
		sValues = sValues.substr(0, (sValues.length - ';;;'.length));
	return sValues;
} // End GetSelectedValues

function MoveItemsBetweenLists(aEmptyOption, oSourceList, oTargetList) {
//************************************************************
//Purpose: To move the selected items from the first list to
//         the second one. It also adds an empty element in
//         case one of the lists gets empty.
//Inputs:  aEmptyOption, oSourceList, oTargetList
//Outputs: oSourceList, oTargetList
//************************************************************
	var i;

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList.options[i].selected)
			if ((! OptionExists(oTargetList, oSourceList.options[i].text, oSourceList.options[i].value)) && ((oSourceList.options[i].text != aEmptyOption[0]) || (oSourceList.options[i].value != aEmptyOption[1])))
				AddItemToList(oSourceList.options[i].text, oSourceList.options[i].value, aEmptyOption, oTargetList);

	RemoveSelectedItemsFromList(aEmptyOption, oSourceList)
	oSourceList.selectedIndex = -1;
	oTargetList.selectedIndex = -1;
} // End MoveItemsBetweenLists

function MoveListItemDown(oList) {
//************************************************************
//Purpose: To move down the selected items from the list.
//Inputs:  oList
//Outputs: oList
//************************************************************
	var i;
	var sTempText = '';
	var sTempValue = '';

	for (i=oList.length-1; i>=0; i--)
		if ((oList.options[i].selected) && (i < oList.length)) {
			sTempText = oList.options[i+1].text;
			sTempValue = oList.options[i+1].value;
			oList.options[i+1].text = oList.options[i].text;
			oList.options[i+1].value = oList.options[i].value;
			oList.options[i].text = sTempText;
			oList.options[i].value = sTempValue;
			oList.options[i+1].selected = true;
			oList.options[i].selected = false;
		}
} // End of MoveListItemDown

function MoveListItemUp(oList) {
//************************************************************
//Purpose: To move up the selected items from the list.
//Inputs:  oList
//Outputs: oList
//************************************************************
	var i;
	var sTempText = '';
	var sTempValue = '';

	for (i=0; i<oList.length; i++)
		if ((oList.options[i].selected) && (i > 0)) {
			sTempText = oList.options[i-1].text;
			sTempValue = oList.options[i-1].value;
			oList.options[i-1].text = oList.options[i].text;
			oList.options[i-1].value = oList.options[i].value;
			oList.options[i].text = sTempText;
			oList.options[i].value = sTempValue;
			oList.options[i-1].selected = true;
			oList.options[i].selected = false;
		}
} // End of MoveListItemUp

function OptionExists(oList, sText, sValue) {
//************************************************************
//Purpose: To check if there is an option with the given value
//         in the list
//Inputs:  oList, sText, sValue
//Outputs: true or false
//************************************************************
	var i;

	for (i=0; i<oList.length; i++)
		if ((oList.options[i].text == sText) && (oList.options[i].value == sValue))
			return true;
	return false;
} // End OptionExists

function RemoveAllItemsFromList(aEmptyOption, oList) {
//************************************************************
//Purpose: To remove all the items from the given list
//Inputs:  oList
//************************************************************
	var i;

	for (i=0; i<oList.length; i++) {
		oList.options[i] = null;
		i--;
	}
	if (oList.length == 0)
		if (aEmptyOption)
			AddItemToList(aEmptyOption[0], aEmptyOption[1], aEmptyOption, oList);
} // End RemoveAllItemsFromList

function RemoveItemByValueFromList(iValue, aEmptyOption, oList) {
//************************************************************
//Purpose: To remove the given items from the list
//Inputs:  iValue, aEmptyOption
//Outputs: oList
//************************************************************
	var i;

	for (i=0; i<oList.length; i++) {
		if (oList.options[i].value == iValue) {
			oList.options[i] = null;
			i--;
		}
	}
	if (oList.length == 0)
		if (aEmptyOption)
			AddItemToList(aEmptyOption[0], aEmptyOption[1], aEmptyOption, oList);
} // End RemoveItemByValueFromList

function RemoveSelectedItemsFromList(aEmptyOption, oList) {
//************************************************************
//Purpose: To remove all the selected items from the given list
//Inputs:  aEmptyOption
//Outputs: oList
//************************************************************
	var i;

	for (i=0; i<oList.length; i++) {
		if (oList.options[i].selected) {
			oList.options[i] = null;
			i--;
		}
	}
	if (oList.length == 0)
		if (aEmptyOption)
			AddItemToList(aEmptyOption[0], aEmptyOption[1], aEmptyOption, oList);
} // End RemoveSelectedItemsFromList

function SelectAllItemsFromList(oList) {
//************************************************************
//Purpose: To select all the items from the list. 
//Inputs:  oList
//Outputs: oList
//************************************************************
	var i;

	for (i=0; i<oList.length; i++)
		oList[i].selected = true;
} // End SelectAllItemsFromList

function SelectItemByText(sText, bSelectAll, oSourceList) {
//************************************************************
//Purpose: To select the items in the combo which text is equal
//         to the first parameter
//Inputs:  sText, bSelectAll, oSourceList
//Outputs: oSourceList
//************************************************************
	var i;
	var bSelected;

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList[i].text == sText) {
			oSourceList[i].selected = true;
			bSelected = true;
			if (! bSelectAll)
				i = oSourceList.length;
		}
		else
			oSourceList[i].selected = false;

	if (! bSelected)
		if (oSourceList[0])
			oSourceList[0].selected = true;
} // End SelectItemByText

function SelectItemByValue(sValue, bSelectAll, oSourceList) {
//************************************************************
//Purpose: To select the items in the combo which value is equal
//         to the first parameter
//Inputs:  sValue, bSelectAll, oSourceList
//Outputs: oSourceList
//************************************************************
	var i;
	var bSelected;

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList[i].value == sValue) {
			oSourceList[i].selected = true;
			bSelected = true;
			if (! bSelectAll)
				i = oSourceList.length;
		}
		else
			oSourceList[i].selected = false;

	if (! bSelected)
		if (oSourceList[0])
			oSourceList[0].selected = true;
} // End SelectItemByValue

function SelectListItemByText(sText, bSelectAll, oSourceList) {
//************************************************************
//Purpose: To select the items in the list which text is equal
//         to the first parameter
//Inputs:  sText, bSelectAll, oSourceList
//Outputs: oSourceList
//************************************************************
	var i;

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList[i].text == sText) {
			oSourceList[i].selected = true;
			if (! bSelectAll)
				i = oSourceList.length;
		}

} // End SelectListItemByText

function SelectListItemByValue(sValue, bSelectAll, oSourceList) {
//************************************************************
//Purpose: To select the items in the list which value is equal
//         to the first parameter
//Inputs:  sValue, bSelectAll, oSourceList
//Outputs: oSourceList
//************************************************************
	var i;

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList[i].value == sValue) {
			oSourceList[i].selected = true;
			if (! bSelectAll)
				i = oSourceList.length;
		}

} // End SelectListItemByValue

function SelectSameItems(oSourceList, oTargetList) {
//************************************************************
//Purpose: To select the items in the target list based on the
//         selected items in the source list. 
//Inputs:  oSourceList
//Outputs: oTargetList
//************************************************************
	var iMaxLength = (oSourceList.length < oTargetList.length) ? oSourceList.length : oTargetList.length;
	var i;

	for (i=0; i<iMaxLength; i++)
		oTargetList[i].selected = oSourceList[i].selected;
} // End SelectSameItems

function UnselectAllItemsFromList(oList) {
//************************************************************
//Purpose: To unselect all the items from the list. 
//Inputs:  oList
//Outputs: oList
//************************************************************
	var i;

	for (i=0; i<oList.length; i++)
		oList[i].selected = false;
} // End SelectAllItemsFromList

function UnSelectItemByValue(sValue, bSelectAll, oSourceList) {
//************************************************************
//Purpose: To unselect the items in the combo which value is equal
//         to the first parameter
//Inputs:  sValue, bSelectAll, oSourceList
//Outputs: oSourceList
//************************************************************
	var i;

	for (i=0; i<oSourceList.length; i++)
		if (oSourceList[i].value == sValue) {
			oSourceList[i].selected = false;
			if (! bSelectAll)
				i = oSourceList.length;
		}
} // End UnSelectItemByValue