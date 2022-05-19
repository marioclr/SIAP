var bIsNetscapeForPopupItem = (document.all) ? 0 : 1;

function HideDisplay(oItemName) {
	var oItem = oItemName.style;

	if (oItem)
		oItem.display = 'none';
}

function HidePopupItem(sItemName, oItem) {
	var oPopupItem;

	if (bIsNetscapeForPopupItem){
		if (oItem)
			oItem.visibility = 'hidden';
	}
	else {
		oPopupItem = document.all(sItemName)
		if (oPopupItem)
			oPopupItem.style.visibility = 'hidden';
	}
}

function IsDisplayed(oItemName) {
	var oItem = null;
	var sState = '';

	if (oItemName.id)
		oItem = oItemName.style;
	else
		oItem = document.all(oItemName).style;

	if (oItem)
		sState = oItem.display;

	return ((sState == 'inline') || (sState == ''));
}

function IsHidden(sItemName, oItem) {
	var oPopupItem;
	var sState = '';

	if (bIsNetscapeForPopupItem)
		if (oItem != null)
			sState = oItem.visibility;
		else {}
	else {
		oPopupItem = document.all(sItemName);

		if (oPopupItem != null)
			sState = oPopupItem.style.visibility;
	}
	return ((sState == 'hidden') || (sState == 'hide') || (sState == ''));
}

function MovePopupItem(sItemName, oItem, iPosX, iPosy) {
	var oPopupItem;

	if (bIsNetscapeForPopupItem){
		if (oItem) {
			oItem.top = iPosy;
			oItem.left = iPosX;
		}
	}
	else {
		oPopupItem = document.all(sItemName)
		if (oPopupItem) {
			oPopupItem.style.top = iPosy;
			oPopupItem.style.left = iPosX;
		}
	}
}

function ShowDisplay(oItemName) {
	var oItem = oItemName.style;

	if (oItem)
		oItem.display = 'inline';
}

function ShowPopupItem(sItemName, oItem, bScroll) {
	var oPopupItem;

	if (bIsNetscapeForPopupItem){
		if (oItem) {
			oItem.visibility = 'visible';
			if (bScroll) {
				oItem.top = parseInt(oItem.top) + document.body.scrollTop;
				oItem.left = parseInt(oItem.left) + document.body.scrollLeft;
			}
		}
	}
	else {
		oPopupItem = document.all(sItemName)
		if (oPopupItem) {
			oPopupItem.style.visibility = 'visible';
			if (bScroll) {
				if (oPopupItem.style.top != '')
					oPopupItem.style.top = parseInt(oPopupItem.style.top) + document.body.scrollTop;
				if (oPopupItem.style.left != '')
					oPopupItem.style.left = parseInt(oPopupItem.style.left) + document.body.scrollLeft;
			}
		}
	}
}

function ToggleDisplay(oItemName) {
	if (IsDisplayed(oItemName))
		HideDisplay(oItemName);
	else
		ShowDisplay(oItemName);
}

function TogglePopupMenu(sItemName, oItem, bScroll) {
	if (IsHidden(sItemName, oItem))
		ShowPopupItem(sItemName, oItem, bScroll);
	else
		HidePopupItem(sItemName, oItem);
}

/*
Notes:
1. The next style must be defined at the top of the page
<STYLE TYPE="text/css"><!--
	.ClassPopupItem	{
					position: absolute;
					visibility: hidden;
					z-index: 100;
					}
--></STYLE>

2. Define a DIV for the item to be hidden and shown as follows:
	<DIV ID="IDMyItemDiv" CLASS="ClassPopupItem" STYLE="width: 120px; height: 80px; left: 127px; top: 93px;"
		onMouseOver="ShowPopupItem('IDMyItemDiv', document.IDMyItemDiv, false)"
		onMouseOut="HidePopupItem('IDMyItemDiv', document.IDMyItemDiv)">
3. Define a LAYER inside the DIV (defined in the previous step) as follows:
	<LAYER WIDTH="120" HEIGHT="80"
		onMouseOver="ShowPopupItem('IDMyItemDiv', parent.document.IDMyItemDiv, false)"
		onMouseOut="HidePopupItem('IDMyItemDiv', parent.document.IDMyItemDiv)">
4. Since the LAYER is defined inside the DIV, the 2nd input parameter must be referenced using the parent object.
*/