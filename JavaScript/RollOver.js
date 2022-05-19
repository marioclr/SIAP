var bIsNetscapeForRollOver = (document.all) ? 0 : 1;

function SwitchItemBGColor(oItem, sBGColor) {
	if (oItem) {
		if (bIsNetscapeForRollOver)
			oItem.backgroundColor = sBGColor;
		else
			oItem.style.backgroundColor = sBGColor;
	}
}