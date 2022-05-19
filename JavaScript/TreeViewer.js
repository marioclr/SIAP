<!--
	var bIsNetscapeForTreeViewer = (document.all) ? 0 : 1;
	var S_IMAGE_NAME = 0;
	var S_CLOSED_IMAGE_SOURCE = 1;
	var S_OPENED_IMAGE_SOURCE = 2;

	function SwitchDisplayForTreeItem(oItem, aImagesToSwitch) {
	// aImagesToSwitch is an array which elements must be 3-elements arrays. Those must have:
	// - the name of the Image to switch
	// - the source of the image to display when closed
	// - the source of the image to display when opened
		if (oItem != null)
			if (oItem.style.display.length == 0) {
				oItem.style.display = 'none';
				for (var i=0; i<aImagesToSwitch.length; i++)
					SwitchImageForTreeItem(aImagesToSwitch[i][S_IMAGE_NAME], aImagesToSwitch[i][S_CLOSED_IMAGE_SOURCE]);
			}
			else {
				oItem.style.display = '';
				for (var i=0; i<aImagesToSwitch.length; i++)
					SwitchImageForTreeItem(aImagesToSwitch[i][S_IMAGE_NAME], aImagesToSwitch[i][S_OPENED_IMAGE_SOURCE]);
			}
	} // End of SwitchDisplayForTreeItem

	function SwitchImageForTreeItem(sImageName, sNewImageSource) {
		var oImage;
		oImage = document.images[sImageName];
		if (oImage != null)
			oImage.src = sNewImageSource;
	} // End of SwitchImageForTreeItem
//-->