var aImages = new Array();

function LoadImages(asImagesSource) {
	var i, j;
	j = aImages.length;

	for(i=0; i<asImagesSource.length; i++) {
		aImages[j] = new Image; 
		aImages[j].src= asImagesSource[i];
		j++;
	}
}

function SwapImage(vImageName, sNewImageSource) {
	var oImage;

	if (typeof vImageName == 'string')
		oImage = document.images[vImageName];
	else
		oImage = vImageName;

	if (oImage != null)
		oImage.src = sNewImageSource;
}

function ToogleImage(vImageName, sImageSource1, sImageSource2) {
	var oImage;
	var sRegExp = eval('/' + sImageSource1.replace(/\//g, '\\/').replace(/\./g, '\\.') + '/g');

	if (typeof vImageName == 'string')
		oImage = document.images[vImageName];
	else
		oImage = vImageName;

	if (oImage != null)
		if (oImage.src.search(sRegExp) != -1)
			oImage.src = sImageSource2;
		else
			oImage.src = sImageSource1;
}