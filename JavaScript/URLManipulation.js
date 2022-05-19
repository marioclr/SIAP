var bIsNetscapeForURLManipulation = (document.all) ? 0 : 1;

var sPageName = GetFileNameFromURL(top.window.location.href);
var sURLRequest = GetRequestSectionFromURL(top.window.location.href);

function BuildURLFromForm(oForm, sExcluded) {
	var i=0;
	var sURL = '';
	var oRegExp = null;

	for (i-0; i<oForm.elements.length; i++) {
		oRegExp = eval('/,' + oForm.elements[i].name + ',/gi');
		if (((','+ sExcluded + ',').search(oRegExp) == -1) && (oForm.elements[i].name != '')) {
			sURL += oForm.elements[i].name + '=';
			sURL += oForm.elements[i].value + '&';
		}
	}
	sURL = sURL.substr(0, sURL.length - '&'.length)
	return sURL;
} //End of BuildURLFromForm

function GetFileNameFromURL(sURLString) {
	var sTemp = sURLString;

	if (sTemp.indexOf('?') > -1)
		sTemp = sTemp.substr(0, sTemp.indexOf('?'));
	if (sTemp.lastIndexOf('\\') > -1)
		sTemp = sTemp.substr(sTemp.lastIndexOf('\\') + '\\'.length);
	else
		sTemp = sTemp.substr(sTemp.lastIndexOf('/') + '/'.length);

	return(sTemp);
} //End of GetFileNameFromURL

function GetRequestSectionFromURL(sURLString) {
	if (sURLString.indexOf('?') > -1)
		return(sURLString.substr((sURLString.indexOf('?') + '?'.length)));
	else
		return('');
} // End of GetRequestSectionFromURL

function GetParameterFromURL(sParameter, sDefaultValue) {
	var iStartPosition = -1;
	var iEndPosition = -1;
	var sParameterValue = '';

	iStartPosition = sURLRequest.indexOf('&' + sParameter + '=');
	if (iStartPosition == -1)
		if (sURLRequest.indexOf(sParameter + '=') == 0)
			iStartPosition = 0;
				
	if (iStartPosition > -1) {
		if (iStartPosition > 0)
			iStartPosition = iStartPosition + '&'.length + sParameter.length + '='.length;
		else
			iStartPosition = iStartPosition + sParameter.length + '='.length;
		iEndPosition = sURLRequest.indexOf('&', iStartPosition);
		if (iEndPosition > -1)
			sParameterValue = sURLRequest.substring(iStartPosition, iEndPosition);
		else
			sParameterValue = sURLRequest.substr(iStartPosition);
	}

	sParameterValue = sParameterValue.replace(/\+/g, " ");

	if (sParameterValue.length == 0)
		return(sDefaultValue);
	else
		return(sParameterValue);
} // End of GetParameterFromURL

function RemoveParameterFromURLRequest(sParameter) {
	var iStartPosition = -1;
	var iEndPosition = -1;
	var sTempURL = sURLRequest;

	iStartPosition = sTempURL.indexOf('&' + sParameter + '=');
	if (iStartPosition == -1)
		if (sTempURL.indexOf(sParameter + '=') == 0)
			iStartPosition = 0;

	if (iStartPosition > -1) {
		if (iStartPosition > 0)
			iEndPosition = sTempURL.indexOf('&', iStartPosition + '&'.length);
		else
			iEndPosition = sTempURL.indexOf('&', iStartPosition);

	if (iEndPosition > -1)
		sTempURL = sTempURL.substring(0, iStartPosition) + sTempURL.substr(iEndPosition);
	else
		sTempURL = sTempURL.substring(0, iStartPosition);
	}

	if (sTempURL.indexOf('&') == 0)
		sTempURL = sTempURL.substr(1);

	return (sTempURL);
} // End of RemoveParameterFromURLRequest

function ReplaceValueInURLRequest(sParameter, sNewValue) {
	var sTempURL = '';

	sTempURL = RemoveParameterFromURLRequest(sParameter);

	if (sTempURL.length > 0)
		return (sTempURL + '&' + sParameter + '=' + sNewValue);
	else
		return (sParameter + '=' + escape(sNewValue));
} // End of ReplaceValueInURLRequest