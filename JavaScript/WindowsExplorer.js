var bIsNetscapeForWindowsExplorer = (document.all) ? 0 : 1;

function AddImageSource(oQuestionText, oImageSrc) {
//************************************************************
//Purpose: To add an image reference.
//Inputs:  oQuestionText, oImageSrc
//Outputs: oQuestionText, oImageSrc
//************************************************************
	var sRange = null;
	if ((oQuestionText) && (oImageSrc)) {
		if (oImageSrc.value != '') {
			oQuestionText.focus();
			sRange = document.selection.createRange();
			sRange.text = '<<' + oImageSrc.value.replace(/\\/gi, '\/') + '>>';
//			oQuestionText.value += '<<' + oImageSrc.value.replace(/\\/gi, '\/') + '>>';
			oImageSrc.value = '';
		}
	}
} // End of AddImageSource

function ShowWindowsExplorer(sFormName, sFieldName, sURL, sFilter){
//************************************************************
//Purpose: To display a window that has the contents of a given
//         folder.
//Inputs:  sFormName, sFieldName, sURL, sFilter
//************************************************************
	var sTempURL = '';
	if (sURL)
		sTempURL += sURL + '&';
	if (sFilter)
		sTempURL += 'FolderFilter=' + sFilter + '&';
	
	OpenNewWindow(('FileExplorer.asp?FormName=' + sFormName + '&FieldName=' + sFieldName + "&" + sTempURL), ('FileExplorer.asp?FormName=' + sFormName + '&FieldName=' + sFieldName + "&" + sTempURL), 'WindowsExplorer', '500', '270', 'no', 'no');
} // End of ShowWindowsExplorer

function ApplyFilter(aFileNames, sFilter, oListToFilter) {
	var i;
	for (i=oListToFilter.options.length-1; i>=0; i--) {
		oListToFilter.options[i] = null;
	}
	for (i=0; i<aFileNames.length-1; i++) {
		if ((sFilter == '') || (aFileNames[i].substr(aFileNames[i].length - sFilter.length) == sFilter))
			AddItemToList(aFileNames[i], aFileNames[i], ['', ''], oListToFilter)
	}
} // End of ApplyFilter

function CheckFolderSelection(oForm){
//************************************************************
//Purpose: To check the windows explorer form fields are complete
//         before send them to be processed.
//Inputs:  oForm
//************************************************************
	if (oForm.FolderName.value == '') {
		alert('Favor de seleccionar un directorio');
		oForm.FolderName.focus();
		return false;
	}
	return true;
}

function CheckFileSelection(oForm){
//************************************************************
//Purpose: To check the windows explorer form fields are complete
//         before send them to be processed.
//Inputs:  oForm
//************************************************************
	if (oForm.FileContents) {
		if (oForm.FileContents.value == '') {
			if (oForm.CourseFile.value == '') {
				alert('Favor de seleccionar un archivo');
				oForm.CourseFile.focus();
				return false;
			}
		}
	}
	if (oForm.CourseFile.value == '') {
		alert('Favor de seleccionar un archivo');
		oForm.CourseFile.focus();
		return false;
	}
	return true;
} // End of CheckFolderSelection