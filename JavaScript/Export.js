var bDummy = true;

function ExportHTMLToWindow(oSource, bHideHTMLTags, bTargetIsDiv, oTarget) {
//************************************************************
//Purpose: To send data to the HTML Container
//Inputs:  oSource, bHideHTMLTags, bTargetIsDiv
//Outputs: oTarget
//************************************************************
	var sHTMLToExport = oSource.innerHTML;
	var sTempHTML = '';
	var iPosition = -1;
	var i = 0;
	var aTags = new Array(Array('<DONT_EXPORT>', '<\\\/DONT_EXPORT>', '</DONT_EXPORT>'), Array('<A', '>', '>'), Array('<\\\/A', '>', '>'))
	var oRegExp;

	if (bHideHTMLTags) {
		for (i=0; i<aTags.length; i++) {
			oRegExp = eval('/' + aTags[i][0] + '/gi');
			iPosition = sHTMLToExport.search(oRegExp);
			while(iPosition > -1) {
				sTempHTML = sHTMLToExport.substr(0, iPosition);
				sHTMLToExport = sHTMLToExport.substr(iPosition);
				oRegExp = eval('/' + aTags[i][1] + '/gi');
				iPosition = sHTMLToExport.search(oRegExp);
				iPosition += aTags[i][2].length;
				sHTMLToExport = sHTMLToExport.substr(iPosition);
				sHTMLToExport = sTempHTML + sHTMLToExport;
				sTempHTML = '';
				oRegExp = eval('/' + aTags[i][0] + '/gi');
				iPosition = sHTMLToExport.search(oRegExp);
			}
		}
	}

	if (bTargetIsDiv)
		oTarget.innerHTML = sHTMLToExport;
	else
		oTarget.value = sHTMLToExport.toString();
} // End of ExportHTMLToWindow

function SendReportToExcel(sURL) {
//************************************************************
//Purpose: To open a new instance of the browser with the given
//         file as document
//Inputs:  sURL_IE, sURL_NC, sTitle, sWidth, sHeight, sScrolls, sStatusBar
//************************************************************
	if (bDummy) {
		oNewWindow = window.open('Export.asp?Dummy=1', 'Dummy', 'toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=no,resizable=no,copyhistory=no,width=0,height=0');
		window.setTimeout("oNewWindow.close()", 1000);
		window.setTimeout("window.open('" + sURL + "', 'ExportReport', 'toolbar=no,location=no,directories=no,status=yes,menubar=yes,scrollbars=yes,resizable=yes,copyhistory=no,width=640,height=480');", 2000);
	}
	else {
		oNewWindow = window.open(sURL, 'ExportReport', 'toolbar=no,location=no,directories=no,status=yes,menubar=yes,scrollbars=yes,resizable=yes,copyhistory=no,width=640,height=480');
	}
} // End of SendReportToExcel

function SendReportToPrint(sSourceContainer, sAccessKey) {
//************************************************************
//Purpose: To open a new window showing the PrintReport.asp page
//Inputs:  sSourceContainer, sAccessKey
//************************************************************
	window.open('Export.asp?Print=1&SourceContainer=' + sSourceContainer + '&AccessKey=' + sAccessKey, 'PrintReport', 'toolbar=no,location=no,directories=no,status=yes,menubar=yes,scrollbars=yes,resizable=no,copyhistory=no,width=612,height=300');
} // End of SendReportToPrint