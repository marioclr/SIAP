var S_MESSAGE = "Sistema Integral de Administración del Personal";
var bMaskKey = false;
var bShiftKey = false;
var oClickTime = 0;

function HelpDown(e) {
//************************************************************
//Purpose: To catch the F1 key event and do nothing.
//************************************************************
	return false;
}

function KeyDown() {
//************************************************************
//Purpose: To catch the key down event and do nothing.
//************************************************************
	if (event.keyCode == 93) {
		alert(S_MESSAGE);
		event.returnValue = false;
		return false;
	}

	if (! bShiftKey && (event.keyCode == 222)) {
		event.returnValue = false;
		return false;
	}

	if (event.keyCode == 16) {
		bShiftKey = true;
		event.returnValue = false;
		return false;
	}
	if ((bMaskKey) && (event.keyCode != 16)) {
		event.returnValue = false;
		return false;
	}

	if (event.keyCode == 17) {
		bMaskKey = true;
		event.returnValue = false;
		return false;
	}
	if ((bMaskKey) && (event.keyCode != 17)) {
		event.returnValue = false;
		return false;
	}
}

function KeyDownForRightClick() {
//************************************************************
//Purpose: To catch the key down event and do nothing.
//************************************************************
	if (event.keyCode == 93) {
		alert(S_MESSAGE);
		event.returnValue = false;
		return false;
	}
	if (event.keyCode == 27) {
		window.close();
	}
}

function KeyUp() {
//************************************************************
//Purpose: To catch the key up event and do nothing.
//************************************************************
	if (event.keyCode == 16) {
		bShiftKey = false;
		event.returnValue = false;
		return false;
	}
	if (event.keyCode == 17) {
		bMaskKey = false;
		event.returnValue = false;
		return false;
	}
}

function MouseDown(e) {
//************************************************************
//Purpose: To catch the click event and do nothing.
//************************************************************
	var oDate = new Date();
	if (document.all)
		if (event.button == 1) {
			if ((oDate.getTime() - oClickTime) < 1000) {
				oClickTime = 0;
				alert('Doble click');
				return false;
			} else {
				oClickTime = oDate.getTime();
				
			}
		}
}
