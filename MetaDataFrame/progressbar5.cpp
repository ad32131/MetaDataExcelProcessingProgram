#include "progressbar5.h"


progressbar5::progressbar5() {
	progressint = 0;
}


BOOL progressbar5::setProgress() {
	SendMessage(PROGRESSBAR_HANDLE, PBM_SETPOS, (WPARAM)progressint, (LPARAM)NULL);
	return TRUE;
}

BOOL progressbar5::getHandle(HWND parent, unsigned int IDC_INT) {
	PROGRESSBAR_HANDLE = GetDlgItem(parent, IDC_INT);

	return TRUE;
}

BOOL progressbar5::setProgressinit() {
	SendMessage(PROGRESSBAR_HANDLE, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
	return TRUE;
}

BOOL progressbar5::setProgressPlus() {
	progressint++;
	setProgress();
	return TRUE;
}

BOOL progressbar5::setProgresszero() {
	progressint = 0;
	setProgress();
	return TRUE;
}

BOOL progressbar5::setProgressFull() {
	progressint = 100;
	setProgress();
	return TRUE;
}

BOOL progressbar5::setProgressInt(unsigned int inputInt) {
	progressint = inputInt;
	setProgress();
	return TRUE;
}


//SendMessage(IDC_PROGRESS1_H, PBM_SETPOS, (WPARAM)progressint, (LPARAM)NULL))