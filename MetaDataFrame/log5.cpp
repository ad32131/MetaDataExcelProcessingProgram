#include "log5.h"

log5::log5() {

}
BOOL log5::getHandle(HWND parent, unsigned int IDC_INT) {
	log_handle = GetDlgItem(parent, IDC_INT);

	return TRUE;
}

BOOL log5::writeLogFile(const wchar_t* writeText) {
	
	fopen_s(&fp,"LogFile.txt", "a");
	fwrite(writeText, 1, wcslen(writeText)*2 + 2, fp);
	fprintf(fp, "\n");
	fclose(fp);
	return TRUE;
}

BOOL log5::writeSettingComplete() {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)L"Setting Complete");
	writeLogFile( L"Setting Complete" );
	return TRUE;
}

BOOL log5::writeFileWriteComplete() {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)L"File Write Complete!");
	writeLogFile(L"File Write Complete!");
	return TRUE;
}

BOOL log5::writeFileRequired() {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)L"This File is Required!");
	writeLogFile(L"This File is Required!");
	return TRUE;
}

BOOL log5::writeFileReadComplete() {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)L"File Read Complete!");
	writeLogFile(L"File Read Complete!");
	return TRUE;
}


BOOL log5::writeReleaseDateIsEmpty() {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)L"Release Date is Empty!");
	writeLogFile(L"Release Date is Empty!");
	return TRUE;
}

BOOL log5::writeStartDateIsEmpty() {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)L"Start Year is Empty!");
	writeLogFile(L"Start Year is Empty!");
	return TRUE;
}

BOOL log5::writeNotMatched() {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)L"Episode Data And Avail Data Is Not Matched!");
	writeLogFile(L"Episode Data And Avail Data Is Not Matched!");
	return TRUE;
}

BOOL log5::writeShortSynopsisOver() {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)L"Short Synopsis is Too long, Short Synopsis's Length limit over 400 characters!");
	writeLogFile(L"Short Synopsis is Too long, Short Synopsis's Length limit over 400 characters!");
	return TRUE;
}
BOOL log5::writeLongSynopsisOver() {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)L"Long Synopsis is Too long, Long Synopsis's Length limit over 2000 characters!");
	writeLogFile(L"Long Synopsis is Too long, Long Synopsis's Length limit over 2000 characters!");
	return TRUE;
}

BOOL log5::writeReleaseDateIsOverd() {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)L"Release Date Is Overd");
	writeLogFile(L"Release Date Is Overd");
	return TRUE;
}

BOOL log5::writeLog(TCHAR* inputstring) {
	SendMessage(log_handle, LB_ADDSTRING, 0, (LPARAM)inputstring);
	writeLogFile(inputstring);
	return TRUE;
}