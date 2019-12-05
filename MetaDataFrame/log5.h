#pragma once
#include <Windows.h>
#include <stdlib.h>
#include <stdio.h>

class log5 {
private:
	HWND log_handle;
	
public:
	FILE *fp;
	log5();
	BOOL getHandle(HWND parent, unsigned int IDC_INT);
	BOOL writeSettingComplete();
	BOOL writeLog(TCHAR* inputstring);
	BOOL writeFileWriteComplete();
	BOOL writeFileRequired();
	BOOL writeFileReadComplete();
	BOOL writeReleaseDateIsEmpty();
	BOOL writeStartDateIsEmpty();
	BOOL writeShortSynopsisOver();
	BOOL writeLongSynopsisOver();
	BOOL writeNotMatched();
	BOOL writeLogFile(const wchar_t* writeText);
	BOOL writeReleaseDateIsOverd();
};