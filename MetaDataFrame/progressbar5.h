#pragma once
#include <Windows.h>
#include <CommCtrl.h> 

class progressbar5
{
private:
	HWND PROGRESSBAR_HANDLE;
	
public:
	unsigned char progressint = 0;
	BOOL getHandle(HWND parent, unsigned int IDC_INT);
	progressbar5();
	BOOL setProgressinit();
	BOOL setProgress();
	BOOL setProgresszero();
	BOOL setProgressPlus();
	BOOL setProgressFull();
	BOOL setProgressInt(unsigned int inputInt);
};

