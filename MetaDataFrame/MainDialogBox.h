#include "framework.h"
#include "MetaDataFrame.h"

#pragma once
class MainDialogBox {

private:
	HINSTANCE hInstance;
	LPWSTR IDDa;
	HWND hwnd_desktop;
public:
	BOOL StartDialogBox();
	MainDialogBox(HINSTANCE hInstance, LPWSTR IDDa, HWND hwnd_desktop);
};