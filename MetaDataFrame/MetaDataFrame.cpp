// MetaDataFrame.cpp : アプリケーションのエントリ ポイントを定義します。
//

#include "framework.h"
#include "MetaDataFrame.h"
#include "MainDialogBox.h"


int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrecInstance, LPSTR lpszCmdParam, int nCmdShow) {
	//DialogBoxW(hInstance, (LPWSTR)IDD_DIALOG1, HWND_DESKTOP, DialogProc);

	MainDialogBox MainDialogBox( hInstance, (LPWSTR)IDD_DIALOG1, HWND_DESKTOP);
	MainDialogBox.StartDialogBox();

}
