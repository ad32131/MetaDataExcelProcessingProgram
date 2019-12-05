#pragma once
#include <Windows.h>

class availFileListData
{
public:
	HWND parent;
	HWND availFileListDataHandle;
	unsigned int max_avail_count;
	int availfilelist_index;
	int dragFileTimes;
	TCHAR getListItemAvailFileName[MAX_PATH];
	availFileListData();
	BOOL getHandle(HWND parent, unsigned int IDC_INT);
	BOOL addFile(unsigned int NameCount, TCHAR* inputData);
	BOOL getCursel();
	int getCount();
	BOOL getText(int availDataProcessingCount);
};

