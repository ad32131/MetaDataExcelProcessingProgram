#include "availFileListData.h"

availFileListData::availFileListData() {
	max_avail_count = 0;
	availfilelist_index = 0;
	dragFileTimes = 0;
}

BOOL availFileListData::getHandle(HWND parent, unsigned int IDC_INT) {
	this->parent = parent;
	availFileListDataHandle = GetDlgItem(parent, IDC_INT);

	return TRUE;
}

BOOL availFileListData::addFile(unsigned int NameCount, TCHAR* inputData) {
	SendMessage(availFileListDataHandle, LB_ADDSTRING, NameCount, (LPARAM)inputData);
	return TRUE;
}

BOOL availFileListData::getCursel() {
	availfilelist_index = SendMessage(availFileListDataHandle, LB_GETCURSEL, 0, 0);
	SendMessage(availFileListDataHandle, LB_DELETESTRING, availfilelist_index, 0);
	return TRUE;
}

//BOOL availFileListData::

int availFileListData::getCount() {
	return SendMessage(availFileListDataHandle, LB_GETCOUNT, 0, 0);
}

BOOL availFileListData::getText(int availDataProcessingCount) {
	memset(getListItemAvailFileName, 0, MAX_PATH);
	SendMessage(availFileListDataHandle, LB_GETTEXT, availDataProcessingCount, (WPARAM)getListItemAvailFileName);
	return TRUE;
}