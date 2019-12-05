#include "episodeListData.h"

episodeListData::episodeListData() {
	inputEpisodeData_index = 0;
	IDC_inputEpisodeData_H_count = 0;
	listMaxFileTimes_Episode = 0;
}

BOOL episodeListData::getHandle(HWND parent, unsigned int IDC_INT) {
	episodeListDataHandle = GetDlgItem(parent, IDC_INT);
	return TRUE;
}

BOOL episodeListData::addFile(unsigned int NameCount, TCHAR* inputData) {
	SendMessage(episodeListDataHandle, LB_ADDSTRING, NameCount, (LPARAM)inputData);
	return TRUE;
}

BOOL episodeListData::getCursel() {
	availfilelist_index = SendMessage(episodeListDataHandle, LB_GETCURSEL, 0, 0);
	SendMessage(episodeListDataHandle, LB_DELETESTRING, availfilelist_index, 0);
	return TRUE;
}

int episodeListData::getCount() {
	IDC_inputEpisodeData_H_count =  SendMessage(episodeListDataHandle, LB_GETCOUNT, 0, 0);
	return IDC_inputEpisodeData_H_count;
}

BOOL episodeListData::getText(int availDataProcessingCount,TCHAR *inputData) {
	memset(inputData, 0, MAX_PATH);
	SendMessage(episodeListDataHandle, LB_GETTEXT, availDataProcessingCount, (WPARAM)inputData);
	return TRUE;
}

/* episodeListData::
BOOL getHandle(HWND parent, unsigned int IDC_INT);
	BOOL addFile(unsigned int NameCount, TCHAR* inputData);
	BOOL getCursel();
	int getCount();
	BOOL getText(int availDataProcessingCount);
	*/