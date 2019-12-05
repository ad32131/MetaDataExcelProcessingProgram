#pragma once
#include <Windows.h>
class episodeListData
{
public:
	//inputEpisodeData
	HWND episodeListDataHandle;
	int inputEpisodeData_index;
	int availfilelist_index;
	unsigned int IDC_inputEpisodeData_H_count;
	int listMaxFileTimes_Episode;
	episodeListData();
	BOOL getHandle(HWND parent, unsigned int IDC_INT);
	BOOL addFile(unsigned int NameCount, TCHAR* inputData);
	BOOL getCursel();
	int getCount();
	BOOL getText(int availDataProcessingCount,TCHAR *inputData);
};

