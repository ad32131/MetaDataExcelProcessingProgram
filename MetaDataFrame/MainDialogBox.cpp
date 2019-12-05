#include "framework.h"
#include "MetaDataFrame.h"
#include "MainDialogBox.h"
#include "excel.h"
#include "log5.h"
#include "progressbar5.h"
#include "availFileListData.h"
#include "episodeListData.h"
#include "MetaData.h"
#include "resource.h"
#include <CommCtrl.h> 
#include <shellapi.h>
#include <shlobj_core.h>
#include <Commdlg.h>
#include <TlHelp32.h>
#include <Shlwapi.h>


#pragma comment(lib, "Shlwapi.lib")

BOOL processAllKill(HWND hDlg,const WCHAR* szProcessName)
{
	MessageBox(hDlg, L"plz exit all excel", L"alert!", MB_OK);
	HANDLE hndl = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	DWORD dwsma = GetLastError();
	HANDLE hHandle;

	DWORD dwExitCode = 0;

	PROCESSENTRY32  procEntry = { 0 };
	procEntry.dwSize = sizeof(PROCESSENTRY32);
	Process32First(hndl, &procEntry);
	while (1)
	{
		if (!wcscmp(procEntry.szExeFile, szProcessName))
		{

			hHandle = ::OpenProcess(PROCESS_ALL_ACCESS, 0, procEntry.th32ProcessID);

			if (::GetExitCodeProcess(hHandle, &dwExitCode))
			{
				if (!::TerminateProcess(hHandle, dwExitCode))
				{
					return FALSE;
				}
			}
		}
		if (!Process32Next(hndl, &procEntry))
		{
			return TRUE;
		}
	}


	return TRUE;
}

int CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM /*lParam*/, LPARAM lpData) {
	if (uMsg == BFFM_INITIALIZED)
	{
		//BROWSEINFO.lParam에서 설정 해준 값이 lpData로넘어온다.
		// LPARAM으로 path를 넘겨주려면 WParam을 TRUE로,
		// PIDL을 넘겨주려면 FALSE로 넘겨준다.
		if (lpData)
			SendMessage(hwnd, BFFM_SETSELECTION, (WPARAM)TRUE, (LPARAM)lpData);
	}
	return 0;
}



BOOL CALLBACK DialogProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {

	//errorlog
	static log5 log_handle;

	//progressBarData
	static progressbar5 progressbar_handle;

	//availFileListdata
	static availFileListData availFileList;

	//episodeFileListData
	static episodeListData episodeList;
	

	static BOOL metaDataTemplate = FALSE;
	static BOOL metaDataStartTemplate = FALSE;

	

	//episodeFileListdata
	static int listMaxFileTimes = 0;
	static HWND IDC_episodefilelist_H;
	static int episodefilelist_index = 0;
	static unsigned int IDC_episodefilelist_H_count = 0;
	//static int dragFileTimes = 0;
	
	//metaDataOutFileListdata
	static HWND IDC_metaDataOutfilelist_H;
	static int metaDataOutfilelist_index = 0;
	//static int dragFileTimes = 0;
	
	static HWND IDC_outputFilePath_H;

	static HWND IDC_maxEpisodeDate_H;

	static TCHAR episodeMaxText[MAX_PATH];
	static int episodeNumber = 0;
	static int seasonNumber = 0;
	unsigned int dateTmpDate = 0;

	//index
	TCHAR episode_id[MAX_PATH];
	static unsigned int max_seriesNumber = 0;
	static unsigned int start_point = 0;
	static unsigned int end_point = 5;

	static BOOL file_read_flag = FALSE;

	typedef struct episodeInformationData {
		TCHAR episode_id[MAX_PATH];
		unsigned int max_seriesNumber;// = 0;
		unsigned int start_point;// = 0;
		unsigned int end_point;// = 5;
	};

	static episodeInformationData *episodeData;

	static HDROP hDrop = 0;
	MetaData* k;
	int endOfFile = 0;
	static int tempInt = 0;
	BOOL proccessingOut = TRUE;

	//excel class
	static excel availsFile;
	static excel episodeFile;
	static excel metaDataFile;

	//OLECHAR data_input[] = OLESTR("dtdtdt");

	const OLECHAR metaDataForm[] = OLESTR("Mthetr-Metadata-Form.xlsx");

	//browse
	static BROWSEINFO browse_OutPutFilePath;
	static OLECHAR outPutFilePath[MAX_PATH];

	//tmpData
	unsigned int episode_check_value_left = 0;
	unsigned int episode_check_value_right = 0;
	TCHAR getDataClipBoard[MAX_PATH];
	TCHAR getData[MAX_PATH];
	TCHAR getData_tmp[MAX_PATH];
	TCHAR getData_cmpstr[MAX_PATH];
	TCHAR getData_inttmp[MAX_PATH];
	TCHAR getTitle_namecheck[MAX_PATH];
	TCHAR metaDataName[MAX_PATH];
	TCHAR metaTitle[MAX_PATH];
	TCHAR episodeDataFileName[MAX_PATH];
	TCHAR metaDataFileName[MAX_PATH];
	TCHAR filenamefromdrag[MAX_PATH];
	TCHAR shortTmpData[4000];
	TCHAR longTmpData[4000];
	TCHAR lastProcessingName[MAX_PATH];
	LPITEMIDLIST pidl;

	//clipboard data
	static HANDLE hMem;
	static TCHAR* p_data;

	switch (message) {
	case WM_INITDIALOG:
		processAllKill(hDlg, L"EXCEL.EXE");
		
		log_handle.getHandle(hDlg, IDC_textLog);

		progressbar_handle.getHandle(hDlg, IDC_PROGRESS1);
		availFileList.getHandle(hDlg, IDC_AvailsDataLIST);

		episodeList.getHandle(hDlg, IDC_inputEpisodeDataList);

		IDC_episodefilelist_H = GetDlgItem(hDlg, IDC_needEpisodeFileList);
		IDC_metaDataOutfilelist_H = GetDlgItem(hDlg, IDC_metaDataOutFileList);
		IDC_maxEpisodeDate_H = GetDlgItem(hDlg, IDC_maxEpisodeDate);
		DragAcceptFiles(hDlg, TRUE);

		memset(outPutFilePath, 0, MAX_PATH);
		SHGetSpecialFolderPath(NULL, outPutFilePath, CSIDL_DESKTOPDIRECTORY, 0);

		IDC_outputFilePath_H = GetDlgItem(hDlg, IDC_outputFilePath);
		SetWindowText(IDC_outputFilePath_H, outPutFilePath);
		break;
	case WM_DESTROY:
		free(episodeData);
		processAllKill(hDlg, L"EXCEL.EXE");
		availsFile.excelclosefile();
		availsFile.excelquit();
		break;
	case WM_DROPFILES:
		availFileList.dragFileTimes = 0;
		memset(filenamefromdrag, 0, MAX_PATH);
		hDrop = (HDROP)wParam;
		availFileList.dragFileTimes = DragQueryFile(hDrop, -1, filenamefromdrag, MAX_PATH);
		for (int fileNameCount = 0; fileNameCount < availFileList.dragFileTimes; fileNameCount++) {
			DragQueryFile(hDrop, fileNameCount, filenamefromdrag, MAX_PATH);
			if( wcsstr(filenamefromdrag, L".xls") && wcsstr(filenamefromdrag, L"Mthetr"))
				availFileList.addFile(fileNameCount, filenamefromdrag);
			else if (wcsstr(filenamefromdrag, L".xls") )
				episodeList.addFile(fileNameCount, filenamefromdrag);
			else
				MessageBox(hDlg, filenamefromdrag, L"This File is Not Excel File", 0);
		}

		//MessageBox(hDlg, filenamefromdrag, L"DRAG_FILE", 0);

		//IDC_AvailsDataLIST
		break;
	case WM_COMMAND:
		switch (LOWORD(wParam)) {
		case IDC_outPutFilePathSet:
			memset(&browse_OutPutFilePath, 0, sizeof(BROWSEINFO));
			memset(outPutFilePath, 0, MAX_PATH);
			browse_OutPutFilePath.hwndOwner = NULL;
			browse_OutPutFilePath.pidlRoot = NULL;
			browse_OutPutFilePath.pszDisplayName = outPutFilePath;
			browse_OutPutFilePath.lpszTitle = L"PATH SET";
			browse_OutPutFilePath.ulFlags = BIF_RETURNONLYFSDIRS | BIF_STATUSTEXT | BIF_VALIDATE;
			browse_OutPutFilePath.lpfn = BrowseCallbackProc;
			browse_OutPutFilePath.lParam = (LPARAM)outPutFilePath;
			browse_OutPutFilePath.ulFlags = BIF_RETURNONLYFSDIRS;

			pidl = SHBrowseForFolder(&browse_OutPutFilePath);
			if (pidl != NULL) {
				SHGetPathFromIDList(pidl, outPutFilePath);
				SetWindowText(IDC_outputFilePath_H, outPutFilePath);
			}


			break;
		case IDC_AvailsDataLIST:
			switch (HIWORD(wParam)) {
			case LBN_SELCHANGE:


				break;
			case LBN_DBLCLK:
				availFileList.getCursel();
				break;
			}
			break;
		case IDC_inputEpisodeDataList:
			switch (HIWORD(wParam)) {
			case LBN_SELCHANGE:


				break;
			case LBN_DBLCLK:
				episodeList.getCursel();
				break;
			}
			break;
		case IDC_needEpisodeFileList:
			switch (HIWORD(wParam)) {
			case LBN_SELCHANGE:


				break;
			case LBN_DBLCLK:
				int copy_index = SendMessage(IDC_episodefilelist_H, LB_GETCURSEL, 0, 0);
				SendMessage(IDC_episodefilelist_H, LB_DELETESTRING, copy_index, 0);

				break;
			}
			break;
		case IDC_SET:

			if (PathFileExists(metaDataForm) != TRUE) {
				MessageBox(hDlg, L"Not Exist MetaData_Form File!", L"Fatal Error!", MB_OK);
				break;
			}
			progressbar_handle.setProgresszero();

			listMaxFileTimes = SendMessage(IDC_episodefilelist_H, LB_GETCOUNT, 0, 0);
			for(int availDataProcessingCount = 0; availDataProcessingCount < listMaxFileTimes; availDataProcessingCount++)
				SendMessage(IDC_episodefilelist_H, LB_DELETESTRING, 0, 0);

			listMaxFileTimes = SendMessage(IDC_metaDataOutfilelist_H, LB_GETCOUNT, 0, 0);
			for (int availDataProcessingCount = 0; availDataProcessingCount < listMaxFileTimes; availDataProcessingCount++)
				SendMessage(IDC_metaDataOutfilelist_H, LB_DELETESTRING, 0, 0);

			listMaxFileTimes = SendMessage(IDC_maxEpisodeDate_H, LB_GETCOUNT, 0, 0);
			for (int availDataProcessingCount = 0; availDataProcessingCount < listMaxFileTimes; availDataProcessingCount++)
				SendMessage(IDC_maxEpisodeDate_H, LB_DELETESTRING, 0, 0);

			listMaxFileTimes = availFileList.getCount();

			//free(episodeData);
			//episodeData = (episodeInformationData*)calloc(listMaxFileTimes, sizeof(episodeInformationData));

			for (int availDataProcessingCount = 0; availDataProcessingCount < listMaxFileTimes; availDataProcessingCount++) {

				availFileList.getText(availDataProcessingCount);




				//memset(episodeData, 0, MAX_PATH);

				memset(getData, 0, MAX_PATH);
				memset(getData_tmp, 0, MAX_PATH);
				memset(metaDataName, 0, MAX_PATH);
				memset(metaTitle, 0, MAX_PATH);

				//memset(episode_id, 0, MAX_PATH);

				endOfFile = 0;
				episodeNumber = 0;
				seasonNumber = 0;
				max_seriesNumber = 0;
				end_point = 5;
				start_point = 0;

				availsFile.excelstart(hDlg);
				availsFile.excelcreatenewwork();
				availsFile.excelreadfile(availFileList.getListItemAvailFileName);

				for (int condition_index = end_point; !endOfFile; condition_index++, progressbar_handle.setProgressPlus()) {

					


					

					if (condition_index == end_point)//first read
					{
						max_seriesNumber = availsFile.excelDataGetRead(0, 0, L"I", condition_index);
						availsFile.excelDataGetRead(metaTitle, MAX_PATH, L"A", condition_index);
						availsFile.excelDataGetRead(metaDataName, MAX_PATH, L"F", condition_index);
						availsFile.excelDataGetRead(episode_id, MAX_PATH, L"R", condition_index);
						start_point = condition_index;
						continue;
					}

					availsFile.excelDataGetRead(getData_tmp, MAX_PATH, L"R", condition_index);
					if (wcslen(getData_tmp)) {
						if (max_seriesNumber < (unsigned int)availsFile.excelDataGetRead(0, 0, L"I", condition_index)) {
							max_seriesNumber = (unsigned int)availsFile.excelDataGetRead(0, 0, L"I", condition_index);
						}
						else if (wcscmp(episode_id, getData_tmp)) {
							end_point = condition_index;
							//endOfFile = 1;
							//wsprintf(episodeDataFileName, L"%s\\%s_%s.xlsx", outPutFilePath, metaDataName, episode_id);
							wsprintf(episodeDataFileName, L"%s_%s.xlsx", metaDataName, episode_id);
							if ((SendMessage(IDC_episodefilelist_H, LB_FINDSTRING, 0, (LPARAM)episodeDataFileName)) < 0) {
								SendMessage(IDC_episodefilelist_H, LB_ADDSTRING, 0, (LPARAM)episodeDataFileName);
								wsprintf(metaDataFileName, L"%s-MetaData-%s.xlsx", metaTitle, episode_id);
								SendMessage(IDC_metaDataOutfilelist_H, LB_ADDSTRING, 0, (LPARAM)metaDataFileName);
								wsprintf(episodeMaxText, L"%d", max_seriesNumber);
								SendMessage(IDC_maxEpisodeDate_H, LB_ADDSTRING, 0, (LPARAM)episodeMaxText);
								
							}

							//wsprintf(metaDataFileName, L"%s\\%s-MetaData-%s.xlsx", outPutFilePath, metaTitle, episode_id);
							/*
							wsprintf(metaDataFileName, L"%s-MetaData-%s.xlsx", metaTitle, episode_id);
							if ((SendMessage(IDC_metaDataOutfilelist_H, LB_FINDSTRING, 0, (LPARAM)metaDataFileName)) < 0){
								SendMessage(IDC_metaDataOutfilelist_H, LB_ADDSTRING, 0, (LPARAM)metaDataFileName);
							}
							*/
						}


						if (wcslen(episode_id) != 0) availsFile.excelDataGetRead(episode_id, MAX_PATH, L"R", condition_index);

						availsFile.excelDataGetRead(getData_tmp, MAX_PATH, L"F", condition_index);
						if (wcscmp(metaDataName, getData_tmp)) availsFile.excelDataGetRead(metaDataName, MAX_PATH, L"F", condition_index);

					}
					else {
						end_point = condition_index;
						endOfFile = 1;

						wsprintf(episodeDataFileName, L"%s_%s.xlsx", metaDataName, episode_id);
						if ((SendMessage(IDC_episodefilelist_H, LB_FINDSTRING, 0, (LPARAM)episodeDataFileName)) < 0) {
							SendMessage(IDC_episodefilelist_H, LB_ADDSTRING, 0, (LPARAM)episodeDataFileName);
							wsprintf(metaDataFileName, L"%s-MetaData-%s.xlsx", metaTitle, episode_id);
							SendMessage(IDC_metaDataOutfilelist_H, LB_ADDSTRING, 0, (LPARAM)metaDataFileName);
							wsprintf(episodeMaxText, L"%d", max_seriesNumber);
							SendMessage(IDC_maxEpisodeDate_H, LB_ADDSTRING, 0, (LPARAM)episodeMaxText);
						}


					}



				}
				
				log_handle.writeSettingComplete();

				

			}
			progressbar_handle.setProgressFull();

			availsFile.excelquit();
			availsFile.excelclosefile();
			break;
			////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		case IDOK:
			
			progressbar_handle.setProgressinit();

			/*Avail Data init*/

			listMaxFileTimes = SendMessage(IDC_episodefilelist_H, LB_GETCOUNT, 0, 0);

			//processing
			for (int processcount = 0; processcount < listMaxFileTimes; processcount++) {
				memset(getData, 0, MAX_PATH);
				memset(getData_tmp, 0, MAX_PATH);
				memset(metaDataName, 0, MAX_PATH);
				memset(metaTitle, 0, MAX_PATH);

				//memset(episode_id, 0, MAX_PATH);
				
				endOfFile = 0;
				episodeNumber = 0;
				seasonNumber = 0;
				max_seriesNumber = 0;

				memset(getData, 0, MAX_PATH);

				
				//getData_tmp
				

				episodeList.getCount();
				for (int episode_index = 0; episode_index < episodeList.IDC_inputEpisodeData_H_count; episode_index++) {
					episodeList.getText(episode_index, getData_tmp);

					IDC_episodefilelist_H_count = (SendMessage(IDC_episodefilelist_H, LB_GETCOUNT, 0, 0));
					for (int episode_count_index = 0; episode_count_index < IDC_episodefilelist_H_count; episode_count_index++) {

						

						SendMessage(IDC_episodefilelist_H, LB_GETTEXT, episode_count_index, (LPARAM)getData);
						
						memset(getData_cmpstr, 0, MAX_PATH);
						wcsncpy_s(getData_cmpstr, getData, wcschr(getData, '_') - getData);

						if ( wcsstr(getData_tmp, getData_cmpstr) ) {
							SendMessage(IDC_episodefilelist_H, LB_DELETESTRING, episode_count_index, 0);

							memset(getData, 0, MAX_PATH);
							SendMessage(IDC_metaDataOutfilelist_H, LB_GETTEXT, episode_count_index, (LPARAM)getData);
							SendMessage(IDC_metaDataOutfilelist_H, LB_DELETESTRING, episode_count_index, 0);


							//episodeDataFileName
							memset(episodeDataFileName, 0, MAX_PATH);
							wsprintf(episodeDataFileName, L"%s", getData_tmp);

							memset(metaDataFileName, 0, MAX_PATH);
							wsprintf(metaDataFileName, L"%s\\%s", outPutFilePath, getData);

							memset(episodeMaxText, 0, MAX_PATH);
							SendMessage(IDC_maxEpisodeDate_H, LB_GETTEXT, episode_count_index, (LPARAM)episodeMaxText);
							SendMessage(IDC_maxEpisodeDate_H, LB_DELETESTRING, episode_count_index, 0);

							max_seriesNumber = _ttoi(episodeMaxText);


							/*MetaData Write Start*/
							episodeFile.excelstart(hDlg);
							episodeFile.excelcreatenewwork();
							episodeFile.excelreadfile(episodeDataFileName);

							DeleteFile(metaDataFileName);
							CopyFile(metaDataForm, metaDataFileName, 0);


							//find
							availsFile.excelstart(hDlg);
							availFileList.max_avail_count = availFileList.getCount();
							
							for (unsigned int avail_index = 0; avail_index < availFileList.max_avail_count; avail_index++) {
								//availsFile read
								metaDataTemplate = FALSE;
								availFileList.getText(avail_index);

								availsFile.excelcreatenewwork();
								availsFile.excelreadfile(availFileList.getListItemAvailFileName);
								for (int dataCount = 1; dataCount < 90000; dataCount++) {
								/* episode end writing*/
								availsFile.excelDataGetRead(getTitle_namecheck, 200, L"F", dataCount);
								
								if (metaDataTemplate == TRUE && wcscmp(lastProcessingName, episodeDataFileName)) {

										max_seriesNumber = dataCount - start_point;


										k = (MetaData*)calloc((max_seriesNumber + 2), sizeof(MetaData));

										for (int dataCountTX = 0; dataCountTX < (max_seriesNumber + 2); dataCountTX++, progressbar_handle.setProgressPlus()) {
											k[dataCountTX].MetaDataInit();

											if (dataCountTX < 2) {
												if (dataCountTX == 0) {
													wcscpy_s(k[dataCountTX].Content_Type, 20, L"TV Series");
													wsprintf(k[dataCountTX].Start_Year, L"%d", (unsigned int)episodeFile.excelDataGetRead(0, 0, L"P", 4));
													if (wcslen(k[dataCountTX].Start_Year) < 1) {
														log_handle.writeLog(episodeDataFileName);
														log_handle.writeStartDateIsEmpty();
													}
												}
												if (dataCountTX == 1)
												{
													wcscpy_s(k[dataCountTX].Content_Type, 20, L"TV Season");

													//episodedata processing

													availsFile.excelDataGetRead(k[dataCountTX].Series_ID_Token1, 100, L"Q", start_point);
													availsFile.excelDataGetRead(k[dataCountTX].Season_Sequence_Number, 100, L"R", start_point);


												}

												availsFile.excelDataGetRead(k[dataCountTX].Title, 200, L"F", start_point);
												episodeFile.excelDataGetRead(k[dataCountTX].Title_pronunciation, 200, L"B", 4);
												episodeFile.excelDataGetRead(shortTmpData, 4000, L"AB", 4);
												episodeFile.excelDataGetRead(longTmpData, 4000, L"AD", 4);
												wsprintf(k[dataCountTX].Short_Synopsis, L"%s%s", shortTmpData, longTmpData);

												if (wcslen(k[dataCountTX].Short_Synopsis) > 400) {
													log_handle.writeLog(episodeDataFileName);
													log_handle.writeShortSynopsisOver();
												}

												wsprintf(k[dataCountTX].Long_Synopsis, L"%s%s", shortTmpData, longTmpData);

												if (wcslen(k[dataCountTX].Long_Synopsis) > 2000) {
													log_handle.writeLog(episodeDataFileName);
													log_handle.writeLongSynopsisOver();
												}

												episodeFile.excelDataGetRead(k[dataCountTX].Actors, 200, L"AF", 4);
												episodeFile.excelDataGetRead(k[dataCountTX].Directors, 200, L"AI", 4);
												episodeFile.excelDataGetRead(k[dataCountTX].Producers, 200, L"AL", 4);
												episodeFile.excelDataGetRead(k[dataCountTX].Writers, 200, L"AR", 4);
											}
											else {

												wcscpy_s(k[dataCountTX].Content_Type, 20, L"TV Episode");


												episodeFile.excelDataGetRead(k[dataCountTX].Title, 200, L"E", (dataCountTX + 2));
												episodeFile.excelDataGetRead(k[dataCountTX].Title_pronunciation, 200, L"F", (dataCountTX + 2));

												episodeFile.excelDataGetRead(shortTmpData, 4000, L"AC", (dataCountTX + 2));
												episodeFile.excelDataGetRead(longTmpData, 4000, L"AD", (dataCountTX + 2));

												wsprintf(k[dataCountTX].Short_Synopsis, L"%s%s", shortTmpData, longTmpData);

												if (wcslen(k[dataCountTX].Short_Synopsis) > 400) {
													log_handle.writeLog(episodeDataFileName);
													log_handle.writeShortSynopsisOver();
												}

												wsprintf(k[dataCountTX].Long_Synopsis, L"%s%s", shortTmpData, longTmpData);

												if (wcslen(k[dataCountTX].Long_Synopsis) > 2000) {
													log_handle.writeLog(episodeDataFileName);
													log_handle.writeLongSynopsisOver();
												}

												episodeFile.excelDataGetRead(k[dataCountTX].Actors, 200, L"AF", (dataCountTX + 2));
												episodeFile.excelDataGetRead(k[dataCountTX].Directors, 200, L"AI", (dataCountTX + 2));
												episodeFile.excelDataGetRead(k[dataCountTX].Producers, 200, L"AL", (dataCountTX + 2));
												episodeFile.excelDataGetRead(k[dataCountTX].Writers, 200, L"AR", (dataCountTX + 2));

												dateTmpDate = (unsigned int)episodeFile.excelDataGetRead(getData_inttmp, MAX_PATH, L"J", (dataCountTX + 2));
												if ( (dateTmpDate == 0) && (wcslen(getData_inttmp) > 8)) {
													log_handle.writeLog(episodeDataFileName);
													log_handle.writeReleaseDateIsOverd();
												}

												wsprintf(k[dataCountTX].Release_Date, L"%d/%d/%d",
													(dateTmpDate - (dateTmpDate % 10000)) / 10000,
													((dateTmpDate - (dateTmpDate - (dateTmpDate % 10000))) - ((dateTmpDate - (dateTmpDate - (dateTmpDate % 10000))) % 100)) / 100,
													(dateTmpDate - (dateTmpDate - (dateTmpDate % 10000))) % 100
												);
												if (wcslen(k[dataCountTX].Release_Date) < 1) {
													log_handle.writeLog(episodeDataFileName);
													log_handle.writeReleaseDateIsEmpty();
												}

											}

											if (dataCountTX < 3) {


												availsFile.excelDataGetRead(k[dataCountTX].Partner_Name, 20, L"A", start_point);

												availsFile.excelDataGetRead(k[dataCountTX].Unique_ID, 40, L"Q", start_point);

												if (dataCountTX == 2) {

													availsFile.excelDataGetRead(k[dataCountTX].Series_ID_Token2, 100, L"Q", start_point + dataCountTX - 2);
													availsFile.excelDataGetRead(k[dataCountTX].Season_ID, 100, L"R", start_point + dataCountTX - 2);

													wsprintf(k[dataCountTX].Episode_Sequence_Number, L"%d", (unsigned int)availsFile.excelDataGetRead(0, 0, L"I", start_point + dataCountTX - 2));
												}
											}
											else
											{
												availsFile.excelDataGetRead(k[dataCountTX].Partner_Name, 20, L"A", start_point + dataCountTX - 2);
												availsFile.excelDataGetRead(k[dataCountTX].Unique_ID, 40, L"Q", start_point + dataCountTX - 2);

												availsFile.excelDataGetRead(k[dataCountTX].Series_ID_Token2, 100, L"Q", start_point + dataCountTX - 2);
												availsFile.excelDataGetRead(k[dataCountTX].Season_ID, 100, L"R", start_point + dataCountTX - 2);

												wsprintf(k[dataCountTX].Episode_Sequence_Number, L"%d", (unsigned int)availsFile.excelDataGetRead(0, 0, L"I", start_point + dataCountTX - 2));

											}


											episodeFile.excelDataGetRead(k[dataCountTX].Copyright_Holder, 200, L"AD", 4);
											wcscpy_s(k[dataCountTX].Rating, 10, L"NR");

											wcscpy_s(k[dataCountTX].Language, 20, L"ja");
											wcscpy_s(k[dataCountTX].OriginalLanguage, 20, L"ja");
											wcscpy_s(k[dataCountTX].CountryOfOrigin, 20, L"JP");
										}
										if ( max_seriesNumber < (unsigned int)episodeFile.excelDataGetRead(0, 0, L"H", (max_seriesNumber + 4)))  {
											log_handle.writeLog(episodeDataFileName);
											log_handle.writeNotMatched();
										}

										log_handle.writeLog(episodeDataFileName);
										log_handle.writeFileReadComplete();



										episodeFile.excelclosefile();

										metaDataFile.excelstart(hDlg);
										metaDataFile.excelcreatenewwork();
										metaDataFile.excelreadfile(metaDataFileName);

										for (int dataCountTX = 0; dataCountTX < (max_seriesNumber + 2); dataCountTX++, progressbar_handle.setProgressPlus()) {

											metaDataFile.excelDataGreenSet(k[dataCountTX].Partner_Name, L"A", dataCountTX + 3);
											metaDataFile.excelDataGreenSet(k[dataCountTX].Unique_ID, L"B", dataCountTX + 3);
											metaDataFile.excelDataGreenSet(k[dataCountTX].Content_Type, L"C", dataCountTX + 3);
											metaDataFile.excelDataGreenSet(k[dataCountTX].Title, L"D", dataCountTX + 3);
											metaDataFile.excelDataGreenSet(k[dataCountTX].Title_pronunciation, L"E", dataCountTX + 3);

											//metaDataFile.excelDataGreenSet(k[dataCountTX].Studio, L"F", dataCountTX + 3);
											metaDataFile.excelDataGreenSet(k[dataCountTX].Rating, L"G", dataCountTX + 3);

											if (dataCountTX == 0) {
												metaDataFile.excelDataYellowSet(k[dataCountTX].Start_Year, L"H", dataCountTX + 3);
												metaDataFile.excelDataGreenSet(k[dataCountTX].End_Year, L"I", dataCountTX + 3);
											}
											else if (dataCountTX == 1) {
												metaDataFile.excelDataGreenSet(k[dataCountTX].Series_ID_Token1, L"J", dataCountTX + 3);
												metaDataFile.excelDataGreenSet(k[dataCountTX].Season_Sequence_Number, L"K", dataCountTX + 3);
											}
											else {
												metaDataFile.excelDataGreenSet(k[dataCountTX].Series_ID_Token2, L"L", dataCountTX + 3);
												metaDataFile.excelDataGreenSet(k[dataCountTX].Season_ID, L"M", dataCountTX + 3);
												metaDataFile.excelDataGreenSet(k[dataCountTX].Episode_Sequence_Number, L"N", dataCountTX + 3);
											}


											metaDataFile.excelDataYellowSet(k[dataCountTX].Short_Synopsis, L"P", dataCountTX + 3);
											metaDataFile.excelDataYellowSet(k[dataCountTX].Long_Synopsis, L"Q", dataCountTX + 3);


											metaDataFile.excelDataGreenSet(k[dataCountTX].Actors, L"R", dataCountTX + 3);
											metaDataFile.excelDataGreenSet(k[dataCountTX].Directors, L"S", dataCountTX + 3);
											metaDataFile.excelDataGreenSet(k[dataCountTX].Producers, L"T", dataCountTX + 3);
											metaDataFile.excelDataGreenSet(k[dataCountTX].Writers, L"U", dataCountTX + 3);


											//metaDataFile.excelDataGreenSet(k[dataCountTX].Release_Date, L"W", dataCountTX + 3);
											metaDataFile.excelDataYellowSet(k[dataCountTX].Release_Date, L"W", dataCountTX + 3);


											metaDataFile.excelDataGreenSet(k[dataCountTX].Copyright_Holder, L"Z", dataCountTX + 3);


											metaDataFile.excelDataGreenSet(k[dataCountTX].Language, L"AB", dataCountTX + 3);
											metaDataFile.excelDataGreenSet(k[dataCountTX].OriginalLanguage, L"AC", dataCountTX + 3);
											metaDataFile.excelDataGreenSet(k[dataCountTX].CountryOfOrigin, L"AD", dataCountTX + 3);


											if (dataCountTX < 5) {
												metaDataFile.excelDataGreenSet(0, L"F", dataCountTX + 3); //Studio
												metaDataFile.excelDataGreenSet(0, L"V", dataCountTX + 3); //createtor
												metaDataFile.excelDataGreenSet(0, L"AA", dataCountTX + 3);//Copyright Year
												metaDataFile.excelDataGreenSet(0, L"AE", dataCountTX + 3);//OriginalLanguageTitle
												metaDataFile.excelDataGreenSet(0, L"Y", dataCountTX + 3);//metaDataFile.excelDataGreenSet(0, L"AE", dataCountTX + 3);//OriginalLanguageTitle
												if (dataCountTX > 1) {

													metaDataFile.excelDataGreenSet(0, L"AF", dataCountTX + 3);//RuntimeInMinutes
													metaDataFile.excelDataGreenSet(0, L"AG", dataCountTX + 3);//Genre
												}
												else {
													metaDataFile.excelDataGreenSet(0, L"W", dataCountTX + 3);//Copyright Year
												}
											}
										}


										metaDataFile.excelsave(2);
										metaDataFile.excelclosefile();
										//metaDataFile.excelquit();

										free(k);
										log_handle.writeLog(metaDataFileName);
										log_handle.writeFileWriteComplete();

										metaDataTemplate = FALSE;
										metaDataStartTemplate = FALSE;

										memset(lastProcessingName, 0, MAX_PATH);
										wcscpy_s(lastProcessingName, episodeDataFileName);
										start_point = 0;
										max_seriesNumber = 0;
									}
								else if (wcsstr(getTitle_namecheck, getData_cmpstr) ) {
									episode_check_value_left = (unsigned int)availsFile.excelDataGetRead(0, 0, L"I", dataCount);
									episode_check_value_right = (unsigned int)availsFile.excelDataGetRead(0, 0, L"I", (dataCount + 1));
									/*episode*/
									if (metaDataStartTemplate == FALSE && wcscmp(lastProcessingName, episodeDataFileName)) {
										start_point = dataCount;
										metaDataStartTemplate = TRUE;
									}
									else if (episode_check_value_left > episode_check_value_right && wcscmp(lastProcessingName, episodeDataFileName)) {
										max_seriesNumber = dataCount - start_point;
										metaDataTemplate = TRUE;
										
									}
									

								}
								else if ( wcslen(getTitle_namecheck) == 0) {

									break;
								}
							}
							}

							//start_point
							//start
							//max_seriesNumber
							//process

							//episodeFile.excelquit();




						}




						/*MetaDataWrite End*/
					}

				}

				
				
			


			}

			listMaxFileTimes = SendMessage(IDC_episodefilelist_H, LB_GETCOUNT, 0, 0);
			if (listMaxFileTimes != 0) {
				for (int list_count = 0; list_count < listMaxFileTimes; list_count++) {
					memset(getData, 0, MAX_PATH);
					SendMessage(IDC_episodefilelist_H, LB_GETTEXT, 0, (LPARAM)getData);
					log_handle.writeLog(getData);
					log_handle.writeFileRequired();
					SendMessage(IDC_episodefilelist_H, LB_DELETESTRING, 0, 0);
				}
			}

			progressbar_handle.setProgressFull();
			MessageBox(hDlg, L"Done!", L"Processing End", 0);

			return FALSE;
			break;
		case IDCANCEL:
			EndDialog(hDlg, TRUE);
			return FALSE;
			break;
		}
		break;
	}
	return FALSE;
}



MainDialogBox::MainDialogBox(HINSTANCE hInstance, LPWSTR IDDa, HWND hwnd_desktop) {
	this->hInstance = hInstance;
	this->IDDa = IDDa;
	this->hwnd_desktop = hwnd_desktop;
}

BOOL MainDialogBox::StartDialogBox() {
	DialogBoxW(hInstance, (LPWSTR)IDDa, hwnd_desktop, DialogProc);
	return TRUE;
}
