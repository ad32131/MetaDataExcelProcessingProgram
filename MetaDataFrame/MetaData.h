#pragma once
#include "framework.h"
#include <stdio.h>

#define _szInt 10
#define _szName 20
#define _szLongName 40
#define _szShortTitle 100
#define _szLongTitle 200
#define _szBigTitle 400
#define _szShortBigData 1000
#define _szLongBigData 4000




class MetaData {
private:
public:
	

	MetaData();
	BOOL MetaDataInit();
	TCHAR Partner_Name[ _szName ];
	TCHAR Unique_ID[_szLongName];
	TCHAR Content_Type[ _szName ];
	TCHAR Title[ _szLongTitle ];
	TCHAR Title_pronunciation[ _szLongTitle ];
	TCHAR Studio[ _szLongTitle ];
	TCHAR Rating[_szInt];
	TCHAR Start_Year[_szInt];
	TCHAR End_Year[_szInt];
	TCHAR Series_ID_Token1[_szShortTitle];
	TCHAR Season_Sequence_Number[_szShortTitle];
	TCHAR Series_ID_Token2[_szShortTitle];
	TCHAR Season_ID[_szShortTitle];
	TCHAR Episode_Sequence_Number[_szInt];
	TCHAR Short_Synopsis[_szShortBigData];
	TCHAR Long_Synopsis[_szLongBigData];
	TCHAR Actors[_szLongTitle];
	TCHAR Directors[_szLongTitle];
	TCHAR Producers[_szLongTitle];
	TCHAR Writers[_szLongTitle];
	TCHAR Release_Date[_szLongTitle];
	TCHAR Copyright_Holder[_szLongTitle];
	TCHAR Language[_szName];
	TCHAR OriginalLanguage[_szName];
	TCHAR CountryOfOrigin[_szName];
	int Episode_Max_Number;
};