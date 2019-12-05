#include "MetaData.h"

MetaData::MetaData() {
	//MetaDataInit();
	/* init public
	const int szInt = _szInt;
	const int szName = _szName;
	const int szLongName = _szLongName;
	const int szShortTitle = _szShortTitle;
	const int szLongTitle = _szLongTitle;
	const int szShortBigData = _szShortBigData;
	const int szLongBigData = _szLongBigData;

	*/
}

BOOL MetaData::MetaDataInit() {
	memset(Partner_Name, 0, _szName);
	memset(Unique_ID, 0, _szLongName);
	memset(Content_Type, 0, _szName);
	memset(Title, 0, _szLongTitle);
	memset(Title_pronunciation, 0, _szLongTitle);
	memset(Studio, 0, _szLongTitle);
	memset(Rating, 0, _szInt);
	memset(Start_Year, 0, _szInt);
	memset(End_Year, 0, _szInt);
	memset(Series_ID_Token1, 0, _szShortTitle);
	memset(Season_Sequence_Number, 0, _szShortTitle);
	memset(Series_ID_Token2, 0, _szShortTitle);
	memset(Season_ID, 0, _szShortTitle);
	memset(Episode_Sequence_Number, 0, _szInt);
	memset(Short_Synopsis, 0, _szShortBigData);
	memset(Long_Synopsis, 0, _szLongBigData);
	memset(Actors, 0, _szLongTitle);
	memset(Directors, 0, _szLongTitle);
	memset(Producers, 0, _szLongTitle);
	memset(Writers, 0, _szLongTitle);
	memset(Release_Date, 0, _szLongTitle);
	memset(Copyright_Holder, 0, _szLongTitle);
	memset(Language, 0, _szName);
	memset(OriginalLanguage, 0, _szName);
	memset(CountryOfOrigin, 0, _szName);
	Episode_Max_Number = 0;

	return TRUE;
}