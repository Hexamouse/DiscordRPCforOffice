#pragma once

#include "../Core/OfficeInfo.h"
#include <windows.h>

long GetExcelColumn(IDispatch* pSelection);
long GetExcelRow(IDispatch* pSelection);
std::string ColumnNumberToLetter(long colNum);
std::string GetBuildVersion();
OfficeAppType GetForegroundOfficeAppType();
OfficeInfo DetectAndGetWordInfo();
OfficeInfo DetectAndGetExcelInfo();
OfficeInfo DetectAndGetPowerPointInfo();
OfficeInfo DetectAndGetOutlookInfo();