#pragma once

#include <windows.h>
#include <string>

// Helper: COM late binding helpers
std::wstring GetStringProperty(IDispatch* pDisp, const wchar_t* propName);
long GetLongProperty(IDispatch* pDisp, const wchar_t* propName);
IDispatch* GetDispatchProperty(IDispatch* pDisp, const wchar_t* propName);
long GetSelectionInfo(IDispatch* pSelection, long infoIndex);
std::string WideToUTF8(const std::wstring& wstr);
std::string FormatNumberWithComma(long number);
long GetWordCount(IDispatch* pDoc, const long WD_STATISTIC_WORDS = 0);
long GetPageCount(IDispatch* pDoc, const long WD_STATISTIC_PAGES = 2);