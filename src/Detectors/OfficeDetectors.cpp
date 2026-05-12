#include "../../include/Helpers/ComHelpers.h"
#include "../../include/Detectors/OfficeVersionDetector.h"
#include <windows.h>
#include <psapi.h>
#include <algorithm>
#include <cwctype>

#pragma comment(lib, "psapi.lib")

// Forward declarations
long GetExcelColumn(IDispatch* pSelection);
long GetExcelRow(IDispatch* pSelection);
std::string ColumnNumberToLetter(long colNum);
std::string GetBuildVersion();

OfficeAppType GetForegroundOfficeAppType() {
    HWND hwnd = GetForegroundWindow();
    if (!hwnd) return OFFICE_NONE;

    wchar_t className[256] = { 0 };
    if (!GetClassNameW(hwnd, className, static_cast<int>(sizeof(className) / sizeof(className[0])))) {
        return OFFICE_NONE;
    }

    std::wstring classStr(className);
    std::transform(classStr.begin(), classStr.end(), classStr.begin(),
        [](wchar_t ch) { return static_cast<wchar_t>(towlower(ch)); });

    if (classStr == L"xlmain" || classStr.find(L"excel") != std::wstring::npos) {
        return OFFICE_EXCEL;
    }

    if (classStr == L"opusapp" || classStr.find(L"word") != std::wstring::npos) {
        return OFFICE_WORD;
    }

    if (classStr == L"pptframeclass" || classStr == L"ppframeclass" || classStr.find(L"powerpnt") != std::wstring::npos) {
        return OFFICE_POWERPOINT;
    }

    if (classStr == L"rctrl_renwnd32" || classStr.find(L"outlook") != std::wstring::npos) {
        return OFFICE_OUTLOOK;
    }

    return OFFICE_NONE;
}

bool IsOfficeWindowActive() {
    return GetForegroundOfficeAppType() != OFFICE_NONE;
}

OfficeInfo DetectAndGetWordInfo() {
    OfficeInfo info;
    info.appType = OFFICE_NONE;

    CLSID clsid;
    IUnknown* pUnk = nullptr;
    IDispatch* pWordApp = nullptr;

    if (SUCCEEDED(CLSIDFromProgID(L"Word.Application", &clsid)) &&
        SUCCEEDED(GetActiveObject(clsid, NULL, &pUnk)) &&
        SUCCEEDED(pUnk->QueryInterface(IID_IDispatch, (void**)&pWordApp))) {

        info.appType = OFFICE_WORD;
        info.appName = "Word";
        info.displayName = "Microsoft Word";
        
        OfficeVersion officeVer = DetectOfficeVersion(pWordApp);
        info.version = officeVer.display;
        info.largeImageKey = SelectLogoKey(OFFICE_WORD, officeVer.year);
        info.smallImageKey = SelectSmallImageKey(OFFICE_WORD);

        IDispatch* pDoc = GetDispatchProperty(pWordApp, L"ActiveDocument");
        IDispatch* pSel = GetDispatchProperty(pWordApp, L"Selection");

        if (pDoc && pSel) {
            std::wstring wDocName = GetStringProperty(pDoc, L"Name");
            info.fileName = WideToUTF8(wDocName);

            const long WD_ACTIVE_END_PAGE_NUMBER = 3;
            info.currentPage = GetSelectionInfo(pSel, WD_ACTIVE_END_PAGE_NUMBER);
            info.totalPages = GetPageCount(pDoc);
            info.wordCount = GetWordCount(pDoc);

            info.details = "Document: " + info.fileName;
            info.displayName += " " + info.version;
        }

        if (pDoc) pDoc->Release();
        if (pSel) pSel->Release();
        if (pWordApp) pWordApp->Release();
        if (pUnk) pUnk->Release();
    }

    return info;
}

OfficeInfo DetectAndGetExcelInfo() {
    OfficeInfo info;
    info.appType = OFFICE_NONE;
    info.excelColumn = 0;
    info.excelRow = 0;

    CLSID clsid;
    IUnknown* pUnk = nullptr;
    IDispatch* pExcelApp = nullptr;

    if (SUCCEEDED(CLSIDFromProgID(L"Excel.Application", &clsid)) &&
        SUCCEEDED(GetActiveObject(clsid, NULL, &pUnk)) &&
        SUCCEEDED(pUnk->QueryInterface(IID_IDispatch, (void**)&pExcelApp))) {

        info.appType = OFFICE_EXCEL;
        info.appName = "Excel";
        info.displayName = "Microsoft Excel";

        OfficeVersion officeVer = DetectOfficeVersion(pExcelApp);
        info.version = officeVer.display;
        info.largeImageKey = SelectLogoKey(OFFICE_EXCEL, officeVer.year);
        info.smallImageKey = SelectSmallImageKey(OFFICE_EXCEL);

        IDispatch* pWorkbook = GetDispatchProperty(pExcelApp, L"ActiveWorkbook");
        
        if (pWorkbook) {
            std::wstring wFileName = GetStringProperty(pWorkbook, L"Name");
            info.fileName = WideToUTF8(wFileName);

            IDispatch* pSheet = GetDispatchProperty(pWorkbook, L"ActiveSheet");
            if (pSheet) {
                std::wstring wSheetName = GetStringProperty(pSheet, L"Name");
                std::string sheetName = WideToUTF8(wSheetName);
                
                info.details = "Workbook: " + info.fileName;
                
                IDispatch* pSelection = GetDispatchProperty(pExcelApp, L"Selection");
                if (pSelection) {
                    info.excelColumn = GetExcelColumn(pSelection);
                    info.excelRow = GetExcelRow(pSelection);
                    
                    std::string colLetter = ColumnNumberToLetter(info.excelColumn);
                    std::string cellReference = colLetter + std::to_string(info.excelRow);
                    
                    info.state = "Sheet: " + sheetName + " | Cell: " + cellReference;
                    
                    pSelection->Release();
                } else {
                    info.state = "Sheet: " + sheetName;
                }
                
                pSheet->Release();
            } else {
                info.details = "Workbook: " + info.fileName;
                info.state = "";
            }

            pWorkbook->Release();
        }

        info.displayName += " " + info.version;

        if (pExcelApp) pExcelApp->Release();
        if (pUnk) pUnk->Release();
    }

    return info;
}

OfficeInfo DetectAndGetPowerPointInfo() {
    OfficeInfo info;
    info.appType = OFFICE_NONE;

    CLSID clsid;
    IUnknown* pUnk = nullptr;
    IDispatch* pPPTApp = nullptr;

    if (SUCCEEDED(CLSIDFromProgID(L"PowerPoint.Application", &clsid)) &&
        SUCCEEDED(GetActiveObject(clsid, NULL, &pUnk)) &&
        SUCCEEDED(pUnk->QueryInterface(IID_IDispatch, (void**)&pPPTApp))) {

        info.appType = OFFICE_POWERPOINT;
        info.appName = "PowerPoint";
        info.displayName = "Microsoft PowerPoint";

        OfficeVersion officeVer = DetectOfficeVersion(pPPTApp);
        info.version = officeVer.display;
        info.largeImageKey = SelectLogoKey(OFFICE_POWERPOINT, officeVer.year);
        info.smallImageKey = SelectSmallImageKey(OFFICE_POWERPOINT);

        IDispatch* pPresentation = GetDispatchProperty(pPPTApp, L"ActivePresentation");
        
        if (pPresentation) {
            std::wstring wFileName = GetStringProperty(pPresentation, L"Name");
            info.fileName = WideToUTF8(wFileName);
            info.details = "Presentation: " + info.fileName;

            IDispatch* pSlides = GetDispatchProperty(pPresentation, L"Slides");
            if (pSlides) {
                long slideCount = GetLongProperty(pSlides, L"Count");
                info.totalPages = slideCount;
                
                IDispatch* pWindow = GetDispatchProperty(pPPTApp, L"ActiveWindow");
                if (pWindow) {
                    long slideIndex = GetLongProperty(pWindow, L"Selection");
                    info.currentPage = slideIndex > 0 ? slideIndex : 1;
                    pWindow->Release();
                }

                pSlides->Release();
            }

            pPresentation->Release();
        }

        info.displayName += " " + info.version;

        if (pPPTApp) pPPTApp->Release();
        if (pUnk) pUnk->Release();
    }

    return info;
}

OfficeInfo DetectAndGetOutlookInfo() {
    OfficeInfo info;
    info.appType = OFFICE_NONE;

    CLSID clsid;
    IUnknown* pUnk = nullptr;
    IDispatch* pOutlookApp = nullptr;

    if (SUCCEEDED(CLSIDFromProgID(L"Outlook.Application", &clsid)) &&
        SUCCEEDED(GetActiveObject(clsid, NULL, &pUnk)) &&
        SUCCEEDED(pUnk->QueryInterface(IID_IDispatch, (void**)&pOutlookApp))) {

        info.appType = OFFICE_OUTLOOK;
        info.appName = "Outlook";
        info.displayName = "Microsoft Outlook";

        OfficeVersion officeVer = DetectOfficeVersion(pOutlookApp);
        info.version = officeVer.display;
        info.largeImageKey = SelectLogoKey(OFFICE_OUTLOOK, officeVer.year);
        info.smallImageKey = SelectSmallImageKey(OFFICE_OUTLOOK);

        IDispatch* pExplorer = GetDispatchProperty(pOutlookApp, L"ActiveExplorer");
        
        if (pExplorer) {
            IDispatch* pFolder = GetDispatchProperty(pExplorer, L"CurrentFolder");
            if (pFolder) {
                std::wstring wFolderName = GetStringProperty(pFolder, L"Name");
                info.details = "Folder: " + WideToUTF8(wFolderName);
                pFolder->Release();
            }

            pExplorer->Release();
        }

        info.displayName += " " + info.version;

        if (pOutlookApp) pOutlookApp->Release();
        if (pUnk) pUnk->Release();
    }

    return info;
}

std::string GetBuildVersion() {
    wchar_t filename[MAX_PATH];
    GetModuleFileNameW(NULL, filename, MAX_PATH);

    DWORD verHandle = 0;
    DWORD verSize = GetFileVersionInfoSizeW(filename, &verHandle);

    if (verSize == 0) {
        return "Build Unknown";
    }

    LPSTR verData = new char[verSize];
    if (!GetFileVersionInfoW(filename, verHandle, verSize, verData)) {
        delete[] verData;
        return "Build Unknown";
    }

    UINT size = 0;
    LPBYTE lpBuffer = NULL;
    if (VerQueryValueW(verData, L"\\", (VOID FAR * FAR*) & lpBuffer, &size)) {
        if (size) {
            VS_FIXEDFILEINFO* verInfo = (VS_FIXEDFILEINFO*)lpBuffer;
            if (verInfo->dwSignature == 0xfeef04bd) {
                int major = (verInfo->dwFileVersionMS >> 16) & 0xffff;
                int minor = (verInfo->dwFileVersionMS >> 0) & 0xffff;
                int revision = (verInfo->dwFileVersionLS >> 16) & 0xffff;
                int build = (verInfo->dwFileVersionLS >> 0) & 0xffff;

                std::string result = "Build " + std::to_string(major) + "." +
                    std::to_string(minor) + "." +
                    std::to_string(revision) + "." +
                    std::to_string(build);
                delete[] verData;
                return result;
            }
        }
    }

    delete[] verData;
    return "Build Unknown";
}

long GetExcelColumn(IDispatch* pSelection) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(L"Column");
    if (FAILED(pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return 0;

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pSelection->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    long column = 0;
    if (result.vt == VT_I4) {
        column = result.lVal;
    }

    VariantClear(&result);
    return column;
}

long GetExcelRow(IDispatch* pSelection) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(L"Row");
    if (FAILED(pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return 0;

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pSelection->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    long row = 0;
    if (result.vt == VT_I4) {
        row = result.lVal;
    }

    VariantClear(&result);
    return row;
}

std::string ColumnNumberToLetter(long colNum) {
    std::string result = "";
    while (colNum > 0) {
        colNum--;
        result = char('A' + (colNum % 26)) + result;
        colNum /= 26;
    }
    return result;
}