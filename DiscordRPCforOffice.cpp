#pragma comment(lib, "discord-rpc.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "version.lib")

#include <iostream>
#include <windows.h>
#include <string>
#include <cmath>
#include "discord_rpc.h"

// =====================================================
// Enum untuk jenis Office Application
// =====================================================
enum OfficeAppType {
    OFFICE_NONE,
    OFFICE_WORD,
    OFFICE_EXCEL,
    OFFICE_POWERPOINT,
    OFFICE_OUTLOOK,
    OFFICE_ACCESS,
    OFFICE_PUBLISHER
};

// =====================================================
// Struktur untuk menyimpan info Office
// =====================================================
struct OfficeInfo {
    OfficeAppType appType;
    std::string appName;
    std::string displayName;
    std::string fileName;
    std::string details;
    std::string state;              // Tambahan untuk state kustom
    std::string largeImageKey;
    int currentPage;
    int totalPages;
    long wordCount;
    std::string version;
    long excelColumn;    // Untuk Excel: nomor kolom
    long excelRow;       // Untuk Excel: nomor baris
};

// =====================================================
// Helper: Ambil Build Information
// =====================================================
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

// =====================================================
// Helper: COM late binding helpers
// =====================================================
std::wstring GetStringProperty(IDispatch* pDisp, const wchar_t* propName) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(propName);
    if (FAILED(pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return L"";

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pDisp->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    std::wstring ret;
    if (result.vt == VT_BSTR && result.bstrVal)
        ret = result.bstrVal;
    VariantClear(&result);
    return ret;
}

long GetLongProperty(IDispatch* pDisp, const wchar_t* propName) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(propName);
    if (FAILED(pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return 0;

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pDisp->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    long ret = 0;
    if (result.vt == VT_I4) ret = result.lVal;
    else if (result.vt == VT_I2) ret = result.iVal;
    VariantClear(&result);
    return ret;
}

IDispatch* GetDispatchProperty(IDispatch* pDisp, const wchar_t* propName) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(propName);
    if (FAILED(pDisp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return nullptr;

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pDisp->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    IDispatch* ret = nullptr;
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        ret = result.pdispVal;
        ret->AddRef();
    }
    VariantClear(&result);
    return ret;
}

long GetSelectionInfo(IDispatch* pSelection, long infoIndex) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(L"Information");
    if (FAILED(pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return 0;

    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_I4;
    arg.lVal = infoIndex;

    DISPPARAMS dp = { &arg, nullptr, 1, 0 };
    VARIANT result;
    VariantInit(&result);
    pSelection->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    long ret = 0;
    if (result.vt == VT_I4) ret = result.lVal;
    else if (result.vt == VT_I2) ret = result.iVal;
    VariantClear(&result);
    return ret;
}

std::string WideToUTF8(const std::wstring& wstr) {
    if (wstr.empty()) return "";
    int size = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string result(size - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), -1, &result[0], size, nullptr, nullptr);
    return result;
}

std::string FormatNumberWithComma(long number) {
    std::string numStr = std::to_string(number);
    std::string result = "";
    int count = 0;

    for (int i = numStr.length() - 1; i >= 0; i--) {
        if (count == 3) {
            result = "," + result;
            count = 0;
        }
        result = numStr[i] + result;
        count++;
    }

    return result;
}

long GetWordCount(IDispatch* pDoc, const long WD_STATISTIC_WORDS = 0) {
    DISPID dispId;
    LPOLESTR statName = const_cast<LPOLESTR>(L"ComputeStatistics");

    if (SUCCEEDED(pDoc->GetIDsOfNames(IID_NULL, &statName, 1, LOCALE_USER_DEFAULT, &dispId))) {
        VARIANT arg;
        VariantInit(&arg);
        arg.vt = VT_I4;
        arg.lVal = WD_STATISTIC_WORDS;

        DISPPARAMS dp = { &arg, nullptr, 1, 0 };
        VARIANT res;
        VariantInit(&res);

        pDoc->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &res, nullptr, nullptr);

        long wordCount = 0;
        if (res.vt == VT_I4) wordCount = res.lVal;
        else if (res.vt == VT_I2) wordCount = res.iVal;

        VariantClear(&res);
        return wordCount;
    }

    return 0;
}

long GetPageCount(IDispatch* pDoc, const long WD_STATISTIC_PAGES = 2) {
    DISPID dispId;
    LPOLESTR statName = const_cast<LPOLESTR>(L"ComputeStatistics");

    if (SUCCEEDED(pDoc->GetIDsOfNames(IID_NULL, &statName, 1, LOCALE_USER_DEFAULT, &dispId))) {
        VARIANT arg;
        VariantInit(&arg);
        arg.vt = VT_I4;
        arg.lVal = WD_STATISTIC_PAGES;

        DISPPARAMS dp = { &arg, nullptr, 1, 0 };
        VARIANT res;
        VariantInit(&res);

        pDoc->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &res, nullptr, nullptr);

        long pageCount = 0;
        if (res.vt == VT_I4) pageCount = res.lVal;
        else if (res.vt == VT_I2) pageCount = res.iVal;

        VariantClear(&res);
        return pageCount;
    }

    return 0;
}

// =====================================================
// Helper: Deteksi Versi Office dari berbagai sumber
// =====================================================

// Struct untuk menyimpan info versi lengkap
struct OfficeVersion {
    int year;           // 2007, 2010, 2013, 2016, 2019, 2021, 2024
    std::string display; // "2007", "2010", dll
    int majorVersion;   // 12, 14, 15, 16, dll (internal Office version)
};

// Cek Registry untuk Office versi ClickToRun (2016+)
OfficeVersion CheckClickToRunVersion() {
    OfficeVersion ver = { 0, "", 0 };

    HKEY hKey;
    const wchar_t* regPath = L"Software\\Microsoft\\Office\\ClickToRun\\Configuration";
    
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, regPath, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        wchar_t version[256] = { 0 };
        DWORD size = sizeof(version);
        
        if (RegQueryValueExW(hKey, L"VersionToReport", nullptr, nullptr, (LPBYTE)version, &size) == ERROR_SUCCESS) {
            std::wstring versionW(version);
            std::string versionA = WideToUTF8(versionW);
            
            // Format: "16.0.xxxxx" atau "17.0.xxxxx" atau "18.0.xxxxx"
            size_t firstDot = versionA.find('.');
            if (firstDot != std::string::npos) {
                try {
                    int major = std::stoi(versionA.substr(0, firstDot));
                    
                    size_t secondDot = versionA.find('.', firstDot + 1);
                    if (secondDot != std::string::npos) {
                        std::string buildStr = versionA.substr(secondDot + 1);
                        int build = std::stoi(buildStr);
                        
                        ver.majorVersion = major;
                        
                        if (major >= 18) {
                            ver.year = 2025;
                            ver.display = "2025";
                        } else if (major == 17) {
                            // Office 17.0 = 2024, tapi cek build untuk 2025 preview
                            if (build >= 18000) {
                                ver.year = 2025;
                                ver.display = "2025";
                            } else {
                                ver.year = 2024;
                                ver.display = "2024";
                            }
                        } else if (major == 16) {
                            if (build >= 17000) {
                                ver.year = 2024;
                                ver.display = "2024";
                            } else if (build >= 14332) {
                                ver.year = 2021;
                                ver.display = "2021";
                            } else if (build >= 10000) {
                                ver.year = 2019;
                                ver.display = "2019";
                            } else {
                                ver.year = 2016;
                                ver.display = "2016";
                            }
                        }
                    }
                }
                catch (...) {}
            }
        }
        
        RegCloseKey(hKey);
    }
    
    return ver;
}

// Cek Registry untuk Office versi MSI (2007, 2010, 2013)
OfficeVersion CheckMSIVersion() {
    OfficeVersion ver = { 0, "", 0 };

    HKEY hKey;
    const wchar_t* regPath = L"Software\\Microsoft\\Office";
    
    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, regPath, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        // Cari semua versi Office yang terinstall
        wchar_t subkeyName[256];
        DWORD index = 0;
        
        while (RegEnumKeyW(hKey, index, subkeyName, sizeof(subkeyName) / sizeof(wchar_t)) == ERROR_SUCCESS) {
            std::wstring subkey = subkeyName;
            
            // Format: "16.0", "15.0", "14.0", "12.0", dll
            if (subkey.find(L".0") != std::wstring::npos) {
                try {
                    int majorVer = std::stoi(subkey);
                    
                    HKEY hSubKey;
                    std::wstring fullPath = std::wstring(regPath) + L"\\" + subkey;
                    
                    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, fullPath.c_str(), 0, KEY_READ, &hSubKey) == ERROR_SUCCESS) {
                        // Cek apakah ada "Common" atau "Word" folder untuk verify installation
                        HKEY hCheckKey;
                        std::wstring checkPath = fullPath + L"\\Common";
                        
                        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, checkPath.c_str(), 0, KEY_READ, &hCheckKey) == ERROR_SUCCESS) {
                            RegCloseKey(hCheckKey);
                            
                            // Versi ditemukan
                            ver.majorVersion = majorVer;
                            
                            if (majorVer >= 16) {
                                ver.year = 2016;
                                ver.display = "2016+";
                            } else if (majorVer == 15) {
                                ver.year = 2013;
                                ver.display = "2013";
                            } else if (majorVer == 14) {
                                ver.year = 2010;
                                ver.display = "2010";
                            } else if (majorVer == 12) {
                                ver.year = 2007;
                                ver.display = "2007";
                            }
                        }
                        
                        RegCloseKey(hSubKey);
                    }
                }
                catch (...) {}
            }
            
            index++;
        }
        
        RegCloseKey(hKey);
    }
    
    return ver;
}

// Deteksi versi dari COM object (fallback)
OfficeVersion DetectVersionFromCOM(IDispatch* pApp) {
    OfficeVersion ver = { 2016, "2016", 16 };
    
    // Ambil property Version dari COM object
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(L"Version");
    if (FAILED(pApp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return ver;

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pApp->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    if (result.vt == VT_BSTR && result.bstrVal) {
        std::wstring versionStr = result.bstrVal;
        
        try {
            size_t firstDot = versionStr.find(L'.');
            if (firstDot != std::wstring::npos) {
                int majorVer = std::stoi(versionStr.substr(0, firstDot));
                
                ver.majorVersion = majorVer;
                
                if (majorVer >= 16) {
                    ver.year = 2016;
                    ver.display = "2016+";
                } else if (majorVer == 15) {
                    ver.year = 2013;
                    ver.display = "2013";
                } else if (majorVer == 14) {
                    ver.year = 2010;
                    ver.display = "2010";
                } else if (majorVer == 12) {
                    ver.year = 2007;
                    ver.display = "2007";
                } else if (majorVer == 11) {
                    ver.year = 2003;
                    ver.display = "2003";
                }
            }
        }
        catch (...) {}
    }
    
    VariantClear(&result);
    return ver;
}

// Master function: Deteksi Office versi dengan urutan prioritas
OfficeVersion DetectOfficeVersion(IDispatch* pApp) {
    // Coba ClickToRun dulu (2016+)
    OfficeVersion ver = CheckClickToRunVersion();
    if (ver.year > 0) return ver;
    
    // Coba MSI version (2007-2013)
    ver = CheckMSIVersion();
    if (ver.year > 0) return ver;
    
    // Fallback ke COM object
    return DetectVersionFromCOM(pApp);
}

// Pilih logo berdasarkan app dan versi (extended untuk 2007+)
std::string SelectLogoKey(OfficeAppType appType, int officeYear) {
    switch (appType) {
        case OFFICE_WORD:
            if (officeYear >= 2025) return "word_2025"; // new era 365
            else if (officeYear >= 2024) return "word_2024"; // end 
			else if (officeYear >= 2021) return "word_2019"; // real 2021 - 2024
            else if (officeYear >= 2019) return "word_2013"; 
            else if (officeYear >= 2016) return "word_2013"; 
            else if (officeYear >= 2013) return "word_2013"; // real 2013 - 2016 - 2019
			else if (officeYear >= 2010) return "word_2010"; // real 2010
			else if (officeYear >= 2007) return "word_2007"; // real 2007
            else return "word_logo";
            
        case OFFICE_EXCEL:
            if (officeYear >= 2025) return "excel_2025";
            else if (officeYear >= 2024) return "excel_2025";
            else if (officeYear >= 2021) return "excel_2021-2024";
            else if (officeYear >= 2019) return "excel_2013-2019";
            else if (officeYear >= 2016) return "excel_2013-2019";
            else if (officeYear >= 2013) return "excel_2013-2019";
            else if (officeYear >= 2010) return "excel_2010";
            else if (officeYear >= 2007) return "excel_2007";
            else return "excel_logo";
            
        case OFFICE_POWERPOINT:
            if (officeYear >= 2025) return "powerpoint_2025";
            else if (officeYear >= 2024) return "powerpoint_2024";
            else if (officeYear >= 2021) return "powerpoint_2021";
            else if (officeYear >= 2019) return "powerpoint_2019";
            else if (officeYear >= 2016) return "powerpoint_2016";
            else if (officeYear >= 2013) return "powerpoint_2013";
            else if (officeYear >= 2010) return "powerpoint_2010";
            else if (officeYear >= 2007) return "powerpoint_2007";
            else return "powerpoint_logo";
            
        case OFFICE_OUTLOOK:
            if (officeYear >= 2025) return "outlook_2025";
            else if (officeYear >= 2024) return "outlook_2024";
            else if (officeYear >= 2021) return "outlook_2021";
            else if (officeYear >= 2019) return "outlook_2019";
            else if (officeYear >= 2016) return "outlook_2016";
            else if (officeYear >= 2013) return "outlook_2013";
            else if (officeYear >= 2010) return "outlook_2010";
            else if (officeYear >= 2007) return "outlook_2007";
            else return "outlook_logo";
            
        default:
            return "office_logo";
    }
}

// =====================================================
// Helper: Excel Column dan Row Detection
// =====================================================

// Helper: Ambil kolom dari Selection Excel
long GetExcelColumn(IDispatch* pSelection) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(L"Column");
    if (FAILED(pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return 0;

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pSelection->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    long ret = 0;
    if (result.vt == VT_I4) ret = result.lVal;
    else if (result.vt == VT_I2) ret = result.iVal;
    VariantClear(&result);
    return ret;
}

// Helper: Ambil baris dari Selection Excel
long GetExcelRow(IDispatch* pSelection) {
    DISPID dispId;
    LPOLESTR name = const_cast<LPOLESTR>(L"Row");
    if (FAILED(pSelection->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispId)))
        return 0;

    VARIANT result;
    VariantInit(&result);
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    pSelection->Invoke(dispId, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);

    long ret = 0;
    if (result.vt == VT_I4) ret = result.lVal;
    else if (result.vt == VT_I2) ret = result.iVal;
    VariantClear(&result);
    return ret;
}

// Helper: Konversi nomor kolom ke huruf (1=A, 2=B, 26=Z, 27=AA, dst)
std::string ColumnNumberToLetter(long colNum) {
    std::string result = "";
    while (colNum > 0) {
        long remainder = (colNum - 1) % 26;
        result = char('A' + remainder) + result;
        colNum = (colNum - 1) / 26;
    }
    return result;
}

// =====================================================
// Office App Detectors - UPDATED
// =====================================================

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
        
        // Deteksi versi dengan master function
        OfficeVersion officeVer = DetectOfficeVersion(pWordApp);
        info.version = officeVer.display;
        info.largeImageKey = SelectLogoKey(OFFICE_WORD, officeVer.year);

        IDispatch* pDoc = GetDispatchProperty(pWordApp, L"ActiveDocument");
        IDispatch* pSel = GetDispatchProperty(pWordApp, L"Selection");

        if (pDoc && pSel) {
            std::wstring wDocName = GetStringProperty(pDoc, L"Name");
            info.fileName = WideToUTF8(wDocName);

            const long WD_ACTIVE_END_PAGE_NUMBER = 3;
            info.currentPage = GetSelectionInfo(pSel, WD_ACTIVE_END_PAGE_NUMBER);
            info.totalPages = GetPageCount(pDoc);
            info.wordCount = GetWordCount(pDoc);

            info.details = "Filename: " + info.fileName;
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

        // Deteksi versi dengan master function
        OfficeVersion officeVer = DetectOfficeVersion(pExcelApp);
        info.version = officeVer.display;
        info.largeImageKey = SelectLogoKey(OFFICE_EXCEL, officeVer.year);

        // Ambil ActiveWorkbook
        IDispatch* pWorkbook = GetDispatchProperty(pExcelApp, L"ActiveWorkbook");
        
        if (pWorkbook) {
            std::wstring wFileName = GetStringProperty(pWorkbook, L"Name");
            info.fileName = WideToUTF8(wFileName);

            // Ambil ActiveSheet
            IDispatch* pSheet = GetDispatchProperty(pWorkbook, L"ActiveSheet");
            if (pSheet) {
                std::wstring wSheetName = GetStringProperty(pSheet, L"Name");
                std::string sheetName = WideToUTF8(wSheetName);
                
                // Gunakan details untuk Workbook name
                info.details = "Workbook: " + info.fileName;
                
                // Ambil Selection untuk mendapatkan kolom dan baris
                IDispatch* pSelection = GetDispatchProperty(pExcelApp, L"Selection");
                if (pSelection) {
                    info.excelColumn = GetExcelColumn(pSelection);
                    info.excelRow = GetExcelRow(pSelection);
                    
                    // Konversi kolom ke huruf
                    std::string colLetter = ColumnNumberToLetter(info.excelColumn);
                    std::string cellReference = colLetter + std::to_string(info.excelRow);
                    
                    // Format: "Sheet: DataSheet | Cell: A1"
                    info.state = "Sheet: " + sheetName + " | Cell: " + cellReference;
                    
                    pSelection->Release();
                } else {
                    // Jika tidak bisa ambil Selection, hanya tampilkan sheet name
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
        
        // Deteksi versi dengan master function
        OfficeVersion officeVer = DetectOfficeVersion(pPPTApp);
        info.version = officeVer.display;
        info.largeImageKey = SelectLogoKey(OFFICE_POWERPOINT, officeVer.year);

        // Ambil ActivePresentation
        IDispatch* pPresentation = GetDispatchProperty(pPPTApp, L"ActivePresentation");
        
        if (pPresentation) {
            std::wstring wFileName = GetStringProperty(pPresentation, L"Name");
            info.fileName = WideToUTF8(wFileName);

            long slideCount = GetLongProperty(pPresentation, L"Slides");
            info.totalPages = slideCount;

            // Ambil SlideShowWindow untuk slide yang aktif
            IDispatch* pSlideShow = GetDispatchProperty(pPPTApp, L"SlideShowWindows");
            if (pSlideShow) {
                // Coba ambil slide pertama/aktif
                info.currentPage = 1;
                pSlideShow->Release();
            }

            info.details = "Presentation: " + info.fileName + " | Slides: " + std::to_string(slideCount);

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
        
        // Deteksi versi dengan master function
        OfficeVersion officeVer = DetectOfficeVersion(pOutlookApp);
        info.version = officeVer.display;
        info.largeImageKey = SelectLogoKey(OFFICE_OUTLOOK, officeVer.year);

        // Ambil ActiveExplorer (Explorer Window yang aktif)
        IDispatch* pExplorer = GetDispatchProperty(pOutlookApp, L"ActiveExplorer");
        
        if (pExplorer) {
            // Ambil Folder yang ditampilkan
            IDispatch* pFolder = GetDispatchProperty(pExplorer, L"CurrentFolder");
            if (pFolder) {
                std::wstring wFolderName = GetStringProperty(pFolder, L"Name");
                info.details = "Folder: " + WideToUTF8(wFolderName);
                pFolder->Release();
            }

            pExplorer->Release();
        } else {
            info.details = "Email Client";
        }

        info.displayName += " " + info.version;

        if (pOutlookApp) pOutlookApp->Release();
        if (pUnk) pUnk->Release();
    }

    return info;
}

// =====================================================
// Discord RPC Handler
// =====================================================
void handleDiscordReady(const DiscordUser* user) {
    std::cout << "Discord siap! Terhubung sebagai: " << user->username << std::endl;
}

// =====================================================
// MAIN
// =====================================================
int main() {
    // Init Discord RPC
    DiscordEventHandlers handlers;
    memset(&handlers, 0, sizeof(handlers));
    handlers.ready = handleDiscordReady;
    Discord_Initialize("YOUR_APPLICATION_ID", &handlers, 1, NULL); // check in discord developers portal

    // Init COM
    if (FAILED(CoInitialize(NULL))) {
        std::cout << "Gagal inisialisasi COM!" << std::endl;
        return 1;
    }

    std::cout << "========================================" << std::endl;
    std::cout << "   Office RPC for Discord sedang jalan... " << std::endl;
    std::cout << "   " << GetBuildVersion() << std::endl;
    std::cout << "========================================" << std::endl;

    static int64_t startTime = (int64_t)time(0);

    while (true) {
        OfficeInfo activeOffice;
        activeOffice.appType = OFFICE_NONE;

        // Deteksi Office apps dalam urutan prioritas
        OfficeInfo wordInfo = DetectAndGetWordInfo();
        if (wordInfo.appType == OFFICE_WORD) {
            activeOffice = wordInfo;
        }
        else {
            OfficeInfo excelInfo = DetectAndGetExcelInfo();
            if (excelInfo.appType == OFFICE_EXCEL) {
                activeOffice = excelInfo;
            }
            else {
                OfficeInfo pptInfo = DetectAndGetPowerPointInfo();
                if (pptInfo.appType == OFFICE_POWERPOINT) {
                    activeOffice = pptInfo;
                }
                else {
                    OfficeInfo outlookInfo = DetectAndGetOutlookInfo();
                    if (outlookInfo.appType == OFFICE_OUTLOOK) {
                        activeOffice = outlookInfo;
                    }
                }
            }
        }

        if (activeOffice.appType != OFFICE_NONE) {
            DiscordRichPresence presence;
            memset(&presence, 0, sizeof(presence));

            presence.details = activeOffice.details.c_str();
            
            std::string stateLabel = "";
            
            // Untuk Excel, gunakan state custom
            if (activeOffice.appType == OFFICE_EXCEL && !activeOffice.state.empty()) {
                stateLabel = activeOffice.state;
            }
            // Untuk aplikasi lain, gunakan format default
            else if (activeOffice.totalPages > 0) {
                stateLabel = "Page " + std::to_string(activeOffice.currentPage) + 
                           " of " + std::to_string(activeOffice.totalPages);
                if (activeOffice.wordCount > 0) {
                    stateLabel += " | " + FormatNumberWithComma(activeOffice.wordCount) + " words";
                }
            }
            
            if (!stateLabel.empty()) {
                presence.state = stateLabel.c_str();
            }

            presence.largeImageKey = activeOffice.largeImageKey.c_str();
            presence.largeImageText = activeOffice.displayName.c_str();
            presence.startTimestamp = startTime;

            Discord_UpdatePresence(&presence);

            std::cout << "\r[Status] " << activeOffice.appName << " | " 
                      << activeOffice.fileName << "     " << std::flush;
        }
        else {
            Discord_ClearPresence();
            std::cout << "\r[Status] No Office application detected...        " << std::flush;
        }

        Discord_RunCallbacks();
        Sleep(2000);
    }

    Discord_Shutdown();
    CoUninitialize();
    return 0;
}
