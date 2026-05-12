#include "../../include/Detectors/OfficeVersionDetector.h"
#include "../../include/Helpers/ComHelpers.h"
#include <windows.h>

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
                        }
                        else if (major == 17) {
                            if (build >= 18000) {
                                ver.year = 2025;
                                ver.display = "2025";
                            }
                            else {
                                ver.year = 2024;
                                ver.display = "2024";
                            }
                        }
                        else if (major == 16) {
                            if (build >= 17000) {
                                ver.year = 2024;
                                ver.display = "2024";
                            }
                            else if (build >= 14332) {
                                ver.year = 2021;
                                ver.display = "2021";
                            }
                            else if (build >= 10000) {
                                ver.year = 2019;
                                ver.display = "2019";
                            }
                            else {
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

OfficeVersion CheckMSIVersion() {
    OfficeVersion ver = { 0, "", 0 };

    HKEY hKey;
    const wchar_t* regPath = L"Software\\Microsoft\\Office";

    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, regPath, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        wchar_t subkeyName[256];
        DWORD index = 0;

        while (RegEnumKeyW(hKey, index, subkeyName, sizeof(subkeyName) / sizeof(wchar_t)) == ERROR_SUCCESS) {
            std::wstring subkey = subkeyName;

            if (subkey.find(L".0") != std::wstring::npos) {
                try {
                    int majorVer = std::stoi(subkey);

                    HKEY hSubKey;
                    std::wstring fullPath = std::wstring(regPath) + L"\\" + subkey;

                    if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, fullPath.c_str(), 0, KEY_READ, &hSubKey) == ERROR_SUCCESS) {
                        HKEY hCheckKey;
                        std::wstring checkPath = fullPath + L"\\Common";

                        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, checkPath.c_str(), 0, KEY_READ, &hCheckKey) == ERROR_SUCCESS) {
                            RegCloseKey(hCheckKey);

                            ver.majorVersion = majorVer;

                            if (majorVer >= 16) {
                                ver.year = 2016;
                                ver.display = "2016+";
                            }
                            else if (majorVer == 15) {
                                ver.year = 2013;
                                ver.display = "2013";
                            }
                            else if (majorVer == 14) {
                                ver.year = 2010;
                                ver.display = "2010";
                            }
                            else if (majorVer == 12) {
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

OfficeVersion DetectVersionFromCOM(IDispatch* pApp) {
    OfficeVersion ver = { 2016, "2016", 16 };

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
                }
                else if (majorVer == 15) {
                    ver.year = 2013;
                    ver.display = "2013";
                }
                else if (majorVer == 14) {
                    ver.year = 2010;
                    ver.display = "2010";
                }
                else if (majorVer == 12) {
                    ver.year = 2007;
                    ver.display = "2007";
                }
                else if (majorVer == 11) {
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

OfficeVersion DetectOfficeVersion(IDispatch* pApp) {
    OfficeVersion ver = CheckClickToRunVersion();
    if (ver.year > 0) return ver;

    ver = CheckMSIVersion();
    if (ver.year > 0) return ver;

    return DetectVersionFromCOM(pApp);
}

std::string SelectLogoKey(OfficeAppType appType, int officeYear) {
    switch (appType) {
    case OFFICE_WORD:
        if (officeYear >= 2025) return "word_2025";
        else if (officeYear >= 2024) return "word_2024";
        else if (officeYear >= 2021) return "word_2021"; // jika tidak ada asset ini, ganti ke word_2019 atau word_logo
        else if (officeYear >= 2019) return "word_2019";
        else if (officeYear >= 2013) return "word_2013";
        else if (officeYear >= 2010) return "word_2010";
        else if (officeYear >= 2007) return "word_2007";
        else return "word_logo";

    case OFFICE_EXCEL:
        if (officeYear >= 2025) return "excel_2025";
        else if (officeYear >= 2021) return "excel_2025";
        else if (officeYear >= 2013) return "excel_2013-2019";
        else if (officeYear >= 2010) return "excel_2010";
        else if (officeYear >= 2007) return "excel_2007";
        else return "office365_logo";

    case OFFICE_POWERPOINT:
        if (officeYear >= 2025) return "powerpoint_2025";
        else if (officeYear >= 2024) return "powerpoint_2024";
        else if (officeYear >= 2021) return "powerpoint_2021";
        else if (officeYear >= 2019) return "powerpoint_2019";
        else if (officeYear >= 2016) return "powerpoint_2016";
        else if (officeYear >= 2013) return "powerpoint_2013";
        else if (officeYear >= 2010) return "powerpoint_2010";
        else if (officeYear >= 2007) return "powerpoint_2007";
        else return "office365_logo";

    case OFFICE_OUTLOOK:
        if (officeYear >= 2025) return "outlook_2025";
        else if (officeYear >= 2024) return "outlook_2024";
        else if (officeYear >= 2021) return "outlook_2021";
        else if (officeYear >= 2019) return "outlook_2019";
        else if (officeYear >= 2016) return "outlook_2016";
        else if (officeYear >= 2013) return "outlook_2013";
        else if (officeYear >= 2010) return "outlook_2010";
        else if (officeYear >= 2007) return "outlook_2007";
        else return "office365_logo";

    default:
        return "office365_logo";
    }
}

std::string SelectSmallImageKey(OfficeAppType appType) {
    switch (appType) {
    case OFFICE_WORD:
        return "office365_logo";
    case OFFICE_EXCEL:
        return "office365_logo";
    case OFFICE_POWERPOINT:
        return "office365_logo";
    case OFFICE_OUTLOOK:
        return "office365_logo";
    default:
        return "office365_logo";
    }
}