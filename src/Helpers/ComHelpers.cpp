#include "../../include/Helpers/ComHelpers.h"

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

    for (int i = static_cast<int>(numStr.length()) - 1; i >= 0; i--) {
        if (count == 3) {
            result = "," + result;
            count = 0;
        }
        result = numStr[i] + result;
        count++;
    }

    return result;
}

long GetWordCount(IDispatch* pDoc, const long WD_STATISTIC_WORDS) {
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

long GetPageCount(IDispatch* pDoc, const long WD_STATISTIC_PAGES) {
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