#pragma once

#include <string>

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
    std::string state;
    std::string largeImageKey;
    std::string smallImageKey;
    int currentPage;
    int totalPages;
    long wordCount;
    std::string version;
    long excelColumn;
    long excelRow;
};

// =====================================================
// Struct untuk menyimpan info versi lengkap
// =====================================================
struct OfficeVersion {
    int year;
    std::string display;
    int majorVersion;
};