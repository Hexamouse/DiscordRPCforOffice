#pragma comment(lib, "discord-rpc.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "version.lib")

#include <iostream>
#include <windows.h>
#include "discord_rpc.h"
#include "../../include/Detectors/OfficeDetectors.h"
#include "../../include/Detectors/OfficeVersionDetector.h"
#include "../../include/Helpers/ComHelpers.h"

bool discordConnected = false;

void handleDiscordReady(const DiscordUser* user) {
    if (user && user->username) {
        std::cout << "[INFO] Discord connected as: " << user->username << std::endl;
        discordConnected = true;
    } else {
        std::cout << "[INFO] Discord connected (user information unavailable)" << std::endl;
        discordConnected = true;
    }
}

int main() {
    DiscordEventHandlers handlers;
    memset(&handlers, 0, sizeof(handlers));
    handlers.ready = handleDiscordReady;
    
    std::cout << "========================================" << std::endl;
    std::cout << "   Office Rich Presence for Discord    " << std::endl;
    std::cout << "   " << GetBuildVersion() << std::endl;
    std::cout << "========================================" << std::endl;
    std::cout << "[INFO] Initializing Discord connection..." << std::endl;
    
    Discord_Initialize("1503853338393837721", &handlers, 1, NULL);

    if (FAILED(CoInitialize(NULL))) {
        std::cout << "[ERROR] Failed to initialize COM interface." << std::endl;
        return 1;
    }

    static int64_t startTime = (int64_t)time(0);
    std::string lastState = "";
    std::string lastDetails = "";
    std::string lastImageKey = "";
    std::string lastAppName = "";
    std::string lastVersion = "";
    int connectionAttempts = 0;

    while (true) {
        Discord_RunCallbacks();
        
        if (!discordConnected && connectionAttempts < 5) {
            connectionAttempts++;
            std::cout << "[INFO] Attempting to establish Discord connection... (" << connectionAttempts << "/5)" << std::endl;
            Sleep(1000);
            continue;
        }

        auto detectOfficeByType = [](OfficeAppType appType) -> OfficeInfo {
            switch (appType) {
            case OFFICE_WORD:
                return DetectAndGetWordInfo();
            case OFFICE_EXCEL:
                return DetectAndGetExcelInfo();
            case OFFICE_POWERPOINT:
                return DetectAndGetPowerPointInfo();
            case OFFICE_OUTLOOK:
                return DetectAndGetOutlookInfo();
            default: {
                OfficeInfo empty{};
                empty.appType = OFFICE_NONE;
                return empty;
            }
            }
        };

        OfficeInfo activeOffice{};
        activeOffice.appType = OFFICE_NONE;

        OfficeAppType foregroundApp = GetForegroundOfficeAppType();
        if (foregroundApp != OFFICE_NONE) {
            activeOffice = detectOfficeByType(foregroundApp);
        }

        if (activeOffice.appType == OFFICE_NONE) {
            OfficeInfo wordInfo = DetectAndGetWordInfo();
            if (wordInfo.appType == OFFICE_WORD) {
                activeOffice = wordInfo;
            } else {
                OfficeInfo excelInfo = DetectAndGetExcelInfo();
                if (excelInfo.appType == OFFICE_EXCEL) {
                    activeOffice = excelInfo;
                } else {
                    OfficeInfo pptInfo = DetectAndGetPowerPointInfo();
                    if (pptInfo.appType == OFFICE_POWERPOINT) {
                        activeOffice = pptInfo;
                    } else {
                        OfficeInfo outlookInfo = DetectAndGetOutlookInfo();
                        if (outlookInfo.appType == OFFICE_OUTLOOK) {
                            activeOffice = outlookInfo;
                        }
                    }
                }
            }
        }

        if (activeOffice.appType != OFFICE_NONE) {
            std::string versionText = activeOffice.version.empty() ? "unknown" : activeOffice.version;

            if (activeOffice.appName != lastAppName || versionText != lastVersion) {
                std::cout << "\n[DETECT] Application: " << activeOffice.appName
                          << " | Version: " << versionText
                          << " | Asset Key: " << activeOffice.largeImageKey
                          << " | Small Asset: " << activeOffice.smallImageKey
                          << std::endl;

                lastAppName = activeOffice.appName;
                lastVersion = versionText;
            }
            
            std::string stateLabel = "";
            
            if (activeOffice.appType == OFFICE_EXCEL && !activeOffice.state.empty()) {
                stateLabel = activeOffice.state;
            }
            else if (activeOffice.totalPages > 0) {
                stateLabel = "Page " + std::to_string(activeOffice.currentPage) + 
                           " of " + std::to_string(activeOffice.totalPages);
                if (activeOffice.wordCount > 0) {
                    stateLabel += " | " + FormatNumberWithComma(activeOffice.wordCount) + " words";
                }
            }

            if (discordConnected && (stateLabel != lastState || activeOffice.details != lastDetails || 
                activeOffice.largeImageKey != lastImageKey)) {
                
                DiscordRichPresence presence;
                memset(&presence, 0, sizeof(presence));

                presence.details = activeOffice.details.c_str();
                if (!stateLabel.empty()) {
                    presence.state = stateLabel.c_str();
                }

                presence.largeImageKey = activeOffice.largeImageKey.c_str();
                presence.largeImageText = activeOffice.displayName.c_str();
                presence.smallImageKey = activeOffice.smallImageKey.c_str();
                presence.smallImageText = activeOffice.version.c_str();
                presence.startTimestamp = startTime;

                Discord_UpdatePresence(&presence);

                std::cout << "[UPDATE] Rich Presence updated successfully." << std::endl;

                lastState = stateLabel;
                lastDetails = activeOffice.details;
                lastImageKey = activeOffice.largeImageKey;
            }

            std::cout << "\r[ACTIVE] " << activeOffice.appName << " " << versionText
                      << " | File: " << activeOffice.fileName << "     " << std::flush;
        }
        else {
            if (discordConnected && (lastState != "" || lastDetails != "")) {
                Discord_ClearPresence();
                lastState = "";
                lastDetails = "";
                lastImageKey = "";
                lastAppName = "";
                lastVersion = "";
                
                std::cout << "\n[INFO] No Office application in foreground. Rich Presence cleared." << std::endl;
            }
            std::cout << "\r[IDLE] Waiting for Office application...        " << std::flush;
        }

        Discord_RunCallbacks();
        Sleep(2000);
    }

    Discord_Shutdown();
    CoUninitialize();
    return 0;
}-