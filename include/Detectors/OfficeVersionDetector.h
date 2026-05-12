#pragma once

#include <windows.h>
#include <string>
#include "../Core/OfficeInfo.h"

// Function declarations
OfficeVersion CheckClickToRunVersion();
OfficeVersion CheckMSIVersion();
OfficeVersion DetectVersionFromCOM(IDispatch* pApp);
OfficeVersion DetectOfficeVersion(IDispatch* pApp);
std::string SelectLogoKey(OfficeAppType appType, int officeYear);
std::string SelectSmallImageKey(OfficeAppType appType);