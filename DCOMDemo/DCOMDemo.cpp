// DCOMDemo.cpp : Implementation of WinMain


#include "pch.h"
#include "framework.h"
#include "resource.h"
#include "DCOMDemo_i.h"

#include <atlbase.h>
#include <iostream>
using namespace ATL;
void PrintErrorMessage(const std::string& message) {
    DWORD errorCode = GetLastError();
    std::cerr << "Error: " << message << ". Error code: " << errorCode << std::endl;
}

std::wstring GetCurrentProcessPath() {
    wchar_t buffer[MAX_PATH];
    GetModuleFileNameW(nullptr, buffer, MAX_PATH);
    std::wstring path(buffer);
    size_t lastSlashPos = path.find_last_of(L"\\/");
    if (lastSlashPos != std::wstring::npos) {
        path = path.substr(0, lastSlashPos);
    }
    return path;
}

std::wstring GetCurrentFolderPath() {
    wchar_t buffer[MAX_PATH];
    GetCurrentDirectoryW(MAX_PATH, buffer);
    return buffer;
}

int SetupRegistyKeys() {
    HKEY hKey;
    LONG result;

    // Open or create the root key
    result = RegCreateKeyExW(HKEY_CLASSES_ROOT, L"Interface\\{A68E05E3-4094-4C62-A050-B75146882D38}", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL);
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to create or open root key");
        return 1;
    }

    // Set the default value
    result = RegSetValueExW(hKey, NULL, 0, REG_SZ, (BYTE*)L"ICOMServer", sizeof(L"ICOMServer"));
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to set default value");
        RegCloseKey(hKey);
        return 1;
    }

    // Create and set the ProxyStubClsid32 subkey
    HKEY subKey;
    result = RegCreateKeyExW(hKey, L"ProxyStubClsid32", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &subKey, NULL);
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to create or open ProxyStubClsid32 subkey");
        RegCloseKey(hKey);
        return 1;
    }
    result = RegSetValueExW(subKey, NULL, 0, REG_SZ, (BYTE*)L"{00020424-0000-0000-C000-000000000046}", sizeof(L"{00020424-0000-0000-C000-000000000046}"));
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to set value for ProxyStubClsid32 subkey");
        RegCloseKey(subKey);
        RegCloseKey(hKey);
        return 1;
    }
    RegCloseKey(subKey);

    // Create and set the TypeLib subkey
    result = RegCreateKeyExW(hKey, L"TypeLib", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &subKey, NULL);
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to create or open TypeLib subkey");
        RegCloseKey(hKey);
        return 1;
    }
    result = RegSetValueExW(subKey, NULL, 0, REG_SZ, (BYTE*)L"{4F0D5B9A-F7CC-4C28-A637-A3AA67A93793}", sizeof(L"{4F0D5B9A-F7CC-4C28-A637-A3AA67A93793}"));
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to set value for TypeLib subkey");
        RegCloseKey(subKey);
        RegCloseKey(hKey);
        return 1;
    }

    result = RegSetValueExW(subKey, L"Version", 0, REG_SZ, (BYTE*)L"1.0", sizeof(L"1.0"));
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to set value for TypeLib\\Version subkey");
        RegCloseKey(subKey);
        RegCloseKey(hKey);
        return 1;
    }

    //// Create and set the 0 subkey
    HKEY subSubKey;
    RegCloseKey(subKey);
    RegCloseKey(hKey);

    // Open or create the TypeLib key
    result = RegCreateKeyExW(HKEY_CLASSES_ROOT, L"TypeLib\\{4F0D5B9A-F7CC-4C28-A637-A3AA67A93793}", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL);
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to create or open TypeLib key");
        return 1;
    }

    // Create and set the 1.0 subkey
    result = RegCreateKeyExW(hKey, L"1.0", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &subKey, NULL);
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to create or open 1.0 subkey");
        RegCloseKey(hKey);
        return 1;
    }
    result = RegSetValueExW(subKey, NULL, 0, REG_SZ, (BYTE*)L"DCOMDemoLib", sizeof(L"DCOMDemoLib"));
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to set value for 1.0 subkey");
        RegCloseKey(subKey);
        RegCloseKey(hKey);
        return 1;
    }

    // Create and set the 0 subkey
    result = RegCreateKeyExW(subKey, L"0", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &subSubKey, NULL);
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to create or open 0 subkey");
        RegCloseKey(subKey);
        RegCloseKey(hKey);
        return 1;
    }

    HKEY subSubSubKey;
    result = RegCreateKeyExW(subSubKey, L"win32", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &subSubSubKey, NULL);
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to create or open win32 subkey");
        RegCloseKey(subSubKey);
        RegCloseKey(subSubSubKey);

        RegCloseKey(hKey);
        return 1;
    }
    std::wstring currentProcessPath = GetCurrentProcessPath();
    result = RegSetValueExW(subSubSubKey, NULL, 0, REG_SZ, (BYTE*)currentProcessPath.c_str(), static_cast<DWORD>((currentProcessPath.length() + 1) * sizeof(wchar_t)));
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to set value for win32 subkey");
        RegCloseKey(subSubSubKey);
        RegCloseKey(subSubKey);
        RegCloseKey(subKey);
        RegCloseKey(hKey);
        return 1;
    }

    RegCloseKey(subSubSubKey);
    RegCloseKey(subSubKey);
    // Create and set the FLAGS subkey
    result = RegCreateKeyExW(subKey, L"FLAGS", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &subSubKey, NULL);
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to create or open FLAGS subkey");
        RegCloseKey(subKey);
        RegCloseKey(hKey);
        return 1;
    }
    result = RegSetValueExW(subSubKey, NULL, 0, REG_SZ, (BYTE*)L"0", sizeof(L"0"));
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to set value for FLAGS subkey");
        RegCloseKey(subSubKey);
        RegCloseKey(subKey);
        RegCloseKey(hKey);
        return 1;
    }
    RegCloseKey(subSubKey);

    // Create and set the HELPDIR subkey
    result = RegCreateKeyExW(subKey, L"HELPDIR", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &subSubKey, NULL);
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to create or open HELPDIR subkey");
        RegCloseKey(subKey);
        RegCloseKey(hKey);
        return 1;
    }






    std::wstring currentFolderPath = GetCurrentFolderPath();
    result = RegSetValueExW(subSubKey, NULL, 0, REG_SZ, (BYTE*)currentFolderPath.c_str(), static_cast<DWORD>((currentFolderPath.length() + 1) * sizeof(wchar_t)));
    if (result != ERROR_SUCCESS) {
        // Handle error
        PrintErrorMessage("Failed to set value for HELPDIR subkey");
        RegCloseKey(subSubKey);
        RegCloseKey(subKey);
        RegCloseKey(hKey);
        return 1;
    }
    RegCloseKey(subSubKey);

    RegCloseKey(subKey);
    RegCloseKey(hKey);

    return 0;
}



class CDCOMDemoModule : public ATL::CAtlExeModuleT< CDCOMDemoModule >
{
public :
	DECLARE_LIBID(LIBID_DCOMDemoLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_DCOMDEMO, "{4f0d5b9a-f7cc-4c28-a637-a3aa67a93793}")
};

CDCOMDemoModule _AtlModule;



//
extern "C" int WINAPI _tWinMain(HINSTANCE /*hInstance*/, HINSTANCE /*hPrevInstance*/,
								LPTSTR /*lpCmdLine*/, int nShowCmd)
{
    SetupRegistyKeys();
	return _AtlModule.WinMain(nShowCmd);
}

