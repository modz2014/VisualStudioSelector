#include <Windows.h>
#include <iostream>
#include <comutil.h>
#include <atlbase.h>
#include <atlcom.h>
#include <comdef.h>
#include <vector>
#include <Setup.Configuration.h>
#include "log.h"
#include <codecvt>
#include <locale>




std::wstring GetYearFromVersion(const std::wstring& version) {
    if (version.substr(0, 2) == L"15") {
        return L"2017";
    }
    else if (version.substr(0, 2) == L"16") {
        return L"2019";
    }
    else if (version.substr(0, 2) == L"17") {
        return L"2022";
    }
    else {
        // handle other versions or return an empty string if the version is not recognized
        return L"";
    }
}

std::wstring string_to_wstring(const std::string& str) {
    std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t> converter;
    return converter.from_bytes(str);
}

std::string wstring_to_string(const std::wstring& wstr) {
    std::wstring_convert<std::codecvt_utf8<wchar_t>, wchar_t> converter;
    return converter.to_bytes(wstr);
}

// Function to get Visual Studio versions
HRESULT GetVisualStudioVersions(std::vector<std::pair<std::wstring, std::wstring>>& instanceInfos) {
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr)) {
        MessageBox(NULL, L"Failed to initialize COM library.", L"Error", MB_OK);
        //Log("Failed to initialize COM library.");
        return hr;
    }

    CComPtr<ISetupConfiguration2> setupConfiguration;
    hr = setupConfiguration.CoCreateInstance(__uuidof(SetupConfiguration));
    if (FAILED(hr)) {
        //Log("Failed to create setup configuration.");
        CoUninitialize();
        return hr;
    }

    CComPtr<IEnumSetupInstances> enumInstances;
    hr = setupConfiguration->EnumAllInstances(&enumInstances);
    if (FAILED(hr)) {
        //Log("Failed to enumerate instances.");
        CoUninitialize();
        return hr;
    }

    CComPtr<ISetupInstance> instance;
    ULONG fetched = 0;
    while (SUCCEEDED(enumInstances->Next(1, &instance, &fetched)) && fetched > 0) {
        CComQIPtr<ISetupInstance2> instance2;
        hr = instance->QueryInterface(__uuidof(ISetupInstance2), (void**)&instance2);
        if (SUCCEEDED(hr) && instance2) {
            BSTR bstrInstallationVersion;
            hr = instance2->GetInstallationVersion(&bstrInstallationVersion);
            if (SUCCEEDED(hr)) {
                // Extract the year from the full version string
                std::wstring versionYear = GetYearFromVersion(bstrInstallationVersion);
                //Log("Version year: " + wstring_to_string(versionYear));

                // Get the installation path
                BSTR bstrInstallationPath;
                hr = instance2->GetInstallationPath(&bstrInstallationPath);
                if (SUCCEEDED(hr)) {
                    std::wstring installationPath(bstrInstallationPath, SysStringLen(bstrInstallationPath));

                    // Append the relative path of devenv.exe to the installation path
                    std::wstring devenvPath = installationPath + L"\\Common7\\IDE\\devenv.exe";

                    //Log("devenv.exe path: " + wstring_to_string(devenvPath));
                    instanceInfos.push_back(std::make_pair(versionYear, devenvPath));
                }


                SysFreeString(bstrInstallationVersion);
                SysFreeString(bstrInstallationPath);
            }
        }
        instance = nullptr;
    }
    CoUninitialize();
    return S_OK;
}
