// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include <wrl/module.h>
#include <wrl/implements.h>
#include <shobjidl_core.h>
#include <wil/resource.h>
#include <Shellapi.h>
#include <Shlwapi.h>
#include <Strsafe.h>
#include <vector>
#include <string>
#include <fstream>
#include <ctime>
#include <sstream>
#include <shlobj.h>
#include "GetVisualStudioInstances.h"
#include "log.h"

using namespace Microsoft::WRL;

HMODULE g_hModule = nullptr;



BOOL APIENTRY DllMain(HMODULE hModule,
    DWORD ul_reason_for_call,
    LPVOID lpReserved)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
        g_hModule = hModule;
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

class VSVersion : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IExplorerCommand, IObjectWithSite>
{
public:
    virtual const EXPCMDFLAGS Flags() { return ECF_DEFAULT; }

    // IExplorerCommand methods
    IFACEMETHODIMP GetTitle(_In_opt_ IShellItemArray* items, _Outptr_result_nullonfailure_ PWSTR* name)
    {
        //Log("Title");
        *name = nullptr;
        auto title = wil::make_cotaskmem_string_nothrow(L"Visual Studio Selector");
        RETURN_IF_NULL_ALLOC(title);
        *name = title.release();
        return S_OK;
    }
    IFACEMETHODIMP GetIcon(_In_opt_ IShellItemArray* items, _Outptr_result_nullonfailure_ PWSTR* iconPath)
    {
        *iconPath = nullptr;
        PWSTR itemPath = nullptr;

        if (items)
        {
            DWORD count;
            RETURN_IF_FAILED(items->GetCount(&count));

            if (count > 0)
            {
                ComPtr<IShellItem> item;
                RETURN_IF_FAILED(items->GetItemAt(0, &item));

                RETURN_IF_FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, &itemPath));
                wil::unique_cotaskmem_string itemPathCleanup(itemPath);

                WCHAR modulePath[MAX_PATH];
                // Get the path to devenv.exe
                std::vector<std::pair<std::wstring, std::wstring>> instanceInfos;
                GetVisualStudioVersions(instanceInfos);
                if (!instanceInfos.empty())
                {
                    // Use the path of the latest version of Visual Studio
                    wcscpy_s(modulePath, instanceInfos.back().second.c_str());
                }

                auto iconPathStr = wil::make_cotaskmem_string_nothrow(modulePath);
                if (iconPathStr)
                {
                    *iconPath = iconPathStr.release();
                }
            }
        }
        return *iconPath ? S_OK : E_FAIL;
    }

    IFACEMETHODIMP GetToolTip(_In_opt_ IShellItemArray*, _Outptr_result_nullonfailure_ PWSTR* infoTip) { *infoTip = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP GetCanonicalName(_Out_ GUID* guidCommandName) { *guidCommandName = GUID_NULL;  return S_OK; }
    IFACEMETHODIMP GetState(_In_opt_ IShellItemArray* selection, _In_ BOOL okToBeSlow, _Out_ EXPCMDSTATE* cmdState)
    {
        *cmdState = ECS_ENABLED;
        return S_OK;
    }    
    IFACEMETHODIMP Invoke(_In_opt_ IShellItemArray* selection, _In_opt_ IBindCtx*) noexcept try
    {
       
        if (!selection)
        {
            // Debug message
            MessageBox(nullptr, L"Invalid argument", L"Debug Info", MB_OK);

            return E_INVALIDARG;
        }

        DWORD count;
        RETURN_IF_FAILED(selection->GetCount(&count));

        if (count == 0)
        {
            // Debug message
            MessageBox(nullptr, L"No items to process", L"Debug Info", MB_OK);

            return S_OK; // No items to process
        }

        ComPtr<IShellItem> item;
        RETURN_IF_FAILED(selection->GetItemAt(0, &item));

        PWSTR filePath;
        RETURN_IF_FAILED(item->GetDisplayName(SIGDN_FILESYSPATH, &filePath));
        wil::unique_cotaskmem_string filePathCleanup(filePath);

        // Get the directory of the DLL
        wchar_t dllPath[MAX_PATH];
        GetModuleFileName(nullptr, dllPath, MAX_PATH);
        PathRemoveFileSpec(dllPath);

        
        
    }
    CATCH_RETURN();
    IFACEMETHODIMP GetFlags(_Out_ EXPCMDFLAGS* flags) { *flags = Flags(); return S_OK; }
    IFACEMETHODIMP EnumSubCommands(_COM_Outptr_ IEnumExplorerCommand** enumCommands) { *enumCommands = nullptr; return E_NOTIMPL; }
    IFACEMETHODIMP SetSite(_In_ IUnknown* site) noexcept { m_site = site; return S_OK; }
    IFACEMETHODIMP GetSite(_In_ REFIID riid, _COM_Outptr_ void** site) noexcept { return m_site.CopyTo(riid, site); }
    void SetTitle(PCWSTR title)
    {
        m_title = wil::make_cotaskmem_string_nothrow(title);
    }

protected:
    ComPtr<IUnknown> m_site;
    wil::unique_cotaskmem_string m_title;
};


class VisualStudioCommand final : public VSVersion
{
public:
    VisualStudioCommand() {}

    VisualStudioCommand(const std::pair<std::wstring, std::wstring>& versionInfo)
        : m_version(versionInfo.first), m_installationPath(versionInfo.second)
    {
    }

        IFACEMETHODIMP GetTitle(_In_opt_ IShellItemArray * items, _Outptr_result_nullonfailure_ PWSTR * name) override
        {

           /* std::ostringstream oss;
            oss << "GetTitle: m_version at start is " << wstring_to_string(m_version);
            Log(oss.str());
            */

            if (!name) {
               // Log("GetTitle: name is null");
                return E_POINTER;
            }
            

            *name = nullptr;
            std::wstring title;

            // Log the value of m_version
            std::string message = "GetTitle: m_version is " + wstring_to_string(m_version);
          //  Log(message.c_str());

            // Use the full version string in the title
            title = L"Visual Studio " + m_version;

            auto titleStr = wil::make_cotaskmem_string_nothrow(title.c_str());
            if (!titleStr) {
                //Log("Failed to allocate memory for title");
                return E_OUTOFMEMORY;
            }

            *name = titleStr.release();
            return S_OK;
        }

        std::string wstring_to_string(const std::wstring & wstr)
        {
            std::string str(wstr.begin(), wstr.end());
            return str;
        }

        IFACEMETHODIMP Invoke(_In_opt_ IShellItemArray* selection, _In_opt_ IBindCtx*) noexcept override
        {
            DWORD count = 0;
            HRESULT hr = selection->GetCount(&count);
            if (FAILED(hr)) {
               // Log("Failed to get item count");
                return hr;
            }

            for (DWORD i = 0; i < count; ++i) {
                CComPtr<IShellItem> item;
                hr = selection->GetItemAt(i, &item);
                if (SUCCEEDED(hr)) {
                    LPWSTR filePath = nullptr;
                    hr = item->GetDisplayName(SIGDN_FILESYSPATH, &filePath);
                    if (SUCCEEDED(hr)) {
                        std::wstring filePathStr(filePath);
                        //Log("File path: " + wstring_to_string(filePathStr));

                        SHELLEXECUTEINFO sei = { 0 };
                        sei.cbSize = sizeof(SHELLEXECUTEINFO);
                        sei.fMask = SEE_MASK_DEFAULT;
                        sei.lpVerb = L"open";
                        sei.lpFile = m_installationPath.c_str();
                        sei.lpParameters = filePath;  // pass the file path to Visual Studio
                        sei.nShow = SW_SHOWNORMAL;

                        if (!ShellExecuteEx(&sei))
                        {
                            //Log("Failed to execute shell command");
                            return HRESULT_FROM_WIN32(GetLastError());
                        }

                        CoTaskMemFree(filePath);
                    }
                    else {
                        //Log("Failed to get display name");
                    }
                }
                else {
                    //Log("Failed to get item at index");
                }
            }

            return S_OK;
        }

private:
    std::wstring m_version;
    std::wstring m_installationPath;


   

};



class VSVersionEnumerator : public RuntimeClass<RuntimeClassFlags<ClassicCom>, IEnumExplorerCommand>
{
public:
    VSVersionEnumerator(const std::vector<std::pair<std::wstring, std::wstring>>& versionInfos)
    {
        for (const auto& versionInfo : versionInfos) {
            m_commands.push_back(Make<VisualStudioCommand>(versionInfo));
        }
        m_current = m_commands.cbegin();
    }
    IFACEMETHODIMP Next(ULONG celt, __out_ecount_part(celt, *pceltFetched) IExplorerCommand** apUICommand, __out_opt ULONG* pceltFetched)
    {
        ULONG fetched{ 0 };
        wil::assign_to_opt_param(pceltFetched, 0ul);

        for (ULONG i = 0; (i < celt) && (m_current != m_commands.cend()); i++)
        {
            m_current->CopyTo(&apUICommand[0]);
            m_current++;
            fetched++;
        }

        wil::assign_to_opt_param(pceltFetched, fetched);
        return (fetched == celt) ? S_OK : S_FALSE;
        
    }
    IFACEMETHODIMP Skip(ULONG /*celt*/) { return E_NOTIMPL; }
    IFACEMETHODIMP Reset()
    {
        m_current = m_commands.cbegin();
        return S_OK;
    }
    IFACEMETHODIMP Clone(__deref_out IEnumExplorerCommand** ppenum) { *ppenum = nullptr; return E_NOTIMPL; }

private:
    std::vector<ComPtr<IExplorerCommand>> m_commands;
    std::vector<ComPtr<IExplorerCommand>>::const_iterator m_current;
};



class __declspec(uuid("7A1E471F-0D43-4122-B1C4-D1AACE76CE9B")) VSVersion1 final : public VSVersion
{

    const EXPCMDFLAGS Flags() override { return ECF_HASSUBCOMMANDS; }

    IFACEMETHODIMP EnumSubCommands(_COM_Outptr_ IEnumExplorerCommand** enumCommands)
    {
        *enumCommands = nullptr;
        std::vector<std::pair<std::wstring, std::wstring>> instanceInfos;
        GetVisualStudioVersions(instanceInfos);
        auto e = Make<VSVersionEnumerator>(instanceInfos);
        return e->QueryInterface(IID_PPV_ARGS(enumCommands));
    }

    std::string wstring_to_string(const std::wstring& wstr)
    {
        std::string str(wstr.begin(), wstr.end());
        return str;
    }
private:
    std::vector<ComPtr<IExplorerCommand>> m_commands;

};


CoCreatableClass(VSVersion1)


STDAPI DllGetActivationFactory(_In_ HSTRING activatableClassId, _COM_Outptr_ IActivationFactory** factory)
{
    return Module<ModuleType::InProc>::GetModule().GetActivationFactory(activatableClassId, factory);
}
STDAPI DllCanUnloadNow()
{
    return Module<InProc>::GetModule().GetObjectCount() == 0 ? S_OK : S_FALSE;
}
STDAPI DllGetClassObject(_In_ REFCLSID rclsid, _In_ REFIID riid, _COM_Outptr_ void** instance)
{
    return Module<InProc>::GetModule().GetClassObject(rclsid, riid, instance);
}
