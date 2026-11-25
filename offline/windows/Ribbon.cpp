#include "Ribbon.hpp"
#include <atomic>
#include <cwchar>
#include <new>
#include <string>
#include <vector>

namespace
{
HMODULE module;
std::atomic<LONG> liveObjects{0};
constexpr wchar_t ribbonXml[] =
    LR"(<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="TrioBridgeTab" label="三桥 AI"><group id="TrioBridgeWork" label="文档工具"><button id="TrioBridgeOpen" label="打开工作台" size="large" imageMso="ReviewNewComment" onAction="OpenWorkbench"/></group></tab></tabs></ribbon></customUI>)";
constexpr DISPID openButton = 41;

std::wstring modulePath(HMODULE image)
{
    std::wstring path(32768, L'\0');
    DWORD length = GetModuleFileNameW(image, path.data(), static_cast<DWORD>(path.size()));
    if (!length || length >= path.size())
        return {};
    path.resize(length);
    return path;
}
const wchar_t *currentHost()
{
    const auto path = modulePath(nullptr);
    const auto pos = path.find_last_of(L"\\/");
    const auto name = path.substr(pos == std::wstring::npos ? 0 : pos + 1);
    if (_wcsicmp(name.c_str(), L"WINWORD.EXE") == 0)
        return L"word";
    if (_wcsicmp(name.c_str(), L"EXCEL.EXE") == 0)
        return L"excel";
    if (_wcsicmp(name.c_str(), L"POWERPNT.EXE") == 0)
        return L"powerpoint";
    return nullptr;
}
HRESULT launchWorkbench()
{
    const auto host = currentHost();
    if (!host)
        return E_UNEXPECTED;
    auto path = modulePath(module);
    const auto slash = path.find_last_of(L"\\/");
    if (slash == std::wstring::npos)
        return E_FAIL;
    const auto folder = path.substr(0, slash);
    const auto exe = folder + L"\\TrioBridge AI.exe";
    std::wstring args = L"\"" + exe + L"\" --office=" + host;
    STARTUPINFOW startup{};
    startup.cb = sizeof(startup);
    PROCESS_INFORMATION process{};
    if (!CreateProcessW(exe.c_str(), args.data(), nullptr, nullptr, FALSE, 0, nullptr, folder.c_str(),
                        &startup, &process))
        return HRESULT_FROM_WIN32(GetLastError());
    CloseHandle(process.hThread);
    CloseHandle(process.hProcess);
    return S_OK;
}
class Ribbon final : public OfficeAddin, public OfficeRibbon
{
    std::atomic<ULONG> refs{1};

  public:
    Ribbon()
    {
        ++liveObjects;
    }
    ~Ribbon()
    {
        --liveObjects;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id, void **result) override
    {
        if (!result)
            return E_POINTER;
        *result = nullptr;
        if (id == IID_IUnknown || id == IID_IDispatch || id == __uuidof(OfficeAddin))
            *result = static_cast<OfficeAddin *>(this);
        else if (id == __uuidof(OfficeRibbon))
            *result = static_cast<OfficeRibbon *>(this);
        else
            return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++refs;
    }
    ULONG STDMETHODCALLTYPE Release() override
    {
        const auto n = --refs;
        if (n == 0)
            delete this;
        return n;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *count) override
    {
        if (!count)
            return E_POINTER;
        *count = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo **) override
    {
        return DISP_E_BADINDEX;
    }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR *names, UINT count, LCID, DISPID *ids) override
    {
        if (!names || !ids)
            return E_POINTER;
        if (count != 1 || !names[0])
            return DISP_E_UNKNOWNNAME;
        static const wchar_t *methods[] = {L"OnConnection", L"OnDisconnection", L"OnAddInsUpdate",
                                           L"OnStartupComplete", L"OnBeginShutdown"};
        for (int i = 0; i < 5; ++i)
            if (_wcsicmp(names[0], methods[i]) == 0)
            {
                ids[0] = i + 1;
                return S_OK;
            }
        if (_wcsicmp(names[0], L"OpenWorkbench") == 0)
        {
            ids[0] = openButton;
            return S_OK;
        }
        if (_wcsicmp(names[0], L"GetCustomUI") == 0)
        {
            ids[0] = 42;
            return S_OK;
        }
        ids[0] = DISPID_UNKNOWN;
        return DISP_E_UNKNOWNNAME;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD flags, DISPPARAMS *, VARIANT *result,
                                     EXCEPINFO *, UINT *) override
    {
        if (!(flags & DISPATCH_METHOD))
            return DISP_E_MEMBERNOTFOUND;
        if (id == openButton)
        {
            try
            {
                return launchWorkbench();
            }
            catch (const std::bad_alloc &)
            {
                return E_OUTOFMEMORY;
            }
        }
        if (id >= 1 && id <= 5)
            return S_OK;
        if (id == 42 && result)
        {
            VariantInit(result);
            result->vt = VT_BSTR;
            return GetCustomUI(nullptr, &result->bstrVal);
        }
        return DISP_E_MEMBERNOTFOUND;
    }
    HRESULT STDMETHODCALLTYPE OnConnection(IDispatch *, LONG, IDispatch *, SAFEARRAY **) override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnDisconnection(LONG, SAFEARRAY **) override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY **) override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY **) override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY **) override
    {
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR, BSTR *xml) override
    {
        if (!xml)
            return E_POINTER;
        *xml = SysAllocString(ribbonXml);
        return *xml ? S_OK : E_OUTOFMEMORY;
    }
};
class Factory final : public IClassFactory
{
    std::atomic<ULONG> refs{1};

  public:
    Factory()
    {
        ++liveObjects;
    }
    ~Factory()
    {
        --liveObjects;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id, void **out) override
    {
        if (!out)
            return E_POINTER;
        *out = nullptr;
        if (id != IID_IUnknown && id != IID_IClassFactory)
            return E_NOINTERFACE;
        *out = static_cast<IClassFactory *>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++refs;
    }
    ULONG STDMETHODCALLTYPE Release() override
    {
        auto n = --refs;
        if (!n)
            delete this;
        return n;
    }
    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown *outer, REFIID id, void **out) override
    {
        if (!out)
            return E_POINTER;
        *out = nullptr;
        if (outer)
            return CLASS_E_NOAGGREGATION;
        auto object = new (std::nothrow) Ribbon();
        if (!object)
            return E_OUTOFMEMORY;
        auto hr = object->QueryInterface(id, out);
        object->Release();
        return hr;
    }
    HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override
    {
        if (lock)
            ++liveObjects;
        else
            --liveObjects;
        return S_OK;
    }
};
} // namespace
BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID)
{
    if (reason == DLL_PROCESS_ATTACH)
        module = instance;
    return TRUE;
}
extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID cls, REFIID id, void **out)
{
    if (!out)
        return E_POINTER;
    *out = nullptr;
    if (cls != WorkbenchClsid)
        return CLASS_E_CLASSNOTAVAILABLE;
    auto f = new (std::nothrow) Factory();
    if (!f)
        return E_OUTOFMEMORY;
    const auto hr = f->QueryInterface(id, out);
    f->Release();
    return hr;
}
extern "C" HRESULT __stdcall DllCanUnloadNow()
{
    return liveObjects == 0 ? S_OK : S_FALSE;
}
