#include "Ribbon.hpp"
#include <cstdio>
#include <cwchar>
static HRESULT verify()
{
    CLSID cls{};
    HRESULT hr = CLSIDFromProgID(L"TrioBridge.Workbench", &cls);
    if (FAILED(hr))
        return hr;
    OfficeRibbon *ribbon = nullptr;
    hr = CoCreateInstance(cls, nullptr, CLSCTX_INPROC_SERVER, __uuidof(OfficeRibbon),
                          reinterpret_cast<void **>(&ribbon));
    if (FAILED(hr))
        return hr;
    OfficeAddin *entry = nullptr;
    hr = ribbon->QueryInterface(__uuidof(OfficeAddin), reinterpret_cast<void **>(&entry));
    if (SUCCEEDED(hr))
    {
        wchar_t button[] = L"OpenWorkbench";
        LPOLESTR names[] = {button};
        DISPID id = DISPID_UNKNOWN;
        hr = entry->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
        if (SUCCEEDED(hr) && id == DISPID_UNKNOWN)
            hr = E_FAIL;
        entry->Release();
    }
    BSTR xml = nullptr;
    if (SUCCEEDED(hr))
        hr = ribbon->GetCustomUI(nullptr, &xml);
    if (SUCCEEDED(hr) && (!xml || !wcsstr(xml, L"OpenWorkbench") || !wcsstr(xml, L"TrioBridgeTab")))
        hr = E_FAIL;
    SysFreeString(xml);
    ribbon->Release();
    return hr;
}
int wmain()
{
    const auto init = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(init))
        return 1;
    auto hr = verify();
    CoUninitialize();
    std::printf("Office COM activation, ribbon XML and callback (%zu-bit): 0x%08lX\n", sizeof(void *) * 8,
                static_cast<unsigned long>(hr));
    return FAILED(hr) ? 1 : 0;
}
