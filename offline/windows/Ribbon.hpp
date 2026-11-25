#pragma once
#include <windows.h>
#include <oaidl.h>

inline constexpr GUID WorkbenchClsid = {0x1ce05ade,0x43df,0x432e,{0xaf,0xf5,0x2d,0xf7,0xd7,0x5a,0xba,0x29}};
struct __declspec(uuid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")) OfficeAddin : IDispatch {
    virtual HRESULT STDMETHODCALLTYPE OnConnection(IDispatch*, LONG, IDispatch*, SAFEARRAY**) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnDisconnection(LONG, SAFEARRAY**) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**) = 0;
};
struct __declspec(uuid("000C0396-0000-0000-C000-000000000046")) OfficeRibbon : IDispatch {
    virtual HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR ribbonId, BSTR* xml) = 0;
};
