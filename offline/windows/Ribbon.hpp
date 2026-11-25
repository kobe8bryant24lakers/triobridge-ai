#pragma once
#include <oaidl.h>
#include <windows.h>

inline constexpr GUID WorkbenchClsid = {
    0x1ce05ade, 0x43df, 0x432e, {0xaf, 0xf5, 0x2d, 0xf7, 0xd7, 0x5a, 0xba, 0x29}};
struct __declspec(uuid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")) OfficeAddin : IDispatch
{
    virtual HRESULT STDMETHODCALLTYPE OnConnection(IDispatch *, LONG, IDispatch *,
                                                   SAFEARRAY **) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnDisconnection(LONG, SAFEARRAY **) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY **) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY **) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY **) = 0;
};
struct __declspec(uuid("000C0396-0000-0000-C000-000000000046")) OfficeRibbon : IDispatch
{
    virtual HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR ribbonId, BSTR *xml) = 0;
};
// Office's dual interfaces: the seven IDispatch slots precede these methods.
// _CustomTaskPane is deliberately opaque; its properties are used via
// IDispatch.
// https://learn.microsoft.com/dotnet/api/microsoft.office.core.ictpfactory
struct __declspec(uuid("000C033B-0000-0000-C000-000000000046")) OfficeTaskPane : IDispatch
{
};
struct __declspec(uuid("000C033D-0000-0000-C000-000000000046")) OfficePaneFactory : IDispatch
{
    virtual HRESULT STDMETHODCALLTYPE CreateCTP(BSTR, BSTR, VARIANT, OfficeTaskPane **) = 0;
};
struct __declspec(uuid("000C033E-0000-0000-C000-000000000046")) OfficePaneConsumer : IDispatch
{
    virtual HRESULT STDMETHODCALLTYPE CTPFactoryAvailable(OfficePaneFactory *) = 0;
};
#ifdef __MINGW32__
// MinGW's __uuidof maps to explicit UUID traits instead of MSVC's attribute.
__CRT_UUID_DECL(OfficeAddin, 0xB65AD801, 0xABAF, 0x11D0, 0xBB, 0x8B, 0x00, 0xA0, 0xC9, 0x0F, 0x27,
                0x44)
__CRT_UUID_DECL(OfficeRibbon, 0x000C0396, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x46)
__CRT_UUID_DECL(OfficeTaskPane, 0x000C033B, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x46)
__CRT_UUID_DECL(OfficePaneFactory, 0x000C033D, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x46)
__CRT_UUID_DECL(OfficePaneConsumer, 0x000C033E, 0x0000, 0x0000, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00,
                0x00, 0x46)
#endif
