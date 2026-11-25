#include "Ribbon.hpp"
#include <cstdio>
#include <cwchar>
#include <functional>
#include <memory>
#include <stdexcept>
#include <string>
#include <vector>
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
    OfficePaneConsumer *consumer = nullptr;
    if (SUCCEEDED(hr))
        hr = ribbon->QueryInterface(__uuidof(OfficePaneConsumer),
                                    reinterpret_cast<void **>(&consumer));
    if (consumer)
        consumer->Release();
    if (SUCCEEDED(hr))
    {
        wchar_t button[] = L"OpenWorkbench";
        LPOLESTR names[] = {button};
        DISPID id = DISPID_UNKNOWN;
        hr = entry->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
        if (SUCCEEDED(hr) && id == DISPID_UNKNOWN)
            hr = E_FAIL;
    }
    if (entry)
        entry->Release();
    BSTR xml = nullptr;
    if (SUCCEEDED(hr))
        hr = ribbon->GetCustomUI(nullptr, &xml);
    if (SUCCEEDED(hr) && (!xml || !wcsstr(xml, L"OpenWorkbench") ||
                          !wcsstr(xml, L"TrioBridgeTab") || !wcsstr(xml, L"OpenTaskPane")))
        hr = E_FAIL;
    SysFreeString(xml);
    ribbon->Release();
    return hr;
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID, REFIID, void **);
extern "C" HRESULT __stdcall DllCanUnloadNow();
namespace
{
void check(bool condition, const char *message)
{
    if (!condition)
        throw std::runtime_error(message);
}
void ok(HRESULT hr, const char *message)
{
    check(SUCCEEDED(hr), message);
}
template <class T> struct Owned
{
    T *p = nullptr;
    ~Owned()
    {
        if (p)
            p->Release();
    }
    T *operator->() const
    {
        return p;
    }
};
void dispatchResult(VARIANT *out, IDispatch *object)
{
    if (!out)
        throw std::runtime_error("missing result");
    VariantInit(out);
    out->vt = VT_DISPATCH;
    out->pdispVal = object;
    object->AddRef();
}
template <class Interface = IDispatch> class Mock : public Interface
{
    std::vector<std::wstring> names;

  public:
    ULONG refs = 1;
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id, void **out) override
    {
        if (!out)
            return E_POINTER;
        *out = nullptr;
        if (id != IID_IUnknown && id != IID_IDispatch && id != __uuidof(Interface))
            return E_NOINTERFACE;
        *out = static_cast<Interface *>(this);
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override
    {
        return ++refs;
    }
    ULONG STDMETHODCALLTYPE Release() override
    {
        check(refs > 0, "unbalanced Release");
        return --refs;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *n) override
    {
        *n = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo **) override
    {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR *name, UINT count, LCID,
                                            DISPID *id) override
    {
        if (count != 1 || !name || !name[0] || !id)
            return E_INVALIDARG;
        for (size_t i = 0; i < names.size(); ++i)
            if (names[i] == name[0])
            {
                *id = static_cast<DISPID>(i + 1);
                return S_OK;
            }
        names.emplace_back(name[0]);
        *id = static_cast<DISPID>(names.size());
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD flags, DISPPARAMS *params,
                                     VARIANT *result, EXCEPINFO *, UINT *) override
    {
        if (id < 1 || static_cast<size_t>(id) > names.size())
            return DISP_E_MEMBERNOTFOUND;
        return call(names[id - 1], flags, params, result);
    }
    virtual HRESULT call(const std::wstring &, WORD, DISPPARAMS *, VARIANT *)
    {
        return DISP_E_MEMBERNOTFOUND;
    }
};
struct Windows : Mock<>
{
    std::vector<IDispatch *> open;
    bool busy = false;
    HRESULT call(const std::wstring &name, WORD, DISPPARAMS *p, VARIANT *out) override
    {
        if (busy)
            return RPC_E_CALL_REJECTED;
        if (name == L"Count")
        {
            out->vt = VT_I4;
            out->lVal = static_cast<LONG>(open.size());
            return S_OK;
        }
        if (name == L"Item" && p->cArgs == 1 && p->rgvarg[0].vt == VT_I4)
        {
            LONG i = p->rgvarg[0].lVal;
            if (i < 1 || static_cast<size_t>(i) > open.size())
                return DISP_E_BADINDEX;
            dispatchResult(out, open[i - 1]);
            return S_OK;
        }
        return DISP_E_MEMBERNOTFOUND;
    }
};
struct Application : Mock<>
{
    Windows windows;
    HRESULT call(const std::wstring &name, WORD, DISPPARAMS *, VARIANT *out) override
    {
        check(name != L"ActiveWindow" && name != L"ActiveDocument", "must not read active target");
        if (name == L"Windows")
        {
            dispatchResult(out, &windows);
            return S_OK;
        }
        return DISP_E_MEMBERNOTFOUND;
    }
};
struct Window : Mock<>
{
    Application *app;
    explicit Window(Application *a) : app(a)
    {
    }
    HRESULT call(const std::wstring &name, WORD, DISPPARAMS *, VARIANT *out) override
    {
        if (name == L"Application")
        {
            dispatchResult(out, app);
            return S_OK;
        }
        return DISP_E_MEMBERNOTFOUND;
    }
};
struct Button : Mock<>
{
    IDispatch *window;
    explicit Button(IDispatch *w) : window(w)
    {
    }
    HRESULT call(const std::wstring &name, WORD, DISPPARAMS *, VARIANT *out) override
    {
        if (name == L"Context" && window)
        {
            dispatchResult(out, window);
            return S_OK;
        }
        return DISP_E_MEMBERNOTFOUND;
    }
};
struct Content : Mock<>
{
    IDispatch *expectedApp, *expectedWindow;
    int attached = 0;
    bool fail = false;
    std::function<void()> duringAttach;
    Content(IDispatch *app, IDispatch *window) : expectedApp(app), expectedWindow(window)
    {
    }
    HRESULT call(const std::wstring &name, WORD flags, DISPPARAMS *p, VARIANT *) override
    {
        if (name != L"Attach")
            return DISP_E_MEMBERNOTFOUND;
        check(flags == DISPATCH_METHOD && p->cArgs == 3, "Attach three arguments");
        check(p->rgvarg[0].vt == VT_BSTR && !wcscmp(p->rgvarg[0].bstrVal, L"word"),
              "Attach host order");
        check(p->rgvarg[1].vt == VT_DISPATCH && p->rgvarg[1].pdispVal == expectedWindow,
              "Attach window order");
        check(p->rgvarg[2].vt == VT_DISPATCH && p->rgvarg[2].pdispVal == expectedApp,
              "Attach app order");
        ++attached;
        if (duringAttach)
            duringAttach();
        return fail ? E_FAIL : S_OK;
    }
};
struct Pane final : Mock<OfficeTaskPane>
{
    Content content;
    int deleted = 0, shown = 0;
    bool failShow = false;
    Pane(IDispatch *app, IDispatch *window) : content(app, window)
    {
    }
    HRESULT call(const std::wstring &name, WORD flags, DISPPARAMS *p, VARIANT *out) override
    {
        if (name == L"ContentControl")
        {
            dispatchResult(out, &content);
            return S_OK;
        }
        if (name == L"Delete")
        {
            ++deleted;
            return S_OK;
        }
        if (name == L"Visible")
        {
            check(flags == DISPATCH_PROPERTYPUT && p->cArgs == 1 && p->cNamedArgs == 1 &&
                      p->rgdispidNamedArgs[0] == DISPID_PROPERTYPUT && p->rgvarg[0].vt == VT_BOOL &&
                      p->rgvarg[0].boolVal == VARIANT_TRUE,
                  "Visible property put");
            ++shown;
            return failShow ? E_FAIL : S_OK;
        }
        if (name == L"DockPosition" || name == L"Width")
        {
            check(flags == DISPATCH_PROPERTYPUT && p->cArgs == 1, "pane geometry property put");
            check(p->rgvarg[0].lVal == (name == L"Width" ? 360 : 2), "pane geometry");
            return S_OK;
        }
        return DISP_E_MEMBERNOTFOUND;
    }
};
struct PaneFactory : Mock<OfficePaneFactory>
{
    IDispatch *app;
    bool failAttach = false;
    std::function<void()> duringAttach;
    std::vector<std::unique_ptr<Pane>> created;
    explicit PaneFactory(IDispatch *a) : app(a)
    {
    }
    HRESULT STDMETHODCALLTYPE CreateCTP(BSTR prog, BSTR title, VARIANT parent,
                                        OfficeTaskPane **out) override
    {
        check(!wcscmp(prog, L"TrioBridge.TaskPane") && SysStringLen(title) > 0, "ActiveX identity");
        check(parent.vt == VT_DISPATCH && parent.pdispVal, "explicit parent window required");
        auto pane = std::make_unique<Pane>(app, parent.pdispVal);
        pane->content.fail = failAttach;
        pane->content.duringAttach = duringAttach;
        *out = pane.get();
        (*out)->AddRef();
        created.push_back(std::move(pane));
        return S_OK;
    }
};
HRESULT action(OfficeAddin *addin, IDispatch *button)
{
    wchar_t name[] = L"OpenTaskPane";
    LPOLESTR names[]{name};
    DISPID id;
    HRESULT hr = addin->GetIDsOfNames(IID_NULL, names, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr))
        return hr;
    VARIANT arg{};
    arg.vt = VT_DISPATCH;
    arg.pdispVal = button;
    DISPPARAMS p{&arg, nullptr, 1, 0};
    return addin->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &p, nullptr, nullptr,
                         nullptr);
}
void behavior()
{
    check(DllCanUnloadNow() == S_OK, "initial module reference count");
    Application app;
    Window first(&app), second(&app);
    app.windows.open = {&first, &second};
    Button one(&first), two(&second), missing(nullptr);
    PaneFactory factory(&app);
    {
        Owned<IClassFactory> cls;
        ok(DllGetClassObject(WorkbenchClsid, IID_IClassFactory, reinterpret_cast<void **>(&cls.p)),
           "class factory");
        Owned<OfficeAddin> addin;
        ok(cls->CreateInstance(nullptr, __uuidof(OfficeAddin), reinterpret_cast<void **>(&addin.p)),
           "addin instance");
        Owned<OfficePaneConsumer> consumer;
        ok(addin->QueryInterface(__uuidof(OfficePaneConsumer),
                                 reinterpret_cast<void **>(&consumer.p)),
           "pane consumer QI");
        Owned<OfficeRibbon> ribbon;
        ok(addin->QueryInterface(__uuidof(OfficeRibbon), reinterpret_cast<void **>(&ribbon.p)),
           "ribbon QI");
        Owned<IUnknown> a, b;
        ok(addin->QueryInterface(IID_IUnknown, reinterpret_cast<void **>(&a.p)), "addin identity");
        ok(consumer->QueryInterface(IID_IUnknown, reinterpret_cast<void **>(&b.p)),
           "consumer identity");
        check(a.p == b.p, "all interfaces share COM identity");
        BSTR xml = nullptr;
        ok(ribbon->GetCustomUI(nullptr, &xml), "ribbon vtable");
        check(wcsstr(xml, L"OpenWorkbench") && wcsstr(xml, L"OpenTaskPane"), "both buttons remain");
        SysFreeString(xml);
        // Both Office interfaces use DISPID 1. Calling the consumer must not be
        // mistaken for IDTExtensibility2.OnConnection or GetCustomUI.
        VARIANT f{};
        f.vt = VT_DISPATCH;
        f.pdispVal = &factory;
        DISPPARAMS p{&f, nullptr, 1, 0};
        ok(consumer->Invoke(1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &p, nullptr, nullptr,
                            nullptr),
           "consumer dispatch");
        ok(addin->OnConnection(&app, 0, nullptr, nullptr), "connection after early factory");
        ok(action(addin.p, &one), "first pane");
        check(factory.created.size() == 1 && factory.created[0]->content.attached == 1,
              "create and attach once");
        ok(action(addin.p, &one), "repeat opens existing pane");
        check(factory.created.size() == 1 && factory.created[0]->shown == 2 &&
                  factory.created[0]->content.attached == 1,
              "repeat only shows");
        ok(action(addin.p, &two), "second window pane");
        check(factory.created.size() == 2, "one pane per window");
        check(FAILED(action(addin.p, &missing)), "missing context rejected");
        Application foreign;
        Window foreignWindow(&foreign);
        Button foreignButton(&foreignWindow);
        check(action(addin.p, &foreignButton) == E_ACCESSDENIED, "foreign application rejected");
        check(factory.created.size() == 2, "invalid contexts never create");
        app.windows.busy = true;
        check(FAILED(action(addin.p, &two)) && factory.created[0]->deleted == 0,
              "busy Office keeps panes");
        app.windows.busy = false;
        app.windows.open = {&second};
        // The UI-thread timer removes a closed window even without a new click.
        const ULONGLONG deadline = GetTickCount64() + 4000;
        while (!factory.created[0]->deleted && GetTickCount64() < deadline)
        {
            MSG message;
            while (PeekMessageW(&message, nullptr, 0, 0, PM_REMOVE))
                DispatchMessageW(&message);
            Sleep(10);
        }
        check(factory.created[0]->deleted == 1 && factory.created[0]->refs == 1,
              "closed-window lifecycle cleanup");
        factory.created[1]->failShow = true;
        ok(action(addin.p, &two), "stale pane recreated");
        check(factory.created[1]->deleted == 1 && factory.created.size() == 3,
              "stale pane deleted once");
        app.windows.open.push_back(&first);
        factory.failAttach = true;
        check(FAILED(action(addin.p, &one)), "Attach failure returned");
        check(factory.created.back()->deleted == 1 && factory.created.back()->refs == 1,
              "failed pane released");
        factory.failAttach = false;
        factory.duringAttach = [&]() { addin->OnDisconnection(0, nullptr); };
        check(FAILED(action(addin.p, &one)), "reentrant disconnect stops creation");
        check(factory.created.back()->deleted == 1, "pending pane deleted on disconnect");
        check(factory.created[2]->deleted == 1, "existing pane removed on disconnect");
        ok(addin->OnBeginShutdown(nullptr), "shutdown is idempotent");
    }
    check(app.refs == 1 && first.refs == 1 && second.refs == 1 && factory.refs == 1,
          "all held Office references released");
    for (const auto &pane : factory.created)
        check(pane->refs == 1 && pane->content.refs == 1, "all pane references released");
    check(DllCanUnloadNow() == S_OK, "module can unload after callbacks and timers");
}
} // namespace
int wmain(int argc, wchar_t **argv)
{
    const auto init = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    if (FAILED(init))
        return 1;
    HRESULT hr = S_OK;
    try
    {
        if (!(argc == 2 && !wcscmp(argv[1], L"--self-test")))
            hr = verify();
        if (SUCCEEDED(hr))
            behavior();
    }
    catch (const std::exception &e)
    {
        std::fprintf(stderr, "Ribbon behavior failed: %s\n", e.what());
        hr = E_FAIL;
    }
    CoUninitialize();
    std::printf("Office COM/ribbon/CTP ABI, window binding and lifecycle "
                "(%zu-bit): 0x%08lX\n",
                sizeof(void *) * 8, static_cast<unsigned long>(hr));
    return FAILED(hr) ? 1 : 0;
}
