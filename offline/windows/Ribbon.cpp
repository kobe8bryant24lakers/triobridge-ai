#include "Ribbon.hpp"
#include <algorithm>
#include <atomic>
#include <cwchar>
#include <new>
#include <string>
#include <utility>
#include <vector>

namespace
{
HMODULE module;
std::atomic<LONG> liveObjects{0};
constexpr wchar_t ribbonXml[] =
    LR"(<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="TrioBridgeTab" label="三桥 AI"><group id="TrioBridgeWork" label="文档工具"><button id="TrioBridgePane" label="打开侧边栏" size="large" imageMso="ReviewNewComment" onAction="OpenTaskPane"/><button id="TrioBridgeOpen" label="打开独立窗口" size="large" imageMso="ReviewNewComment" onAction="OpenWorkbench"/></group></tab></tabs></ribbon></customUI>)";
constexpr DISPID openButton = 41;
constexpr DISPID paneButton = 43;

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
#ifdef TRIO_RIBBON_TEST
    return L"word";
#else
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
#endif
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
    if (!CreateProcessW(exe.c_str(), args.data(), nullptr, nullptr, FALSE, 0, nullptr,
                        folder.c_str(), &startup, &process))
        return HRESULT_FROM_WIN32(GetLastError());
    CloseHandle(process.hThread);
    CloseHandle(process.hProcess);
    return S_OK;
}
// Small owning wrappers keep all COM references balanced on failure paths.
template <class T> class ComPtr
{
    T *value = nullptr;

  public:
    ComPtr() = default;
    explicit ComPtr(T *p) : value(p)
    {
        if (value)
            value->AddRef();
    }
    ComPtr(const ComPtr &p) : ComPtr(p.value)
    {
    }
    ComPtr(ComPtr &&p) noexcept : value(std::exchange(p.value, nullptr))
    {
    }
    ~ComPtr()
    {
        if (value)
            value->Release();
    }
    ComPtr &operator=(ComPtr p) noexcept
    {
        std::swap(value, p.value);
        return *this;
    }
    T *get() const
    {
        return value;
    }
    T *operator->() const
    {
        return value;
    }
    explicit operator bool() const
    {
        return value != nullptr;
    }
    T **put()
    {
        if (value)
            value->Release();
        value = nullptr;
        return &value;
    }
};
struct Value
{
    VARIANT v;
    Value()
    {
        VariantInit(&v);
    }
    ~Value()
    {
        VariantClear(&v);
    }
    Value(const Value &) = delete;
    Value &operator=(const Value &) = delete;
};
HRESULT invoke(IDispatch *object, const wchar_t *name, WORD flags, VARIANT *args = nullptr,
               UINT count = 0, VARIANT *result = nullptr)
{
    if (!object)
        return E_POINTER;
    LPOLESTR member = const_cast<LPOLESTR>(name);
    DISPID id;
    HRESULT hr = object->GetIDsOfNames(IID_NULL, &member, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr))
        return hr;
    DISPID put = DISPID_PROPERTYPUT;
    DISPPARAMS params{args, (flags & DISPATCH_PROPERTYPUT) ? &put : nullptr, count,
                      (flags & DISPATCH_PROPERTYPUT) ? 1U : 0U};
    EXCEPINFO error{};
    hr = object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, flags, &params, result, &error, nullptr);
    if (error.pfnDeferredFillIn)
        error.pfnDeferredFillIn(&error);
    SysFreeString(error.bstrSource);
    SysFreeString(error.bstrDescription);
    SysFreeString(error.bstrHelpFile);
    return hr;
}
HRESULT asDispatch(const VARIANT &value, ComPtr<IDispatch> &result)
{
    Value copy;
    HRESULT hr = VariantCopyInd(&copy.v, const_cast<VARIANT *>(&value));
    if (FAILED(hr))
        return hr;
    IUnknown *object = copy.v.vt == VT_DISPATCH  ? copy.v.pdispVal
                       : copy.v.vt == VT_UNKNOWN ? copy.v.punkVal
                                                 : nullptr;
    return object ? object->QueryInterface(IID_IDispatch, reinterpret_cast<void **>(result.put()))
                  : DISP_E_TYPEMISMATCH;
}
HRESULT property(IDispatch *object, const wchar_t *name, ComPtr<IDispatch> &result)
{
    Value value;
    HRESULT hr = invoke(object, name, DISPATCH_PROPERTYGET, nullptr, 0, &value.v);
    return FAILED(hr) ? hr : asDispatch(value.v, result);
}
ComPtr<IUnknown> identity(IUnknown *object)
{
    ComPtr<IUnknown> result;
    if (object)
        object->QueryInterface(IID_IUnknown, reinterpret_cast<void **>(result.put()));
    return result;
}
HRESULT putLong(IDispatch *object, const wchar_t *name, LONG number)
{
    VARIANT value{};
    value.vt = VT_I4;
    value.lVal = number;
    return invoke(object, name, DISPATCH_PROPERTYPUT, &value, 1);
}
HRESULT showPane(IDispatch *pane)
{
    VARIANT value{};
    value.vt = VT_BOOL;
    value.boolVal = VARIANT_TRUE;
    return invoke(pane, L"Visible", DISPATCH_PROPERTYPUT, &value, 1);
}
class Ribbon;
// Separate dispatch implementations avoid DISPID 1 collisions between Office's
// dual interfaces. Every interface still shares the same canonical IUnknown.
template <class I> class DispatchPart : public I
{
  protected:
    Ribbon *owner;

  public:
    explicit DispatchPart(Ribbon *p) : owner(p)
    {
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID, void **) override;
    ULONG STDMETHODCALLTYPE AddRef() override;
    ULONG STDMETHODCALLTYPE Release() override;
    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *n) override
    {
        if (!n)
            return E_POINTER;
        *n = 0;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo **) override
    {
        return DISP_E_BADINDEX;
    }
};
class RibbonPart final : public DispatchPart<OfficeRibbon>
{
  public:
    using DispatchPart::DispatchPart;
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR *, UINT, LCID, DISPID *) override;
    HRESULT STDMETHODCALLTYPE Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS *, VARIANT *,
                                     EXCEPINFO *, UINT *) override;
    HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR, BSTR *) override;
};
class ConsumerPart final : public DispatchPart<OfficePaneConsumer>
{
  public:
    using DispatchPart::DispatchPart;
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR *, UINT, LCID, DISPID *) override;
    HRESULT STDMETHODCALLTYPE Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS *, VARIANT *,
                                     EXCEPINFO *, UINT *) override;
    HRESULT STDMETHODCALLTYPE CTPFactoryAvailable(OfficePaneFactory *) override;
};
class Ribbon final : public OfficeAddin
{
    std::atomic<ULONG> refs{1};
    RibbonPart ribbon{this};
    ConsumerPart consumer{this};
    ComPtr<IDispatch> application;
    ComPtr<OfficePaneFactory> factory;
    struct Pane
    {
        ComPtr<IUnknown> windowId;
        ComPtr<IDispatch> window;
        ComPtr<OfficeTaskPane> control;
    };
    std::vector<Pane> panes;
    bool working = false, disconnected = false;
    UINT_PTR timer = 0;
    inline static thread_local std::vector<std::pair<UINT_PTR, Ribbon *>> timers;

    void stopTimer()
    {
        if (!timer)
            return;
        KillTimer(nullptr, timer);
        timers.erase(std::remove_if(timers.begin(), timers.end(),
                                    [this](const auto &item) { return item.first == timer; }),
                     timers.end());
        timer = 0;
    }
    static void CALLBACK tick(HWND, UINT, UINT_PTR id, DWORD)
    {
        auto found = std::find_if(timers.begin(), timers.end(),
                                  [id](const auto &item) { return item.first == id; });
        if (found == timers.end())
            return;
        Ribbon *self = found->second;
        self->AddRef();
        if (!self->working && !self->disconnected)
        {
            self->working = true;
            try
            {
                self->prune();
            }
            catch (...)
            { /* Office may be shutting down. */
            }
            self->working = false;
            if (self->disconnected)
                self->clear();
        }
        self->Release();
    }
    HRESULT startTimer()
    {
        if (timer)
            return S_OK;
        timer = SetTimer(nullptr, 0, 1500, tick);
        if (!timer)
            return GetLastError() ? HRESULT_FROM_WIN32(GetLastError()) : E_FAIL;
        try
        {
            timers.emplace_back(timer, this);
        }
        catch (...)
        {
            KillTimer(nullptr, timer);
            timer = 0;
            throw;
        }
        return S_OK;
    }
    void clear()
    {
        const bool previousWorking = working;
        working = true;
        stopTimer();
        auto old = std::move(panes);
        panes.clear();
        factory = {};
        application = {};
        for (auto &pane : old)
            invoke(pane.control.get(), L"Delete", DISPATCH_METHOD);
        working = previousWorking;
    }
    // A closed window is removed from Application.Windows even if a held COM
    // reference keeps its wrapper alive. Do not delete panes on cancellable
    // BeforeClose events or when Office temporarily rejects an automation call.
    HRESULT prune()
    {
        if (panes.empty() || !application)
            return S_OK;
        ComPtr<IDispatch> windows;
        HRESULT hr = property(application.get(), L"Windows", windows);
        if (FAILED(hr))
            return hr;
        Value count;
        hr = invoke(windows.get(), L"Count", DISPATCH_PROPERTYGET, nullptr, 0, &count.v);
        if (FAILED(hr))
            return hr;
        hr = VariantChangeType(&count.v, &count.v, 0, VT_I4);
        if (FAILED(hr) || count.v.lVal < 0 || count.v.lVal > 2048)
            return E_FAIL;
        std::vector<ComPtr<IUnknown>> open;
        for (LONG n = 1; n <= count.v.lVal; ++n)
        {
            VARIANT index{};
            index.vt = VT_I4;
            index.lVal = n;
            Value value;
            hr = invoke(windows.get(), L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET, &index, 1,
                        &value.v);
            if (FAILED(hr))
                return hr;
            ComPtr<IDispatch> window;
            hr = asDispatch(value.v, window);
            if (FAILED(hr))
                return hr;
            auto key = identity(window.get());
            if (!key)
                return E_NOINTERFACE;
            open.push_back(std::move(key));
        }
        if (disconnected)
            return RPC_E_DISCONNECTED;
        // Remove first; Delete can re-enter Office and the add-in.
        for (size_t i = panes.size(); i > 0; --i)
        {
            auto &pane = panes[i - 1];
            if (std::none_of(open.begin(), open.end(),
                             [&pane](const auto &key) { return key.get() == pane.windowId.get(); }))
            {
                auto removed = std::move(pane.control);
                panes.erase(panes.begin() + static_cast<ptrdiff_t>(i - 1));
                invoke(removed.get(), L"Delete", DISPATCH_METHOD);
                if (disconnected)
                    return RPC_E_DISCONNECTED;
            }
        }
        if (panes.empty())
            stopTimer();
        return S_OK;
    }
    HRESULT openPane(IDispatch *button)
    {
        if (!application || !factory || disconnected)
            return E_UNEXPECTED;
        const wchar_t *host = currentHost();
        if (!host)
            return E_UNEXPECTED;
        ComPtr<IDispatch> window, windowApplication;
        HRESULT hr = property(button, L"Context", window);
        if (FAILED(hr))
            return hr; // Never fall back to ActiveWindow.
        hr = property(window.get(), L"Application", windowApplication);
        if (FAILED(hr))
            return hr;
        auto appKey = identity(application.get()), ownerKey = identity(windowApplication.get());
        auto key = identity(window.get());
        if (!appKey || !ownerKey || !key || appKey.get() != ownerKey.get())
            return E_ACCESSDENIED;
        hr = prune();
        if (FAILED(hr) || disconnected)
            return FAILED(hr) ? hr : RPC_E_DISCONNECTED;
        for (size_t i = 0; i < panes.size(); ++i)
            if (panes[i].windowId.get() == key.get())
            {
                hr = showPane(panes[i].control.get());
                if (SUCCEEDED(hr))
                    return hr;
                auto removed = std::move(panes[i].control);
                panes.erase(panes.begin() + static_cast<ptrdiff_t>(i));
                invoke(removed.get(), L"Delete", DISPATCH_METHOD);
                break;
            }
        if (disconnected)
            return RPC_E_DISCONNECTED;
        if (panes.size() >= 512)
            return E_OUTOFMEMORY;
        Value prog, title, parent;
        prog.v.vt = title.v.vt = VT_BSTR;
        prog.v.bstrVal = SysAllocString(L"TrioBridge.TaskPane");
        title.v.bstrVal = SysAllocString(L"三桥 AI");
        if (!prog.v.bstrVal || !title.v.bstrVal)
            return E_OUTOFMEMORY;
        parent.v.vt = VT_DISPATCH;
        parent.v.pdispVal = window.get();
        window->AddRef();
        ComPtr<OfficeTaskPane> pane;
        hr = factory->CreateCTP(prog.v.bstrVal, title.v.bstrVal, parent.v, pane.put());
        if (FAILED(hr) || !pane)
        {
            if (pane)
                invoke(pane.get(), L"Delete", DISPATCH_METHOD);
            return FAILED(hr) ? hr : E_FAIL;
        }
        // Any failure after CreateCTP removes the incomplete, owned pane.
        struct Pending
        {
            OfficeTaskPane *pane;
            ~Pending()
            {
                if (pane)
                    invoke(pane, L"Delete", DISPATCH_METHOD);
            }
        } pending{pane.get()};
        if (disconnected)
            return RPC_E_DISCONNECTED;
        ComPtr<IDispatch> content;
        hr = property(pane.get(), L"ContentControl", content);
        if (FAILED(hr))
            return hr;
        Value hostName;
        hostName.v.vt = VT_BSTR;
        hostName.v.bstrVal = SysAllocString(host);
        if (!hostName.v.bstrVal)
            return E_OUTOFMEMORY;
        VARIANT arguments[3]{}; // COM dispatch parameters are right-to-left.
        arguments[0] = hostName.v;
        arguments[1] = parent.v;
        arguments[2].vt = VT_DISPATCH;
        arguments[2].pdispVal = application.get();
        hr = invoke(content.get(), L"Attach", DISPATCH_METHOD, arguments, 3);
        if (FAILED(hr))
            return hr;
        if (disconnected)
            return RPC_E_DISCONNECTED;
        hr = putLong(pane.get(), L"DockPosition", 2); // msoCTPDockPositionRight
        if (SUCCEEDED(hr))
            hr = putLong(pane.get(), L"Width", 360);
        if (SUCCEEDED(hr))
            hr = showPane(pane.get());
        if (FAILED(hr) || disconnected)
            return FAILED(hr) ? hr : RPC_E_DISCONNECTED;
        hr = startTimer();
        if (FAILED(hr))
            return hr;
        try
        {
            panes.push_back({std::move(key), std::move(window), pane});
        }
        catch (...)
        {
            if (panes.empty())
                stopTimer();
            throw;
        }
        pending.pane = nullptr;
        return S_OK;
    }

  public:
    Ribbon()
    {
        ++liveObjects;
    }
    ~Ribbon()
    {
        disconnected = true;
        clear();
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
            *result = static_cast<OfficeRibbon *>(&ribbon);
        else if (id == __uuidof(OfficePaneConsumer))
            *result = static_cast<OfficePaneConsumer *>(&consumer);
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
        auto n = --refs;
        if (!n)
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
    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR *names, UINT count, LCID,
                                            DISPID *ids) override
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
        if (_wcsicmp(names[0], L"OpenTaskPane") == 0)
        {
            ids[0] = paneButton;
            return S_OK;
        }
        ids[0] = DISPID_UNKNOWN;
        return DISP_E_UNKNOWNNAME;
    }
    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD flags, DISPPARAMS *params,
                                     VARIANT *, EXCEPINFO *, UINT *) override
    {
        if (!(flags & DISPATCH_METHOD))
            return DISP_E_MEMBERNOTFOUND;
        try
        {
            if (id == openButton)
                return launchWorkbench();
            if (id == paneButton)
            {
                if (!params || params->cArgs != 1)
                    return DISP_E_BADPARAMCOUNT;
                if (working)
                    return RPC_E_CALL_REJECTED;
                ComPtr<IDispatch> button;
                HRESULT hr = asDispatch(params->rgvarg[0], button);
                if (FAILED(hr))
                    return hr;
                ComPtr<OfficeAddin> hold(this);
                struct Operation
                {
                    Ribbon *owner;
                    explicit Operation(Ribbon *p) : owner(p)
                    {
                        owner->working = true;
                    }
                    ~Operation()
                    {
                        owner->working = false;
                        if (owner->disconnected)
                            owner->clear();
                    }
                } operation(this);
                return openPane(button.get());
            }
            if (id == 1)
            {
                if (!params || params->cArgs != 4)
                    return DISP_E_BADPARAMCOUNT;
                ComPtr<IDispatch> app;
                HRESULT hr = asDispatch(params->rgvarg[3], app);
                return FAILED(hr) ? hr : OnConnection(app.get(), 0, nullptr, nullptr);
            }
            if (id == 2)
                return OnDisconnection(0, nullptr);
            if (id == 3)
                return OnAddInsUpdate(nullptr);
            if (id == 4)
                return OnStartupComplete(nullptr);
            if (id == 5)
                return OnBeginShutdown(nullptr);
            return DISP_E_MEMBERNOTFOUND;
        }
        catch (const std::bad_alloc &)
        {
            return E_OUTOFMEMORY;
        }
        catch (...)
        {
            return E_FAIL;
        }
    }
    HRESULT STDMETHODCALLTYPE OnConnection(IDispatch *app, LONG, IDispatch *, SAFEARRAY **) override
    {
        if (!app)
            return E_POINTER;
        if (working)
            return RPC_E_CALL_REJECTED;
        // Office can supply its factory before OnConnection on initial load.
        auto earlyFactory = !application ? std::move(factory) : ComPtr<OfficePaneFactory>();
        disconnected = true;
        clear();
        application = ComPtr<IDispatch>(app);
        factory = std::move(earlyFactory);
        disconnected = false;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE OnDisconnection(LONG, SAFEARRAY **) override
    {
        disconnected = true;
        if (!working)
            clear();
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
        return OnDisconnection(0, nullptr);
    }
    HRESULT setFactory(OfficePaneFactory *value)
    {
        if (!value)
            return E_POINTER;
        if (working || !panes.empty())
            return RPC_E_CALL_REJECTED;
        factory = ComPtr<OfficePaneFactory>(value);
        return S_OK;
    }
};
template <class I> HRESULT DispatchPart<I>::QueryInterface(REFIID id, void **out)
{
    return owner->QueryInterface(id, out);
}
template <class I> ULONG DispatchPart<I>::AddRef()
{
    return owner->AddRef();
}
template <class I> ULONG DispatchPart<I>::Release()
{
    return owner->Release();
}
HRESULT singleName(LPOLESTR *names, UINT count, DISPID *ids, const wchar_t *name)
{
    if (!names || !ids)
        return E_POINTER;
    if (count != 1 || !names[0] || _wcsicmp(names[0], name))
        return DISP_E_UNKNOWNNAME;
    ids[0] = 1;
    return S_OK;
}
HRESULT RibbonPart::GetIDsOfNames(REFIID, LPOLESTR *n, UINT c, LCID, DISPID *ids)
{
    return singleName(n, c, ids, L"GetCustomUI");
}
HRESULT RibbonPart::GetCustomUI(BSTR, BSTR *xml)
{
    if (!xml)
        return E_POINTER;
    *xml = SysAllocString(ribbonXml);
    return *xml ? S_OK : E_OUTOFMEMORY;
}
HRESULT RibbonPart::Invoke(DISPID id, REFIID, LCID, WORD flags, DISPPARAMS *p, VARIANT *result,
                           EXCEPINFO *, UINT *)
{
    if (id != 1 || !(flags & DISPATCH_METHOD))
        return DISP_E_MEMBERNOTFOUND;
    if (!p || p->cArgs != 1)
        return DISP_E_BADPARAMCOUNT;
    if (!result)
        return E_POINTER;
    VariantInit(result);
    result->vt = VT_BSTR;
    return GetCustomUI(nullptr, &result->bstrVal);
}
HRESULT ConsumerPart::GetIDsOfNames(REFIID, LPOLESTR *n, UINT c, LCID, DISPID *ids)
{
    return singleName(n, c, ids, L"CTPFactoryAvailable");
}
HRESULT ConsumerPart::CTPFactoryAvailable(OfficePaneFactory *factory)
{
    return owner->setFactory(factory);
}
HRESULT ConsumerPart::Invoke(DISPID id, REFIID, LCID, WORD flags, DISPPARAMS *p, VARIANT *,
                             EXCEPINFO *, UINT *)
{
    if (id != 1 || !(flags & DISPATCH_METHOD))
        return DISP_E_MEMBERNOTFOUND;
    if (!p || p->cArgs != 1)
        return DISP_E_BADPARAMCOUNT;
    ComPtr<IDispatch> object;
    HRESULT hr = asDispatch(p->rgvarg[0], object);
    if (FAILED(hr))
        return hr;
    ComPtr<OfficePaneFactory> factory;
    hr = object->QueryInterface(__uuidof(OfficePaneFactory),
                                reinterpret_cast<void **>(factory.put()));
    return FAILED(hr) ? hr : CTPFactoryAvailable(factory.get());
}
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
