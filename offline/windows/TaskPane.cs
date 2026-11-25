using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

[assembly: ComVisible(false)]
[assembly: AssemblyVersion("1.0.0.0")]

namespace TrioBridge
{
[ComVisible(true)]
[Guid("B810F627-276D-4B11-82BC-E7A294068FA0")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface ITaskPane
{
    [DispId(1)]
    void Attach([MarshalAs(UnmanagedType.IDispatch)] object application,
                [MarshalAs(UnmanagedType.IDispatch)] object window, string hostName);
}

[ComVisible(true)]
[Guid("640D4723-A40B-46E5-A862-30C4254606D4")]
[ProgId("TrioBridge.TaskPane")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(ITaskPane))]
public sealed class TaskPane : UserControl, ITaskPane
{
    internal const string PageUrl = "https://triobridge.local/offline.html";
    static readonly string PaneDirectory = Path.GetDirectoryName(typeof(TaskPane).Assembly.Location);
    static readonly object LoaderLock = new object();
    static bool loaderConfigured;
    readonly int uiThread = Thread.CurrentThread.ManagedThreadId;
    readonly SemaphoreSlim outputLock = new SemaphoreSlim(1, 1);
    readonly HashSet<string> pending = new HashSet<string>();
    readonly Dictionary<string, string> requests = new Dictionary<string, string>();
    object officeApplication, officeWindow;
    string host, profileDirectory;
    Control browserControl;
    Process worker;
    StreamWriter workerInput;
    int stopped, generation;
    bool attached, initializing, loaded;

    static TaskPane()
    {
        // RegAsm /codebase loads this assembly outside Office's executable folder.
        // Resolve only the two SDK assemblies shipped alongside this control.
        AppDomain.CurrentDomain.AssemblyResolve += ResolveWebViewAssembly;
    }
    static Assembly ResolveWebViewAssembly(object sender, ResolveEventArgs args)
    {
        string name = new AssemblyName(args.Name).Name;
        if (name != "Microsoft.Web.WebView2.Core" && name != "Microsoft.Web.WebView2.WinForms")
            return null;
        string file = Path.Combine(PaneDirectory, name + ".dll");
        return File.Exists(file) ? Assembly.LoadFrom(file) : null;
    }
    public TaskPane()
    {
        BackColor = Color.White;
        Dock = DockStyle.Fill;
        SetStatus("正在打开三桥 AI…");
    }
    public void Attach(object application, object window, string hostName)
    {
        RequireUiThread();
        if (stopped != 0) throw new InvalidOperationException("侧边栏已关闭，请重新打开。");
        if (!ValidHost(hostName) || application == null || window == null)
            throw new InvalidOperationException("无法绑定 Office 窗口。");
        if (attached)
        {
            if (host != hostName || !Office.SameInstance(officeApplication, application) ||
                !Office.SameInstance(officeWindow, window))
                throw new InvalidOperationException("侧边栏已绑定另一个 Office 窗口。");
            return;
        }
        dynamic suppliedWindow = window;
        if (!Office.SameInstance(application, (object)suppliedWindow.Application))
            throw new InvalidOperationException("Office 应用与窗口不一致。");
        Office.CheckWindow(application, window);
        officeApplication = application;
        officeWindow = window;
        host = hostName;
        attached = true;
        // Creating the handle on the Office thread installs WinForms' synchronization
        // context before any async WebView initialization or pipe callback.
        if (!IsHandleCreated) CreateControl();
        StartWhenReady();
    }
    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        StartWhenReady();
    }
    void StartWhenReady()
    {
        if (!attached || initializing || stopped != 0 || !IsHandleCreated) return;
        initializing = true;
        int session = generation;
        BeginInvoke(new Action(() => {
            try { Initialize(session); }
            catch (Exception) { Fail("侧边栏组件缺失，请关闭 Office 后重新运行安装程序。"); }
        }));
    }
    void RequireUiThread()
    {
        if (Thread.CurrentThread.ManagedThreadId != uiThread ||
            Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
            throw new InvalidOperationException("请从 Office 窗口打开侧边栏。");
    }
    internal static bool ValidHost(string value)
    {
        return value == "word" || value == "excel" || value == "powerpoint";
    }
    internal static bool TrustedPage(string value)
    {
        if (value == null) return false;
        int hash = value.IndexOf('#');
        return (hash < 0 ? value : value.Substring(0, hash)) == PageUrl;
    }
    internal static bool LocalResource(string value)
    {
        Uri uri;
        return Uri.TryCreate(value, UriKind.Absolute, out uri) &&
               uri.Scheme == "https" && uri.Host == "triobridge.local" &&
               uri.IsDefaultPort && uri.UserInfo.Length == 0;
    }
    internal static bool LocalBlob(string value)
    {
        return value != null && value.StartsWith("blob:https://triobridge.local/", StringComparison.Ordinal);
    }
    [DllImport("kernel32.dll", SetLastError = true)]
    static extern bool IsWow64Process2(IntPtr process, out ushort processMachine, out ushort nativeMachine);
    static string ProcessArchitecture()
    {
        try
        {
            ushort processMachine, nativeMachine;
            if (IsWow64Process2(new IntPtr(-1), out processMachine, out nativeMachine))
            {
                ushort machine = processMachine == 0 ? nativeMachine : processMachine;
                if (machine == 0xaa64) return "arm64";
                if (machine == 0x8664) return "x64";
                if (machine == 0x014c) return "x86";
            }
        }
        catch (EntryPointNotFoundException) { }
        return IntPtr.Size == 8 ? "x64" : "x86";
    }
    [MethodImpl(MethodImplOptions.NoInlining)]
    async void Initialize(int session)
    {
        try
        {
            RequireUiThread();
            if (stopped != 0 || session != generation) return;
            string root = Directory.GetParent(PaneDirectory).FullName;
            string ui = Path.Combine(root, "resources", "app", "ui");
            if (!File.Exists(Path.Combine(ui, "offline.html")))
                throw new InvalidOperationException("侧边栏文件缺失，请重新解压安装包并运行安装程序。");
            lock (LoaderLock)
            {
                if (!loaderConfigured)
                {
                    string loader = Path.Combine(PaneDirectory, "runtimes", "win-" + ProcessArchitecture(), "native");
                    if (!File.Exists(Path.Combine(loader, "WebView2Loader.dll")))
                        throw new InvalidOperationException("侧边栏组件缺失，请重新解压安装包并运行安装程序。");
                    try { CoreWebView2Environment.SetLoaderDllFolderPath(loader); }
                    catch (InvalidOperationException)
                    {
                        // Another control in this Office process may already have
                        // initialized the same SDK's loader. Verify that it works.
                        CoreWebView2Environment.GetAvailableBrowserVersionString();
                    }
                    loaderConfigured = true;
                }
            }
            // A separate profile avoids cross-window permission/session sharing.
            profileDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                                            "TrioBridge AI", "PaneCache", Guid.NewGuid().ToString("N"));
            var environment = await CoreWebView2Environment.CreateAsync(null, profileDirectory);
            if (stopped != 0 || session != generation) return;
            var browser = new WebView2 { Dock = DockStyle.Fill, AllowExternalDrop = false };
            browserControl = browser;
            foreach (Control child in new List<Control>(ControlItems())) child.Dispose();
            Controls.Add(browser);
            await browser.EnsureCoreWebView2Async(environment);
            if (stopped != 0 || session != generation) return;
            var core = browser.CoreWebView2;
            core.Settings.AreDefaultContextMenusEnabled = false;
            core.Settings.AreDevToolsEnabled = false;
            core.Settings.IsStatusBarEnabled = false;
            core.Settings.AreBrowserAcceleratorKeysEnabled = false;
            core.Settings.IsPasswordAutosaveEnabled = false;
            core.Settings.IsGeneralAutofillEnabled = false;
            core.Settings.IsWebMessageEnabled = true;
            core.SetVirtualHostNameToFolderMapping("triobridge.local", ui, CoreWebView2HostResourceAccessKind.Deny);
            core.NavigationStarting += (sender, e) => {
                e.Cancel = !TrustedPage(e.Uri) || loaded;
            };
            core.NavigationCompleted += (sender, e) => {
                if (e.IsSuccess) loaded = true;
                else if (stopped == 0 && e.WebErrorStatus != CoreWebView2WebErrorStatus.OperationCanceled)
                    Fail("无法加载侧边栏，请关闭并重新打开。");
            };
            core.FrameNavigationStarting += (sender, e) => { e.Cancel = true; };
            core.NewWindowRequested += (sender, e) => { e.Handled = true; };
            core.DownloadStarting += Download;
            core.PermissionRequested += (sender, e) => {
                e.State = CoreWebView2PermissionState.Deny;
                e.SavesInProfile = false;
                e.Handled = true;
            };
            core.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All);
            core.WebResourceRequested += (sender, e) => {
                if (!LocalResource(e.Request.Uri) && !LocalBlob(e.Request.Uri) &&
                    !(e.Request.Uri.StartsWith("data:", StringComparison.Ordinal) &&
                      e.ResourceContext == CoreWebView2WebResourceContext.Image))
                    e.Response = environment.CreateWebResourceResponse(null, 403, "Forbidden", "");
            };
            core.ProcessFailed += (sender, e) => {
                if (e.ProcessFailedKind == CoreWebView2ProcessFailedKind.BrowserProcessExited ||
                    e.ProcessFailedKind == CoreWebView2ProcessFailedKind.RenderProcessExited)
                    Fail("侧边栏浏览器已停止，请关闭并重新打开。");
            };
            core.WebMessageReceived += WebMessage;
            await core.AddScriptToExecuteOnDocumentCreatedAsync(BridgeScript(host));
            if (stopped != 0 || session != generation) return;
            StartWorker(root, session);
            core.Navigate(PageUrl);
        }
        catch (WebView2RuntimeNotFoundException)
        {
            if (session != generation) return;
            Fail("缺少 Microsoft Edge WebView2 Runtime。请关闭 Office，运行原 ZIP 根目录的 " +
                 "MicrosoftEdgeWebView2RuntimeInstallerX64.exe（ARM64 安装包使用 " +
                 "MicrosoftEdgeWebView2RuntimeInstallerARM64.exe），然后重新打开侧边栏。");
        }
        catch (Exception e)
        {
            if (session != generation) return;
            Fail(e is InvalidOperationException ? e.Message :
                 "无法启动侧边栏，请关闭 Office 后重新运行安装程序。");
        }
    }
    internal static string BridgeScript(string hostName)
    {
        if (!ValidHost(hostName)) throw new ArgumentException("hostName");
        return @"(() => {
  'use strict';
  if (window !== window.top || location.href.split('#')[0] !== 'https://triobridge.local/offline.html') return;
  const channel = window.chrome.webview;
  const prefix = crypto.randomUUID();
  const pending = new Map();
  let counter = 0, disconnected = '';
  function invoke(method, fields) {
    if (disconnected) return Promise.reject(new Error(disconnected));
    const id = prefix + ':' + (++counter);
    return new Promise((resolve, reject) => {
      pending.set(id, {resolve, reject});
      try { channel.postMessage(Object.assign({id, method}, fields)); }
      catch (error) { pending.delete(id); reject(error); }
    });
  }
  channel.addEventListener('message', event => {
    const message = event.data;
    if (!message || typeof message !== 'object') return;
    if (message.type === 'disconnect') {
      disconnected = message.message || '侧边栏已断开，请关闭并重新打开。';
      for (const call of pending.values()) call.reject(new Error(disconnected));
      pending.clear();
      return;
    }
    const call = pending.get(message.id);
    if (!call) return;
    pending.delete(message.id);
    if (message.ok === true) call.resolve(message.value);
    else call.reject(new Error(message.message || '侧边栏未能完成操作。'));
  });
  Object.defineProperty(window, 'trioWorkbench', {value: Object.freeze({
    host: '" + hostName + @"',
    model: (requestId, action, payload) => invoke('model', {requestId, action, payload}),
    cancel: requestId => { if (!disconnected) channel.postMessage({method: 'cancel', requestId}); },
    office: (action, data) => invoke('office', {action, data})
  }), writable: false, configurable: false});
})();";
    }
    static JavaScriptSerializer Serializer()
    {
        return new JavaScriptSerializer { MaxJsonLength = int.MaxValue, RecursionLimit = 80 };
    }
    static string Field(Dictionary<string, object> message, string key)
    {
        object value;
        return message.TryGetValue(key, out value) ? value as string : null;
    }
    internal static bool ValidRequest(Dictionary<string, object> message)
    {
        string method = Field(message, "method"), requestId = Field(message, "requestId");
        string[] keys = method == "cancel" ? new[] { "method", "requestId" }
            : method == "model" ? new[] { "id", "method", "requestId", "action", "payload" }
            : new[] { "id", "method", "action", "data" };
        foreach (string key in message.Keys)
            if (Array.IndexOf(keys, key) < 0) return false;
        Guid guid;
        if (method == "cancel") return Guid.TryParse(requestId, out guid);
        string id = Field(message, "id"), action = Field(message, "action");
        if (string.IsNullOrEmpty(id) || id.Length > 100) return false;
        if (method == "model")
            return Guid.TryParse(requestId, out guid) && Field(message, "payload") != null &&
                   (action == "state" || action == "complete" || action == "activate" ||
                    action == "delete" || action == "import" || action == "upsert");
        return method == "office" &&
               (action == "release" || action == "selection" || action == "target" ||
                action == "presentation" || action == "presentationTarget" || action == "write" ||
                action == "edit" || action == "create" || action == "append");
    }
    async void WebMessage(object sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        if (stopped != 0 || !TrustedPage(e.Source) || !TrustedPage(((WebView2)browserControl).CoreWebView2.Source))
            return;
        int session = generation;
        string id = null;
        try
        {
            var message = Serializer().DeserializeObject(e.WebMessageAsJson) as Dictionary<string, object>;
            if (message == null || !ValidRequest(message)) return;
            id = Field(message, "id");
            if (Field(message, "method") == "cancel")
            {
                // Only this pane's outstanding model calls can be cancelled.
                if (requests.ContainsValue(Field(message, "requestId"))) await SendWorker(message);
                return;
            }
            if (pending.Contains(id)) return;
            if (pending.Count >= 32)
                throw new InvalidOperationException("请求繁忙，请稍后重试。");
            pending.Add(id);
            if (Field(message, "method") == "model") requests[id] = Field(message, "requestId");
            await SendWorker(message);
        }
        catch (Exception e2)
        {
            if (id != null && stopped == 0 && session == generation)
            {
                pending.Remove(id);
                requests.Remove(id);
                Post(new { id = id, ok = false, message = e2 is InvalidOperationException ? e2.Message : "侧边栏已断开，请关闭并重新打开。" });
            }
        }
    }
    void StartWorker(string root, int session)
    {
        string executable = Path.Combine(root, "TrioBridge AI.exe");
        string script = Path.Combine(root, "resources", "app", "pane-worker.mjs");
        if (!File.Exists(executable) || !File.Exists(script))
            throw new InvalidOperationException("工作台文件缺失，请重新解压安装包并运行安装程序。");
        var start = new ProcessStartInfo(executable, "\"" + script + "\" " + host) {
            WorkingDirectory = root, UseShellExecute = false, CreateNoWindow = true,
            RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
            StandardOutputEncoding = new UTF8Encoding(false), StandardErrorEncoding = new UTF8Encoding(false)
        };
        start.EnvironmentVariables["ELECTRON_RUN_AS_NODE"] = "1";
        // Node options inherited from Office must not inject modules into the worker.
        start.EnvironmentVariables.Remove("NODE_OPTIONS");
        start.EnvironmentVariables.Remove("NODE_PATH");
        var process = new Process { StartInfo = start };
        if (!process.Start()) throw new InvalidOperationException("无法启动工作台后台进程。");
        worker = process;
        workerInput = new StreamWriter(process.StandardInput.BaseStream, new UTF8Encoding(false));
        Task.Run(() => ReadWorker(process, session));
        Task.Run(async () => {
            // Drain stderr without retaining model contents or blocking stdout.
            var buffer = new char[4096];
            try { while (await process.StandardError.ReadAsync(buffer, 0, buffer.Length) > 0) { } }
            catch (Exception) { }
        });
    }
    async Task SendWorker(object message)
    {
        string line = Serializer().Serialize(message);
        int session = generation;
        var writer = workerInput;
        await outputLock.WaitAsync().ConfigureAwait(false);
        try
        {
            if (stopped != 0 || session != generation || writer == null)
                throw new InvalidOperationException("侧边栏已断开，请关闭并重新打开。");
            await writer.WriteLineAsync(line).ConfigureAwait(false);
            await writer.FlushAsync().ConfigureAwait(false);
        }
        finally { outputLock.Release(); }
    }
    async Task ReadWorker(Process process, int session)
    {
        try
        {
            string line;
            while (stopped == 0 && session == generation && (line = await process.StandardOutput.ReadLineAsync().ConfigureAwait(false)) != null)
            {
                var message = Serializer().DeserializeObject(line) as Dictionary<string, object>;
                if (message == null) throw new InvalidDataException();
                OnUi(() => ReceiveWorker(message), session);
            }
        }
        catch (Exception) { }
        OnUi(() => Fail("工作台后台进程已断开，请关闭并重新打开侧边栏。"), session);
    }
    async void ReceiveWorker(Dictionary<string, object> message)
    {
        if (stopped != 0) return;
        int session = generation;
        string call = Field(message, "call");
        if (call != null)
        {
            object reply;
            try
            {
                RequireUiThread();
                string op = Field(message, "op");
                object value;
                if (op == "native")
                {
                    object request;
                    var q = message.TryGetValue("request", out request) ? request as Dictionary<string, object> : null;
                    if (q == null) throw new InvalidOperationException("无效的 Office 操作。");
                    value = Office.Dispatch(officeApplication, officeWindow, host, q);
                }
                else if (op == "saveDeck") value = SaveDeck();
                else throw new InvalidOperationException("未知的侧边栏操作。");
                reply = new { reply = call, ok = true, value = value };
            }
            catch (Exception e)
            {
                reply = new { reply = call, ok = false, message = e is InvalidOperationException ? e.Message :
                    "无法完成 Office 操作。请打开可编辑文档，检查当前选区，然后重试。" };
            }
            try { await SendWorker(reply); }
            catch (Exception) { if (session == generation) Fail("工作台后台进程已断开，请关闭并重新打开侧边栏。"); }
            return;
        }
        string id = Field(message, "id");
        if (id != null && pending.Remove(id))
        {
            requests.Remove(id);
            Post(message);
        }
    }
    [DllImport("user32.dll")]
    static extern IntPtr GetAncestor(IntPtr window, uint flags);
    sealed class DialogOwner : IWin32Window
    {
        readonly IntPtr handle;
        public DialogOwner(IntPtr value) { handle = value; }
        public IntPtr Handle { get { return handle; } }
    }
    IWin32Window OwnerWindow()
    {
        IntPtr owner = GetAncestor(Handle, 2);
        return new DialogOwner(owner == IntPtr.Zero ? Handle : owner);
    }
    void Download(object sender, CoreWebView2DownloadStartingEventArgs e)
    {
        e.Cancel = true;
        e.Handled = true;
        if (stopped != 0 || !LocalBlob(e.DownloadOperation.Uri) ||
            !TrustedPage(((WebView2)browserControl).CoreWebView2.Source)) return;
        int session = generation;
        var deferral = e.GetDeferral();
        // WebView2 forbids nested modal loops inside its own event callback.
        BeginInvoke(new Action(() => {
            try
            {
                if (stopped != 0 || session != generation) return;
                Office.CheckWindow(officeApplication, officeWindow);
                using (var dialog = new SaveFileDialog())
                {
                    dialog.Title = "保存文件";
                    dialog.FileName = Path.GetFileName(e.ResultFilePath);
                    dialog.OverwritePrompt = true;
                    dialog.RestoreDirectory = true;
                    if (dialog.ShowDialog(OwnerWindow()) == DialogResult.OK)
                    {
                        e.ResultFilePath = dialog.FileName;
                        e.Cancel = false;
                    }
                }
            }
            catch (Exception) { e.Cancel = true; }
            finally { deferral.Complete(); }
        }));
    }
    string SaveDeck()
    {
        Office.CheckWindow(officeApplication, officeWindow);
        using (var dialog = new SaveFileDialog())
        {
            dialog.Title = "保存 PowerPoint 文件";
            dialog.FileName = "三桥演示稿.pptx";
            dialog.DefaultExt = "pptx";
            dialog.Filter = "PowerPoint 演示文稿 (*.pptx)|*.pptx";
            dialog.AddExtension = true;
            dialog.OverwritePrompt = true;
            dialog.RestoreDirectory = true;
            return dialog.ShowDialog(OwnerWindow()) == DialogResult.OK
                ? dialog.FileName : null;
        }
    }
    void OnUi(Action action, int session)
    {
        if (stopped != 0 || session != generation || !IsHandleCreated) return;
        try { BeginInvoke(new Action(() => { if (stopped == 0 && session == generation) action(); })); }
        catch (InvalidOperationException) { }
    }
    void Post(object message)
    {
        if (browserControl != null) PostToBrowser(message);
    }
    [MethodImpl(MethodImplOptions.NoInlining)]
    void PostToBrowser(object message)
    {
        var browser = browserControl as WebView2;
        if (browser != null && browser.CoreWebView2 != null && TrustedPage(browser.CoreWebView2.Source))
        {
            try { browser.CoreWebView2.PostWebMessageAsJson(Serializer().Serialize(message)); }
            catch (InvalidOperationException) { }
            catch (COMException) { }
        }
    }
    void SetStatus(string message)
    {
        var label = new Label { Dock = DockStyle.Fill, AutoSize = false, Padding = new Padding(20),
                                Text = message, ForeColor = Color.FromArgb(45, 55, 72) };
        Controls.Add(label);
        label.BringToFront();
    }
    void Fail(string message)
    {
        if (stopped != 0) return;
        Post(new { type = "disconnect", message = message });
        Stop(false);
        SetStatus(message);
        var retry = new Button { Text = "重新打开侧边栏", Dock = DockStyle.Bottom, Height = 38 };
        retry.Click += (sender, e) => {
            if (officeApplication == null || officeWindow == null) return;
            stopped = 0;
            initializing = false;
            loaded = false;
            profileDirectory = null;
            foreach (Control child in new List<Control>(ControlItems())) child.Dispose();
            SetStatus("正在打开三桥 AI…");
            StartWhenReady();
        };
        Controls.Add(retry);
        retry.BringToFront();
    }
    IEnumerable<Control> ControlItems()
    {
        foreach (Control control in Controls) yield return control;
    }
    void Stop(bool releaseBinding = true)
    {
        if (releaseBinding) { officeApplication = null; officeWindow = null; }
        if (Interlocked.Exchange(ref stopped, 1) != 0) return;
        generation++;
        pending.Clear();
        requests.Clear();
        var process = worker;
        var input = workerInput;
        worker = null;
        workerInput = null;
        if (browserControl != null)
        {
            browserControl.Dispose();
            browserControl = null;
        }
        // These RCWs may be shared with other Office components. Never forcibly
        // release them; discard only our own references on the owning UI thread.
        string cache = profileDirectory;
        Task.Run(async () => {
            bool acquired = await outputLock.WaitAsync(250).ConfigureAwait(false);
            if (acquired)
            {
                try { if (input != null) input.Dispose(); }
                catch (Exception) { }
                finally { outputLock.Release(); }
            }
            else if (process != null)
            {
                // A worker that stopped reading cannot hold teardown behind a full pipe.
                try { if (input != null) input.BaseStream.Dispose(); } catch (Exception) { }
                try { process.Kill(); } catch (Exception) { }
            }
            if (process != null)
            {
                try { if (!process.WaitForExit(1500)) process.Kill(); }
                catch (Exception) { }
                finally { process.Dispose(); }
            }
            // WebView's renderer exits asynchronously after its control is disposed.
            for (int attempt = 0; cache != null && attempt < 5; attempt++)
            {
                try { if (Directory.Exists(cache)) Directory.Delete(cache, true); break; }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
                await Task.Delay(500).ConfigureAwait(false);
            }
        });
    }
    protected override void OnHandleDestroyed(EventArgs e)
    {
        if (!RecreatingHandle) Stop();
        base.OnHandleDestroyed(e);
    }
    protected override void Dispose(bool disposing)
    {
        if (disposing) Stop();
        base.Dispose(disposing);
    }
}
}
