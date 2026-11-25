using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;
using Microsoft.Web.WebView2.WinForms;

namespace TrioBridge
{
// Uses the registered ActiveX control without automating or writing an Office file.
public sealed class PaneSmokeApplication
{
    public PaneSmokeWindow ActiveWindow { get; set; }
}
public sealed class PaneSmokeWindow
{
    public PaneSmokeApplication Application { get; set; }
}
internal static class TaskPaneTests
{
    static void Check(bool condition, string message)
    {
        if (!condition) throw new Exception(message);
    }
    static Dictionary<string, object> Message(string json)
    {
        return (Dictionary<string, object>)new JavaScriptSerializer().DeserializeObject(json);
    }
    static void Policies()
    {
        Check(TaskPane.TrustedPage(TaskPane.PageUrl) && TaskPane.TrustedPage(TaskPane.PageUrl + "#local"), "Trusted document rejected.");
        foreach (string url in new[] { "file:///offline.html", "https://triobridge.local/other.html",
            "https://triobridge.local/offline.html?remote=1", "https://triobridge.local.evil/offline.html",
            "https://user@triobridge.local/offline.html", "http://triobridge.local/offline.html", null })
            Check(!TaskPane.TrustedPage(url), "Untrusted document accepted: " + url);
        Check(TaskPane.LocalResource("https://triobridge.local/assets/main.js"), "Packaged resource rejected.");
        foreach (string url in new[] { "http://127.0.0.1/private", "https://example.com/model", "file:///C:/secret",
            "https://triobridge.local:8443/assets/main.js", "https://user@triobridge.local/assets/main.js" })
            Check(!TaskPane.LocalResource(url), "External resource accepted: " + url);
        Check(TaskPane.LocalBlob("blob:https://triobridge.local/123"), "Local export/PDF blob rejected.");
        Check(!TaskPane.LocalBlob("blob:https://triobridge.local.evil/123") &&
              !TaskPane.LocalBlob("blob:https://example.com/123"), "Foreign blob accepted.");
        string model = "{\"id\":\"pane:1\",\"method\":\"model\",\"requestId\":\"11111111-1111-4111-8111-111111111111\",\"action\":\"state\",\"payload\":\"{}\"}";
        Check(TaskPane.ValidRequest(Message(model)), "Valid model request rejected.");
        var injected = Message(model); injected["reply"] = "reverse-call";
        Check(!TaskPane.ValidRequest(injected), "A web request must not impersonate a worker reply.");
        injected = Message(model); injected["call"] = "native";
        Check(!TaskPane.ValidRequest(injected), "A web request must not impersonate a native callback.");
        injected = Message(model); injected["action"] = "readFile";
        Check(!TaskPane.ValidRequest(injected), "Unknown model capability accepted.");
        Check(TaskPane.ValidRequest(Message("{\"id\":\"pane:2\",\"method\":\"office\",\"action\":\"write\",\"data\":{}}")), "Office write rejected.");
        Check(!TaskPane.ValidRequest(Message("{\"id\":\"pane:2\",\"method\":\"native\",\"action\":\"write\"}")), "Direct native web access accepted.");
        Check(TaskPane.ValidRequest(Message("{\"method\":\"cancel\",\"requestId\":\"11111111-1111-4111-8111-111111111111\"}")), "Cancellation rejected.");
        Console.WriteLine("Windows task pane policies passed.");
    }
    static void ComInterfaces(object pane)
    {
        IntPtr unknown = Marshal.GetIUnknownForObject(pane);
        try
        {
            foreach (string id in new[] { "00020400-0000-0000-C000-000000000046", "00000112-0000-0000-C000-000000000046" })
            {
                Guid iid = new Guid(id);
                IntPtr result;
                int hr = Marshal.QueryInterface(unknown, ref iid, out result);
                Check(hr == 0 && result != IntPtr.Zero, "The ActiveX control does not expose " + id);
                Marshal.Release(result);
            }
        }
        finally { Marshal.Release(unknown); }
    }
    static void RegisteredSmoke(bool webView)
    {
        // Use the same merged registry view as the actual x86/x64 COM client.
        const string paneServer = @"CLSID\{640D4723-A40B-46E5-A862-30C4254606D4}\InprocServer32";
        using (var server = Registry.ClassesRoot.OpenSubKey(paneServer))
        {
            Check(server != null && string.Equals(server.GetValue("") as string, "mscoree.dll",
                  StringComparison.OrdinalIgnoreCase), "Missing managed class registration in the client registry view.");
            var codeBase = new Uri((string)server.GetValue("CodeBase"), UriKind.Absolute);
            Check(codeBase.IsFile && File.Exists(codeBase.LocalPath), "Registered task pane DLL is missing.");
            var identity = AssemblyName.GetAssemblyName(codeBase.LocalPath);
            using (var version = server.OpenSubKey(identity.Version.ToString()))
                foreach (var entry in new[] { server, version })
                {
                    Check(entry != null && (string)entry.GetValue("Assembly") == identity.FullName &&
                          (string)entry.GetValue("Class") == "TrioBridge.TaskPane" &&
                          (string)entry.GetValue("RuntimeVersion") == "v4.0.30319" &&
                          (string)entry.GetValue("CodeBase") == codeBase.AbsoluteUri,
                          "Managed class registration does not match the installed assembly.");
                }
            foreach (string name in new[] { "Microsoft.Web.WebView2.Core.dll", "Microsoft.Web.WebView2.WinForms.dll" })
                Check(File.Exists(Path.Combine(Path.GetDirectoryName(codeBase.LocalPath), name)), "WebView2 dependency missing: " + name);
        }
        Console.WriteLine("Verified the installed managed class, both registration records and WebView2 dependencies.");
        // Isolate backend configuration and credentials from the current account.
        string directory = Path.Combine(Path.GetTempPath(), "triobridge-pane-smoke-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string previousLocalAppData = Environment.GetEnvironmentVariable("LOCALAPPDATA");
        Environment.SetEnvironmentVariable("LOCALAPPDATA", directory);
        object instance = Activator.CreateInstance(Type.GetTypeFromProgID("TrioBridge.TaskPane", true));
        var control = instance as Control;
        Check(control != null, "The registered pane is not a WinForms ActiveX control.");
        try
        {
            ComInterfaces(instance);
            Console.WriteLine("Registered Windows task pane exposes IDispatch and IOleObject.");
            if (!webView) return;
            using (var form = new Form { Width = 520, Height = 800, Text = "TrioBridge pane smoke test" })
            {
                form.Controls.Add(control);
                control.Dock = DockStyle.Fill;
                var application = new PaneSmokeApplication();
                var window = new PaneSmokeWindow { Application = application };
                application.ActiveWindow = window;
                Exception failure = null;
                form.Shown += async (sender, e) => {
                    try
                    {
                        instance.GetType().InvokeMember("Attach", BindingFlags.InvokeMethod, null, instance,
                                                        new object[] { application, window, "word" });
                        DateTime deadline = DateTime.UtcNow.AddSeconds(40);
                        bool ready = false, requested = false;
                        while (DateTime.UtcNow < deadline)
                        {
                            await Task.Delay(200);
                            WebView2 browser = null;
                            foreach (Control child in control.Controls) if (child is WebView2) browser = (WebView2)child;
                            if (browser == null || browser.CoreWebView2 == null) continue;
                            if (!requested)
                            {
                                string result = await browser.ExecuteScriptAsync("Boolean(window.trioWorkbench && document.querySelector('#root')?.childElementCount)");
                                if (result != "true") continue;
                                await browser.ExecuteScriptAsync("window.__paneSmoke = 'pending'; window.trioWorkbench.model(crypto.randomUUID(), 'state', '{}').then(r => window.__paneSmoke = r.status === 200 ? 'ok' : 'failed', () => window.__paneSmoke = 'failed');");
                                requested = true;
                            }
                            string reply = await browser.ExecuteScriptAsync("window.__paneSmoke");
                            if (reply == "\"ok\"") { ready = true; break; }
                            Check(reply != "\"failed\"", "The packaged worker failed its isolated model-state request.");
                        }
                        Check(ready, "The real WebView2 pane did not load the packaged workbench in 40 seconds.");
                        Console.WriteLine("Registered Windows task pane loaded real WebView2, workbench and isolated worker.");
                        var field = instance.GetType().GetField("worker", BindingFlags.Instance | BindingFlags.NonPublic);
                        var worker = (Process)field.GetValue(instance);
                        int workerId = worker.Id;
                        // Closing the ActiveX control must end its worker, even without
                        // an explicit Office disconnect notification from the test host.
                        control.Dispose();
                        bool exited = false;
                        for (int attempt = 0; attempt < 30; attempt++)
                        {
                            await Task.Delay(100);
                            try { using (var process = Process.GetProcessById(workerId)) exited = process.HasExited; }
                            catch (ArgumentException) { exited = true; }
                            if (exited) break;
                        }
                        Check(exited, "Disposing the pane leaked its background worker.");
                        Console.WriteLine("Registered Windows task pane disposed its worker.");
                    }
                    catch (Exception error) { failure = error; }
                    finally { form.Close(); }
                };
                Application.Run(form);
                if (failure != null) throw failure;
            }
        }
        finally
        {
            if (control != null) control.Dispose();
            Environment.SetEnvironmentVariable("LOCALAPPDATA", previousLocalAppData);
            try { Directory.Delete(directory, true); } catch (IOException) { }
        }
    }
    [STAThread]
    public static int Main(string[] args)
    {
        try
        {
            if (args.Length > 0 && args[0] == "--registered") RegisteredSmoke(args.Length > 1 && args[1] == "--webview");
            else if (args.Length > 0 && args[0] == "--script") Console.Write(TaskPane.BridgeScript("word"));
            else Policies();
            return 0;
        }
        catch (Exception e) { Console.Error.WriteLine(e); return 1; }
    }
}
}
