using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Win32;

namespace TrioBridge
{
class Setup
{
    const string Id = "{1CE05ADE-43DF-432E-AFF5-2DF7D75ABA29}", Prog = "TrioBridge.Workbench",
                 Owner = "ai.triobridge.workbench";
    const string Clsid = @"Software\Classes\CLSID\" + Id, ProgKey = @"Software\Classes\" + Prog;
    const string PaneId = "{640D4723-A40B-46E5-A862-30C4254606D4}", PaneProg = "TrioBridge.TaskPane";
    const string PaneClsid = @"Software\Classes\CLSID\" + PaneId, PaneProgKey = @"Software\Classes\" + PaneProg;
    static readonly string[] Hosts = { "Word", "Excel", "PowerPoint" };
    static readonly string Parent = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "TrioBridgeAI");
    static readonly string Dest = Path.Combine(Parent, "Workbench");
    static string PaneAssembly(string root)
    {
        return Path.Combine(root, "pane", "TrioBridge.TaskPane.dll");
    }
    static string[] RegistrationKeys()
    {
        return new[] { Clsid, ProgKey, PaneClsid, PaneProgKey }.Concat(Hosts.Select(Addin)).ToArray();
    }
    static readonly RegistryView[] Views = Environment.Is64BitOperatingSystem
                                               ? new[] { RegistryView.Registry64, RegistryView.Registry32 }
                                               : new[] { RegistryView.Registry32 };
    static string Addin(string h)
    {
        return @"Software\Microsoft\Office\" + h + @"\Addins\" + Prog;
    }
    static bool Owned(string d)
    {
        return File.Exists(Path.Combine(d, "workbench.id")) &&
               File.ReadAllText(Path.Combine(d, "workbench.id")).Trim() == Owner;
    }
    static void CopyTree(string from, string to)
    {
        if ((File.GetAttributes(from) & FileAttributes.ReparsePoint) != 0)
            throw new IOException("安装目录包含不支持的链接。");
        Directory.CreateDirectory(to);
        foreach (string f in Directory.GetFiles(from))
        {
            if ((File.GetAttributes(f) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("安装文件包含链接。");
            File.Copy(f, Path.Combine(to, Path.GetFileName(f)));
        }
        foreach (string d in Directory.GetDirectories(from))
            CopyTree(d, Path.Combine(to, Path.GetFileName(d)));
    }
    static void VerifyKeys()
    {
        foreach (var v in Views)
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, v))
            {
                foreach (var path in RegistrationKeys())
                    using (
                        var k = root.OpenSubKey(
                            path)) if (k != null &&
                                       Convert.ToString(k.GetValue("TrioBridgeOwner")) !=
                                           Owner) throw new IOException("Office 注册位置已由其他组件占用。");
            }
    }
    static RegistryKey Key(RegistryKey root, string path)
    {
        var k = root.CreateSubKey(path);
        k.SetValue("TrioBridgeOwner", Owner);
        return k;
    }
    static void Register()
    {
        foreach (var v in Views)
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, v))
            {
                using (var cls = Key(root, Clsid))
                {
                    cls.SetValue("", "TrioBridge Office Workbench");
                    using (var p = cls.CreateSubKey("ProgId")) p.SetValue("", Prog);
                    using (var server = cls.CreateSubKey("InprocServer32"))
                    {
                        string binary = v == RegistryView.Registry64 ? "TrioBridge.Ribbon.x64.dll"
                                                                     : "TrioBridge.Ribbon.x86.dll";
                        server.SetValue("", Path.Combine(Dest, binary));
                        server.SetValue("ThreadingModel", "Apartment");
                    }
                }
                using (var p = Key(root, ProgKey)) using (var c = p.CreateSubKey("CLSID")) c.SetValue("", Id);
                // Rollback may restore a standalone-only version of the workbench.
                if (File.Exists(PaneAssembly(Dest))) RegisterPane(root);
                foreach (string h in Hosts)
                    using (var k = Key(root, Addin(h)))
                    {
                        k.SetValue("FriendlyName", "三桥 AI");
                        k.SetValue("Description", "在侧边栏或独立窗口中打开工作台");
                        k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    }
            }
    }
    static void RegisterPane(RegistryKey root)
    {
        string file = PaneAssembly(Dest);
        AssemblyName identity = AssemblyName.GetAssemblyName(file);
        string codeBase = new Uri(file).AbsoluteUri;
        using (var cls = Key(root, PaneClsid))
        {
            cls.SetValue("", "三桥 AI 工作台侧边栏");
            using (var prog = cls.CreateSubKey("ProgId")) prog.SetValue("", PaneProg);
            using (cls.CreateSubKey("Control")) { }
            using (cls.CreateSubKey(@"Implemented Categories\{40FC6ED4-2438-11CF-A3DB-080036F12502}")) { }
            using (var server = cls.CreateSubKey("InprocServer32"))
            {
                server.SetValue("", "mscoree.dll");
                server.SetValue("ThreadingModel", "Both");
                // Match RegAsm /codebase for each registry view. The CLR shim
                // resolves the managed class from its version-specific registration.
                using (var version = server.CreateSubKey(identity.Version.ToString()))
                    foreach (var entry in new[] { server, version })
                    {
                        entry.SetValue("Class", PaneProg);
                        entry.SetValue("Assembly", identity.FullName);
                        entry.SetValue("RuntimeVersion", "v4.0.30319");
                        entry.SetValue("CodeBase", codeBase);
                    }
            }
        }
        using (var p = Key(root, PaneProgKey))
        using (var c = p.CreateSubKey("CLSID")) c.SetValue("", PaneId);
    }
    static void RemoveKeys()
    {
        foreach (var v in Views)
            using (var root = RegistryKey.OpenBaseKey(
                       RegistryHive.CurrentUser,
                       v)) foreach (string path in RegistrationKeys())
            {
                bool owned;
                using (var key = root.OpenSubKey(path)) owned =
                    key != null && Convert.ToString(key.GetValue("TrioBridgeOwner")) == Owner;
                if (owned)
                    root.DeleteSubKeyTree(path, false);
            }
    }
    static void CheckClosed()
    {
        foreach (string name in new[] { "WINWORD", "EXCEL", "POWERPNT", "TrioBridge AI" })
            if (Process.GetProcessesByName(name).Any(p => p.SessionId ==
                                                          Process.GetCurrentProcess().SessionId))
                throw new IOException(
                    "请保存文档，完全退出 Word、Excel、PowerPoint 和三桥 AI，然后重新运行安装或卸载。");
    }
    static void Install()
    {
        string src = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Workbench");
        if (!Owned(src) || !File.Exists(Path.Combine(src, "TrioBridge AI.exe")) ||
            !File.Exists(PaneAssembly(src)) ||
            !File.Exists(Path.Combine(src, "resources", "app", "pane-worker.mjs")) ||
            !File.Exists(Path.Combine(src, "TrioBridge.Ribbon.x86.dll")) ||
            !File.Exists(Path.Combine(src, "TrioBridge.Ribbon.x64.dll")))
            throw new IOException("请先完整解压安装包，再打开 Install。");
        if (Directory.Exists(Dest) && !Owned(Dest))
            throw new IOException("目标文件夹中存在其他程序，请检查安装目录。");
        VerifyKeys();
        Directory.CreateDirectory(Parent);
        string stage = Path.Combine(Parent, ".setup-" + Guid.NewGuid().ToString("N")),
               old = stage + "-previous";
        bool backed = false, moved = false;
        try
        {
            CopyTree(src, stage);
            if (Directory.Exists(Dest))
            {
                Directory.Move(Dest, old);
                backed = true;
            }
            Directory.Move(stage, Dest);
            moved = true;
            Register();
        }
        catch
        {
            if (moved)
            {
                RemoveKeys();
                Directory.Delete(Dest, true);
            }
            if (backed)
            {
                Directory.Move(old, Dest);
                Register();
            }
            throw;
        }
        finally
        {
            if (Directory.Exists(stage))
                Directory.Delete(stage, true);
        }
        if (backed)
            Directory.Delete(old, true);
    }
    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            CheckClosed();
            if (args.SequenceEqual(new[] { "--uninstall" }))
            {
                RemoveKeys();
                if (Owned(Dest))
                    Directory.Delete(Dest, true);
                Console.WriteLine("已卸载。你的模型配置和凭据已保留。");
            }
            else if (args.Length == 0)
            {
                Install();
                Console.WriteLine(
                    "安装完成。打开 Word、Excel 或 PowerPoint，在“三桥 AI”功能区选择“打开侧边栏”或“打开独立窗口”。");
            }
            else
                return 2;
            return 0;
        }
        catch (Exception e)
        {
            Console.Error.WriteLine("操作未完成。" + e.Message);
            return 1;
        }
    }
}
}
