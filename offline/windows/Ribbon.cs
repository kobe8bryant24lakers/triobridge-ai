using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace TrioBridge {
  [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
  public interface IAddin {
    [DispId(1)] void OnConnection([MarshalAs(UnmanagedType.IDispatch)]object application,int mode,[MarshalAs(UnmanagedType.IDispatch)]object addin,[In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(2)] void OnDisconnection(int mode,[In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(3)] void OnAddInsUpdate([In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(4)] void OnStartupComplete([In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(5)] void OnBeginShutdown([In,MarshalAs(UnmanagedType.SafeArray,SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
  }
  [ComImport,Guid("000C0396-0000-0000-C000-000000000046"),InterfaceType(ComInterfaceType.InterfaceIsDual)]
  public interface IRibbon { [DispId(1)] [return:MarshalAs(UnmanagedType.BStr)] string GetCustomUI([MarshalAs(UnmanagedType.BStr)]string ribbonId); }
  [ComVisible(true),Guid("669B5B4E-7F51-47A8-A7D4-DAEC93C7015E"),InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
  public interface IButtons { [DispId(41)] void OpenWorkbench([MarshalAs(UnmanagedType.IDispatch)]object control); }
  [ComVisible(true),Guid("1CE05ADE-43DF-432E-AFF5-2DF7D75ABA29"),ProgId("TrioBridge.Workbench"),ClassInterface(ClassInterfaceType.None),ComDefaultInterface(typeof(IButtons))]
  public class Ribbon : IAddin, IRibbon, IButtons {
    string host;
    public void OnConnection(object instance,int mode,object addin,ref Array custom) {
      string name=Convert.ToString(instance.GetType().InvokeMember("Name",BindingFlags.GetProperty,null,instance,null));
      foreach(string id in new[]{"Word","Excel","PowerPoint"}) if(name.IndexOf(id,StringComparison.OrdinalIgnoreCase)>=0)host=id.ToLowerInvariant();
    }
    public string GetCustomUI(string id) {
      return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'><ribbon><tabs><tab id='TrioBridgeTab' label='三桥 AI'><group id='TrioBridgeWork' label='文档工具'><button id='TrioBridgeOpen' label='打开工作台' size='large' imageMso='ReviewNewComment' onAction='OpenWorkbench'/></group></tab></tabs></ribbon></customUI>";
    }
    public void OpenWorkbench(object control) {
      if(host==null)return;
      string dir=Path.GetDirectoryName(typeof(Ribbon).Assembly.Location);
      Process.Start(new ProcessStartInfo(Path.Combine(dir,"TrioBridge AI.exe"),"--office="+host){UseShellExecute=false,WorkingDirectory=dir});
    }
    public void OnDisconnection(int mode,ref Array custom){host=null;}
    public void OnAddInsUpdate(ref Array custom){}
    public void OnStartupComplete(ref Array custom){}
    public void OnBeginShutdown(ref Array custom){}
  }
}
