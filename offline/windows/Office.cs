using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Web.Script.Serialization;

namespace TrioBridge {
  internal class Office {
    static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength=9000000,RecursionLimit=80 };
    static object[] A(object value) { return ((IEnumerable)value).Cast<object>().ToArray(); }
    static Dictionary<string,object> D(object v){return (Dictionary<string,object>)v;}
    static string S(object v){return Convert.ToString(v,CultureInfo.InvariantCulture);}
    static int I(object v){return Convert.ToInt32(v,CultureInfo.InvariantCulture);}
    static object Get(Dictionary<string,object> d,string k){return d.ContainsKey(k)?d[k]:null;}
    static void Check(bool ok,string message){if(!ok)throw new InvalidOperationException(message);}
    static void Same(object a,object b){Check(Json.Serialize(a)==Json.Serialize(b),"文档、选区或内容已变化，请重新读取。");}
    static string Bound(string s){Check(s.Length<=48000,"选区过大，请缩小范围。");return s;}
    static object Guard(dynamic doc,dynamic window){return new {name=S(doc.FullName),window=S(window.HWND)};}
    static object Word(dynamic app,Dictionary<string,object> q,string op,string kind){
      dynamic doc=app.ActiveDocument, range=app.Selection.Range;
      string text=Bound(S(range.Text));object guard=new {doc=Guard(doc,app.ActiveWindow),start=(int)range.Start,end=(int)range.End,story=(int)range.StoryType,text};
      bool normal=range.Tables.Count==0 && range.Fields.Count==0 && range.InlineShapes.Count==0 && range.ContentControls.Count==0 && range.OMaths.Count==0 && range.Hyperlinks.Count==0 && range.Footnotes.Count==0 && range.Endnotes.Count==0 && !text.Contains("\a");
      if(op=="selection"){Check(text.Trim().Length>0,"请在 Word 中先选中文字。");return new {guard,text,label="Word · 当前选区",writable=normal,warning=normal?null:"选区包含文档对象，可保留原文并在选区后插入结果。"};}
      if(op=="target")return new {guard,text,label="Word · 在当前选区后插入",writable=true};
      Same(q["guard"],guard);
      Check(!doc.ReadOnly,"文档为只读，请打开可编辑副本。");
      if(kind=="replace"){Check(normal,"请只选择普通文字。");range.Text=S(q["value"]);return null;}
      range.Collapse(0);
      if(kind=="insert")range.InsertAfter(S(q["value"]));
      else if(kind=="document")range.InsertFile(S(q["file"]),Type.Missing,false,false,false);
      else throw new InvalidOperationException("未知 Word 操作。");
      return null;
    }
    static object ToMatrix(object value){
      var a=value as Array;
      if(a==null || a.Rank!=2)return value;
      var rows=new List<object[]>();
      for(int i=a.GetLowerBound(0);i<=a.GetUpperBound(0);i++){
        var row=new List<object>(); for(int j=a.GetLowerBound(1);j<=a.GetUpperBound(1);j++)row.Add(a.GetValue(i,j));rows.Add(row.ToArray());
      }return rows;
    }
    static object Excel(dynamic app,Dictionary<string,object> q,string op,string kind){
      dynamic book=app.ActiveWorkbook, sheet=app.ActiveSheet;
      object key=Guard(book,app);
      if(kind=="table"){
        var names=new List<string>();foreach(dynamic item in book.Sheets)names.Add(S(item.Name));
        var guard=new {doc=key,sheets=names};
        if(op=="target")return new {guard,text="新建工作表，保留已有数据",label="Excel · 新建数据表",writable=true};
        Same(q["guard"],guard);Check(!book.ReadOnly,"工作簿为只读。");
        var t=D(q["value"]);var rows=A(t["matrix"]);var formats=A(t["formats"]);int cols=formats.Length;
        Check(rows.Length>1 && rows.Length<=1001 && cols>0 && cols<=40,"数据表尺寸无效。");
        var tableValues=new object[rows.Length,cols];
        for(int i=0;i<rows.Length;i++){var row=A(rows[i]);Check(row.Length==cols,"数据表的列数不一致。");for(int j=0;j<cols;j++)tableValues[i,j]=row[j];}
        string seed=S(t["name"]);foreach(char c in ":\\/?*[]")seed=seed.Replace(c,' ');seed=seed.Trim().Trim('\'');if(seed.Length>24)seed=seed.Substring(0,24);if(seed.Length==0)seed="AI 数据";
        string name=seed;int suffix=2;while(names.Contains(name,StringComparer.OrdinalIgnoreCase))name=seed+" "+suffix++;
        dynamic added=book.Worksheets.Add(After:book.Sheets[book.Sheets.Count]);
        try{
          added.Name=name;
          dynamic at=added.Range["A1"].Resize[rows.Length,cols];at.Formula=tableValues;
          for(int j=0;j<cols;j++) added.Cells[2,j+1].Resize[rows.Length-1,1].NumberFormat=S(formats[j]);
          added.ListObjects.Add(1,at,Type.Missing,1);at.Columns.AutoFit();
        }catch{bool alerts=app.DisplayAlerts;try{app.DisplayAlerts=false;added.Delete();}finally{app.DisplayAlerts=alerts;}throw;}
        return null;
      }
      dynamic range=kind=="formula"?app.ActiveCell:app.Selection;
      Check(Convert.ToDouble(range.CountLarge)<=2000 && range.Areas.Count==1,"请选择不超过 2,000 个单元格的一块连续区域。");
      object values=ToMatrix(range.Value2),formula=ToMatrix(range.Formula);
      string address=S(range.Address);object anchor=new {doc=key,sheet=S(sheet.Name),address,values,formula};
      string text=Bound(Json.Serialize(values)+(Json.Serialize(values)==Json.Serialize(formula)?"":"\n公式\n"+Json.Serialize(formula)));
      if(op=="selection"||op=="target")return new {guard=anchor,text,label="Excel · "+sheet.Name+"!"+address,writable=kind=="formula"};
      Same(q["guard"],anchor);Check(kind=="formula" && range.Count==1 && !book.ReadOnly && !range.MergeCells && !range.HasArray,"请选择可编辑的普通单元格。");
      range.Formula=S(q["value"]);return null;
    }
    static int Rgb(string hex){return Convert.ToInt32(hex.Substring(1,2),16) | (Convert.ToInt32(hex.Substring(3,2),16)<<8) | (Convert.ToInt32(hex.Substring(5,2),16)<<16);}
    static string Hex(int n){return "#"+(n&255).ToString("X2")+((n>>8)&255).ToString("X2")+((n>>16)&255).ToString("X2");}
    static Dictionary<string,object> Snapshot(dynamic app,string scope){
      dynamic doc=app.ActivePresentation,win=app.ActiveWindow,sel=win.Selection;
      var ids=new List<string>();foreach(dynamic slide in doc.Slides)ids.Add(S(slide.SlideID));
      Check(ids.Count<=500,"演示稿超过 500 页。");var chosen=new List<string>();
      if(scope=="all")chosen.AddRange(ids);else{try{foreach(dynamic sl in sel.SlideRange)chosen.Add(S(sl.SlideID));}catch(COMException){}if(chosen.Count==0)chosen.Add(S(win.View.Slide.SlideID));}
      Check(chosen.Count<=60,"每次最多编辑 60 页。");var selected=new List<string>();
      if(scope=="selected" && chosen.Count==1 && (sel.Type==2||sel.Type==3)){try{foreach(dynamic sh in sel.ShapeRange)selected.Add(S(sh.Id));}catch(COMException){}}
      var slides=new List<object>();int skipped=0,count=0;
      foreach(string id in chosen){
        dynamic sl=doc.Slides.FindBySlideID(I(id));var shapes=new List<object>();string title=null,first=null;
        foreach(dynamic sh in sl.Shapes){
          Check(++count<=350,"选中的对象过多，请缩小范围。");
          var shape=new Dictionary<string,object>{{"id",S(sh.Id)},{"name",S(sh.Name)},{"type",S(sh.Type)},{"left",(double)sh.Left},{"top",(double)sh.Top},{"width",(double)sh.Width},{"height",(double)sh.Height}};
          if(sh.HasTextFrame==-1){
            dynamic r=sh.TextFrame.TextRange;string text=Bound(S(r.Text));bool isTitle=sh.Type==14 && (sh.PlaceholderFormat.Type==1||sh.PlaceholderFormat.Type==3);
            shape["role"]=isTitle?"title":"body";shape["text"]=text;
            dynamic f=r.Font;int align=I(r.ParagraphFormat.Alignment);
            shape["font"]=new {name=S(f.Name),size=(double)f.Size,color=Hex(I(f.Color.RGB)),bold=I(f.Bold)==-1};
            shape["alignment"]=align==1?"Left":align==2?"Center":align==3?"Right":null;
            if(first==null && text.Length>0)first=text;if(isTitle)title=text;
          }else skipped++;
          shapes.Add(shape);
        }
        slides.Add(new {id,index=ids.IndexOf(id),title=title??first??("第 "+sl.SlideIndex+" 页"),shapes});
      }
      return new Dictionary<string,object>{{"scope",scope},{"slides",slides},{"slideIds",ids},{"selectedShapeIds",selected},{"skipped",skipped},{"size",new {width=(double)doc.PageSetup.SlideWidth,height=(double)doc.PageSetup.SlideHeight}}};
    }
    static object PowerPoint(dynamic app,Dictionary<string,object> q,string op,string kind){
      if(op=="openDeck"){app.Presentations.Open(S(q["file"]));return null;}
      dynamic doc=app.ActivePresentation;object guard=Guard(doc,app.ActiveWindow);
      if(op=="presentation")return new {guard,snapshot=Snapshot(app,S(q["scope"]))};
      if(op=="selection"){
        Dictionary<string,object> snap=Snapshot(app,"selected");var canonical=D(Json.DeserializeObject(Json.Serialize(snap)));
        var texts=from sl in A(canonical["slides"]) from sh in A(D(sl)["shapes"]) where D(sh).ContainsKey("text") select S(D(sh)["text"]);
        string text=Bound(string.Join("\n",texts));Check(text.Trim().Length>0,"所选页面没有可读取的文字。");return new {guard,text,label="PowerPoint · 所选页面",writable=false};
      }
      if(op=="target"||op=="write"){
        var ids=new List<string>();foreach(dynamic s in doc.Slides)ids.Add(S(s.SlideID));var anchor=new {doc=guard,ids};
        if(op=="target")return new {guard=anchor,text="保留现有页面",label="PowerPoint · 在末尾添加页面",writable=true};
        Same(q["guard"],anchor);int before=doc.Slides.Count;
        try{foreach(object v in A(q["value"])){var spec=D(v);dynamic sl=doc.Slides.Add(doc.Slides.Count+1,2);sl.Shapes.Title.TextFrame.TextRange.Text=S(spec["title"]);sl.Shapes.Placeholders[2].TextFrame.TextRange.Text=string.Join("\r",A(spec["bullets"]).Select(S));}}
        catch{while(doc.Slides.Count>before)doc.Slides[doc.Slides.Count].Delete();throw;}return null;
      }
      Same(q["guard"],guard);var snapshot=D(q["snapshot"]);Same(snapshot,Snapshot(app,S(snapshot["scope"])));
      if(op=="append"){
        int before=doc.Slides.Count;try{int n=doc.Slides.InsertFromFile(S(q["file"]),before);Check(n==I(q["count"]),"导入页数与预览不一致。");}catch{while(doc.Slides.Count>before)doc.Slides[doc.Slides.Count].Delete();throw;}return null;
      }
      Check(op=="edit","未知 PowerPoint 操作。");var proposal=D(q["proposal"]);var changes=A(proposal["changes"]);var undo=new Stack<Action>();
      var previousOrder=A(snapshot["slideIds"]).Select(S).ToArray();app.StartNewUndoEntry();
      try{
        foreach(object obj in changes){
          var p=D(obj);dynamic sl=doc.Slides.FindBySlideID(I(p["slideId"]));dynamic shape=null;
          foreach(dynamic candidate in sl.Shapes)if(S(candidate.Id)==S(p["shapeId"])){shape=candidate;break;}
          Check(shape!=null,"对象已移动。");dynamic r=shape.HasTextFrame==-1?shape.TextFrame.TextRange:null;
          if(p.ContainsKey("text")){string old=S(r.Text);r.Text=S(p["text"]);undo.Push(()=>{r.Text=old;});}
          if(p.ContainsKey("fontName")){string old=S(r.Font.Name);r.Font.Name=S(p["fontName"]);undo.Push(()=>{r.Font.Name=old;});}
          if(p.ContainsKey("fontSize")){float old=r.Font.Size;r.Font.Size=Convert.ToSingle(p["fontSize"]);undo.Push(()=>{r.Font.Size=old;});}
          if(p.ContainsKey("fontColor")){int old=r.Font.Color.RGB;r.Font.Color.RGB=Rgb(S(p["fontColor"]));undo.Push(()=>{r.Font.Color.RGB=old;});}
          if(p.ContainsKey("bold")){int old=r.Font.Bold;r.Font.Bold=(bool)p["bold"]?-1:0;undo.Push(()=>{r.Font.Bold=old;});}
          if(p.ContainsKey("alignment")){int old=r.ParagraphFormat.Alignment;string v=S(p["alignment"]);r.ParagraphFormat.Alignment=v=="Left"?1:v=="Center"?2:3;undo.Push(()=>{r.ParagraphFormat.Alignment=old;});}
          if(p.ContainsKey("left")){float old=shape.Left;shape.Left=Convert.ToSingle(p["left"]);undo.Push(()=>{shape.Left=old;});}
          if(p.ContainsKey("top")){float old=shape.Top;shape.Top=Convert.ToSingle(p["top"]);undo.Push(()=>{shape.Top=old;});}
          if(p.ContainsKey("width")){float old=shape.Width;shape.Width=Convert.ToSingle(p["width"]);undo.Push(()=>{shape.Width=old;});}
          if(p.ContainsKey("height")){float old=shape.Height;shape.Height=Convert.ToSingle(p["height"]);undo.Push(()=>{shape.Height=old;});}
        }
        if(proposal.ContainsKey("slideOrder")){
          var scoped=A(proposal["slideOrder"]).Select(S).ToList();var ids=new HashSet<string>(scoped);int pos=0;
          var order=previousOrder.Select(id=>ids.Contains(id)?scoped[pos++]:id).ToArray();
          undo.Push(()=>{for(int i=0;i<previousOrder.Length;i++)doc.Slides.FindBySlideID(I(previousOrder[i])).MoveTo(i+1);});
          for(int i=0;i<order.Length;i++)doc.Slides.FindBySlideID(I(order[i])).MoveTo(i+1);
        }
      }catch{bool failed=false;while(undo.Count>0){try{undo.Pop()();}catch{failed=true;}}if(failed)throw new InvalidOperationException("请在 PowerPoint 中撤销部分改动并检查目标页面。");throw;}return null;
    }
    [STAThread] static void Main(){
      Console.InputEncoding=new UTF8Encoding(false);Console.OutputEncoding=new UTF8Encoding(false);
      try{
        var q=D(Json.DeserializeObject(Console.In.ReadToEnd()));string h=S(q["host"]),op=S(q["op"]),kind=S(Get(q,"kind"));
        string id=h=="word"?"Word.Application":h=="excel"?"Excel.Application":h=="powerpoint"?"PowerPoint.Application":null;Check(id!=null,"未知 Office 应用。");
        dynamic app=Marshal.GetActiveObject(id);
        object result=h=="word"?Word(app,q,op,kind):h=="excel"?Excel(app,q,op,kind):PowerPoint(app,q,op,kind);
        Console.Write(Json.Serialize(new {ok=true,value=result}));
      }catch(Exception e){Console.Write(Json.Serialize(new {ok=false,message=e is InvalidOperationException?e.Message:"无法完成 Office 操作。请打开可编辑文档，检查当前选区，然后重试。"}));}
    }
  }
}
