using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web.Script.Serialization;

namespace TrioBridge
{
public class TestSlide
{
    public int SlideID { get; set; }
    public string Text { get; set; }
    public TestSlides Owner { get; set; }
    public object Shapes { get { throw new Exception("Insertion must not read slide objects."); } }
    public void Delete() { Owner.Items.Remove(this); }
}
public class TestSlides : IEnumerable
{
    public readonly List<TestSlide> Items = new List<TestSlide>();
    public int ImportCalls;
    public int ImportIndex;
    public int Count { get { return Items.Count; } }
    public TestSlide this[int index] { get { return Items[index - 1]; } }
    public IEnumerator GetEnumerator() { return Items.GetEnumerator(); }
    public void Add(int id) { Items.Add(new TestSlide { SlideID = id, Owner = this }); }
    public int InsertFromFile(string file, int index)
    {
        if (file != "generated.pptx") throw new Exception("Unexpected input file.");
        ImportCalls++;
        ImportIndex = index;
        Add(900);
        Add(901);
        return 2;
    }
}
public class TestPageSetup
{
    public float SlideWidth { get; set; }
    public float SlideHeight { get; set; }
}
public class TestPresentation
{
    public string FullName { get; set; }
    public TestSlides Slides { get; set; }
    public TestPageSetup PageSetup { get; set; }
}
public class TestWindow
{
    // PowerPoint.DocumentWindow does not expose Word/Excel's HWND property.
    public string Caption { get; set; }
    public object Selection { get { throw new Exception("Insertion must not read a selection."); } }
}
public class TestApplication
{
    public TestPresentation ActivePresentation { get; set; }
    public TestWindow ActiveWindow { get; set; }
}
public class TestWordCount
{
    public int Count { get; set; }
}
public class TestWordWindow
{
    public int HWND { get; set; }
    public Action OnActivate;
    public int ActivateCalls;
    public void Activate() { ActivateCalls++; if (OnActivate != null) OnActivate(); }
}
public class TestWordSelection
{
    public TestWordRange Range { get; set; }
}
public class TestWordDocument
{
    public string FullName { get; set; }
    public bool ReadOnly { get; set; }
    public string Body { get; set; }
    public int Writes;
    public int PlainWrites;
    public int RichWrites;
    public int CloseCalls;
    public int TableCount;
    public readonly List<string> Operations = new List<string>();
    public TestWordRange Content { get { return new TestWordRange(this, 0, Body.Length); } }
    public TestWordRange Range(int start, int end) { return new TestWordRange(this, start, end); }
    public void Close(bool save)
    {
        if (save) throw new Exception("The imported temporary document must never be saved.");
        CloseCalls++;
    }
}
public class TestWordParagraph
{
    public TestWordRange Range { get; set; }
}
public class TestWordParagraphs
{
    readonly TestWordRange range;
    public TestWordParagraphs(TestWordRange value) { range = value; }
    public TestWordParagraph this[int index]
    {
        get
        {
            if (index < 1) throw new Exception("Word paragraph indexes are one-based.");
            string text = range.Owner.Body;
            int start = range.Start == 0 ? 0 : text.LastIndexOf('\r', range.Start - 1) + 1;
            for (int i = 1; i < index; i++)
            {
                int next = text.IndexOf('\r', start);
                if (next < 0) throw new Exception("Paragraph index exceeds the document.");
                start = next + 1;
            }
            int marker = text.IndexOf('\r', start);
            int end = marker < 0 ? text.Length : marker + 1;
            if (range.Owner.TableCount > 0 && end < text.Length && text[end] == '\a') end++;
            return new TestWordParagraph { Range = new TestWordRange(range.Owner, start, end) };
        }
    }
}
public class TestWordRange
{
    public TestWordDocument Owner;
    public int Start { get; set; }
    public int End { get; set; }
    public int StoryType { get; set; }
    public TestWordCount Tables = new TestWordCount();
    public TestWordCount Fields = new TestWordCount();
    public TestWordCount InlineShapes = new TestWordCount();
    public TestWordCount ContentControls = new TestWordCount();
    public TestWordCount OMaths = new TestWordCount();
    public TestWordCount Hyperlinks = new TestWordCount();
    public TestWordCount Footnotes = new TestWordCount();
    public TestWordCount Endnotes = new TestWordCount();
    public TestWordRange(TestWordDocument owner, int start, int end)
    {
        Owner = owner; Start = start; End = end; StoryType = 1;
        Tables.Count = owner.TableCount;
    }
    public TestWordRange Duplicate { get { return new TestWordRange(Owner, Start, End); } }
    public TestWordParagraphs Paragraphs { get { return new TestWordParagraphs(this); } }
    public string Text
    {
        get { return Owner.Body.Substring(Start, End - Start); }
        set { Owner.PlainWrites++; Write(value, "Text"); }
    }
    public object FormattedText
    {
        get { return this; }
        set
        {
            Owner.RichWrites++;
            Write(((TestWordRange)value).Text, "FormattedText");
        }
    }
    void Write(string value, string operation)
    {
        Owner.Writes++;
        Owner.Operations.Add(operation + ":" + Start + ":" + End);
        Owner.Body = Owner.Body.Substring(0, Start) + value + Owner.Body.Substring(End);
        End = Start + value.Length;
    }
    public void SetRange(int start, int end) { Start = start; End = end; }
    public void Collapse(int direction) { if (direction == 0) Start = End; else End = Start; }
    public void InsertAfter(string value) { Start = End; Write(value, "InsertAfter"); }
    public void InsertBefore(string value) { Write(value + Text, "InsertBefore"); }
}
public class TestWordDocuments
{
    public string ExpectedFile;
    public string ImportedText = "new text\r";
    public int OpenCalls;
    public bool FailOpen;
    public Action OnOpen;
    public TestWordDocument Imported;
    public TestWordDocument Open(string FileName, bool ConfirmConversions, bool ReadOnly,
                                 bool AddToRecentFiles, bool Visible)
    {
        if (FileName != ExpectedFile || ConfirmConversions || ReadOnly || AddToRecentFiles || Visible)
            throw new Exception("HTML must be opened privately without prompts or recent-file entries.");
        OpenCalls++;
        if (FailOpen) throw new InvalidOperationException("Synthetic import failure.");
        Imported = new TestWordDocument { FullName = FileName, Body = ImportedText };
        if (OnOpen != null) OnOpen();
        return Imported;
    }
}
public class TestWordApplication
{
    public TestWordDocument ActiveDocument { get; set; }
    public TestWordWindow ActiveWindow { get; set; }
    public TestWordSelection Selection { get; set; }
    public TestWordDocuments Documents = new TestWordDocuments();
}
internal class WindowsOfficeTests
{
    static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
    static readonly MethodInfo InvokePowerPoint = typeof(Office).GetMethod(
        "PowerPoint", BindingFlags.Static | BindingFlags.NonPublic);
    static readonly MethodInfo InvokeWord = typeof(Office).GetMethod(
        "Word", BindingFlags.Static | BindingFlags.NonPublic);
    static void Check(bool value, string message)
    {
        if (!value) throw new Exception(message);
    }
    static TestApplication App(bool empty = false)
    {
        var slides = new TestSlides();
        if (!empty) { slides.Add(256); slides.Add(257); }
        return new TestApplication {
            ActivePresentation = new TestPresentation {
                FullName = "Synthetic.pptx", Slides = slides,
                PageSetup = new TestPageSetup { SlideWidth = 960, SlideHeight = 540 }
            },
            ActiveWindow = new TestWindow { Caption = "Synthetic.pptx" }
        };
    }
    static object Call(TestApplication app, string op, Dictionary<string, object> input)
    {
        try { return InvokePowerPoint.Invoke(null, new object[] { app, input, op, "" }); }
        catch (TargetInvocationException e) { throw e.InnerException; }
    }
    static Dictionary<string, object> Prepare(TestApplication app)
    {
        // Round-trip the actual helper result through the IPC JSON representation.
        return (Dictionary<string, object>)Json.DeserializeObject(Json.Serialize(
            Call(app, "presentationTarget", new Dictionary<string, object>())));
    }
    static void Append(TestApplication app, Dictionary<string, object> prepared, int count = 2)
    {
        Call(app, "append", new Dictionary<string, object> {
            { "guard", prepared["guard"] }, { "target", prepared["target"] },
            { "file", "generated.pptx" }, { "count", count }
        });
    }
    static void RejectChangedTarget(Action<TestApplication> change, string description)
    {
        var app = App();
        var prepared = Prepare(app);
        change(app);
        bool rejected = false;
        try { Append(app, prepared); }
        catch (InvalidOperationException) { rejected = true; }
        Check(rejected, description + " must invalidate the insertion target.");
        Check(app.ActivePresentation.Slides.ImportCalls == 0,
              description + " must be rejected before insertion.");
    }
    static TestWordApplication WordApp(string text = "before selected after\r", int start = 7, int end = 15)
    {
        var doc = new TestWordDocument { FullName = "Synthetic.docx", Body = text };
        return new TestWordApplication {
            ActiveDocument = doc,
            ActiveWindow = new TestWordWindow { HWND = 123 },
            Selection = new TestWordSelection { Range = new TestWordRange(doc, start, end) }
        };
    }
    static object WordCall(TestWordApplication app, string op, string kind, Dictionary<string, object> input)
    {
        try { return InvokeWord.Invoke(null, new object[] { app, input, op, kind }); }
        catch (TargetInvocationException e) { throw e.InnerException; }
    }
    static Dictionary<string, object> WordPrepare(TestWordApplication app, string kind = "replace")
    {
        return (Dictionary<string, object>)Json.DeserializeObject(Json.Serialize(
            WordCall(app, kind == "replace" ? "selection" : "target", kind,
                     new Dictionary<string, object>())));
    }
    static void WordWrite(TestWordApplication app, Dictionary<string, object> prepared,
                          string kind, string value)
    {
        WordCall(app, "write", kind, new Dictionary<string, object> {
            { "guard", prepared["guard"] }, { "value", value }
        });
    }
    static void WordRichWrite(TestWordApplication app, Dictionary<string, object> prepared,
                              string kind, bool block)
    {
        string file = Path.GetTempFileName();
        try
        {
            app.Documents.ExpectedFile = file;
            WordCall(app, "write", kind, new Dictionary<string, object> {
                { "guard", prepared["guard"] }, { "file", file }, { "block", block }
            });
        }
        finally { File.Delete(file); }
    }
    static void CheckRichWrite(TestWordApplication app, string expected, string description)
    {
        Check(app.ActiveDocument.Body == expected, description + " must preserve exact text and paragraph boundaries.");
        Check(app.ActiveDocument.Writes == 1 && app.ActiveDocument.RichWrites == 1 &&
              app.ActiveDocument.PlainWrites == 0, description + " must use one formatted destination write.");
        Check(app.Documents.OpenCalls == 1 && app.Documents.Imported.CloseCalls == 1,
              description + " must open and close one private import document.");
    }
    static void RejectWordChange(Action<TestWordApplication> change, string description)
    {
        var app = WordApp();
        var prepared = WordPrepare(app);
        change(app);
        string before = app.ActiveDocument.Body;
        bool rejected = false;
        try { WordWrite(app, prepared, "replace", "replacement"); }
        catch (InvalidOperationException) { rejected = true; }
        Check(rejected, description + " must reject Word replacement.");
        Check(app.ActiveDocument.Body == before && app.ActiveDocument.Writes == 0,
              description + " must not mutate Word content.");
    }
    static void RejectWordDuringImport(Action<TestWordApplication> change, string description)
    {
        var app = WordApp();
        var original = app.ActiveDocument;
        var prepared = WordPrepare(app);
        app.Documents.OnOpen = () => change(app);
        bool rejected = false;
        try { WordRichWrite(app, prepared, "replace", true); }
        catch (InvalidOperationException) { rejected = true; }
        Check(rejected && original.Body == "before selected after\r" && original.Writes == 0,
              description + " during import must be checked before the single destination write.");
        Check(app.Documents.Imported.CloseCalls == 1,
              description + " during import must close the temporary document.");
    }
    static void WordScenarios()
    {
        var plain = WordApp();
        WordWrite(plain, WordPrepare(plain), "replace", "new text");
        Check(plain.ActiveDocument.Body == "before new text after\r" &&
              plain.ActiveDocument.Writes == 1 && plain.ActiveDocument.PlainWrites == 1,
              "Plain selection replacement must preserve surrounding text in one write.");
        RejectWordChange(a => a.ActiveDocument.FullName = "Other.docx", "Document change");
        RejectWordChange(a => a.ActiveWindow.HWND = 456, "Window change");
        RejectWordChange(a => a.Selection.Range.Start++, "Selection change");
        RejectWordChange(a => a.ActiveDocument.Body = "before modified after\r", "Selected text change");
        RejectWordChange(a => a.ActiveDocument.ReadOnly = true, "Read-only document");
        RejectWordChange(a => a.Selection.Range.Tables.Count = 1, "Protected table selection");
        RejectWordChange(a => a.Selection.Range.Fields.Count = 1, "Protected field selection");

        var inline = WordApp();
        WordRichWrite(inline, WordPrepare(inline), "replace", false);
        CheckRichWrite(inline, "before new text after\r", "Inline selection replacement");

        var block = WordApp();
        block.Documents.ImportedText = "Heading\rBody\r";
        WordRichWrite(block, WordPrepare(block), "replace", true);
        CheckRichWrite(block, "before \rHeading\rBody\r after\r", "Block selection replacement");
        Check(block.Documents.Imported.Operations.Contains("InsertBefore:0:13") &&
              block.Documents.Imported.Operations.Contains("FormattedText:0:1"),
              "The leading paragraph boundary and inherited prefix formatting belong to the temporary document.");

        var insert = WordApp();
        WordRichWrite(insert, WordPrepare(insert, "insert"), "insert", false);
        CheckRichWrite(insert, "before selectednew text after\r", "Inline insertion at the exact selection end");
        Check(insert.ActiveDocument.Operations[0] == "FormattedText:15:15",
              "Insertion must not replace the selection or drift to its paragraph end.");

        var cursor = WordApp("prefixSUFFIX\r", 6, 6);
        WordRichWrite(cursor, WordPrepare(cursor, "insert"), "insert", false);
        CheckRichWrite(cursor, "prefixnew textSUFFIX\r", "Inline insertion at a collapsed cursor");

        var paragraphEnd = WordApp("picked\rNEXT\r", 0, 7);
        paragraphEnd.Documents.ImportedText = "Heading\rBody\r";
        WordRichWrite(paragraphEnd, WordPrepare(paragraphEnd, "insert"), "insert", true);
        CheckRichWrite(paragraphEnd, "picked\rHeading\rBody\rNEXT\r", "Block insertion after a selected paragraph marker");

        var wholeBlock = WordApp("before\rselected\rafter\r", 7, 16);
        wholeBlock.Documents.ImportedText = "Heading\rBody\r";
        WordRichWrite(wholeBlock, WordPrepare(wholeBlock), "replace", true);
        CheckRichWrite(wholeBlock, "before\rHeading\rBody\rafter\r", "Whole-paragraph block replacement");

        var textOnlyBlock = WordApp("before\rselected\rafter\r", 7, 15);
        textOnlyBlock.Documents.ImportedText = "Heading\rBody\r";
        WordRichWrite(textOnlyBlock, WordPrepare(textOnlyBlock), "replace", true);
        CheckRichWrite(textOnlyBlock, "before\rHeading\rBody\rafter\r",
                       "Block replacement ending just before the original paragraph marker");
        Check(textOnlyBlock.ActiveDocument.Operations[0] == "FormattedText:7:16",
              "The verified following paragraph marker must join the single formatted replacement.");

        var emptyParagraph = WordApp("before\r\rafter\r", 7, 7);
        emptyParagraph.Documents.ImportedText = "Heading\rBody\r";
        WordRichWrite(emptyParagraph, WordPrepare(emptyParagraph, "insert"), "insert", true);
        CheckRichWrite(emptyParagraph, "before\rHeading\rBody\rafter\r", "Block insertion into an empty paragraph");
        Check(emptyParagraph.ActiveDocument.Operations[0] == "FormattedText:7:8",
              "An empty destination paragraph must not leave an extra blank paragraph after import.");

        var paragraphTail = WordApp("prefix\rafter\r", 6, 6);
        paragraphTail.Documents.ImportedText = "Heading\rBody\r";
        WordRichWrite(paragraphTail, WordPrepare(paragraphTail, "insert"), "insert", true);
        CheckRichWrite(paragraphTail, "prefix\rHeading\rBody\rafter\r", "Block insertion at a paragraph text end");
        Check(paragraphTail.ActiveDocument.Operations[0] == "FormattedText:6:7",
              "Paragraph-end insertion must isolate the prefix without duplicating the existing final marker.");

        var followingBlank = WordApp("selected\r\rNEXT\r", 0, 9);
        followingBlank.Documents.ImportedText = "Heading\rBody\r";
        WordRichWrite(followingBlank, WordPrepare(followingBlank), "replace", true);
        CheckRichWrite(followingBlank, "Heading\rBody\r\rNEXT\r", "Block replacement before an unselected blank paragraph");
        Check(followingBlank.ActiveDocument.Operations[0] == "FormattedText:0:9",
              "A selection already containing its paragraph marker must not consume the following blank paragraph.");

        var cellEnd = WordApp("\r\aNEXT\r", 0, 0);
        cellEnd.ActiveDocument.TableCount = 1;
        cellEnd.Selection.Range.Tables.Count = 1;
        cellEnd.Documents.ImportedText = "Heading\rBody\r";
        WordRichWrite(cellEnd, WordPrepare(cellEnd, "insert"), "insert", true);
        CheckRichWrite(cellEnd, "Heading\rBody\r\r\aNEXT\r", "Block insertion preserves cell-end structure");
        Check(cellEnd.ActiveDocument.Operations[0] == "FormattedText:0:0",
              "Paragraph deduplication must not consume a table cell-end marker.");

        var wholeInline = WordApp("before\rselected\rafter\r", 7, 16);
        WordRichWrite(wholeInline, WordPrepare(wholeInline), "replace", false);
        CheckRichWrite(wholeInline, "before\rnew text\rafter\r", "Whole-paragraph inline replacement");

        var failedImport = WordApp();
        var failedGuard = WordPrepare(failedImport);
        failedImport.Documents.FailOpen = true;
        bool importRejected = false;
        try { WordRichWrite(failedImport, failedGuard, "replace", true); }
        catch (InvalidOperationException) { importRejected = true; }
        Check(importRejected && failedImport.ActiveDocument.Body == "before selected after\r" &&
              failedImport.ActiveDocument.Writes == 0,
              "An import failure must leave the original selection untouched.");

        var staleImport = WordApp();
        var staleGuard = WordPrepare(staleImport);
        staleImport.Documents.OnOpen = () => staleImport.ActiveDocument.Body = "before modified after\r";
        bool staleRejected = false;
        try { WordRichWrite(staleImport, staleGuard, "replace", true); }
        catch (InvalidOperationException) { staleRejected = true; }
        Check(staleRejected && staleImport.ActiveDocument.Body == "before modified after\r" &&
              staleImport.ActiveDocument.Writes == 0 && staleImport.Documents.Imported.CloseCalls == 1,
              "A changed selection during import must be rejected and the temporary document closed.");

        RejectWordDuringImport(a => a.Selection.Range = a.ActiveDocument.Range(0, 6), "A new selection object");
        RejectWordDuringImport(a => a.ActiveDocument = new TestWordDocument {
            FullName = "Other.docx", Body = "Other document\r"
        }, "Active document switch");
        RejectWordDuringImport(a => a.ActiveWindow = new TestWordWindow { HWND = 456 }, "Active window switch");
        RejectWordDuringImport(a => a.Selection.Range.Fields.Count = 1, "A newly protected selection");
        RejectWordDuringImport(a => a.ActiveDocument.ReadOnly = true, "New read-only protection");

        var focus = WordApp();
        var focusDoc = focus.ActiveDocument;
        var focusWindow = focus.ActiveWindow;
        var focusSelection = focus.Selection;
        focusWindow.OnActivate = () => {
            focus.ActiveDocument = focusDoc;
            focus.ActiveWindow = focusWindow;
            focus.Selection = focusSelection;
        };
        focus.Documents.OnOpen = () => {
            focus.ActiveDocument = focus.Documents.Imported;
            focus.ActiveWindow = new TestWordWindow { HWND = 999 };
            focus.Selection = new TestWordSelection { Range = focus.Documents.Imported.Content };
        };
        WordRichWrite(focus, WordPrepare(focus), "replace", false);
        CheckRichWrite(focus, "before new text after\r", "Temporary importer focus restoration");
        Check(focusWindow.ActivateCalls == 1, "Only the original destination window may be reactivated.");

        foreach (string property in new[] { "Tables", "Fields", "InlineShapes", "ContentControls",
                                          "OMaths", "Hyperlinks", "Footnotes", "Endnotes" })
        {
            var protectedApp = WordApp();
            var count = (TestWordCount)typeof(TestWordRange).GetField(property).GetValue(protectedApp.Selection.Range);
            count.Count = 1;
            var protectedGuard = WordPrepare(protectedApp);
            Check((bool)protectedGuard["writable"] == false, property + " must be reported as non-replaceable.");
            bool rejected = false;
            try { WordRichWrite(protectedApp, protectedGuard, "replace", true); }
            catch (InvalidOperationException) { rejected = true; }
            Check(rejected && protectedApp.ActiveDocument.Writes == 0 && protectedApp.Documents.OpenCalls == 0,
                  property + " must reject formatted replacement before even opening the import.");
        }

        foreach (string imported in new[] { "\r", "missing final paragraph marker" })
        {
            var invalid = WordApp();
            invalid.Documents.ImportedText = imported;
            bool rejected = false;
            try { WordRichWrite(invalid, WordPrepare(invalid), "replace", false); }
            catch (InvalidOperationException) { rejected = true; }
            Check(rejected && invalid.ActiveDocument.Body == "before selected after\r" &&
                  invalid.ActiveDocument.Writes == 0 && invalid.Documents.Imported.CloseCalls == 1,
                  "An empty or invalid inline import must not clear the selection.");
        }

        var protectedInsert = WordApp();
        protectedInsert.Selection.Range.Tables.Count = 1;
        WordRichWrite(protectedInsert, WordPrepare(protectedInsert, "insert"), "insert", false);
        CheckRichWrite(protectedInsert, "before selectednew text after\r", "Insertion preserves a protected selection");
        Console.WriteLine("Windows Word writing: 32 scenarios passed.");
    }
    static void RejectDispatch(Action action, string description)
    {
        bool rejected = false;
        try { action(); } catch (InvalidOperationException) { rejected = true; }
        Check(rejected, description);
    }
    static void PaneDispatchScenarios()
    {
        var word = WordApp();
        var request = new Dictionary<string, object> { { "host", "word" }, { "op", "selection" } };
        var result = (Dictionary<string, object>)Json.DeserializeObject(Json.Serialize(
            Office.Dispatch(word, word.ActiveWindow, "word", request)));
        Check(result.ContainsKey("guard") && !result.ContainsKey("ok"),
              "The embedded entry must return the raw native result, without a second envelope.");
        RejectDispatch(() => Office.Dispatch(word, new TestWordWindow { HWND = 123 }, "word", request),
                       "Matching HWND text must not replace exact COM window identity.");
        RejectDispatch(() => Office.Dispatch(WordApp(), word.ActiveWindow, "word", request),
                       "Another application instance must not inherit the pane's target window.");
        RejectDispatch(() => Office.Dispatch(word, word.ActiveWindow, "excel", request),
                       "Cross-host native requests must be rejected.");
        RejectDispatch(() => Office.CheckWindow(word, null), "Closed windows must fail closed.");
        var ppt = App();
        var target = new Dictionary<string, object> { { "host", "powerpoint" }, { "op", "presentationTarget" } };
        Check(Office.Dispatch(ppt, ppt.ActiveWindow, "powerpoint", target) != null,
              "PowerPoint panes must support the selection-independent insertion target.");
        RejectDispatch(() => Office.Dispatch(ppt, new TestWindow { Caption = ppt.ActiveWindow.Caption },
                                              "powerpoint", target),
                       "Matching PowerPoint captions must not replace exact window identity.");
        var original = word.ActiveWindow;
        word.ActiveWindow = new TestWordWindow { HWND = 456 };
        RejectDispatch(() => Office.Dispatch(word, original, "word", request),
                       "Switching active windows must invalidate pane operations before selection is read.");
        Check(word.ActiveDocument.Writes == 0, "Rejected pane dispatches must never write.");
        Console.WriteLine("Windows pane dispatch: 8 scenarios passed.");
    }
    public static void Main()
    {
        var app = App();
        var prepared = Prepare(app);
        app.ActivePresentation.Slides.Items[0].Text = "Unrelated existing content changed";
        Append(app, prepared);
        Check(app.ActivePresentation.Slides.ImportIndex == 2, "Append must target the end.");
        Check(app.ActivePresentation.Slides.Items.Select(s => s.SlideID)
                  .SequenceEqual(new[] { 256, 257, 900, 901 }), "Existing slide order must remain.");

        var empty = App(true);
        Append(empty, Prepare(empty));
        Check(empty.ActivePresentation.Slides.ImportIndex == 0 &&
              empty.ActivePresentation.Slides.Count == 2, "Empty decks must accept slides.");

        RejectChangedTarget(a => a.ActivePresentation.FullName = "Other.pptx", "Document change");
        RejectChangedTarget(a => a.ActiveWindow.Caption = "Synthetic.pptx:2", "Window change");
        RejectChangedTarget(a => a.ActivePresentation.Slides.Items.Reverse(), "Slide reorder");
        RejectChangedTarget(a => a.ActivePresentation.Slides.Add(258), "Slide addition");
        RejectChangedTarget(a => a.ActivePresentation.PageSetup.SlideWidth = 720, "Width change");
        RejectChangedTarget(a => a.ActivePresentation.PageSetup.SlideHeight = 720, "Height change");

        var rollback = App();
        bool failed = false;
        try { Append(rollback, Prepare(rollback), 3); }
        catch (InvalidOperationException) { failed = true; }
        Check(failed && rollback.ActivePresentation.Slides.Items.Select(s => s.SlideID)
                  .SequenceEqual(new[] { 256, 257 }), "Count mismatch must roll back new slides.");
        Console.WriteLine("Windows PowerPoint insertion: 9 scenarios passed.");
        WordScenarios();
        PaneDispatchScenarios();
    }
}
}
