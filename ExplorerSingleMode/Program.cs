using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Windows.Automation;

using ExplorerSingleMode;

class Program
{
    static void Main(string[] args)
    {
        var winElmMap = new Dictionary<IntPtr, AutomationElement>();
        var tabNumMap = new Dictionary<IntPtr, int>();

        var shellType = Type.GetTypeFromProgID("Shell.Application");
        if (shellType is null) return;
        dynamic shellInstance = Activator.CreateInstance(shellType);
        try
        {
            var shellWindows = shellInstance.Windows();
            var comList = new List<dynamic>();
            // ExplorerのCOMオブジェクトリスト作成
            for (int idx = 0; idx < shellWindows.Count; idx++)
            {
                var shellItem = shellWindows.Item(idx);
                if (shellItem == null) continue;
                var path = System.IO.Path.GetFileName((string)shellItem.FullName);
                if (path.ToLower() == "explorer.exe" && shellItem.Document.FocusedItem is not null) comList.Add(shellItem);
            }
            // タブ部分のHWNDとAutomationElementマップ作成
            foreach (var comObj in comList)
            {
                if (winElmMap.ContainsKey((IntPtr)comObj.Hwnd)) { continue; }
                var ExplorerInfo = ExplorerSingleMode.WindowManager.GetExprolerInfo((IntPtr)comObj.Hwnd);
                if (ExplorerInfo is null) continue;
                winElmMap.Add((IntPtr)comObj.Hwnd, ExplorerInfo.Item1);
                tabNumMap.Add((IntPtr)comObj.Hwnd, ExplorerInfo.Item2);
            }
        }
        finally
        {
            Marshal.FinalReleaseComObject(shellInstance);
        }

        // 母艦にするウィンドウを決定する
        IntPtr baseHwd = IntPtr.Zero;
        int tabMax = 0;
        foreach(var item in tabNumMap.ToList())
        {
            if (tabMax < item.Value)
            {
                tabMax = item.Value;
                baseHwd = item.Key;
            }
        }
        AutomationElement tgt = null;
        if (baseHwd != IntPtr.Zero)
        {
            // 母艦とその他を分離
            tgt = winElmMap[baseHwd];
            winElmMap.Remove((IntPtr)baseHwd);

            // 現存する全ウィンドウをマージ
            foreach (var item in winElmMap.ToList())
            {
                // タブ個数の分だけ繰り返し
                for (int idx = 0; idx < tabNumMap[item.Key]; idx++)
                {
                    var src = item.Value;
                    if (idx > 0)
                    {
                        // タブが減るとHWDが振り直されるため再取得
                        src = ExplorerSingleMode.WindowManager.GetExprolerInfo(item.Key, false).Item1;
                    }
                    ExplorerSingleMode.WindowManager.DragExplorerTab(src, tgt);
                }
            }
        }
        winElmMap.Clear();
        tabNumMap.Clear();

        Automation.AddAutomationEventHandler(
            WindowPattern.WindowOpenedEvent,
            AutomationElement.RootElement,
            TreeScope.Subtree,
            (sender, e) =>
            {
                if (e.EventId == WindowPattern.WindowOpenedEvent)
                {
                    AutomationElement element = sender as AutomationElement;
                    // TODO: コントロールパネルの除外
                    if (element.Current.ClassName == "CabinetWClass")
                    {
                        Console.WriteLine("エクスプローラーウィンドウ({0:x})が開かれました", element.Current.NativeWindowHandle);
                        EventQueue.Add((IntPtr)element.Current.NativeWindowHandle);
                    }
                }
            });

        Console.CancelKeyPress += new ConsoleCancelEventHandler(OnCanceled);
        while (true)
        {
            try
            {
                IntPtr hwnd = EventQueue.Take(Cancel.Token);
                var ExplorerInf = ExplorerSingleMode.WindowManager.GetExprolerInfo(hwnd, false);
                if (ExplorerInf is not null) ExplorerSingleMode.WindowManager.DragExplorerTab(ExplorerInf.Item1, tgt);
            }
            catch(OperationCanceledException)
            {
                Console.WriteLine("teminate request accepted.");
                break;
            }
            catch (NoTargetException ex)
            {
                Console.WriteLine(ex.ToString());
                tgt = AutomationElement.FromHandle(ex.Hwnd);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        Console.WriteLine("teminated.");
        Cancel.Dispose();
        EventQueue.Dispose();
    }

    private static void OnCanceled(object sender, ConsoleCancelEventArgs e)
    {
        Console.WriteLine("teminate requested.");
        e.Cancel = true;
        Cancel.Cancel();
    }

    private static AutomationElementCollection FindElements(AutomationElement rootElement, string automationClass)
    {
        return rootElement.FindAll(
            TreeScope.Subtree,
            new PropertyCondition(
              AutomationElement.ClassNameProperty,
              automationClass));
    }

    private static BlockingCollection<IntPtr> EventQueue = new BlockingCollection<IntPtr>();
    private static CancellationTokenSource Cancel = new CancellationTokenSource();
}
