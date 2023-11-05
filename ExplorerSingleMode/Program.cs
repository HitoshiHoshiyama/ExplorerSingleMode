using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Automation;
using NLog;

using ExplorerSingleMode;

class Program
{
    static void Main(string[] args)
    {
        logger.Info("Start.");
        ExplorerSingleMode.WindowManager.SetLogger(logger);

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

        // 母艦にするウィンドウを決定する(存在すれば)
        IntPtr baseHwd = IntPtr.Zero;
        int tabMax = 0;
        foreach (var item in tabNumMap.ToList())
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
                    if (idx > 0) src = ExplorerSingleMode.WindowManager.GetExprolerInfo(item.Key, false).Item1; // タブが減るとHWDが振り直されるため再取得
                    ExplorerSingleMode.WindowManager.DragExplorerTab(src, tgt);
                }
            }
        }
        winElmMap.Clear();
        tabNumMap.Clear();
        logger.Info("Tab merge complete.");

        // イベントハンドラ設定
        Automation.AddAutomationEventHandler(WindowPattern.WindowOpenedEvent, AutomationElement.RootElement, TreeScope.Subtree, OnOpenExplorer);
        Process.GetCurrentProcess().EnableRaisingEvents = true;
        Process.GetCurrentProcess().Exited += new EventHandler(OnTermination);

        // エクスプローラの起動イベント待ちループ
        var dupeCheck = new List<IntPtr>();
        while (true)
        {
            try
            {
                IntPtr hwnd = EventQueue.Take(Cancel.Token);
                // ウィンドウひとつにつきイベントが2回発生するため重複チェック
                if (dupeCheck.Contains(hwnd))
                {
                    dupeCheck.Remove(hwnd);
                    logger.Debug($"Duplicate Hwnd:0x{hwnd:x8}");
                    continue;
                }
                var ExplorerInf = ExplorerSingleMode.WindowManager.GetExprolerInfo(hwnd, false);
                // エクスプローラではないウィンドウハンドルだった場合はドラッグアンドドロップをスキップ
                if (ExplorerInf is not null)
                {
                    ExplorerSingleMode.WindowManager.DragExplorerTab(ExplorerInf.Item1, tgt);
                    dupeCheck.Add(hwnd);
                }
            }
            catch (OperationCanceledException)       // 終了通知
            {
                logger.Debug("Teminate request accepted.");
                break;
            }
            catch (NoTargetException ex)            // ドロップターゲット消失
            {
                logger.Info("Lost drop target.");
                tgt = AutomationElement.FromHandle(ex.Hwnd);    // ドロップソースを新たなドロップターゲットに設定
            }
            catch (Exception ex)                    // その他例外はログ出力
            {
                logger.Warn(ex.ToString());
            }
        }
        Automation.RemoveAllEventHandlers();
        Cancel.Dispose();
        EventQueue.Dispose();
        logger.Info("Teminated.");
    }

    /// <summary>
    /// 終了イベントハンドラ
    /// </summary>
    /// <param name="sender">イベント通知元が設定される</param>
    /// <param name="e">イベント引数が設定される</param>
    private static void OnTermination(object sender, EventArgs e)
    {
        logger.Debug("teminate requested.");
        Cancel.Cancel();
        Process.GetCurrentProcess().Exited -= new EventHandler(OnTermination);
    }

    /// <summary>
    /// エクスプローラオープンハンドラ
    /// </summary>
    /// <param name="sender">イベント通知元が設定される</param>
    /// <param name="e">イベント引数が設定される</param>
    private static void OnOpenExplorer(object sender, AutomationEventArgs e)
    {
        if (e.EventId == WindowPattern.WindowOpenedEvent)
        {
            AutomationElement element = sender as AutomationElement;
            if (element.Current.ClassName == "CabinetWClass")
            {
                // エクスプローラのウィンドウが開かれた可能性があるためハンドルを通知
                logger.Info($"Detect open explorer window(0x{element.Current.NativeWindowHandle:x8}).");
                EventQueue.Add((IntPtr)element.Current.NativeWindowHandle);
            }
        }
    }

    /// <summary>イベントキュー</summary>
    private static BlockingCollection<IntPtr> EventQueue = new BlockingCollection<IntPtr>();
    /// <summary>イベントキューのキャンセルオブジェクト</summary>
    private static CancellationTokenSource Cancel = new CancellationTokenSource();

    private static Logger logger = LogManager.GetCurrentClassLogger();

}
