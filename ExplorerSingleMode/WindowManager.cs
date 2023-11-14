using System.Runtime.InteropServices;
using System.Windows.Automation;
using NLog;
using ExplorerSingleMode;
using System.Runtime.CompilerServices;
using System.Windows;

namespace ExplorerSingleMode
{
    /// <summary>エクスプローラのタブ操作を行うクラス。</summary>
    static internal class WindowManager
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll", SetLastError = true)]
        static extern void SetCursorPos(int X, int Y);
        [DllImport("user32.dll", SetLastError = true)]
        static extern void SetWindowPos(IntPtr Hwnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);
        [DllImport("user32.dll", EntryPoint = "SystemParametersInfo", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern Boolean SystemParametersInfoGet(UInt32 action, UInt32 param, ref UInt32 vparam, UInt32 init);
        [DllImport("user32.dll", EntryPoint = "SystemParametersInfo", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern Boolean SystemParametersInfoSet(UInt32 action, UInt32 param, UInt32 vparam, UInt32 init);
        [DllImport("user32.dll", SetLastError = true)]
        static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetSystemMetrics(int smIndex);
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool IsZoomed(IntPtr hWnd);
        [System.Runtime.InteropServices.DllImportAttribute("user32.dll", EntryPoint = "BlockInput")]
        [return: System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool BlockInput([System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.Bool)] bool fBlockIt);

        private const int MOUSEEVENTF_ABSOLUTE = 0x8000;
        private const int MOUSEEVENTF_MOVE = 0x1;
        private const int MOUSEEVENTF_LEFTDOWN = 0x2;
        private const int MOUSEEVENTF_LEFTUP = 0x4;

        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;

        private const int SW_SHOWNORMAL = 1;
        private const int SW_MAXIMIZE = 3;
        private const int SW_SHOW = 5;
        private const int SW_MINIMIZE = 6;

        private const int HWND_TOP = 0;
        private const int SWP_NOSIZE = 0x0001;

        private const uint SPI_GETFOREGROUNDLOCKTIMEOUT = 0x2000;
        private const uint SPI_SETFOREGROUNDLOCKTIMEOUT = 0x2001;
        private const int SPIF_SENDCHANGE = 2;

        // TODO: WINDOW_TABLIST_HEIGHTをちゃんとAutomationElementから取得する
        /// <summary>エクスプローラタブエリア高さ(拡大とかあるから本当はAutomationElementから取得すべき)</summary>
        private const int WINDOW_TABLIST_HEIGHT = 48;
        /// <summary>ドラッグ時のX座標補正値。</summary>
        private const int DRAG_OFFSET_X = 50;
        /// <summary>ドラッグ時のY座標補正値。</summary>
        private const int DRAG_OFFSET_Y = 20;
        /// <summary>ドラッグ操作時のX軸マウス移動量。</summary>
        private const int DRAG_SLIDE_X = 10;
        /// <summary>ドロップ座標のX座標補正値。</summary>
        private const int DROP_OFFSET_X = 50;
        /// <summary>ドロップ座標のY座標補正値。</summary>
        private const int DROP_OFFSET_Y = 30;
        /// <summary>フォーカス変更完了待ちのタイマ上限(msec)。</summary>
        private const int SET_FOREGROUND_WAIT_MAX = 1000;

        /// <summary>
        /// <br>ロガーを設定する。</br>
        /// <br>未設定の場合はログ出力しない。</br>
        /// </summary>
        /// <param name="logger">ログ出力に使用するロガーを指定する。</param>
        public static void SetLogger(Logger logger)
        {
            WindowManager.LoggerInstance = logger;
        }

        /// <summary>
        /// エクスプローラのウィンドウハンドルからAutomationElementとタブ数を取得する。
        /// </summary>
        /// <param name="Hwnd">エクスプローラのウィンドウハンドルを指定する。</param>
        /// <param name="NeedTabCount">タブ数をカウントする場合はtrueを指定する。省略可能(既定値はtrue)。</param>
        /// <returns>
        /// <br>タブ部分のコントロールハンドルとタブ数のTupleを返す。</br>
        /// <br>エクスプローラのではないウィンドウハンドルを渡された場合はnullを返す。</br>
        /// </returns>
        public static Tuple<AutomationElement, int> GetExprolerInfo(IntPtr Hwnd, bool NeedTabCount = true)
        {
            var WinElm = AutomationElement.FromHandle(Hwnd);
            var TitleElm = FindElements(WinElm, "TITLE_BAR_SCAFFOLDING_WINDOW_CLASS");
            bool IsMin = false;
            try
            {
                if (TitleElm is null || TitleElm.Count == 0)
                {
                    // コントロールパネルはここで排除(Countが0になる)
                    LoggerInstance?.Debug($"HWND:0x{Hwnd:x8} TITLE_BAR_SCAFFOLDING_WINDOW_CLASS not found.");
                    return null;
                }
                if (NeedTabCount)
                {
                    IsMin = IsIconic(Hwnd);
                    if (IsMin) ShowWindow(Hwnd, SW_SHOWNORMAL); // 省くと最小化されたウィンドウのタブが0とカウントされる
                }
                var elmCount = FindElements(WinElm, "ShellTabWindowClass");
                var TabNum = NeedTabCount ? (elmCount.Count > 0 ? elmCount.Count : 1) : 1;
                if (TabNum == 0)
                {
                    LoggerInstance?.Debug($"HWND:0x{Hwnd:x8} ShellTabWindowClass not found.");
                    return null;
                }
                LoggerInstance?.Debug($"HWND:0x{Hwnd:x8} AutomationElement:0x{TitleElm[0].Current.NativeWindowHandle:x8} Tabs:{TabNum}");
                return new Tuple<AutomationElement, int>(TitleElm[0], TabNum);
            }
            finally { if (IsMin) ShowWindow(Hwnd, SW_MINIMIZE); }
        }

        /// <summary>
        /// エクスプローラのタブをドラッグアンドドロップでウィンドウ間移動させる。
        /// </summary>
        /// <param name="Source">移動するタブの有るウィンドウのAutomationElementを指定する。</param>
        /// <param name="Target">移動先ウィンドウのAutomationElementとウィンドウハンドルのTupleを指定する。</param>
        /// <exception cref="NoTargetException">移動先のエクスプローラが閉じられていた場合に発生する。</exception>
        public static void DragExplorerTab(Tuple<AutomationElement, IntPtr> Source, Tuple<AutomationElement, IntPtr> Target, DummyForm dummyForm)
        {
            uint winSwitchTime = 0;
            try
            {
                SystemParametersInfoGet(SPI_GETFOREGROUNDLOCKTIMEOUT, 0, ref winSwitchTime, 0);
                SystemParametersInfoSet(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, 0, SPIF_SENDCHANGE);
                BlockInput(true);

                // 移動先のウインドウがない場合はSourceを次Target候補として例外で通知
                if (Target is null || Target.Item2 == Source.Item2) throw new NoTargetException((IntPtr)Source.Item1.Current.NativeWindowHandle, Source.Item2);
                LoggerInstance.Debug($"Move tab source:0x{Source.Item1.Current.NativeWindowHandle} target:0x{Target.Item1.Current.NativeWindowHandle}");
                // 移動先ウィンドウが最小化(最大化)されている場合は戻す
                var IsMin = IsIconic(Target.Item2);
                if (IsMin)
                {
                    ShowWindow(Target.Item2, SW_SHOW);
                    System.Threading.Thread.Sleep(100 + OperationWaitOffset);
                }
                if (!Target.Item1.Current.IsEnabled) throw new NoTargetException((IntPtr)Source.Item1.Current.NativeWindowHandle, Source.Item2);
                var IsMax = IsZoomed(Target.Item2);
                if (IsMax)
                {
                    ShowWindow(Target.Item2, SW_SHOWNORMAL);
                    System.Threading.Thread.Sleep(100 + OperationWaitOffset);
                }

                // ウィンドウが重なっていた場合ずらす
                var SrcMovePont = GetWindowMovableSpacePosition(Source.Item1, Target.Item1);
                SetWindowPos(Source.Item2, (IntPtr)HWND_TOP, SrcMovePont.X, SrcMovePont.Y, 0, 0, SWP_NOSIZE);
                System.Threading.Thread.Sleep(100 + OperationWaitOffset);

                // 移動先ウィンドウの座標情報取得
                var TgtRect = Target.Item1.Current.BoundingRectangle;

                // 移動するウィンドウの座標情報取得
                var Parent = AutomationElement.FromHandle(Source.Item2);
                var SrcRect = Parent.Current.BoundingRectangle;
                var DragX = (int)SrcRect.X + DRAG_OFFSET_X;
                var DragY = (int)SrcRect.Y + DRAG_OFFSET_Y;

                // ドラッグアンドドロップ操作
                SetCursorPos(DragX, DragY);
                AutomationElement.FromHandle(Source.Item2).SetFocus();
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                System.Threading.Thread.Sleep(100 + OperationWaitOffset);
                mouse_event(MOUSEEVENTF_MOVE, DRAG_SLIDE_X, 0, 0, 0);
                System.Threading.Thread.Sleep(100 + OperationWaitOffset);
                AutomationElement.FromHandle(Target.Item2).SetFocus();
                var smx = GetSystemMetrics(SM_CXSCREEN);
                var smy = GetSystemMetrics(SM_CYSCREEN);
                var DropX = ((int)TgtRect.Left + DROP_OFFSET_X) * (65535 / smx);
                var DropY = ((int)TgtRect.Y + DROP_OFFSET_Y) * (65535 / smy);
                mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, DropX, DropY, 0, 0);
                System.Threading.Thread.Sleep(200 + OperationWaitOffset);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                System.Threading.Thread.Sleep(900 + OperationWaitOffset);

                // 移動先ウィンドウを元の状態に戻す
                if (IsMax) ShowWindow(Target.Item2, SW_MAXIMIZE);
                if (IsMin) ShowWindow(Target.Item2, SW_MINIMIZE);
                LoggerInstance?.Debug($"Mouse move:({DragX},{DragY})->({DropX},{DropY}), operation wait offset:{OperationWaitOffset}(msec)");
            }
            catch(Exception e) { LoggerInstance.Warn(e); }
            finally
            {
                BlockInput(false);
                SystemParametersInfoSet(SPI_SETFOREGROUNDLOCKTIMEOUT, 0, winSwitchTime, SPIF_SENDCHANGE);
            }
        }

        public static void StoreMousePosition()
        {
            PointStack.Push(Cursor.Position);
        }
        public static void RestoreMousePosition()
        {
            if (PointStack.Count > 0) Cursor.Position = PointStack.Pop();
        }

        /// <summary>マウス操作の合間に挿入する待ち時間の補正値。</summary>
        public static int OperationWaitOffset { get; set; } = 0;

        /// <summary>
        /// 親ウィンドウのAutomationElementを起点として指定クラスのAutomationElementを検索する。
        /// </summary>
        /// <param name="rootElement">検索の起点とするAutomationElementを指定する。</param>
        /// <param name="automationClass">検索するウィンドウクラス名を指定する。</param>
        /// <param name="NeedCheckSibling">rootElementの兄弟要素を確認するか指定する。既定値は false。</param>
        /// <returns>クラス名が一致した最初のAutomationElementのリストを返す。</returns>
        private static List<AutomationElement> FindElements(AutomationElement rootElement, string automationClass, bool NeedCheckSibling = false)
        {
            var result = new List<AutomationElement>();
            try
            {
                var child = TreeWalker.ContentViewWalker.GetFirstChild(rootElement);
                if (child != null)
                {
                    if (child.Current.ClassName == automationClass) result.Add(child);
                    var nextLevel = FindElements(child, automationClass, true);
                    if (nextLevel is not null) result.AddRange(nextLevel);
                }
                if (NeedCheckSibling)
                {
                    if (rootElement.Current.ClassName == automationClass) result.Add(rootElement);
                    var nextElement = TreeWalker.ContentViewWalker.GetNextSibling(rootElement);
                    if (nextElement != null) result.AddRange(FindElements(nextElement, automationClass, true));
                }
            }
            catch(Exception ex) { LoggerInstance.Warn(ex.ToString()); }
            return result;
        }

        /// <summary>
        /// <br>Sourceウィンドウの移動可能な座標を取得する。</br>
        /// <br>適切な移動先が見付からなかった場合は、現在の座標を返す。</br>
        /// </summary>
        /// <param name="Src">タブ移動元ウィンドウのAutomationElementを指定する。</param>
        /// <param name="Tgt">タブ移動先ウィンドウのAutomationElementを指定する。</param>
        /// <returns>移動先または現在の座標を返す。</returns>
        private static System.Drawing.Point GetWindowMovableSpacePosition(AutomationElement Src, AutomationElement Tgt)
        {
            if (!Src.Current.BoundingRectangle.IntersectsWith(Tgt.Current.BoundingRectangle))
                // 両者が重ならない場合は元の座標をそのまま返す
                return new System.Drawing.Point((int)Src.Current.BoundingRectangle.X, (int)Src.Current.BoundingRectangle.Y);
            var tgtRect = Tgt.Current.BoundingRectangle;
            var srcRect = Src.Current.BoundingRectangle;
            var tgtScreen = Screen.FromRectangle(new Rectangle((int)tgtRect.X, (int)tgtRect.Y, (int)tgtRect.Width, (int)tgtRect.Height));
            var left = (tgtRect.Left - tgtScreen.WorkingArea.Left) >= srcRect.Width;
            var right = (tgtScreen.WorkingArea.Right - tgtRect.Right) >= srcRect.Width;
            if (left) return new System.Drawing.Point((int)tgtScreen.WorkingArea.Left, (int)srcRect.Top);   // Tgtの左に配置
            else if (right) return new System.Drawing.Point((int)tgtRect.Right, (int)srcRect.Top);          // Tgtの右に配置
            var upper = (tgtRect.Top - tgtScreen.WorkingArea.Top) >= srcRect.Height;
            var lower = (tgtScreen.WorkingArea.Bottom - tgtRect.Bottom) >= WINDOW_TABLIST_HEIGHT;                // タブリスト分の高さだけあればヨシ
            if (upper) return new System.Drawing.Point((int)srcRect.Left, (int)tgtScreen.WorkingArea.Top);  // Tgtの上に配置
            else if (lower) return new System.Drawing.Point((int)tgtRect.Right, (int)srcRect.Top);          // Tgtの下に配置
            // Tgtの存在するScreenにはスペースが無かった
            foreach (var screen in Screen.AllScreens)
            {
                if (screen.Bounds == tgtScreen.Bounds) continue;    // さっき見た
                if (screen.WorkingArea.Width > srcRect.Width && screen.WorkingArea.Height > srcRect.Height)
                    // 収まる別のScreenがあった
                    return new System.Drawing.Point((int)screen.WorkingArea.Left, (int)screen.WorkingArea.Top);
            }
            // 移動は無し
            return new System.Drawing.Point((int)Src.Current.BoundingRectangle.X, (int)Src.Current.BoundingRectangle.Y);
        }

#nullable enable
        /// <summary>ロガーのインスタンス。</summary>
        private static Logger? LoggerInstance { get; set; }
#nullable disable

        /// <summary>座標のスタック。</summary>
        private static readonly Stack<System.Drawing.Point> PointStack = new Stack<System.Drawing.Point>();
    }

    /// <summary>ドロップターゲットが存在しなかった場合にスローされる例外。</summary>
    internal class NoTargetException : Exception
    {
        /// <summary>
        /// コンストラクタ。
        /// </summary>
        /// <param name="ElementHwnd">ドロップソースのウィンドウハンドルを指定する。</param>
        /// <param name="ParentHwnd">ドロップソースの親ウィンドウハンドルを指定する。</param>
        public NoTargetException(IntPtr ElementHwnd, IntPtr ParentHwnd)
        {
            this.ElementHwnd = ElementHwnd;
            this.ParentHwnd = ParentHwnd;
        }
        /// <summary>ドロップソースのウィンドウハンドルを取得する。</summary>
        public IntPtr ElementHwnd { get; private set; }
        /// <summary>ドロップソースの親ウィンドウハンドルを取得する。</summary>
        public IntPtr ParentHwnd { get; private set; }
    }
}
