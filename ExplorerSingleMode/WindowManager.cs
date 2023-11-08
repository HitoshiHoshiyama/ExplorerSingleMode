using System.Runtime.InteropServices;
using System.Windows.Automation;
using NLog;

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
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetSystemMetrics(int smIndex);
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool IsIconic(IntPtr hWnd);
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
        private const int SW_MINIMIZE = 6;

        /// <summary>ドラッグ時のY座標補正値。</summary>
        private const int DRAG_OFFSET_Y = 20;
        /// <summary>ドラッグ操作時のX軸マウス移動量。</summary>
        private const int DRAG_SLIDE_X = 10;
        /// <summary>ドロップ座標のY座標補正値。</summary>
        private const int DROP_OFFSET_Y = 30;

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
                    if (IsMin) ShowWindow(Hwnd, SW_SHOWNORMAL);// 省くと最小化されたウィンドウのタブが0とカウントされる
                }
                var TabNum = NeedTabCount ? FindElements(WinElm, "ShellTabWindowClass").Count : 1;
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
        public static void DragExplorerTab(Tuple<AutomationElement, IntPtr> Source, Tuple<AutomationElement, IntPtr> Target)
        {
            try
            {
                BlockInput(true);

                // 移動先のウインドウがない場合はSourceを次Target候補として例外で通知
                if (Target is null) throw new NoTargetException((IntPtr)Source.Item1.Current.NativeWindowHandle, Source.Item2);
                // 移動先ウィンドウが最小化されている場合は戻す
                var IsMin = IsIconic(Target.Item2);
                if (IsMin) ShowWindow(Target.Item2, SW_SHOWNORMAL);
                if (!Target.Item1.Current.IsEnabled) throw new NoTargetException((IntPtr)Source.Item1.Current.NativeWindowHandle, Source.Item2);

                // 移動先ウィンドウの座標情報取得(タスクバーが上/左にあった場合の補正情報込み)
                var TgtRect = Target.Item1.Current.BoundingRectangle;
                var TgtScreen = Screen.FromHandle((IntPtr)Target.Item1.Current.NativeWindowHandle);
                var TgtPosCorrect = new Point(TgtScreen.WorkingArea.Left - TgtScreen.Bounds.X, TgtScreen.WorkingArea.Top - TgtScreen.Bounds.Y);

                // 移動するウィンドウの座標情報取得(タスクバーが上/左にあった場合の補正情報込み)
                var SrcRect = Source.Item1.Current.BoundingRectangle;
                var SrcScreen = Screen.FromHandle((IntPtr)Source.Item1.Current.NativeWindowHandle);
                var SrcPosCorrect = new Point(SrcScreen.WorkingArea.Left - SrcScreen.Bounds.X, SrcScreen.WorkingArea.Top - SrcScreen.Bounds.Y);
                var DragX = (int)SrcRect.X - SrcPosCorrect.X;
                var DragY = (int)SrcRect.Y - SrcPosCorrect.Y + DRAG_OFFSET_Y;

                // ドラッグアンドドロップ操作
                SetCursorPos(DragX, DragY);
                SetForegroundWindow((IntPtr)Source.Item1.Current.NativeWindowHandle);
                System.Threading.Thread.Sleep(100 + OperationWaitOffset);
                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                System.Threading.Thread.Sleep(100 + OperationWaitOffset);
                mouse_event(MOUSEEVENTF_MOVE, DRAG_SLIDE_X, 0, 0, 0);
                System.Threading.Thread.Sleep(100 + OperationWaitOffset);
                SetForegroundWindow(Target.Item2);
                var smx = GetSystemMetrics(SM_CXSCREEN);
                var smy = GetSystemMetrics(SM_CYSCREEN);
                var DropX = ((int)TgtRect.Right - TgtPosCorrect.X) * (65535 / smx);
                var DropY = ((int)TgtRect.Y - TgtPosCorrect.Y + DROP_OFFSET_Y) * (65535 / smy);
                mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, DropX, DropY, 0, 0);
                System.Threading.Thread.Sleep(200 + OperationWaitOffset);
                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                System.Threading.Thread.Sleep(900 + OperationWaitOffset);

                // 移動先ウィンドウを元の状態に戻す
                if (IsMin) ShowWindow(Target.Item2, SW_MINIMIZE);
                LoggerInstance?.Debug($"Mouse move:({DragX},{DragY})->({DropX},{DropY}), operation wait offset:{OperationWaitOffset}(msec)");
            }
            finally
            {
                BlockInput(false);
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
        /// <returns>クラス名が一致したAutomationElementのAutomationElementCollectionを返す。</returns>
        private static AutomationElementCollection FindElements(AutomationElement rootElement, string automationClass)
        {
            return rootElement.FindAll(TreeScope.Subtree, new PropertyCondition(AutomationElement.ClassNameProperty, automationClass));
        }

#nullable enable
        /// <summary>ロガーのインスタンス。</summary>
        private static Logger? LoggerInstance { get; set; }
#nullable disable

        /// <summary>座標のスタック。</summary>
        private static Stack<Point> PointStack = new Stack<Point>();
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
