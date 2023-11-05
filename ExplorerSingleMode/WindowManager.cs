using System.Runtime.InteropServices;
using System.Windows.Automation;
using NLog;

namespace ExplorerSingleMode
{
    /// <summary>エクスプローラのタブ操作を行うクラス</summary>
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

        private const int MOUSEEVENTF_ABSOLUTE = 0x8000;
        private const int MOUSEEVENTF_MOVE = 0x1;
        private const int MOUSEEVENTF_LEFTDOWN = 0x2;
        private const int MOUSEEVENTF_LEFTUP = 0x4;

        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;

        private const int SW_SHOWNORMAL = 1;
        private const int SW_MINIMIZE = 6;

        /// <summary>ドラッグ時のY座標補正値</summary>
        private const int DRAG_OFFSET_Y = 20;
        /// <summary>ドラッグ操作時のX軸マウス移動量</summary>
        private const int DRAG_SLIDE_X = 10;
        /// <summary>ドロップ座標のY座標補正値</summary>
        private const int DROP_OFFSET_Y = 30;

        public static void SetLogger(Logger logger)
        {
            WindowManager.logger = logger;
        }

        /// <summary>
        /// エクスプローラのウィンドウハンドルからAutomationElementとタブ数を取得
        /// </summary>
        /// <param name="Hwnd">エクスプローラのウィンドウハンドル</param>
        /// <param name="NeedTabCount">タブ数をカウントする場合はtrue</param>
        /// <returns>
        /// <br>タブ部分のコントロールハンドルとタブ数のTupleを返す。</br>
        /// <br>エクスプローラのではないウィンドウハンドルを渡された場合はnullを返す。</br>
        /// </returns>
        public static Tuple<AutomationElement, int> GetExprolerInfo(IntPtr Hwnd, bool NeedTabCount = true)
        {
            var WinElm = AutomationElement.FromHandle(Hwnd);
            var TitleElm = FindElements(WinElm, "TITLE_BAR_SCAFFOLDING_WINDOW_CLASS");
            if (TitleElm is null || TitleElm.Count == 0)
            {
                // コントロールパネルはここで排除(Countが0になる)
                logger.Debug($"HWND:0x{Hwnd:x8} TITLE_BAR_SCAFFOLDING_WINDOW_CLASS not found.");
                return null;
            }
            if (NeedTabCount) ShowWindow(Hwnd, SW_SHOWNORMAL);          // 省くと最小化されたウィンドウのタブが0とカウントされる
            var TabNum = NeedTabCount ? FindElements(WinElm, "ShellTabWindowClass").Count : 1;
            if (TabNum == 0)
            {
                logger.Debug($"HWND:0x{Hwnd:x8} ShellTabWindowClass not found.");
                return null;
            }
            logger.Debug($"HWND:0x{Hwnd:x8} AutomationElement:0x{TitleElm[0].Current.NativeWindowHandle:x8} Tabs:{TabNum}");
            return new Tuple<AutomationElement, int>(TitleElm[0], TabNum);
        }

        /// <summary>
        /// エクスプローラのタブをドラッグアンドドロップでウィンドウ間移動させる
        /// </summary>
        /// <param name="Source">移動するタブの有るウィンドウのAutomationElement</param>
        /// <param name="Target">移動先ウィンドウのAutomationElement</param>
        /// <exception cref="NoTargetException">移動先のエクスプローラが閉じられていた場合に発生</exception>
        public static void DragExplorerTab(AutomationElement Source, AutomationElement Target)
        {
            // 移動先のウインドウがない場合はSourceを次Target候補として例外で通知
            if (Target is null) throw new NoTargetException((IntPtr)Source.Current.NativeWindowHandle);
            // 移動先ウィンドウが最小化されている場合は戻す
            var IsMin = IsIconic((IntPtr)Target.Current.NativeWindowHandle);
            if (IsMin) ShowWindow((IntPtr)Target.Current.NativeWindowHandle, SW_SHOWNORMAL);
            if(!Target.Current.IsEnabled)throw new NoTargetException((IntPtr)Source.Current.NativeWindowHandle);

            // 移動先ウィンドウの座標情報取得(タスクバーが上/左にあった場合の補正情報込み)
            var TgtRect = Target.Current.BoundingRectangle;
            var TgtScreen = Screen.FromHandle((IntPtr)Target.Current.NativeWindowHandle);
            var TgtPosCorrect = new Point(TgtScreen.WorkingArea.Left - TgtScreen.Bounds.X, TgtScreen.WorkingArea.Top - TgtScreen.Bounds.Y);

            // 移動するウィンドウの座標情報取得(タスクバーが上/左にあった場合の補正情報込み)
            var SrcRect = Source.Current.BoundingRectangle;
            var SrcScreen = Screen.FromHandle((IntPtr)Source.Current.NativeWindowHandle);
            var SrcPosCorrect = new Point(SrcScreen.WorkingArea.Left - SrcScreen.Bounds.X, SrcScreen.WorkingArea.Top - SrcScreen.Bounds.Y);
            var DragX = (int)SrcRect.X - SrcPosCorrect.X;
            var DragY = (int)SrcRect.Y - SrcPosCorrect.Y + DRAG_OFFSET_Y;

            // ドラッグアンドドロップ操作
            SetCursorPos(DragX, DragY);
            SetForegroundWindow((IntPtr)Source.Current.NativeWindowHandle);
            System.Threading.Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            System.Threading.Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_MOVE, DRAG_SLIDE_X, 0, 0, 0);
            System.Threading.Thread.Sleep(100);
            SetForegroundWindow((IntPtr)Target.Current.NativeWindowHandle);
            var smx = GetSystemMetrics(SM_CXSCREEN);
            var smy = GetSystemMetrics(SM_CYSCREEN);
            var DropX = ((int)TgtRect.Right - TgtPosCorrect.X) * (65535 / smx);
            var DropY = ((int)TgtRect.Y - TgtPosCorrect.Y + DROP_OFFSET_Y) * (65535 / smy);
            mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, DropX, DropY, 0, 0);
            System.Threading.Thread.Sleep(200);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            System.Threading.Thread.Sleep(900);

            // 移動先ウィンドウを元の状態に戻す
            if (IsMin) ShowWindow((IntPtr)Target.Current.NativeWindowHandle, SW_MINIMIZE);
            logger.Debug($"Mouse move:({DragX},{DragY})->({DropX},{DropY})");
        }

        /// <summary>
        /// 親ウィンドウのAutomationElementを起点として指定クラスのAutomationElementを検索
        /// </summary>
        /// <param name="rootElement">検索の起点とするAutomationElementを指定</param>
        /// <param name="automationClass">検索するウィンドウクラス名を指定</param>
        /// <returns>クラス名が一致したAutomationElementのAutomationElementCollectionを返す</returns>
        private static AutomationElementCollection FindElements(AutomationElement rootElement, string automationClass)
        {
            return rootElement.FindAll(TreeScope.Subtree, new PropertyCondition(AutomationElement.ClassNameProperty, automationClass));
        }

        private static Logger logger { get; set; }
    }

    /// <summary>ドロップターゲットが存在しなかった場合にスローされる例外</summary>
    internal class NoTargetException : Exception
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="Hwnd">ドロップソースのウィンドウハンドルを指定</param>
        public NoTargetException(IntPtr Hwnd) { this.Hwnd = Hwnd; }
        /// <summary>ドロップソースのウィンドウハンドルを取得</summary>
        public IntPtr Hwnd { get; private set; }
    }
}
