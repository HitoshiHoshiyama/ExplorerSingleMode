using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace ExplorerSingleMode
{
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

        public static Tuple<AutomationElement, int> GetExprolerInfo(IntPtr Hwnd, bool NeedTabCount = true)
        {
            var WinElm = AutomationElement.FromHandle(Hwnd);
            var TitleElm = FindElements(WinElm, "TITLE_BAR_SCAFFOLDING_WINDOW_CLASS");
            if (TitleElm is null || TitleElm.Count == 0) return null;
            if (NeedTabCount) ShowWindow(Hwnd, SW_SHOWNORMAL);          // 省くと最小化されたウィンドウのタブが0とカウントされる
            var TabNum = NeedTabCount ? FindElements(WinElm, "ShellTabWindowClass").Count : 1;
            if (TabNum == 0) return null;
            return new Tuple<AutomationElement, int>(TitleElm[0], TabNum);
        }

        public static void DragExplorerTab(AutomationElement Src, AutomationElement Target)
        {
            if (Target is null) throw new NoTargetException((IntPtr)Src.Current.NativeWindowHandle);
            var IsMin = IsIconic((IntPtr)Target.Current.NativeWindowHandle);
            if (IsMin) ShowWindow((IntPtr)Target.Current.NativeWindowHandle, SW_SHOWNORMAL);
            if(!Target.Current.IsEnabled)throw new NoTargetException((IntPtr)Src.Current.NativeWindowHandle);
            var TgtRect = Target.Current.BoundingRectangle;
            var TgtScreen = Screen.FromHandle((IntPtr)Target.Current.NativeWindowHandle);
            var TgtPosCorrect = new Point(TgtScreen.WorkingArea.Left - TgtScreen.Bounds.X, TgtScreen.WorkingArea.Top - TgtScreen.Bounds.Y);

            var SrcRect = Src.Current.BoundingRectangle;
            var SrcScreen = Screen.FromHandle((IntPtr)Src.Current.NativeWindowHandle);
            var SrcPosCorrect = new Point(SrcScreen.WorkingArea.Left - SrcScreen.Bounds.X, SrcScreen.WorkingArea.Top - SrcScreen.Bounds.Y);
            var x = (int)SrcRect.X - SrcPosCorrect.X;
            var y = (int)SrcRect.Y - SrcPosCorrect.Y + 20;
            SetCursorPos(x, y);
            SetForegroundWindow((IntPtr)Src.Current.NativeWindowHandle);
            System.Threading.Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            System.Threading.Thread.Sleep(50);
            mouse_event(MOUSEEVENTF_MOVE, 10, 0, 0, 0);
            System.Threading.Thread.Sleep(100);
            SetForegroundWindow((IntPtr)Target.Current.NativeWindowHandle);
            var smx = GetSystemMetrics(SM_CXSCREEN);
            var smy = GetSystemMetrics(SM_CYSCREEN);
            x = ((int)TgtRect.Right - TgtPosCorrect.X) * (65535 / smx);
            y = ((int)TgtRect.Y - TgtPosCorrect.Y + 30) * (65535 / smy);
            mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, x, y, 0, 0);
            System.Threading.Thread.Sleep(150);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            System.Threading.Thread.Sleep(900);
            if (IsMin) ShowWindow((IntPtr)Target.Current.NativeWindowHandle, SW_MINIMIZE);
        }

        private static AutomationElementCollection FindElements(AutomationElement rootElement, string automationClass)
        {
            return rootElement.FindAll(TreeScope.Subtree, new PropertyCondition(AutomationElement.ClassNameProperty, automationClass));
        }
    }

    internal class NoTargetException : Exception
    {
        public NoTargetException(IntPtr Hwnd) { this.Hwnd = Hwnd; }
        public IntPtr Hwnd { get; private set; }
    }
}
