// See https://aka.ms/new-console-template for more information
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace Test
{
    internal class Program
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll", SetLastError = true)]
        static extern void SetCursorPos(int X, int Y);
        [DllImport("user32.dll", SetLastError = true)]
        static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetSystemMetrics(int smIndex);

        private const int MOUSEEVENTF_ABSOLUTE = 0x8000;
        private const int MOUSEEVENTF_MOVE = 0x1;
        private const int MOUSEEVENTF_LEFTDOWN = 0x2;
        private const int MOUSEEVENTF_LEFTUP = 0x4;

        private const int SM_CXSCREEN = 0;
        private const int SM_CYSCREEN = 1;

        static void Main(string[] args)
        {
            var t = Type.GetTypeFromProgID("Shell.Application");
            if (t is null) return;
            dynamic o = Activator.CreateInstance(t);
            try
            {
                var ws = o.Windows();
                var coms = new List<dynamic>();
                for (int i = 0; i < ws.Count; i++)
                {
                    var ie = ws.Item(i);
                    if (ie == null) continue;
                    var path = System.IO.Path.GetFileName((string)ie.FullName);
                    if (path.ToLower() == "explorer.exe")
                    {
                        if (ie.Document.FocusedItem is not null)
                        {
                            var explorepath = System.IO.Path.GetDirectoryName(ie.Document.FocusedItem.path);
                            Console.WriteLine(explorepath);
                        }
                        coms.Add(ie);
                    }
                }
                var winElmMap = new Dictionary<IntPtr, AutomationElement>();
                foreach (var comObj in coms)
                {
                    if (comObj is not IDispatch dispatch) continue;
                    var typeInfo = dispatch.GetTypeInfo(0, 0);
                    if (GetWindowThreadProcessId((IntPtr)comObj.Hwnd, out int pid) == 0) continue;
                    Console.WriteLine("{0}:{1:X} {2:D}", Marshal.GetTypeInfoName(typeInfo), ((IntPtr)comObj.Hwnd).ToInt64(), pid);
                    if(winElmMap.ContainsKey((IntPtr)comObj.Hwnd)) { continue; }
                    var winElm = AutomationElement.FromHandle((IntPtr)comObj.Hwnd);
                    var titleElm = FindElements(winElm, "TITLE_BAR_SCAFFOLDING_WINDOW_CLASS");
                    if (titleElm is null || titleElm.Count == 0) continue;
                    winElmMap.Add((IntPtr)comObj.Hwnd, titleElm[0]);
                }
                if(winElmMap.Count > 1)
                {
                    var src = winElmMap.Last().Value;
                    var tgt = winElmMap.First().Value;
                    var srcRect = src.Current.BoundingRectangle;
                    var tgtRect = tgt.Current.BoundingRectangle;
                    SetForegroundWindow((IntPtr)src.Current.NativeWindowHandle);
                    var x = (int)srcRect.X - 250;
                    var y = (int)srcRect.Y + 20;
                    SetCursorPos(x, y);
                    //System.Threading.Thread.Sleep(100);
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                    System.Threading.Thread.Sleep(100);
                    //SetCursorPos((int)x + 100, (int)y);
                    mouse_event(MOUSEEVENTF_MOVE, 10, 0, 0, 0);
                    System.Threading.Thread.Sleep(100);
                    SetForegroundWindow((IntPtr)tgt.Current.NativeWindowHandle);
                    var smx = GetSystemMetrics(SM_CXSCREEN);
                    var smy = GetSystemMetrics(SM_CYSCREEN);
                    x = ((int)tgtRect.Right - 190) * (65535 / smx);
                    y = ((int)tgtRect.Y + 30) * (65535 / smy);
                    //SetCursorPos(x, y);
                    mouse_event(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, x, y, 0, 0);
                    System.Threading.Thread.Sleep(100);
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                }
            }
            finally
            {
                Marshal.FinalReleaseComObject(o);
            }
        }

        [ComImport]
        [Guid("00020400-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IDispatch
        {
            int GetTypeInfoCount();
            ITypeInfo GetTypeInfo(int iTInfo, int lcid);
        }
        private static AutomationElementCollection FindElements(AutomationElement rootElement, string automationClass)
        {
            return rootElement.FindAll(
                TreeScope.Subtree,
                new PropertyCondition(
                  AutomationElement.ClassNameProperty,
                  automationClass));
        }
    }
}
