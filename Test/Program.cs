// See https://aka.ms/new-console-template for more information
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace Test
{
    internal class Program
    {
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        static void Main(string[] args)
        {
            var t = Type.GetTypeFromProgID("Shell.Application");
            if (t is null) return;
            dynamic o = Activator.CreateInstance(t);
            try
            {
                var ws = o.Windows();
                var coms = new List<dynamic>();
                var autoElms = new List<AutomationElement>();
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
                foreach (var comObj in coms)
                {
                    if (comObj is not IDispatch dispatch) continue;
                    var typeInfo = dispatch.GetTypeInfo(0, 0);
                    int pid;
                    var hwnd2pid = GetWindowThreadProcessId((IntPtr)comObj.Hwnd, out pid);
                    Console.WriteLine("{0}:{1:X} {2:D}", Marshal.GetTypeInfoName(typeInfo), ((IntPtr)comObj.Hwnd).ToInt64(), pid);
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
    }
}
