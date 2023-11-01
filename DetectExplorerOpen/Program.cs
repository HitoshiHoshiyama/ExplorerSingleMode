using System;
using System.Windows.Automation;

class Program
{
    static void Main(string[] args)
    {
        Automation.AddAutomationEventHandler(
            WindowPattern.WindowOpenedEvent,
            AutomationElement.RootElement,
            TreeScope.Subtree,
            (sender, e) =>
            {
                if (e.EventId == WindowPattern.WindowOpenedEvent)
                {
                    AutomationElement element = sender as AutomationElement;
                    if (element.Current.ClassName == "CabinetWClass")
                    {
                        Console.WriteLine("エクスプローラーウィンドウ({0:x})が開かれました", element.Current.NativeWindowHandle);
                    }
                }
            });
        Console.ReadLine();
    }
}
