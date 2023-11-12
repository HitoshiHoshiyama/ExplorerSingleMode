using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using NLog;

namespace ExplorerSingleMode
{
    internal partial class DummyForm : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        public DummyForm()
        {
            InitializeComponent();

            this.Width = 150;
            this.Height = 100;
            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.AllowTransparency = true;
            this.TopMost = true;
#if DEBUG
            this.Opacity = 0.5;
#else
            this.Opacity = 0.0;
#endif

            this.Shown += new EventHandler(OnShown);
            this.Closed += new EventHandler(OnClosed);
        }

        /// <summary>
        /// Shownイベントハンドラ。
        /// </summary>
        /// <param name="sender">イベント発生元が指定されている。</param>
        /// <param name="e">イベント引数が設定されている。</param>
        public void OnShown(object sender, EventArgs e) { this.logger?.Info($"Dummy form shown."); }

        /// <summary>
        /// Closedイベントハンドラ。
        /// </summary>
        /// <param name="sender">イベント発生元が指定されている。</param>
        /// <param name="e">イベント引数が設定されている。</param>
        public void OnClosed(object sender, EventArgs e) { this.logger?.Info($"Dummy form close."); }

        /// <summary>
        /// ロガーを設定する。
        /// </summary>
        /// <param name="logger">ロガーを指定する。</param>
        public void SetLogger(Logger logger) { this.logger = logger; }

        private Logger? logger = null;
    }
}
