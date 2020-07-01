using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace ISC_Win_WinForm_GUI
{
    static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        static String[] DeprecatedProductName = { "DLP-NIR-Win-SDK", "ISC-NIRScan", "ISC-WinForm-SDK", "ISC-WinForm-GUI" };
        static String TiEvmProductName = "NirscanNanoGUI";
        static String WinISCGuiProductName = "Win-ISC-GUI";
        [STAThread]
        static void Main()
        {
            Process currentProc = Process.GetCurrentProcess();
            String currentProcName = currentProc.ProcessName;
            String currentProdName = Application.ProductName;

            if (System.Diagnostics.Process.GetProcessesByName(currentProcName).Length > 1)
            {
                DialogResult ret = Message.ShowQuestion("Existing ISC NIRScan SDK GUI detected!\n\nClose all existig GUI?");

                if (ret == DialogResult.Yes)
                {
                    Process[] p = Process.GetProcessesByName(currentProcName).ToArray();
                    foreach (Process thisProc in p)
                    {
                        if (thisProc.Id != currentProc.Id)
                            thisProc.Kill();
                    }
                }
                else
                {
                    Message.ShowError("GUI launch failed!\n\nPlease close other related GUI's and retry again.");
                    return;
                }
            }

            if (System.Diagnostics.Process.GetProcessesByName(TiEvmProductName).Length > 0)
            {
                DialogResult ret = Message.ShowQuestion("Existing TI EVM GUI detected!\n\nClose all existig App?");

                if (ret == DialogResult.Yes)
                {
                    Process[] p = Process.GetProcessesByName(TiEvmProductName).ToArray();
                    foreach (Process thisProc in p)
                       thisProc.Kill();
                }
                else
                {
                    Message.ShowError("GUI launch failed!\n\nPlease close other related GUI's and retry again.");
                    return;
                }
            }

            if (System.Diagnostics.Process.GetProcessesByName(WinISCGuiProductName).Length > 0)
            {
                DialogResult ret = Message.ShowQuestion("Existing ISC GUI detected!\n\nClose all existig App?");

                if (ret == DialogResult.Yes)
                {
                    Process[] p = Process.GetProcessesByName(WinISCGuiProductName).ToArray();
                    foreach (Process thisProc in p)
                        thisProc.Kill();
                }
                else
                {
                    Message.ShowError("GUI launch failed!\n\nPlease close other related GUI's and retry again.");
                    return;
                }
            }

            for (int i = 0; i < DeprecatedProductName.Length; i++)
            {
                String pattern = @DeprecatedProductName[i];
                foreach (Process p in Process.GetProcesses("."))
                {
                    MatchCollection matches = Regex.Matches(p.ProcessName, pattern);
                    if (matches.Count > 0)
                    {
                        DialogResult ret;
                        if (p.MainModule.FileVersionInfo.ProductName == currentProdName && p.Id == currentProc.Id)
                            continue;

                        if (p.ProcessName.Contains("Qt"))
                            ret = Message.ShowQuestion("Existing ISC Factory GUI detected!\n\nClose the existig App?");
                        else
                            ret = Message.ShowQuestion("Existing old ISC NIRScan GUI detected!\n\nClose the existig App?");

                        if (ret == DialogResult.Yes)
                            p.Kill();
                        else
                        {
                            Message.ShowError("GUI launch failed!\n\nPlease close other related GUI's and retry again.");
                            return;
                        }
                    }
                }
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainWindow());
        }
    }
}
