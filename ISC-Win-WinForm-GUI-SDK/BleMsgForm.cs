using ISC_Win_CS_LIB;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ISC_Win_WinForm_GUI
{
    public partial class BleMsgForm : Form
    {
        private Boolean UsbWait;
        private BackgroundWorker bwUsbBusyCheck;
        public BleMsgForm(bool wait)
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            if ( wait )
            {
                btn_quit.Visible = false;
                btn_wait.Visible = false;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                lb_msg.Text = "\n\n\t\t\tWaiting for Bluetooth closing...";
                UsbWait = wait;
                bwUsbBusyCheck = new BackgroundWorker();
                bwUsbBusyCheck.DoWork += new DoWorkEventHandler(bwUsbBusyCheck_DoWork);
                bwUsbBusyCheck.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bwUsbBusyCheck_Complete);
            }
        }

        private void btn_quit_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Abort;
            this.Close();
        }

        private void btn_wait_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void BleMsgForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(this.DialogResult == System.Windows.Forms.DialogResult.Cancel)
                this.DialogResult = System.Windows.Forms.DialogResult.Abort;
        }

        private void BleMsgForm_Shown(object sender, EventArgs e)
        {
            if ( UsbWait )
            {
                bwUsbBusyCheck.RunWorkerAsync();
            }
        }
        private void bwUsbBusyCheck_DoWork(object sender, DoWorkEventArgs e)
        {
            while (SDK.IsUsbConnectionBusy) { Thread.Sleep(500); }
        }
        private void bwUsbBusyCheck_Complete(object sender, RunWorkerCompletedEventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }
    }
}
