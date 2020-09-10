using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;

namespace ISC_Win_WinForm_GUI
{
    public partial class ProgressBar : Form
    {
        System.Timers.Timer onTopTimer = new System.Timers.Timer();
        public static event Action UserCancelRequest = null;
        internal static bool SendUserCancelRequest { set { UserCancelRequest(); } }
        public ProgressBar(String Title, String Content, Boolean Cancellable)
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            MainWindow.RequestPBWClose += new Action(ProgressCompleted);
            MainWindow.RequestPBWContentChange += new Action<string>(UpdateContent);
            this.Text = Title;
            tb_content.AutoSize = false;
            Point content_start_pos = new Point();
            content_start_pos.Y = tb_content.Location.Y;
            content_start_pos.X = (int) (this.Location.X + this.Width * 0.025);
            tb_content.Location = content_start_pos;
            tb_content.Width = (int)(this.Width * 0.9);
            tb_content.Text = Content;
            if(!Cancellable)
            {
                button_cancel.Visible = false;
                this.Height = (int)(this.Height * 0.8);
            }
            onTopTimer.Interval = 250;
            onTopTimer.Enabled = true;
            onTopTimer.Elapsed += onTopTimer_Elapsed;
            onTopTimer.Start();
        }

        private void onTopTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            this.BringToFront();
        }

        private void ProgressCompleted()
        {
            this.Close();
        }

        private void UpdateContent(String newContent)
        {
            tb_content.Text = newContent;
        }

        private void button_cancel_Click(object sender, EventArgs e)
        {
            button_cancel.Width = button_cancel.Width * 2;
            Point newLocation = new Point();
            newLocation.X = this.Width / 2 - button_cancel.Width / 2;
            newLocation.Y = button_cancel.Location.Y;
            button_cancel.Location = newLocation;
            button_cancel.Text = "Cancelling...";
            button_cancel.Enabled = false;

            SendUserCancelRequest = true;
        }

        private const int CP_NOCLOSE_BUTTON = 0x200;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams myCp = base.CreateParams;
                myCp.ClassStyle = myCp.ClassStyle | CP_NOCLOSE_BUTTON;
                return myCp;
            }
        }
    }
}
