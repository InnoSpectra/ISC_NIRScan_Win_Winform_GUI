using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ISC_Win_WinForm_GUI
{
    public partial class ProgressBar : Form
    {
        public ProgressBar(String Title, String Content, Boolean Cancellable)
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            MainWindow.RequestPBWClose += new Action(ProgressCompleted);
            this.Text = Title;
            label_content.AutoSize = false;
            Point content_start_pos = new Point();
            content_start_pos.Y = label_content.Location.Y;
            content_start_pos.X = this.Location.X;
            label_content.Location = content_start_pos;
            label_content.Width = (int)(this.Width * 0.9);
            label_content.TextAlign = ContentAlignment.MiddleCenter;
            label_content.Text = Content;
            if(!Cancellable)
            {
                button_cancel.Visible = false;
                this.Height = (int)(this.Height * 0.8);
            }
        }

        private void ProgressCompleted()
        {
            this.Close();
        }

        private void button_cancel_Click(object sender, EventArgs e)
        {
            MainWindow.UserCancelScan = true;
            
            button_cancel.Width = button_cancel.Width * 2;
            Point newLocation = new Point();
            newLocation.X = this.Width / 2 - button_cancel.Width / 2;
            newLocation.Y = button_cancel.Location.Y;
            button_cancel.Location = newLocation;
            button_cancel.Text = "Cancelling...";
            button_cancel.Enabled = false;
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
