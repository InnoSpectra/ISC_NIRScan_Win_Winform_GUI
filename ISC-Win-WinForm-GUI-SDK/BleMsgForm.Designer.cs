namespace ISC_Win_WinForm_GUI
{
    partial class BleMsgForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lb_msg = new System.Windows.Forms.Label();
            this.btn_wait = new System.Windows.Forms.Button();
            this.btn_quit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lb_msg
            // 
            this.lb_msg.AutoSize = true;
            this.lb_msg.Location = new System.Drawing.Point(32, 21);
            this.lb_msg.Name = "lb_msg";
            this.lb_msg.Size = new System.Drawing.Size(202, 36);
            this.lb_msg.TabIndex = 0;
            this.lb_msg.Text = "Device is in Bluetooth mode...\n\nUSB connection is currently not available!";
            // 
            // btn_wait
            // 
            this.btn_wait.Location = new System.Drawing.Point(132, 68);
            this.btn_wait.Name = "btn_wait";
            this.btn_wait.Size = new System.Drawing.Size(108, 32);
            this.btn_wait.TabIndex = 1;
            this.btn_wait.Text = "Wait for Bluetooth";
            this.btn_wait.UseVisualStyleBackColor = true;
            this.btn_wait.Click += new System.EventHandler(this.btn_wait_Click);
            // 
            // btn_quit
            // 
            this.btn_quit.Location = new System.Drawing.Point(30, 68);
            this.btn_quit.Name = "btn_quit";
            this.btn_quit.Size = new System.Drawing.Size(71, 32);
            this.btn_quit.TabIndex = 2;
            this.btn_quit.Text = "Exit GUI";
            this.btn_quit.UseVisualStyleBackColor = true;
            this.btn_quit.Click += new System.EventHandler(this.btn_quit_Click);
            // 
            // BleMsgForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(265, 112);
            this.Controls.Add(this.btn_quit);
            this.Controls.Add(this.btn_wait);
            this.Controls.Add(this.lb_msg);
            this.Name = "BleMsgForm";
            this.Text = "Device in Bluetooth Mode";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.BleMsgForm_FormClosing);
            this.Shown += new System.EventHandler(this.BleMsgForm_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lb_msg;
        private System.Windows.Forms.Button btn_wait;
        private System.Windows.Forms.Button btn_quit;
    }
}