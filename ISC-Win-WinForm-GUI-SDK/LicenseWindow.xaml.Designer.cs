namespace ISC_Win_WinForm_GUI
{
    partial class LicenseWindow
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LicenseWindow));
            this.button_LicenseOk = new System.Windows.Forms.Button();
            this.TextBlock_License = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // button_LicenseOk
            // 
            this.button_LicenseOk.Location = new System.Drawing.Point(697, 526);
            this.button_LicenseOk.Name = "button_LicenseOk";
            this.button_LicenseOk.Size = new System.Drawing.Size(75, 23);
            this.button_LicenseOk.TabIndex = 1;
            this.button_LicenseOk.Text = "OK";
            this.button_LicenseOk.UseVisualStyleBackColor = true;
            this.button_LicenseOk.Click += new System.EventHandler(this.button_LicenseOk_Click);
            // 
            // TextBlock_License
            // 
            this.TextBlock_License.Location = new System.Drawing.Point(-1, -1);
            this.TextBlock_License.Name = "TextBlock_License";
            this.TextBlock_License.Size = new System.Drawing.Size(787, 521);
            this.TextBlock_License.TabIndex = 0;
            this.TextBlock_License.Text = "";
            // 
            // LicenseWindow
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(784, 561);
            this.Controls.Add(this.button_LicenseOk);
            this.Controls.Add(this.TextBlock_License);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "LicenseWindow";
            this.Text = "LicenseWindow";
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button button_LicenseOk;
        private System.Windows.Forms.RichTextBox TextBlock_License;
    }
}