using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;

namespace ISC_Win_WinForm_GUI
{
    class Message
    {
        public static void ShowInfo(string Text, string Caption = null)
        {
            if (string.IsNullOrEmpty(Caption))
                Caption = "Information";

            MessageBox.Show(Text, Caption, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
        }

        public static void ShowWarning(string Text, string Caption = null)
        {
            if (string.IsNullOrEmpty(Caption))
                Caption = "Warning";

            MessageBox.Show(Text, Caption, MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
        }

        public static void ShowError(string Text, string Caption = null)
        {
            if (string.IsNullOrEmpty(Caption))
                Caption = "Error";

            MessageBox.Show(Text, Caption, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification);
        }

        public static DialogResult ShowQuestion(string Text, string Caption = null, MessageBoxButtons Button = MessageBoxButtons.YesNo)
        {
            MessageBoxDefaultButton DefaultBtn;

            if (string.IsNullOrEmpty(Caption))
                Caption = "Question";

            if (Button == MessageBoxButtons.OKCancel || Button == MessageBoxButtons.RetryCancel || Button == MessageBoxButtons.YesNo)
                DefaultBtn = MessageBoxDefaultButton.Button2;
            else if (Button == MessageBoxButtons.YesNoCancel)
                DefaultBtn = MessageBoxDefaultButton.Button3;
            else
                DefaultBtn = MessageBoxDefaultButton.Button1;

            return MessageBox.Show(Text, Caption, Button, MessageBoxIcon.Question, DefaultBtn);
        }
    }
}

