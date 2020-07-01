using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Windows;


namespace ISC_Win_WinForm_GUI
{
    public partial class LicenseWindow : Form
    {
        public LicenseWindow()
        {
            InitializeComponent();
            ShowLicense();
        }

        private void ShowLicense()
        {
            String msg = String.Empty;

        
            // ISC GUI
            //TextBlock_License.Inlines.Add(new Run("Inno Spectra GUI\n\n") { FontWeight = FontWeights.Bold, FontSize = 20 });
            msg = "Copyright © 2018 Inno Spectra Corp. All Rights Reserved.\n\n\n";
            TextBlock_License.AppendText(msg);

            // LiveCharts
            //TextBlock_License.Inlines.Add(new Run("LiveCharts\n\n") { FontWeight = FontWeights.Bold, FontSize = 20 });

            msg = "The MIT License (MIT)\n\n" +
                  "Copyright © 2016 Alberto Rodriguez & LiveCharts contributors\n\n" +
                  "Permission is hereby granted, free of charge, to any person obtaining a copy " +
                  "of this software and associated documentation files (the \"Software\"), to deal " +
                  "in the Software without restriction, including without limitation the rights " +
                  "to use, copy, modify, merge, publish, distribute, sublicense, and/ or sell " +
                  "copies of the Software, and to permit persons to whom the Software is " +
                  "furnished to do so, subject to the following conditions:\n\n" +
                  "The above copyright notice and this permission notice shall be included in all " +
                  "copies or substantial portions of the Software.\n\n" +
                  "THE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR " +
                  "IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, " +
                  "FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE " +
                  "AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER " +
                  "LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, " +
                  "OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.\n\n\n";
            TextBlock_License.AppendText(msg);

            // HIDAPI
            //TextBlock_License.Inlines.Add(new Run("HIDAPI\n\n") { FontWeight = FontWeights.Bold, FontSize = 20 });

            msg = "Multi-Platform library for communication with HID devices.\n\n" +
                  "Copyright © 2009, Alan Ott, Signal 11 Software. All Rights Reserved.\n\n" +
                  "At the discretion of the user of this library, this software may be licensed under the " +
                  "terms of the GNU General Public License v3, a BSD-Style license, or the original HIDAPI " +
                  "license as outlined in the LICENSE.txt, LICENSE-gpl3.txt, LICENSE-bsd.txt, " +
                  "and LICENSE-orig.txt files located at the root of the source distribution. " +
                  "These files may also be found in the public source code repository located at:\n";
            TextBlock_License.AppendText(msg);

            /* Hyperlink link = new Hyperlink() { NavigateUri = new Uri("https://github.com/signal11/hidapi") };
             link.Inlines.Add(new Run("https://github.com/signal11/hidapi \n\n\n"));
             link.RequestNavigate += HIDAPI_Uri;
             TextBlock_License.Inlines.Add(link);*/
            msg = "https://github.com/signal11/hidapi \n\n\n";
            TextBlock_License.AppendText(msg);
            // TPL
            //TextBlock_License.Inlines.Add(new Run("TPL\n\n") { FontWeight = FontWeights.Bold, FontSize = 20 });

            msg = "The BSD License\n\n" +
                  "Copyright © 2005-2013, Troy D. Hanson  ";
            TextBlock_License.AppendText(msg);

            /* link = new Hyperlink() { NavigateUri = new Uri("http://troydhanson.github.com/tpl/") };
             link.Inlines.Add(new Run("http://troydhanson.github.com/tpl/ \n"));
             link.RequestNavigate += TPL_Uri;
             TextBlock_License.Inlines.Add(link);*/
            msg = "http://troydhanson.github.com/tpl/ \n";
            TextBlock_License.AppendText(msg);

            msg = "All rights reserved.\n\n" +
                  "Redistribution and use in source and binary forms, with or without " +
                  "modification, are permitted provided that the following conditions are met:\n\n" +
                  "* Redistributions of source code must retain the above copyright " +
                  "notice, this list of conditions and the following disclaimer.\n\n" +
                  "THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS \"AS " +
                  "IS\" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED " +
                  "TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A " +
                  "PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER " +
                  "OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, " +
                  "EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, " +
                  "PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR " +
                  "PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF " +
                  "LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING " +
                  "NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS " +
                  "SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.";
            TextBlock_License.AppendText(msg);
        }

        private void button_LicenseOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }

     
    }
}
