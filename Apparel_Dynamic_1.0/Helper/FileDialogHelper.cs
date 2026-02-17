using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Apparel_Dynamic_1._0.Helper
{
    class FileDialogHelper
    {
        public static string ShowFileDialog()
        {
            string filePath = null;

            Thread t = new Thread(() =>
            {
                using (var ofd = new System.Windows.Forms.OpenFileDialog())
                {
                    ofd.Title = "Select a file";
                    ofd.Filter = "All files (*.*)|*.*";
                    if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        filePath = ofd.FileName;
                    }
                }
            });

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            return filePath;
        }
    }
}
