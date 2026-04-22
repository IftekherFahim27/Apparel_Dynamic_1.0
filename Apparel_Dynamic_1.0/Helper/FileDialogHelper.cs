using System;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace Apparel_Dynamic_1._0.Helper
{
    class FileDialogHelper
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        public static string ShowFileDialog()
        {
            string filePath = null;

            // Get SAP B1 window handle
            IntPtr sapHandle = GetForegroundWindow();

            Thread t = new Thread(() =>
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Title = "Select a file";
                    ofd.Filter = "All files (*.*)|*.*";
                    ofd.Multiselect = false;

                    WindowWrapper wrapper = new WindowWrapper(sapHandle);

                    if (ofd.ShowDialog(wrapper) == DialogResult.OK)
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

    public class WindowWrapper : IWin32Window
    {
        private readonly IntPtr _handle;

        public WindowWrapper(IntPtr handle)
        {
            _handle = handle;
        }

        public IntPtr Handle
        {
            get { return _handle; }
        }
    }
}