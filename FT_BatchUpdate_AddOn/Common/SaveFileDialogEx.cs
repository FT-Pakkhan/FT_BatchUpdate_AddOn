using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;


namespace FTS.Common
{
    class SaveFileDialogEx
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        SaveFileDialog _oFileDialog;

        // Properties 
        public string FileName
        {
            get { return _oFileDialog.FileName; }
            set { _oFileDialog.FileName = value; }
        }

        public string Filter
        {
            get { return _oFileDialog.Filter; }
            set { _oFileDialog.Filter = value; }
        }

        public string Title
        {
            get { return _oFileDialog.Title; }
            set { _oFileDialog.Title = value; }
        }

        public string InitialDirectory
        {
            get { return _oFileDialog.InitialDirectory; }
            set { _oFileDialog.InitialDirectory = value; }
        }

        // Constructor 
        public SaveFileDialogEx() { _oFileDialog = new SaveFileDialog(); }

        // Methods 
        public void GetFileName()
        {
            IntPtr ptr = GetForegroundWindow();
            WindowWrapper oWindow = new WindowWrapper(ptr);
            if (_oFileDialog.ShowDialog(oWindow) != DialogResult.OK)
            { _oFileDialog.FileName = string.Empty; }
            oWindow = null;
        }
    }
}
