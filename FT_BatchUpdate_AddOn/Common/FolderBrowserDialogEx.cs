using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace FTS.Common
{
    class FolderBrowserDialogEx
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
        FolderBrowserDialog _oFileDialog;

        // Properties 
        public string SelectedPath
        {
            get { return _oFileDialog.SelectedPath; }
            set { _oFileDialog.SelectedPath = value; }
        }

        public string Description
        {
            get { return _oFileDialog.Description; }
            set { _oFileDialog.Description = value; }
        }

        public bool ShowNewFolderButton
        {
            get { return _oFileDialog.ShowNewFolderButton; }
            set { _oFileDialog.ShowNewFolderButton = value; }
        }

        // Constructor 
        public FolderBrowserDialogEx() { _oFileDialog = new FolderBrowserDialog(); }

        // Methods 
        public void GetFileName()
        {
            IntPtr ptr = GetForegroundWindow();
            WindowWrapper oWindow = new WindowWrapper(ptr);
            if (_oFileDialog.ShowDialog(oWindow) != DialogResult.OK)
            { _oFileDialog.SelectedPath = string.Empty; }
            oWindow = null;
        }
    }
}
