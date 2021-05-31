using System;
using System.Diagnostics;

namespace FTS.Common
{
    public class SystemHandler
    {
        static public void ExecuteBatFile(string batFilePath, string batFileName, string arguments, ref string batOutput, ref string batError, ref string batExitCode)
        {
            int ExitCode;
            ProcessStartInfo ProcessInfo;
            Process process;

            try
            {
                ProcessInfo = new ProcessStartInfo(batFilePath + "\\" + batFileName, arguments);
                ProcessInfo.CreateNoWindow = true;
                ProcessInfo.UseShellExecute = false;
                ProcessInfo.WorkingDirectory = batFilePath;
                // *** Redirect the output ***
                ProcessInfo.RedirectStandardError = true;
                ProcessInfo.RedirectStandardOutput = true;

                process = Process.Start(ProcessInfo);
                process.WaitForExit();

                // *** Read the streams ***
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                ExitCode = process.ExitCode;

                //System.Windows.Forms.MessageBox.Show("output>>" + (String.IsNullOrEmpty(output) ? "(none)" : output));
                //System.Windows.Forms.MessageBox.Show("error>>" + (String.IsNullOrEmpty(error) ? "(none)" : error));
                //System.Windows.Forms.MessageBox.Show("ExitCode: " + ExitCode.ToString(), "ExecuteCommand");
                batOutput = (String.IsNullOrEmpty(output) ? "" : output);
                batError = (String.IsNullOrEmpty(error) ? "" : error);
                batExitCode = ExitCode.ToString();

                process.Close();
            }
            catch (Exception ex)
            {
                batError = ex.Message;
            }
        }


        static public void ExecuteCommand(string command, ref string batOutput, ref string batError, ref string batExitCode)
        {
            ProcessStartInfo ProcessInfo;
            Process process;
            int ExitCode;

            try
            {
                ProcessInfo = new System.Diagnostics.ProcessStartInfo("cmd", "/c " + command);
                // Do not create the black window.
                ProcessInfo.CreateNoWindow = true;
                ProcessInfo.UseShellExecute = false;
                // *** Redirect the output ***
                ProcessInfo.RedirectStandardError = true;
                ProcessInfo.RedirectStandardOutput = true;

                // Now we create a process, assign its ProcessStartInfo and start it
                process = new System.Diagnostics.Process();
                process.StartInfo = ProcessInfo;
                process.Start();

                // Get the output into a string
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                ExitCode = process.ExitCode;

                //System.Windows.Forms.MessageBox.Show("output>>" + (String.IsNullOrEmpty(output) ? "(none)" : output));
                //System.Windows.Forms.MessageBox.Show("error>>" + (String.IsNullOrEmpty(error) ? "(none)" : error));
                //System.Windows.Forms.MessageBox.Show("ExitCode: " + ExitCode.ToString(), "ExecuteCommand");
                batOutput = (String.IsNullOrEmpty(output) ? "" : output);
                batError = (String.IsNullOrEmpty(error) ? "" : error);
                batExitCode = ExitCode.ToString();

            }
            catch (Exception err)
            {
                throw new Exception(err.Message);
            }
        }
    }
}
