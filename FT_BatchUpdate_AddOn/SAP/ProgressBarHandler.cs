using System;

namespace FTS.SAP
{
    public class ProgressBarHandler
    {
        static private SAPbouiCOM.ProgressBar AppProgressBar = null;
        static private int CurStep = 0;
        static private int Maximum = 0;
        static private Boolean Stoppable = false;
        static private int RetryStart = 0;
        static public Boolean ProgressBarExists = false;

        static public Boolean Start(string text, int curStep, int maximum, Boolean stoppable)
        {
            try
            {
                CurStep = curStep;
                Maximum = maximum;
                Stoppable = stoppable;
                AppProgressBar = AddOn.ApplicationInstance.StatusBar.CreateProgressBar("[" + curStep.ToString() + "/" + maximum.ToString() + "] " + text, maximum, stoppable);
                ProgressBarExists = true;
                RetryStart = 0;
                return true;
            }
            catch
            {
                if (RetryStart < 3) // try 3 restart before return error
                {
                    RetryStart++;
                    Stop();
                    return Start(text, curStep, maximum, stoppable);
                }
                else
                {
                    return false;
                }

            }
        }

        static public void Stop()
        {
            try
            {
                if (ProgressBarExists == true)
                {
                    AppProgressBar.Stop();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(AppProgressBar);
                    AppProgressBar = null;
                    ProgressBarExists = false;
                }
            }
            catch
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(AppProgressBar);
                AppProgressBar = null;
                ProgressBarExists = false;
            }
        }

        static public void Increment(string text, int step)
        {
            try
            {
                if (ProgressBarExists)
                {
                    CurStep += step;
                    AppProgressBar.Value = CurStep;
                    AppProgressBar.Text = "[" + CurStep.ToString() + "/" + Maximum.ToString() + "] " + text;
                }
                else
                {
                    CurStep += step;
                    AddOn.ApplicationInstance.SetStatusBarMessage("[" + CurStep.ToString() + "/" + Maximum.ToString() + "] " + text, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                }
            }
            catch
            {
                Stop();
                Start(text, CurStep, Maximum, Stoppable);
            }
        }
    }
}
