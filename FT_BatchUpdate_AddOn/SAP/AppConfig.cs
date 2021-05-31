using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FTS.SAP
{
    public class AppConfig
    {
        public bool UseDIAPI { get; set; }
        public bool EnableMenuEvent { get; set; }
        public bool EnableItemEvent { get; set; }
        public bool EnableFormDataEvent { get; set; }
        public bool EnableProgressBarEvent { get; set; }
        public bool EnableStatusBarEvent { get; set; }
        public bool EnableRightClickEvent { get; set; }
        public bool EnablePrintEvent { get; set; }
        public bool EnableLayoutKeyEvent { get; set; }

        public AppConfig()
        {
            this.UseDIAPI = true;
            this.EnableMenuEvent = false;
            this.EnableItemEvent = false;
            this.EnableFormDataEvent = false;
            this.EnableProgressBarEvent = false;
            this.EnableStatusBarEvent = false;
            this.EnableRightClickEvent = false;
            this.EnablePrintEvent = false;
            this.EnableLayoutKeyEvent = false;
        }
    }
}
