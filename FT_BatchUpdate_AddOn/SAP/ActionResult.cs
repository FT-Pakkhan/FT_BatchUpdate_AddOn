using System;
using System.Collections.Generic;

namespace FTS.SAP
{
    public class ActionResult
    {
        public string Status { get; set; }
        public string Key { get; set; }
        public string Reason { get; set; }

        public ActionResult(string status, string key, string reason)
        {
            this.Status = (status.Length > 100 ? status.Substring(0, 100) : status);
            this.Key = (key.Length > 50 ? key.Substring(0, 50) : key);
            this.Reason = (reason.Length > 254 ? reason.Substring(0, 254) : reason);
        }
    }
}
