using System;

namespace FTS.SAP
{
    public class UserPermissionItem
    {
        public string PermissionID { get; set; }
        public string PermissionName { get; set; }
        public SAPbobsCOM.BoUPTOptions PermissionOptions { get; set; }
        public string ParentID { get; set; }
        public string FormType { get; set; }

        public UserPermissionItem(string permissionID, string permissionName, SAPbobsCOM.BoUPTOptions permissionOptions, string parentID, string formType)
        {
            this.PermissionID = permissionID;
            this.PermissionName = permissionName;
            this.PermissionOptions = permissionOptions;
            this.ParentID = parentID;
            this.FormType = formType;
        }
    }
}
