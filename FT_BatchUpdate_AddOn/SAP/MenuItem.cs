using System;
using System.Collections.Generic;

namespace FTS.SAP
{
    public class MenuItem
    {
        public string UniqueID { get; set; }
        public string MenuName { get; set; }
        public string FatherMenuID { get; set; }
        public int Position { get; set; } // 0 will be on the top, use large amount (eg:100) to add at bottom
        public SAPbouiCOM.BoMenuType MenuType { get; set; }
        public bool Enabled { get; set; }
        public string Image { get; set; } //Supported image formats are BMP and JPG

        public MenuItem(string uniqueID, string menuName, string fatherMenuId, int position, SAPbouiCOM.BoMenuType menuType, bool enabled, string image)
        {
            this.UniqueID = uniqueID;
            this.MenuName = menuName;
            this.FatherMenuID = fatherMenuId;
            this.Position = position;
            this.MenuType = menuType;
            this.Enabled = enabled;
            this.Image = image;
        }
    }
}
