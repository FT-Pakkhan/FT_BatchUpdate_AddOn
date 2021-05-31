using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System.Runtime.Serialization;

using System.Threading;
using SAPbouiCOM;
using SAPbobsCOM;
using FTS.SAP.Extension;

namespace FTS.SAP
{
    [Serializable]
    public class Warning : Exception
    {
        public Warning()
            : base() { }

        public Warning(string message)
            : base(message) { }

        public Warning(string format, params object[] args)
            : base(string.Format(format, args)) { }

        public Warning(string message, Exception innerException)
            : base(message, innerException) { }

        public Warning(string format, Exception innerException, params object[] args)
            : base(string.Format(format, args), innerException) { }

        protected Warning(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }

    public class AddOnSysParm
    {
        // AddOn System Parameter
        // Stored in FTS_SYSPARM table
        public string Version { get; set; }
        public string Parm1 { get; set; }
        public string Parm2 { get; set; }
        public string Parm3 { get; set; }
        public string AddOnKey { get; set; }
        public string SystemVersion { get; set; }

        public AddOnSysParm()
        {
            Version = "";
            Parm1 = "";
            Parm2 = "";
            Parm3 = "";
            AddOnKey = "";
            SystemVersion = "";
        }
    }

    public class UDOParm
    {
        public string UdoName { get; set; }
        public string UdoDescription { get; set; }
        public SAPbobsCOM.BoUDOObjType ObjType { get; set; }
        public string TableName { get; set; }
        public string ChildTableList { get; set; }
        public SAPbobsCOM.BoYesNoEnum CanFind{ get; set; }
        public string FindColumns { get; set; }
        public SAPbobsCOM.BoYesNoEnum ManageSeries{ get; set; }
        public SAPbobsCOM.BoYesNoEnum CanCancel{ get; set; }
        public SAPbobsCOM.BoYesNoEnum CanClose{ get; set; }
        public SAPbobsCOM.BoYesNoEnum CanDelete{ get; set; }
        public SAPbobsCOM.BoYesNoEnum Log{ get; set; }
        public string LogName{ get; set; }
        public SAPbobsCOM.BoYesNoEnum CanYearTransfer{ get; set; }
        public SAPbobsCOM.BoYesNoEnum CanCreateDefaultForm{ get; set; }
        public string FormColumns{ get; set; }
        public string ChildFormColumns{ get; set; }
        public SAPbobsCOM.BoYesNoEnum UseUniqueFormType{ get; set; }
        public SAPbobsCOM.BoYesNoEnum EnableEnhancedForm{ get; set; }
        public SAPbobsCOM.BoYesNoEnum AddToMainMenu{ get; set; }
        public int FatherMenuID{ get; set; }
        public int PositionInMenu{ get; set; }
        public string MenuID{ get; set; }
        public string MenuName { get; set; }

        public UDOParm()
        {
            ObjType = BoUDOObjType.boud_MasterData;
            ChildTableList = "";
            CanFind = BoYesNoEnum.tNO;
            FindColumns = "";
            ManageSeries = BoYesNoEnum.tNO;
            CanCancel = BoYesNoEnum.tNO;
            CanClose = BoYesNoEnum.tNO;
            CanDelete = BoYesNoEnum.tNO;
            Log = BoYesNoEnum.tNO;
            LogName = "";
            CanYearTransfer = BoYesNoEnum.tNO;
            CanCreateDefaultForm = BoYesNoEnum.tNO;
            FormColumns = "";
            ChildFormColumns = "";
            UseUniqueFormType = BoYesNoEnum.tNO;
            EnableEnhancedForm = BoYesNoEnum.tNO;
            AddToMainMenu = BoYesNoEnum.tNO;
            FatherMenuID = 0;
            PositionInMenu = 0;
            MenuID = "";
            MenuName = "";
        }
    }


    public class UDFParm
    {
        public string TableName { get; set; }
        public string FieldName { get; set; }
        public string FieldDescription { get; set; }
        public SAPbobsCOM.BoFieldTypes FieldType { get; set; }
        public int Size { get; set; }
        public string DefaultValue { get; set; }
        public Boolean Mandatory { get; set; }
        public SAPbobsCOM.BoFldSubTypes SubType { get; set; }
        public string ValidValues { get; set; }
        public string LinkedTable { get; set; }

        public UDFParm()
        {
            Mandatory = false;
            FieldType = BoFieldTypes.db_Alpha;
            SubType = BoFldSubTypes.st_None;
            Size = 50;
            DefaultValue = "";
            ValidValues = "";
            LinkedTable = "";
        }
    }

    public class AddOn
    {
        public string AddOnName { get; set; }
        public List<MenuItem> AddOnMenuItemParams = new List<MenuItem>();
        public List<UserPermissionItem> AddOnUserPermissionItems = new List<UserPermissionItem>();
        public AppConfig AddOnConfig = new AppConfig();
        public Boolean LoginAsSuperUser = false;
        
        public static SAPbouiCOM.Application ApplicationInstance;
        public static SAPbobsCOM.Company CompanyInstance;
        public static SAPbobsCOM.CompanyService CompanyServiceInstance;
        public static SAPbobsCOM.AdminInfo AdminInfoInstance;
        public static SAPbouiCOM.ProgressBar ProgressBarInstance;
        public static AddOnSysParm AddOnSysParms = new AddOnSysParm();


        public string StartApplication()
        {
            SAPbouiCOM.EventFilters eventFilters = new SAPbouiCOM.EventFilters();
            return StartApplication(eventFilters, true);
        }

        public string StartApplication(SAPbouiCOM.EventFilters eventFilters)
        {
            return StartApplication(eventFilters, true);
        }

        public string StartApplication(Boolean licenseRequired)
        {
            SAPbouiCOM.EventFilters eventFilters = new SAPbouiCOM.EventFilters();
            return StartApplication(eventFilters, licenseRequired);
        }

        public string StartApplication(SAPbouiCOM.EventFilters eventFilters, Boolean licenseRequired)
        {
            try
            {
                //SAPbobsCOM.Users oUsers = null;

                #region Establish UIAPI Connection
                SAPbouiCOM.SboGuiApi sboGuiApi = new SAPbouiCOM.SboGuiApi();
                string sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));

                //// Fast Track SBOi AddOn License Key
                if(licenseRequired) sboGuiApi.AddonIdentifier = "56455230354241534953303030363030303439303A4C30353436383833333837BE2D8E3EA0DBD35826EF326077F8A12A43680561";
                sboGuiApi.Connect(sConnectionString);
                #endregion

                // Get an instantialized application object
                ApplicationInstance = sboGuiApi.GetApplication(-1);

                // Apply Event Filters
                if (eventFilters.Count > 0) ApplicationInstance.SetFilter(eventFilters);

                ApplicationInstance.StatusBar.SetText(AddOnName + " establishing DIAPI connection...", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning);
                #region Establish DIAPI Connection
                CompanyInstance = new SAPbobsCOM.Company();
                CompanyInstance = (SAPbobsCOM.Company)ApplicationInstance.Company.GetDICompany();
                #endregion

                ApplicationInstance.StatusBar.SetText(AddOnName + " preparing environment...", BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Warning);
                #region Prepare Services
                CompanyServiceInstance = (SAPbobsCOM.CompanyService)CompanyInstance.GetCompanyService();
                AdminInfoInstance = (SAPbobsCOM.AdminInfo)((SAPbobsCOM.CompanyService)CompanyInstance.GetCompanyService()).GetAdminInfo();
                #endregion

                // Add Menu Items defined in MenuItemParams
                AddMenuItems();

                // Get User's Superuser Status
                try
                {
                    //oUsers = (SAPbobsCOM.Users)CompanyInstance.GetBusinessObject(BoObjectTypes.oUsers);
                    //oUsers.GetByKey(CompanyInstance.UserSignature);
                    //LoginAsSuperUser = (oUsers.Superuser == BoYesNoEnum.tYES);
                    //oUsers = null;

                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)CompanyInstance.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //rs.DoQuery(FTS.SAP.Extension.FTExtension.ParseSQLSelect("SELECT [SUPERUSER] FROM [OUSR] WHERE [INTERNAL_K] = " + CompanyInstance.UserSignature));
                    rs.DoQuery_FTExt("SELECT [SUPERUSER] FROM [OUSR] WHERE [INTERNAL_K] = " + CompanyInstance.UserSignature);
                    if (rs.EoF == false) LoginAsSuperUser = (rs.Fields.Item("SUPERUSER").Value.ToString() == "Y");
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); // Must ensure no other objects alive when add UDF, UDT
                    rs = null;
                    GC.Collect();

                    //LoginAsSuperUser = true;
                }
                catch(Exception ex)
                {
                    ApplicationInstance.StatusBar.SetText("Error evaluate superuser right. " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }

                // Add User Permission Tree Items defined in UserPermissionItems
                if (LoginAsSuperUser) AddUserPermissionItems();

                #region Add deligates to events
                ApplicationInstance.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(AppEvent);
                if (AddOnConfig.EnableMenuEvent) ApplicationInstance.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(MenuEvent);
                if (AddOnConfig.EnableItemEvent) ApplicationInstance.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(ItemEvent);
                if (AddOnConfig.EnableFormDataEvent) ApplicationInstance.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(ref FormDataEvent);
                if (AddOnConfig.EnableProgressBarEvent) ApplicationInstance.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(ProgressBarEvent);
                if (AddOnConfig.EnableStatusBarEvent) ApplicationInstance.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(StatusBarEvent);
                if (AddOnConfig.EnableRightClickEvent) ApplicationInstance.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(RightClickEvent);
                if (AddOnConfig.EnablePrintEvent) ApplicationInstance.PrintEvent += new SAPbouiCOM._IApplicationEvents_PrintEventEventHandler(PrintEvent);
                if (AddOnConfig.EnableLayoutKeyEvent) ApplicationInstance.LayoutKeyEvent += new SAPbouiCOM._IApplicationEvents_LayoutKeyEventEventHandler(LayoutKeyEvent);
                #endregion

                // Display status
                ApplicationInstance.StatusBar.SetText(AddOnName + " started.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                return "";
            }
            catch (Exception ex)
            {
                // DLL cannot return message
                //System.Windows.Forms.MessageBox.Show(ex.Message);
                //System.Environment.Exit(0);
                return ex.Message;
            }
        }

        public virtual void AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    System.Environment.Exit(0);
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    System.Environment.Exit(0);
                    break;

                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    AddMenuItems();
                    break;
            }
        }

        public virtual void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //if (!pVal.BeforeAction) MenuEvent.processMenuEvent(ref pVal);
        }

        public virtual void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //ItemEvent.processItemEvent(FormUID, ref pVal, ref BubbleEvent);
        }

        public virtual void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //FormDataEvent.process_FormDataEvent(ref BusinessObjectInfo,ref BubbleEvent);
        }

        public virtual void ProgressBarEvent(ref SAPbouiCOM.ProgressBarEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public virtual void PrintEvent(ref SAPbouiCOM.PrintEventInfo printEventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public virtual void LayoutKeyEvent(ref SAPbouiCOM.LayoutKeyInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public virtual void StatusBarEvent(string Text, SAPbouiCOM.BoStatusBarMessageType MessageType)
        {
            //SBO_Application.MessageBox(@"Status bar event with message: """ + Text + @""" has been sent", 1, "Ok", "", "");
        }

        public virtual void RightClickEvent(ref SAPbouiCOM.ContextMenuInfo EventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public void AddMenuItems()
        {
            SAPbouiCOM.MenuItem fatherMenuItem;
            SAPbouiCOM.MenuCreationParams menuCreationParams = (SAPbouiCOM.MenuCreationParams)ApplicationInstance.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

            foreach (MenuItem menuItem in AddOnMenuItemParams)
            {
                // If menu already exists, remove it
                if (ApplicationInstance.Menus.Exists(menuItem.UniqueID))
                {
                    try { ApplicationInstance.Menus.RemoveEx(menuItem.UniqueID); }
                    catch { }
                }

                menuCreationParams.UniqueID = menuItem.UniqueID;
                menuCreationParams.String = menuItem.MenuName;
                menuCreationParams.Enabled = menuItem.Enabled;
                menuCreationParams.Position = menuItem.Position;
                menuCreationParams.Type = menuItem.MenuType;
                if(menuItem.FatherMenuID == "")
                {
                    fatherMenuItem = ApplicationInstance.Menus.Item("43520");
                }
                else
                {
                    fatherMenuItem = ApplicationInstance.Menus.Item(menuItem.FatherMenuID);
                }

                try
                {
                    fatherMenuItem.SubMenus.AddEx(menuCreationParams);
                }
                catch (Exception ex)
                {
                    ApplicationInstance.MessageBox("Error adding Menu: " + ex.Message, 1, "&Ok", "", ""); 
                    break; 
                } 
            }
        }

        public void AddUserPermissionItems()
        {
            SAPbobsCOM.UserPermissionTree userPermissionTree = (SAPbobsCOM.UserPermissionTree) CompanyInstance.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree);

            foreach (UserPermissionItem permissionItem in AddOnUserPermissionItems)
            {
                try
                {
                    if (!userPermissionTree.GetByKey(permissionItem.PermissionID))
                    {
                        userPermissionTree.PermissionID = permissionItem.PermissionID;
                        userPermissionTree.Name = permissionItem.PermissionName;
                        userPermissionTree.Options = permissionItem.PermissionOptions;
                        if (permissionItem.ParentID != "")
                        {
                            userPermissionTree.ParentID = permissionItem.ParentID;
                        }
                        if (permissionItem.FormType != "")
                        {
                            userPermissionTree.UserPermissionForms.FormType = permissionItem.FormType;
                        }
                        if (userPermissionTree.Add() != 0)
                        {
                            ApplicationInstance.MessageBox("Error modifying General Authorization: " + CompanyInstance.GetLastErrorDescription(), 1, "&Ok", "", "");
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    ApplicationInstance.MessageBox("Error modifying General Authorization: " + ex.Message, 1, "&Ok", "", "");
                    break;
                } // Catch error in case menu already exists
            }

            System.Runtime.InteropServices.Marshal.ReleaseComObject(userPermissionTree);
            userPermissionTree = null;
            GC.Collect();
        }

        static public bool AddUDT(string tableName, string tableDescription, SAPbobsCOM.BoUTBTableType tableType)
        {
            /// <summary>
            /// Create User Defined Table
            /// </summary>
            /// <param name="tableName">User Defined Table Name start with @</param>
            /// <param name="tableDescription">UDT Description</param>
            /// <param name="tableType">UDT Type</param>
            /// <returns></returns>
            /// 
            SAPbobsCOM.UserTablesMD oUserTableMD = ((SAPbobsCOM.UserTablesMD)(CompanyInstance.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));

            try
            {
                int IRetCode = 0;
                if (!oUserTableMD.GetByKey(tableName))
                {
                    ApplicationInstance.StatusBar.SetText("Creating UDT [" + tableName + "]... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    oUserTableMD.TableName = tableName;
                    oUserTableMD.TableDescription = tableDescription;
                    oUserTableMD.TableType = tableType;
                    // Must ensure no other objects alive before call Add()
                    IRetCode = oUserTableMD.Add();
                    ApplicationInstance.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
                    if (IRetCode != 0)
                    {
                        ApplicationInstance.StatusBar.SetText("Error creating UDT [" + tableName + "]: " + CompanyInstance.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTableMD);
                        oUserTableMD = null;
                        GC.Collect();
                        return false;
                    }

                    ApplicationInstance.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTableMD);
                oUserTableMD = null;
                GC.Collect();
                return true;
            }
            catch (Exception err)
            {
                ApplicationInstance.StatusBar.SetText("Error creating UDT [" + tableName + "]: " + err.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTableMD);
                oUserTableMD = null;
                GC.Collect();
                return false;
            }
        }


        static public Boolean AddUDF(UDFParm udfParm)
        {
            return AddUDF(udfParm.TableName, udfParm.FieldName, udfParm.FieldDescription, udfParm.FieldType, udfParm.Size, udfParm.DefaultValue, udfParm.Mandatory, udfParm.SubType, udfParm.ValidValues, udfParm.LinkedTable);
        }

        static public Boolean AddUDF(string tableName, string fieldName, string fieldDescription, SAPbobsCOM.BoFieldTypes fieldType, int size, string defaultValue, Boolean mandatory, SAPbobsCOM.BoFldSubTypes subType, string validValues, string linkedTable)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)(CompanyInstance.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields));

            try
            {
                int IRetCode = 0, IvalidValues = 0, recordCount = 0;

                SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)CompanyInstance.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //oRec.DoQuery(FTExtension.ParseSQLSelect("SELECT [AliasID] FROM [CUFD] WHERE [TableID] = '" + tableName + "' AND [AliasID] = '" + fieldName + "'"));
                oRec.DoQuery_FTExt("SELECT [AliasID] FROM [CUFD] WHERE [TableID] = '" + tableName + "' AND [AliasID] = '" + fieldName + "'");
                recordCount = oRec.RecordCount;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec); // Must ensure no other objects alive when add UDF, UDT
                oRec = null;
                GC.Collect();

                if (recordCount == 0) // UDF Not Exists
                {
                    ApplicationInstance.StatusBar.SetText("Creating UDF [" + tableName + "].[" + fieldName + "]... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    oUserFieldsMD.TableName = tableName.Replace("@", "");
                    oUserFieldsMD.Name = fieldName;
                    oUserFieldsMD.Description = fieldDescription;
                    oUserFieldsMD.Type = fieldType;
                    if (subType != SAPbobsCOM.BoFldSubTypes.st_None) { oUserFieldsMD.SubType = subType; }
                    if (fieldType != SAPbobsCOM.BoFieldTypes.db_Numeric && size > 0) oUserFieldsMD.Size = size;
                    if (fieldType == SAPbobsCOM.BoFieldTypes.db_Numeric && size > 0) { oUserFieldsMD.EditSize = size; }
                    if (defaultValue != "")
                    {
                        oUserFieldsMD.DefaultValue = defaultValue;
                    }
                    if (linkedTable != "") oUserFieldsMD.LinkedTable = linkedTable;

                    if (validValues != "")
                    {
                        foreach (string value in validValues.Split('|'))
                        {
                            IvalidValues++;
                            string[] parm = value.Split(':');
                            if (IvalidValues != 1) oUserFieldsMD.ValidValues.Add();
                            oUserFieldsMD.ValidValues.SetCurrentLine(IvalidValues - 1);
                            oUserFieldsMD.ValidValues.Value = parm[0];
                            oUserFieldsMD.ValidValues.Description = parm[1];
                        }
                    }
                    if (mandatory) oUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;

                    // Must ensure no other objects alive before call Add()
                    IRetCode = oUserFieldsMD.Add();
                    if (IRetCode != 0)
                    {
                        ApplicationInstance.StatusBar.SetText("Error creating UDF [" + tableName + "].[" + fieldName + "]: " + CompanyInstance.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                        oUserFieldsMD = null;
                        GC.Collect();
                        return false;
                    }

                    ApplicationInstance.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();
                return true;
            }
            catch (Exception err)
            {
                ApplicationInstance.StatusBar.SetText("Error creating UDF [" + tableName + "].[" + fieldName + "]: " + err.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();
                return false;
            }
        }


        static public Boolean UpdateUDFDescription(string tableName, string fieldName, string fieldDescription)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)(CompanyInstance.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields));

            try
            {
                int fieldID = 0;

                SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)CompanyInstance.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //oRec.DoQuery(FTExtension.ParseSQLSelect("SELECT [FieldID] FROM [CUFD] WHERE [TableID] ='" + tableName + "' AND [AliasID] = '" + fieldName + "'"));
                oRec.DoQuery_FTExt("SELECT [FieldID] FROM [CUFD] WHERE [TableID] ='" + tableName + "' AND [AliasID] = '" + fieldName + "'");
                if (oRec.EoF != true) fieldID = Int32.Parse(oRec.Fields.Item("FieldID").Value.ToString());
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec); // Must ensure no other objects alive when add UDF, UDT
                oRec = null;
                GC.Collect();

                if (fieldID > 0) // UDF Exists
                {

                    if (oUserFieldsMD.GetByKey(tableName.Replace("@", ""), fieldID))
                    {
                        ApplicationInstance.StatusBar.SetText("Updating UDF [" + tableName + "].[" + fieldName + "]... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        oUserFieldsMD.Description = fieldDescription;
                        if (oUserFieldsMD.Update() != 0) throw new Exception(CompanyInstance.GetLastErrorDescription());

                        ApplicationInstance.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();
                return true;
            }
            catch (Exception err)
            {
                ApplicationInstance.StatusBar.SetText("Error creating UDF [" + tableName + "].[" + fieldName + "]: " + err.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
                oUserFieldsMD = null;
                GC.Collect();
                return false;
            }
        }

        

        static public Boolean AddUDO(UDOParm udoParm)
        {
            return AddUDO(udoParm.UdoName, udoParm.UdoDescription, udoParm.ObjType, udoParm.TableName, udoParm.ChildTableList, udoParm.CanFind, udoParm.FindColumns, udoParm.ManageSeries, udoParm.CanCancel, udoParm.CanClose, udoParm.CanDelete, udoParm.Log, udoParm.LogName, udoParm.CanYearTransfer, udoParm.CanCreateDefaultForm, udoParm.FormColumns, udoParm.ChildFormColumns, udoParm.UseUniqueFormType, udoParm.EnableEnhancedForm, udoParm.AddToMainMenu, udoParm.FatherMenuID, udoParm.PositionInMenu, udoParm.MenuID, udoParm.MenuName);
        }

        static public Boolean AddUDO(string udoName, string udoDescription, SAPbobsCOM.BoUDOObjType objType, string tableName, string childTableList, SAPbobsCOM.BoYesNoEnum canFind, string findColumns, SAPbobsCOM.BoYesNoEnum manageSeries, SAPbobsCOM.BoYesNoEnum canCancel, SAPbobsCOM.BoYesNoEnum canClose, SAPbobsCOM.BoYesNoEnum canDelete, SAPbobsCOM.BoYesNoEnum log, string logName, SAPbobsCOM.BoYesNoEnum canYearTransfer, SAPbobsCOM.BoYesNoEnum canCreateDefaultForm, string formColumns, string childFormColumns, SAPbobsCOM.BoYesNoEnum useUniqueFormType)
        {
            return AddUDO(udoName, udoDescription, objType, tableName, childTableList, canFind, findColumns, manageSeries, canCancel, canClose, canDelete, log, logName, canYearTransfer, canCreateDefaultForm, formColumns, childFormColumns, useUniqueFormType, BoYesNoEnum.tNO, BoYesNoEnum.tNO, 0, 0, "", "");
        }

        static public Boolean AddUDO(string udoName, string udoDescription, SAPbobsCOM.BoUDOObjType objType, string tableName, string childTableList, SAPbobsCOM.BoYesNoEnum canFind, string findColumns, SAPbobsCOM.BoYesNoEnum manageSeries, SAPbobsCOM.BoYesNoEnum canCancel, SAPbobsCOM.BoYesNoEnum canClose, SAPbobsCOM.BoYesNoEnum canDelete, SAPbobsCOM.BoYesNoEnum log, string logName, SAPbobsCOM.BoYesNoEnum canYearTransfer, SAPbobsCOM.BoYesNoEnum canCreateDefaultForm, string formColumns, string childFormColumns, SAPbobsCOM.BoYesNoEnum useUniqueFormType, SAPbobsCOM.BoYesNoEnum enableEnhancedForm, SAPbobsCOM.BoYesNoEnum addToMainMenu, int fatherMenuID, int positionInMenu, string menuID, string menuName)
        {           
            int IRetCode = 0, Iindex = 0, childCnt = 0;
            Boolean rtn = true;
            string[] column;

            SAPbobsCOM.UserObjectsMD oUserObjectMD;
            oUserObjectMD = (SAPbobsCOM.UserObjectsMD)CompanyInstance.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

            try
            {
                if (!oUserObjectMD.GetByKey(udoName))
                {
                    ApplicationInstance.StatusBar.SetText("Creating UDO [" + udoName + "]...", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);

                    oUserObjectMD.Code = udoName;
                    oUserObjectMD.Name = udoDescription;
                    oUserObjectMD.ObjectType = objType;
                    oUserObjectMD.TableName = tableName;
                    oUserObjectMD.CanCancel = canCancel;
                    oUserObjectMD.CanClose = canClose;
                    oUserObjectMD.CanDelete = canDelete;
                    oUserObjectMD.CanFind = canFind;// SAPbobsCOM.BoYesNoEnum.tYES;
                    oUserObjectMD.CanLog = log;
                    oUserObjectMD.LogTableName = logName;
                    oUserObjectMD.CanYearTransfer = canYearTransfer;// SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.ExtensionName = "";
                    oUserObjectMD.CanCreateDefaultForm = canCreateDefaultForm;// SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.UseUniqueFormType = useUniqueFormType;

                    if (childTableList != "")
                    {
                        foreach (string childTable in childTableList.Split('|'))
                        {
                            Iindex++;
                            if (Iindex > 1) oUserObjectMD.ChildTables.Add();
                            oUserObjectMD.ChildTables.SetCurrentLine(Iindex - 1);

                            column = childTable.Split(':');
                            oUserObjectMD.ChildTables.TableName = column[0];
                            if (column.GetUpperBound(0) > 0) oUserObjectMD.ChildTables.LogTableName = column[1];
                        }
                    }
                    if (manageSeries == SAPbobsCOM.BoYesNoEnum.tYES && objType == SAPbobsCOM.BoUDOObjType.boud_Document)
                        oUserObjectMD.ManageSeries = manageSeries;
                    else oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;

                    Iindex = 0;
                    if (findColumns != "")
                    {
                        foreach (string colName in findColumns.Split('|'))
                        {
                            column = colName.Split(':');
                            oUserObjectMD.FindColumns.ColumnAlias = column[0];
                            if (column.GetUpperBound(0) > 0) oUserObjectMD.FindColumns.ColumnDescription = column[1];
                            oUserObjectMD.FindColumns.Add();
                        }
                    }

                    Iindex = 0;
                    if (formColumns != "")
                    {
                        foreach (string colName in formColumns.Split('|'))
                        {
                            column = colName.Split(':');
                            oUserObjectMD.FormColumns.FormColumnAlias = column[0];
                            if (column.GetUpperBound(0) > 0) oUserObjectMD.FormColumns.FormColumnDescription = column[1];
                            if (CompanyInstance.Version >= 902001) oUserObjectMD.FormColumns.SetEditable_FTExt(BoYesNoEnum.tYES); //oUserObjectMD.FormColumns.Editable = BoYesNoEnum.tYES;

                            oUserObjectMD.FormColumns.Add();
                        }
                    }

                    childCnt = 0;
                    Iindex = 0;
                    if (childFormColumns != "")
                    {
                        foreach (string colNameByForm in childFormColumns.Split(','))
                        {
                            childCnt++;
                            foreach (string colName in colNameByForm.Split('|'))
                            {
                                oUserObjectMD.FormColumns.SonNumber = childCnt;

                                column = colName.Split(':');
                                oUserObjectMD.FormColumns.FormColumnAlias = column[0];
                                if (column.GetUpperBound(0) > 0) oUserObjectMD.FormColumns.FormColumnDescription = column[1];
                                if (CompanyInstance.Version >= 902001) oUserObjectMD.FormColumns.SetEditable_FTExt(BoYesNoEnum.tYES); //oUserObjectMD.FormColumns.Editable = BoYesNoEnum.tYES;

                                oUserObjectMD.FormColumns.Add();
                            }
                        }
                    }

                    if (CompanyInstance.Version >= 902001)
                    {
                        oUserObjectMD.SetEnableEnhancedForm_FTExt(enableEnhancedForm);
                        oUserObjectMD.AddFormToMenu_FTExt(addToMainMenu, fatherMenuID, positionInMenu, menuID, menuName);
                    }

                    IRetCode = oUserObjectMD.Add();
                    if (IRetCode != 0)
                    {
                        ApplicationInstance.StatusBar.SetText("Error creating UDO [" + udoName + "]: " + CompanyInstance.GetLastErrorDescription(), BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        rtn = false;
                    }

                    ApplicationInstance.StatusBar.SetText("", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_None);
                }
            }
            catch (Exception ex)
            {
                ApplicationInstance.StatusBar.SetText("Error creating UDO [" + udoName + "]: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                rtn = false;
            }

            // Following code is required by SAP SDK Best Practice // oUserObjectMD = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
            return rtn;
        }

        static public bool AddForm(string xmlPath, out SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.FormCreationParams oFormCreationPackage = (SAPbouiCOM.FormCreationParams)ApplicationInstance.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);

                System.Xml.XmlDocument oXMLDocument = new System.Xml.XmlDocument();
                string addOnDirectory = System.Windows.Forms.Application.StartupPath;
                oXMLDocument.Load(addOnDirectory + xmlPath);

                oFormCreationPackage.UniqueID = "FT_" + DateTime.Now.ToString("yyMMddHHmmssff");
                oFormCreationPackage.XmlData = oXMLDocument.InnerXml;     // Load form from xml 
                oForm = ApplicationInstance.Forms.AddEx(oFormCreationPackage);

                oFormCreationPackage = null;
                return true;
            }
            catch (Exception err)
            {
                ApplicationInstance.StatusBar.SetText(err.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                oForm = null;
                return false;
            }
        }


        static public bool UpdateFormByXML(string xmlPath, string formUID)
        {
            try
            {
                string sXPath = "Application//forms//action//form//@uid";
                string sPath = System.Environment.CurrentDirectory + "\\";

                System.Xml.XmlDocument oXMLDocument = new System.Xml.XmlDocument();
                string addOnDirectory = System.Windows.Forms.Application.StartupPath;
                oXMLDocument.Load(addOnDirectory + xmlPath);

                System.Xml.XmlNode xNode = oXMLDocument.SelectSingleNode(sXPath);
                xNode.InnerText = formUID;

                string sXML = oXMLDocument.InnerXml.ToString();
                ApplicationInstance.LoadBatchActions(ref sXML);
                return true;
            }
            catch (Exception err)
            {
                ApplicationInstance.StatusBar.SetText(err.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                return false;
            }
        }


        static public string ShowOpenFileDialog(string title, string filter, string initialDirectory, string fileName, bool multiSelect)
        {
            string selectedFileName = "";

            FTS.Common.OpenFileDialogEx oOpenFileDialog = new FTS.Common.OpenFileDialogEx();
            oOpenFileDialog.Title = title;
            oOpenFileDialog.Filter = filter;
            oOpenFileDialog.InitialDirectory = initialDirectory; //Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            oOpenFileDialog.FileName = fileName;
            oOpenFileDialog.MultiSelect = multiSelect;

            Thread openFileDialogThread = new Thread(new ThreadStart(oOpenFileDialog.GetFileName));
            openFileDialogThread.SetApartmentState(ApartmentState.STA);

            try
            {
                openFileDialogThread.Start();
                while (!openFileDialogThread.IsAlive) ;
                // Wait for thread to get started 
                Thread.Sleep(1);
                // Wait a sec more 
                openFileDialogThread.Join();
                // Wait for thread to end 
                selectedFileName = oOpenFileDialog.FileName;
            }
            catch (Exception ex)
            {
                ApplicationInstance.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            openFileDialogThread = null;
            oOpenFileDialog = null;

            return selectedFileName;
        }

        static public string ShowFolderBrowserDialog(string title, string filePath, bool showNewDirectoryButton)
        {
            string selectedFileName = "";

            FTS.Common.FolderBrowserDialogEx oFolderBrowserDialog = new FTS.Common.FolderBrowserDialogEx();
            oFolderBrowserDialog.Description = title;
            oFolderBrowserDialog.SelectedPath = filePath;
            oFolderBrowserDialog.ShowNewFolderButton = showNewDirectoryButton;

            Thread openFileDialogThread = new Thread(new ThreadStart(oFolderBrowserDialog.GetFileName));
            openFileDialogThread.SetApartmentState(ApartmentState.STA);

            try
            {
                openFileDialogThread.Start();
                while (!openFileDialogThread.IsAlive) ;
                // Wait for thread to get started 
                Thread.Sleep(1);
                // Wait a sec more 
                openFileDialogThread.Join();
                // Wait for thread to end 
                selectedFileName = oFolderBrowserDialog.SelectedPath;
            }
            catch (Exception ex)
            {
                ApplicationInstance.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            openFileDialogThread = null;
            oFolderBrowserDialog = null;

            return selectedFileName;
        }

        static public string ShowSaveFileDialog(string title, string filter, string initialDirectory, string fileName)
        {
            string selectedFileName = "";

            FTS.Common.SaveFileDialogEx oSaveFileDialog = new FTS.Common.SaveFileDialogEx();
            oSaveFileDialog.Title = title;
            oSaveFileDialog.Filter = filter;
            oSaveFileDialog.InitialDirectory = initialDirectory; //Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            oSaveFileDialog.FileName = fileName;

            Thread openFileDialogThread = new Thread(new ThreadStart(oSaveFileDialog.GetFileName));
            openFileDialogThread.SetApartmentState(ApartmentState.STA);

            try
            {
                openFileDialogThread.Start();
                while (!openFileDialogThread.IsAlive) ;
                // Wait for thread to get started 
                Thread.Sleep(1);
                // Wait a sec more 
                openFileDialogThread.Join();
                // Wait for thread to end 
                selectedFileName = oSaveFileDialog.FileName;
            }
            catch (Exception ex)
            {
                ApplicationInstance.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            openFileDialogThread = null;
            oSaveFileDialog = null;

            return selectedFileName;
        }

        public void StopProgressBar()
        {
            try { ProgressBarInstance.Stop(); }
            catch { }
        }

        static public int CompareDatabaseAddOnVersion(AddOnSysParm sysParm)
        {
            // < 0 Database Outdated
            // > 0 AddOn Outdated
            // = 0 Same Version

            //SAPbobsCOM.Recordset rs = null;
            Assembly assembly;
            System.Diagnostics.FileVersionInfo fileVersionInfo = null;
            //string dbVersion = "";
            int result = 0;

            try
            {
                // Get Add-On File version
                assembly = Assembly.GetExecutingAssembly();
                fileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
                sysParm.SystemVersion = fileVersionInfo.FileVersion;

                result = String.Compare(sysParm.Version, fileVersionInfo.FileVersion);

                if (result > 0)
                {
                    ApplicationInstance.StatusBar.SetText("Add-On requires update. Add-On Version: " + fileVersionInfo.FileVersion + " Database Version: " + sysParm.Version + ".", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else if (result < 0)
                {
                    ApplicationInstance.StatusBar.SetText("Database requires update. Add-On Version: " + fileVersionInfo.FileVersion + " Database Version: " + sysParm.Version + ".", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
                }
                else
                {
                    ApplicationInstance.StatusBar.SetText("Database check successful. Add-On Version: " + fileVersionInfo.FileVersion + ".", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
                }

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Error when determine Add-On Version", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                ApplicationInstance.StatusBar.SetText("Database check failed. Add-On terminated.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                System.Environment.Exit(0);
            }


            return result;
        }

        static public AddOnSysParm GetAddOnSysParm(string addOn)
        {
            SAPbobsCOM.Recordset rs = null;
            AddOnSysParm sysParm = null;

            try
            {
                sysParm = new AddOnSysParm();

                // Get current Database version
                rs = (SAPbobsCOM.Recordset)AddOn.CompanyInstance.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                try
                {
                    //rs.DoQuery(FTExtension.ParseSQLSelect("SELECT [Version], [Parm1], [Parm2], [Parm3], [AddOnKey] FROM [FTS_SYSPARM] WHERE [AddOn] = '" + addOn + "'"));
                    rs.DoQuery_FTExt("SELECT [Version], [Parm1], [Parm2], [Parm3], [AddOnKey] FROM [FTS_SYSPARM] WHERE [AddOn] = '" + addOn + "'");
                    if (rs.EoF == false)
                    {
                        sysParm.Version = rs.Fields.Item("Version").Value.ToString();
                        sysParm.Parm1 = rs.Fields.Item("Parm1").Value.ToString();
                        sysParm.Parm2 = rs.Fields.Item("Parm2").Value.ToString();
                        sysParm.Parm3 = rs.Fields.Item("Parm3").Value.ToString();
                        sysParm.AddOnKey = rs.Fields.Item("AddOnKey").Value.ToString();
                    }
                    else // AddOn not found, insert entry
                    {
                        //rs.DoQuery(FTExtension.ParseSQLSelect("INSERT INTO [FTS_SYSPARM] ([AddOn], [Version]) VALUES ('" + addOn + "', '')"));
                        rs.DoQuery_FTExt("INSERT INTO [FTS_SYSPARM] ([AddOn], [Version]) VALUES ('" + addOn + "', '')");
                    }
                }
                catch
                {
                    // Error reading table, reset the table

                    if (FTS.SAP.Extension.FTExtension.IsHANA())//(AddOn.CompanyInstance.DbServerType == BoDataServerTypes.dst_HANADB)
                    {
                        //rs.DoQuery("DECLARE TBLEXIST INT = 0; " +
                        //    "SELECT COUNT(*) INTO TBLEXIST FROM \"PUBLIC\".\"M_TABLES\" WHERE \"SCHEMA_NAME\" = CURRENT_SCHEMA AND \"TABLE_NAME\" = 'FTGST_tb_EXCLUDE'; " +
                        //    "IF (:TBLEXIST <= 0) THEN " +
                        //    "   CREATE TABLE \"FTS_SYSPARM\" (\"AddOn\" NVARCHAR(50), \"Version\" NVARCHAR(30), \"Parm1\" NVARCHAR(100), \"Parm2\" NVARCHAR(100), \"Parm3\" NVARCHAR(100), \"AddOnKey\" NVARCHAR(100)); " +
                        //    "END IF;");
                        //rs.DoQuery("DROP TABLE \"FTS_SYSPARM\"; ");
                        rs.DoQuery("EXECUTE IMMEDIATE 'CREATE TABLE \"FTS_SYSPARM\" (\"AddOn\" NVARCHAR(50), \"Version\" NVARCHAR(30), \"Parm1\" NVARCHAR(100), \"Parm2\" NVARCHAR(100), \"Parm3\" NVARCHAR(100), \"AddOnKey\" NVARCHAR(100))'; ");
                    }
                    else
                    {
                        rs.DoQuery("IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.Tables WHERE TABLE_NAME = 'FTS_SYSPARM') DROP TABLE FTS_SYSPARM");
                        rs.DoQuery("DECLARE @Statement NVARCHAR(2000); SET @Statement = 'CREATE TABLE FTS_SYSPARM (AddOn NVARCHAR(50), Version NVARCHAR(30), Parm1 NVARCHAR(100), Parm2 NVARCHAR(100), Parm3 NVARCHAR(100), AddOnKey NVARCHAR(100))'; EXECUTE  sp_executesql @Statement");
                    }

                    //rs.DoQuery(FTExtension.ParseSQLSelect("INSERT INTO [FTS_SYSPARM] ([AddOn], [Version]) VALUES ('" + addOn + "', '')"));
                    rs.DoQuery_FTExt("INSERT INTO [FTS_SYSPARM] ([AddOn], [Version]) VALUES ('" + addOn + "', '')");
                }
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); // Must ensure no other objects alive when add UDF, UDT
                rs = null;
                GC.Collect();

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "Error when determine Add-On Version", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                ApplicationInstance.StatusBar.SetText("Database check failed. Add-On terminated.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                System.Environment.Exit(0);
            }


            return sysParm;
        }

        static public void SetAddOnSysParm(string addOn, string column, string value)
        {
            SAPbobsCOM.Recordset rs = null;
            rs = (SAPbobsCOM.Recordset)AddOn.CompanyInstance.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //rs.DoQuery(FTExtension.ParseSQLSelect("UPDATE [FTS_SYSPARM] SET [" + column + "] = '" + value + "' WHERE [AddOn] = '" + addOn + "'"));
            rs.DoQuery("UPDATE [FTS_SYSPARM] SET [" + column + "] = '" + value + "' WHERE [AddOn] = '" + addOn + "'");
            rs = null;
        }

 

    }

}
