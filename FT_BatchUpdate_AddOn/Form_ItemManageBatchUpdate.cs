using FTS.SAP;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FT_BatchUpdate_AddOn
{
    class Form_ItemManageBatchUpdate
    {
        internal static void ItemEvent_before(string formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
        }

        internal static void ItemEvent_after(string formUID, ref ItemEvent pVal)
        {
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.IChooseFromListEvent oCFLEvento = null;
            SAPbouiCOM.DataTable oCFLDataTable = null;

            switch (pVal.EventType)
            {
                case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                    oForm = SAPAddOn.ApplicationInstance.Forms.Item(formUID);
                    switch (pVal.ItemUID)
                    {
                        case "btnRet":
                            LoadData(oForm);
                            break;
                        case "btnUpdate":
                            UpdateData(oForm);
                            break;
                    }
                    break;
                case BoEventTypes.et_CHOOSE_FROM_LIST:
                    oForm = SAPAddOn.ApplicationInstance.Forms.Item(formUID);

                    switch (pVal.ItemUID)
                    {
                        case "txtItemFrm":
                        case "txtItemTo":
                            oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                            oCFLDataTable = (SAPbouiCOM.DataTable)oCFLEvento.SelectedObjects;
                            if (oCFLDataTable != null && oCFLDataTable.IsEmpty == false)
                            {
                                try
                                {
                                    for (int i = 0; i < oCFLDataTable.Rows.Count; i++)
                                    {
                                        if (pVal.ItemUID == "txtItemFrm")
                                            oForm.DataSources.UserDataSources.Item("ItemFrm").Value = oCFLDataTable.GetValue("ItemCode", i).ToString().Trim();
                                        if (pVal.ItemUID == "txtItemTo")
                                            oForm.DataSources.UserDataSources.Item("ItemTo").Value = oCFLDataTable.GetValue("ItemCode", i).ToString().Trim();

                                    }

                                }
                                catch (Exception ex)
                                {
                                    AddOn.ApplicationInstance.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                }
                            }
                            break;
                    }

                    //Dave way
                    //AutoFillChooseFromList(oForm, "txtItemFrm", pVal);
                    break;
            }

        }

        internal static void FormDataEvent_before(ref BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
        }

        internal static void FormDataEvent_after(ref BusinessObjectInfo businessObjectInfo)
        {
        }

        internal static void MenuEvent_before(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
        }

        public static void LoadData(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.DataTable dt1 = null;
            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.EditTextColumn oEditCol;

            try
            {
                dt1 = oForm.DataSources.DataTables.Item("dt1");
                oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid1").Specific;
                oGrid.DataTable = dt1;
                string status = oForm.DataSources.UserDataSources.Item("Status").Value;
                string itemFrm = oForm.DataSources.UserDataSources.Item("ItemFrm").Value;
                string itemTo = oForm.DataSources.UserDataSources.Item("ItemTo").Value;
                string query = "SELECT 'N' AS [Select], AbsEntry, ItemCode AS [Item Code], ItemName AS [Description], DistNumber AS [Batch No.] FROM OBTN WHERE 0 = 0 ";

                // Filtering
                if (status != "") query += "AND Status = '" + status + "' "; else throw new Exception("Please select Status to filter!");
                if (itemFrm != "") query += "AND ItemCode >= '" + itemFrm + "' ";
                if (itemTo != "") query += "AND ItemCode <= '" + itemTo + "' ";


                // Sort list at last
                query += "ORDER BY ItemCode, AbsEntry ";
                dt1.ExecuteQuery(query);

                foreach (SAPbouiCOM.GridColumn col in oGrid.Columns)
                {
                    oEditCol = (SAPbouiCOM.EditTextColumn)col;
                    col.Editable = false;
                    switch (col.UniqueID)
                    {
                        case "Select":
                            col.Editable = true;
                            col.Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                            break;
                        case "AbsEntry":
                            oEditCol.LinkedObjectType = "10000044";
                            col.Visible = true;
                            col.TitleObject.Caption = "";
                            col.Width = 15;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                SAPAddOn.ApplicationInstance.StatusBar.SetSystemMessage(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        public static void UpdateData(SAPbouiCOM.Form oForm)
        {
            SAPbobsCOM.BatchNumberDetailsService oBatch;
            SAPbobsCOM.BatchNumberDetailParams oBatchParam;
            SAPbobsCOM.BatchNumberDetail oBatchDetails;
            SAPbouiCOM.DataTable dt1;
            SAPbouiCOM.ProgressBar oProgBar;

            oBatch = SAPAddOn.CompanyServiceInstance.GetBusinessService(SAPbobsCOM.ServiceTypes.BatchNumberDetailsService);
            oBatchParam = oBatch.GetDataInterface(SAPbobsCOM.BatchNumberDetailsServiceDataInterfaces.bndsBatchNumberDetailParams);

            dt1 = oForm.DataSources.DataTables.Item("dt1");

            string status = oForm.DataSources.UserDataSources.Item("Status").Value;
            string updatestatus = oForm.DataSources.UserDataSources.Item("UStatus").Value;


            oProgBar = SAPAddOn.ApplicationInstance.StatusBar.CreateProgressBar("Updating...", dt1.Rows.Count, false);
            try
            {
                if (updatestatus == "")
                    throw new Exception("Please select a status to update!");

                if (status == updatestatus)
                    throw new Exception("Please select a different status to update!");

                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    if (dt1.GetValue("Select", i).ToString() != "Y") continue;
                    int absentry = (int)dt1.GetValue("AbsEntry", i);

                    oBatchParam.DocEntry = absentry;

                    oBatchDetails = oBatch.Get(oBatchParam);

                    if (updatestatus == "0")
                        oBatchDetails.Status = BoDefaultBatchStatus.dbs_Released;
                    else if (updatestatus == "1")
                        oBatchDetails.Status = BoDefaultBatchStatus.dbs_NotAccessible;
                    else if (updatestatus == "2")
                        oBatchDetails.Status = BoDefaultBatchStatus.dbs_Locked;
                }

                LoadData(oForm);

                oProgBar.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar);
                oProgBar = null;
                SAPAddOn.ApplicationInstance.StatusBar.SetText("Action completed successfully", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SAPAddOn.ApplicationInstance.StatusBar.SetSystemMessage(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
            if (oProgBar != null)
            {
                oProgBar.Stop();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar);
                oProgBar = null;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBatchParam);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oBatch);
            oBatchParam = null;
            oBatch = null;
        }

        private static void AutoFillChooseFromList(SAPbouiCOM.Form oForm, string currentId, ItemEvent pVal)
        {
            var dt = ((SAPbouiCOM.IChooseFromListEvent)pVal).SelectedObjects;

            if (dt == null) return;

            var item = oForm.Items.Item(currentId);

            try
            {
                string code;
                string tablename;
                string alias;

                switch (item.Type)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                        var txt = item.Specific as SAPbouiCOM.EditText;
                        code = dt.GetValue(txt.ChooseFromListAlias, 0).ToString();
                        tablename = txt.DataBind.TableName;
                        alias = txt.DataBind.Alias;
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        var cbox = item.Specific as SAPbouiCOM.ComboBox;
                        code = dt.GetValue(0, 0).ToString();
                        tablename = cbox.DataBind.TableName;
                        alias = cbox.DataBind.Alias;
                        break;
                    default:
                        return;
                }

                if (alias == null) return;

                if (tablename == null)
                {
                    if (alias == String.Empty || !oForm.HasUserSource(alias)) return;

                    oForm.SetUserSourceValue(alias, code);
                }
                else if (oForm.HasDataSource(tablename))
                {
                    oForm.DataSources.DBDataSources.Item(tablename).SetValue(alias, 0, code);
                }
                else if (oForm.HasDataTable(tablename))
                {
                    oForm.DataSources.DataTables.Item(tablename).SetValue(alias, 0, code);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dt);
                dt = null;
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(item);
                item = null;
                GC.Collect();
            }
        }
    }
}
