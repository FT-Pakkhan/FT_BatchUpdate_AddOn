using System;
using FTS.SAP;
using SAPbobsCOM;
using SAPbouiCOM;

namespace FT_BatchUpdate_AddOn
{
    class SAPAddOn : FTS.SAP.AddOn
    {
        public SAPAddOn()
        {
            #region Customize AddOn Settings
            this.AddOnName = "Item Management Batch Update Add-On";
            this.AddOnConfig.EnableMenuEvent = true;
            this.AddOnConfig.EnableItemEvent = true;
            this.AddOnConfig.EnableProgressBarEvent = true;
            this.AddOnConfig.EnableFormDataEvent = true;
            #endregion

            #region Customize Menu (Optional)
            this.AddOnMenuItemParams.Add(new FTS.SAP.MenuItem("FTS00IMBU", "Item Management Batch Update", "15872", 2, SAPbouiCOM.BoMenuType.mt_STRING, true, ""));
            #endregion

            #region Define Event Filters (Optional)
            SAPbouiCOM.EventFilters oEventFilters = new SAPbouiCOM.EventFilters();
            SAPbouiCOM.EventFilter oEventFilter;
            oEventFilter = oEventFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
            oEventFilter.AddEx("1292"); // Add Row
            oEventFilter.AddEx("1281"); // Find
            oEventFilter.AddEx("1282"); // Add

            oEventFilter = oEventFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);

            ////Unable to populate ChooseFromList data into System Form
            //oEventFilter.AddEx("1470000200"); // Purchase Request 

            //oEventFilter = oEventFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
            /* --Unable to handle Validate of UDF in System Form
            oEventFilter.AddEx("1470000200"); // Purchase Request
            */

            oEventFilter = oEventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            oEventFilter.AddEx("FTS00IMBU"); // Item Management Batch Update

            oEventFilter = oEventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
            oEventFilter.AddEx("FTS00IMBU"); // Item Management Batch Update

            oEventFilter = oEventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
            oEventFilter.AddEx("FTS00IMBU"); // Item Management Batch Update

            oEventFilter = oEventFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);
            oEventFilter.AddEx("FTS00IMBU"); // Item Management Batch Update

            oEventFilter = oEventFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            oEventFilter.AddEx("FTS00IMBU"); // Item Management Batch Update

            oEventFilter = oEventFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
            oEventFilter.AddEx("FTS00IMBU"); // Item Management Batch Update
            #endregion

            #region Start AddOn
            try
            {
                string errMsg = StartApplication(oEventFilters);
                if (errMsg != "")
                {
                    System.Windows.Forms.MessageBox.Show(errMsg);
                    System.Environment.Exit(0);
                }
            }
            catch (Exception err)
            {
                System.Windows.Forms.MessageBox.Show("Error connecting to SAP Business One: " + err.Message);
                System.Environment.Exit(0);
            }
            #endregion
        }

        public override void MenuEvent(ref MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.ComboBox oComboBox = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.DataTable dt1 = null;
            SAPbouiCOM.ChooseFromList oCFL = null;

            if (pVal.BeforeAction == true)
            {
                switch (pVal.MenuUID)
                {
                    case "FTS00IMBU": // Item Management Batch Update
                        try
                        {
                            if (AddForm("\\ItemManageBatchUpdate.xml", out oForm))
                            {
                                dt1 = oForm.DataSources.DataTables.Add("dt1");

                                oForm.DataSources.UserDataSources.Add("Status", SAPbouiCOM.BoDataType.dt_SHORT_TEXT).Value = "";
                                oForm.DataSources.UserDataSources.Add("UStatus", SAPbouiCOM.BoDataType.dt_SHORT_TEXT).Value = "";
                                oForm.DataSources.UserDataSources.Add("ItemFrm", SAPbouiCOM.BoDataType.dt_SHORT_TEXT).Value = "";
                                oForm.DataSources.UserDataSources.Add("ItemTo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT).Value = "";

                                oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbStatus").Specific;
                                oForm.Items.Item("cmbStatus").DisplayDesc = true;
                                oComboBox.ExpandType = BoExpandType.et_DescriptionOnly;
                                oComboBox.ValidValues.Add("0", "Released");
                                oComboBox.ValidValues.Add("1", "Not Accessible");
                                oComboBox.ValidValues.Add("2", "Locked");
                                oComboBox.DataBind.SetBound(true, "", "Status");

                                oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("cmbUStatus").Specific;
                                oForm.Items.Item("cmbUStatus").DisplayDesc = true;
                                oComboBox.ExpandType = BoExpandType.et_DescriptionOnly;
                                oComboBox.ValidValues.Add("0", "Released");
                                oComboBox.ValidValues.Add("1", "Not Accessible");
                                oComboBox.ValidValues.Add("2", "Locked");
                                oComboBox.DataBind.SetBound(true, "", "UStatus");

                                oCFL = oForm.ChooseFromLists.Item("CFL_OITM1");
                                oCFL.SetConditions(null);
                                SAPbouiCOM.Conditions oConds = oCFL.GetConditions();
                                SAPbouiCOM.Condition oCond = oConds.Add();

                                oCond.Alias = "ManBtchNum";
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCond.CondVal = "Y";
                                oCFL.SetConditions(oConds);

                                oCFL = oForm.ChooseFromLists.Item("CFL_OITM2");
                                oCFL.SetConditions(null);
                                oConds = oCFL.GetConditions();
                                oCond = oConds.Add();

                                oCond.Alias = "ManBtchNum";
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCond.CondVal = "Y";
                                oCFL.SetConditions(oConds);

                                oItem = oForm.Items.Item("txtItemFrm");
                                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
                                oEditText.DataBind.SetBound(true, "", "ItemFrm");
                                oEditText.ChooseFromListUID = "CFL_OITM1";
                                oEditText.ChooseFromListAlias = "ItemCode";

                                oItem = oForm.Items.Item("txtItemTo");
                                oEditText = ((SAPbouiCOM.EditText)(oItem.Specific));
                                oEditText.DataBind.SetBound(true, "", "ItemTo");
                                oEditText.ChooseFromListUID = "CFL_OITM2";
                                oEditText.ChooseFromListAlias = "ItemCode";

                                oForm.Visible = true;
                            }
                            oForm = null;
                        }
                        catch (Exception ex)
                        {
                            AddOn.ApplicationInstance.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                        }
                        break;
                }
            }
            else // After event
            {
            }
        }

        public override void FormDataEvent(ref BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            // Use override keyword to customize Event Handler
            BubbleEvent = true;
            if (BusinessObjectInfo.BeforeAction)
            {
            }
            else
            {
                switch (BusinessObjectInfo.FormTypeEx)
                {
                    case "FTS00IMBU": // Item Management Batch Update
                        Form_ItemManageBatchUpdate.FormDataEvent_after(ref BusinessObjectInfo);
                        break;
                }
            }
        }


        public override void ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            // Use override keyword to customize Event Handler
            BubbleEvent = true;

            if (pVal.Before_Action)
            {
                switch (pVal.FormTypeEx)
                {
                    case "FTS00IMBU": // Item Management Batch Update
                        Form_ItemManageBatchUpdate.ItemEvent_before(FormUID, ref pVal, out BubbleEvent);
                        break;
                }
            }
            else
            {
                switch (pVal.FormTypeEx)
                {
                    case "FTS00IMBU": // Item Management Batch Update
                        Form_ItemManageBatchUpdate.ItemEvent_after(FormUID, ref pVal);
                        break;
                }
            }
        }
    }
}
