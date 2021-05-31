using System;

using SAPbouiCOM;

namespace FTS.SAP
{
    /// <summary>
    /// This class convert SAP MessageBox into Dialogbox with single field.
    /// User can customized the field according to data type and field type.
    /// To use the class user need to:
    /// 1. Add message box to ItemPressed Event(et_ITEM_PRESSED) eg: "oEventFilter.AddEx("0"); // MessageBox"
    /// 2. Add message box to FormLoad Event(et_FORM_LOAD) eg: "oEventFilter.AddEx("0"); // MessageBox"
    /// 3. Redirect ItemEvent Before_Action to UserInput Class "case "0": if (UserInput.Active) UserInput.ItemEvent_before(FormUID, ref pVal, out BubbleEvent)"
    /// 4. Show the dialog box eg: "userInput = UserInput.Show("Cancellation Options", "Rejected payments will be cancelled at", "SysDate:Current System Date|DocDate:Original Document Date", BoFormItemTypes.it_COMBO_BOX)"
    /// 
    /// </summary>
    static class UserInput
    {
        static public Boolean Active = false;
        static private string Title { get; set; }
        static private string ValidValues { get; set; } 
        static private string Result { get; set; }
        static private SAPbouiCOM.BoFormItemTypes FormItemType { get; set; }
        static private SAPbouiCOM.BoDataType DataType { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="title">Dialog Box Title, eg: "Cancellation Options"</param>
        /// <param name="message">Dialog Box Message, eg: "Rejected payments will be cancelled at"</param>
        /// <param name="validValues">Valid Value, For ComboBox use "SysDate:Current System Date|DocDate:Original Document Date"</param>
        /// <returns>Selected value</returns>
        static public string Show(string title, string message, string validValues, SAPbouiCOM.BoFormItemTypes formItemtype, SAPbouiCOM.BoDataType dataType = SAPbouiCOM.BoDataType.dt_LONG_TEXT)
        {
            Active = true;
            Result = "";
            Title = title;
            ValidValues = validValues;
            FormItemType = formItemtype;
            DataType = dataType;
            AddOn.ApplicationInstance.MessageBox(message, 1, "OK", "Cancel", "");

            return Result;
        }

        internal static void ItemEvent_before(string FormUID, ref ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.Form oForm = null;
            SAPbouiCOM.EditText oEditText = null;
            SAPbouiCOM.ComboBox oCombo = null;
            SAPbouiCOM.OptionBtn oOptionBtn = null;
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Item oItemRef = null;
            int optionCnt = 0;

            switch (pVal.EventType)
            {
                case BoEventTypes.et_ITEM_PRESSED:
                    try
                    {
                        if (pVal.ItemUID == "1") // OK
                        {
                            UserInput.Active = false;
                            Result = "";

                            oForm = AddOn.ApplicationInstance.Forms.Item(FormUID);
                            switch (FormItemType)
                            {
                                case BoFormItemTypes.it_EDIT:
                                    UserInput.Result = ((SAPbouiCOM.EditText)oForm.Items.Item("result").Specific).Value;
                                    break;

                                case BoFormItemTypes.it_EXTEDIT:
                                    UserInput.Result = ((SAPbouiCOM.EditText)oForm.Items.Item("result").Specific).Value;
                                    break;

                                case BoFormItemTypes.it_COMBO_BOX:
                                    UserInput.Result = ((SAPbouiCOM.ComboBox)oForm.Items.Item("result").Specific).Value;
                                    break;

                                case BoFormItemTypes.it_OPTION_BUTTON:
                                    //UserInput.Result = oForm.DataSources.UserDataSources.Item("data").ValueEx; 
                                    UserInput.Result = (((SAPbouiCOM.OptionBtn)oForm.Items.Item("result1").Specific).Selected ? "1" : "2");
                                    break;
                            }
                        }
                        else if (pVal.ItemUID == "2") // Cancel
                        {
                            UserInput.Active = false;
                            Result = "";
                        }
                    }
                    catch (Exception ex)
                    {
                        AddOn.ApplicationInstance.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }
                    break;

                case BoEventTypes.et_FORM_LOAD:
                    try
                    {
                        oForm = AddOn.ApplicationInstance.Forms.Item(FormUID);
                        oForm.Title = UserInput.Title;
                        oForm.DataSources.UserDataSources.Add("data", DataType);// BoDataType.dt_LONG_TEXT);

                        oItemRef = oForm.Items.Item("7");

                        switch (FormItemType)
                        {
                            case BoFormItemTypes.it_EDIT:
                                oItem = oForm.Items.Add("result", BoFormItemTypes.it_EDIT);                        
                                oItem.Top = oItemRef.Top + oItemRef.Height + 10;
                                oItem.Left = oItemRef.Left;
                                if (DataType == BoDataType.dt_LONG_TEXT)
                                {
                                    oItem.Width = oItemRef.Width;
                                }
                                else
                                {
                                    oItem.Width = 120; // Shorter width for number and date
                                }
                                oItem.Height = 18;
                                oItem.DisplayDesc = true;
                                break;

                            case BoFormItemTypes.it_EXTEDIT:
                                oItem = oForm.Items.Add("result", BoFormItemTypes.it_EXTEDIT);                        
                                oItem.Top = oItemRef.Top + oItemRef.Height + 10;
                                oItem.Left = oItemRef.Left;
                                oItem.Width = oItemRef.Width;
                                oItem.Height = 36;
                                oItem.DisplayDesc = true;
                                break;

                            case BoFormItemTypes.it_COMBO_BOX:
                                oItem = oForm.Items.Add("result", BoFormItemTypes.it_COMBO_BOX);                        
                                oItem.Top = oItemRef.Top + oItemRef.Height + 10;
                                oItem.Left = oItemRef.Left;
                                oItem.Width = oItemRef.Width;
                                oItem.Height = 18;
                                oItem.DisplayDesc = true;
                                break;

                            case BoFormItemTypes.it_OPTION_BUTTON:

                                foreach (string value in UserInput.ValidValues.Split('|'))
                                {
                                    string[] parm = value.Split(':');
                                    optionCnt++;
                                    oItem = oForm.Items.Add("result" + optionCnt.ToString(), BoFormItemTypes.it_OPTION_BUTTON);
                                    oItem.Top = oItemRef.Top + oItemRef.Height + 10;
                                    oItem.Left = oItemRef.Left;
                                    oItem.Width = oItemRef.Width;
                                    oItem.Height = 18;
                                    oItem.DisplayDesc = true;

                                    oItemRef = oItem;
                                    oOptionBtn = (SAPbouiCOM.OptionBtn)oItem.Specific;
                                    oOptionBtn.Caption = parm[1];
                                    oOptionBtn.ValOn = parm[0];
                                    oOptionBtn.ValOff = "";

                                    if (optionCnt == 1)
                                    {
                                        oOptionBtn.DataBind.SetBound(true, "", "data");
                                        oForm.DataSources.UserDataSources.Item("data").Value = parm[0];
                                    }
                                    else
                                    {
                                        oOptionBtn.GroupWith("result" + (optionCnt - 1).ToString());
                                        oOptionBtn.DataBind.SetBound(true, "", "data");
                                    }
                                }

                                break;
                        }


                        switch (FormItemType)
                        {
                            case BoFormItemTypes.it_EDIT:
                                oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                                oEditText.DataBind.SetBound(true, "", "data");

                                oForm.DataSources.UserDataSources.Item("data").Value = UserInput.ValidValues;
                                break;

                            case BoFormItemTypes.it_EXTEDIT:
                                oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                                oEditText.DataBind.SetBound(true, "", "data");

                                oForm.DataSources.UserDataSources.Item("data").Value = UserInput.ValidValues;
                                break;

                            case BoFormItemTypes.it_COMBO_BOX:
                                oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                                oCombo.DataBind.SetBound(true, "", "data");

                                foreach (string value in UserInput.ValidValues.Split('|'))
                                {
                                    string[] parm = value.Split(':');
                                    oCombo.ValidValues.Add(parm[0], parm[1]);
                                    if (oCombo.ValidValues.Count == 1) oForm.DataSources.UserDataSources.Item("data").Value = parm[0];
                                }
                                break;
                        }
                        

                    }
                    catch (Exception ex)
                    {
                        AddOn.ApplicationInstance.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                    }

                    break;
            }
        }
    }
}
