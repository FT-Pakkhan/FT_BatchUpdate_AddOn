using System;
using System.Collections.Generic;

namespace FTS.SAP
{
    public class FormHandler
    {
        static public SAPbouiCOM.ChooseFromList AddChooseFromList(ref SAPbouiCOM.Form oForm, SAPbouiCOM.BoLinkedObject chooseFromObject, string cflUID, bool multiSelection = false)
        {
            /*
             *  2019-08-29  Change oForm to ref
             * */
            SAPbouiCOM.ChooseFromListCollection oCFLs = oForm.ChooseFromLists;
            SAPbouiCOM.ChooseFromListCreationParams oCFLparms = null;
            SAPbouiCOM.ChooseFromList oCFL = null;

            try
            {
                // Create Choose From List - Item
                oCFLparms = (SAPbouiCOM.ChooseFromListCreationParams)AddOn.ApplicationInstance.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLparms.MultiSelection = multiSelection;
                oCFLparms.ObjectType = chooseFromObject.GetHashCode().ToString();
                oCFLparms.UniqueID = cflUID;
                oCFL = oCFLs.Add(oCFLparms);
            }
            catch (Exception ex)
            {
                oCFLparms = null;
                oCFL = null;
                oCFLs = null;
                throw new Exception(ex.Message); }

            oCFLparms = null;
            oCFL = null;
            oCFLs = null;
            return oCFL;
        }

        static public SAPbouiCOM.ChooseFromList AddChooseFromList(ref SAPbouiCOM.Form oForm, string chooseFromObjectName, string cflUID, bool multiSelection = false)
        {
            /*
             *  2019-08-29  Change oForm to ref
             * */
            SAPbouiCOM.ChooseFromListCollection oCFLs = oForm.ChooseFromLists;
            SAPbouiCOM.ChooseFromListCreationParams oCFLparms = null;
            SAPbouiCOM.ChooseFromList oCFL = null;

            try
            {
                // Create Choose From List - Item
                oCFLparms = (SAPbouiCOM.ChooseFromListCreationParams)AddOn.ApplicationInstance.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLparms.MultiSelection = multiSelection;
                oCFLparms.ObjectType = chooseFromObjectName;
                oCFLparms.UniqueID = cflUID;
                oCFL = oCFLs.Add(oCFLparms);
            }
            catch (Exception ex)
            {
                oCFLparms = null;
                oCFL = null;
                oCFLs = null;
                throw new Exception(ex.Message); }

            oCFLparms = null;
            oCFL = null;
            oCFLs = null;
            return oCFL;
        }

        static public void AddRow(ref SAPbouiCOM.Form oForm, string matrixName, string datasourceName, Boolean setFocusAfter)
        {
            // Add row in matrix
            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixName).Specific;
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item(datasourceName);

                oMatrix.FlushToDataSource();
                ds.InsertRecord(ds.Size);
                //if (ds.GetValue("DocEntry", 0) == "") ds.RemoveRecord(0);
                //ds.SetValue("DocEntry", ds.Size - 1, ds.Size.ToString()); // To handle DocEntry NULL error
                oMatrix.LoadFromDataSource();


                // NEW FEATURE IN SAP 8.81
                if (setFocusAfter) oMatrix.SetCellFocus(oMatrix.RowCount, 1);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

                oForm.Freeze(false);
                oMatrix = null;
                ds = null;
            }
            catch (Exception ex)
            {
                oForm.Freeze(false);
                AddOn.ApplicationInstance.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        static public void ShowActionResult(List<ActionResult> actionResults)
        {
            try
            {
                SAPbouiCOM.Form oForm;
                if (AddOn.AddForm("\\Form_ActionResult.xml", out oForm))
                {
                    SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("ACTIONRESULT");

                    for (int i = 0; i < actionResults.Count; i++)
                    {
                        oDataTable.SetValue(0, i, i + 1);
                        oDataTable.SetValue(1, i, actionResults[i].Status);
                        oDataTable.SetValue(2, i, actionResults[i].Key);
                        oDataTable.SetValue(3, i, actionResults[i].Reason);
                        oDataTable.Rows.Add();
                    }
                    SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid1").Specific;
                    oGrid.AutoResizeColumns();

                    oForm.Visible = true;
                    oForm = null;
                    oGrid = null;
                    oDataTable = null;
                }
            }
            catch (Exception ex)
            {
                AddOn.ApplicationInstance.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }


        public static Boolean TrySetFormEditText(ref SAPbouiCOM.Form oForm, string fieldID, string fieldName, string data, bool silent = false)
        {
            // Use this function to assign value to Edit Text in Form
            // It will prompt error if the field is protected or does not exists / visible
            try
            {
                ((SAPbouiCOM.EditText)oForm.Items.Item(fieldID).Specific).Value = data;
            }
            catch (Exception ex)
            {
                if (silent == false) AddOn.ApplicationInstance.StatusBar.SetText("Fail to set value to '" + fieldName + "' in form. (FieldID: " + fieldID + ")" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            return true;
        }

        public static Boolean TrySetMatrixEditText(ref SAPbouiCOM.Matrix oMatrix, string columnID, string columnName, int row, string data, bool silent = false)
        {
            // Use this function to assign value to Edit Text in Matrix
            // It will prompt error if the column is protected or does not exists / visible
            try
            {
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item(columnID).Cells.Item(row).Specific).String = data;
            }
            catch (Exception ex)
            {
                if(silent == false) AddOn.ApplicationInstance.StatusBar.SetText("Fail to set value to '" + columnName + "' in matrix. (ColumnID: " + columnID + ")" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            return true;
        }

        public static Boolean TrySetMatrixComboBox(ref SAPbouiCOM.Matrix oMatrix, string columnID, string columnName, int row, string data, SAPbouiCOM.BoSearchKey comboSearchBy, bool silent = false)
        {
            // Use this function to assign value to ComboBox in Matrix
            // It will prompt error if the column is protected or does not exists / visible
            try
            {
                ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item(columnID).Cells.Item(row).Specific).Select(data, comboSearchBy);
            }
            catch (Exception ex)
            {
                if (silent == false) AddOn.ApplicationInstance.StatusBar.SetText("Fail to set value to '" + columnName + "' in matrix. (ColumnID: " + columnID + ")" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            return true;
        }


        //static public Matrix CastToNetMatrix(SAPbouiCOM.Matrix oMatrix)
        //{
        //    Matrix xmlMatrix = null;

        //    try
        //    {
        //        //System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
        //        System.Xml.Serialization.XmlSerializer XmlSerializer = new System.Xml.Serialization.XmlSerializer(typeof(Matrix));

        //        string matrixXML = oMatrix.SerializeAsXML((SAPbouiCOM.BoMatrixXmlSelect.mxs_All));
        //        //xmlDoc.LoadXml(matrixXML);
        //        //string xMLFileName = System.AppDomain.CurrentDomain.BaseDirectory + "Matrix.xml";
        //        //xmlDoc.Save(xMLFileName);

        //        //// Load Matrix from XML file to .Net matrix object 
        //        //System.IO.FileStream myFileStream = new System.IO.FileStream(xMLFileName, System.IO.FileMode.Open);
        //        //xmlMatrix = (Matrix)XmlSerializer.Deserialize(myFileStream);

        //        using (var reader = new System.IO.StringReader(matrixXML))
        //        {
        //            xmlMatrix = (Matrix)XmlSerializer.Deserialize(reader);
        //        }
        //    }
        //    catch { }

        //    return xmlMatrix;
        //}

    }
}
