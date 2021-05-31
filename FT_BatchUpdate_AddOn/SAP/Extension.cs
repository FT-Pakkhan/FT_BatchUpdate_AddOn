using System;
using SAPbouiCOM;
using SAPbobsCOM;


namespace FTS.SAP.Extension 
{
    public static class FTExtension
    {

        // Handle missing function prior 8.81 PL8
        public static int CancelbyCurrentSystemDate_FTExt(this SAPbobsCOM.Payments oPV)
        {
            return oPV.CancelbyCurrentSystemDate();
        }


        public static Matrix CastToNetMatrix_FTExt(this SAPbouiCOM.Matrix oMatrix)
        {
            Matrix xmlMatrix = null;

            try
            {
                //System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                System.Xml.Serialization.XmlSerializer XmlSerializer = new System.Xml.Serialization.XmlSerializer(typeof(Matrix));

                string matrixXML = oMatrix.SerializeAsXML((SAPbouiCOM.BoMatrixXmlSelect.mxs_All));
                //xmlDoc.LoadXml(matrixXML);
                //string xMLFileName = System.AppDomain.CurrentDomain.BaseDirectory + "Matrix.xml";
                //xmlDoc.Save(xMLFileName);

                //// Load Matrix from XML file to .Net matrix object 
                //System.IO.FileStream myFileStream = new System.IO.FileStream(xMLFileName, System.IO.FileMode.Open);
                //xmlMatrix = (Matrix)XmlSerializer.Deserialize(myFileStream);

                using (var reader = new System.IO.StringReader(matrixXML))
                {
                    xmlMatrix = (Matrix)XmlSerializer.Deserialize(reader);
                }
            }
            catch { }

            return xmlMatrix;
        }


        // Handle missing option prior 9.0
        public static void SetEditable_FTExt(this UserObjectMD_FormColumns oCol, BoYesNoEnum value)
        {
            oCol.Editable = value;
        }

        // Handle missing option prior 9.0
        public static void SetEnableEnhancedForm_FTExt(this UserObjectsMD oUserObjectMD, BoYesNoEnum value)
        {
            // EnableEnhancedForm = tNO for Matrix layout
            oUserObjectMD.EnableEnhancedForm = value;
        }

        // Handle missing option prior 9.0
        public static void AddFormToMenu_FTExt(this UserObjectsMD oUserObjectMD, BoYesNoEnum addToMainMenu, int fatherMenuID, int positionInMenu, string menuID, string menuName)
        {
            if (addToMainMenu == BoYesNoEnum.tYES && fatherMenuID > 0 && menuID != "" && menuName != "")
            {
                // Add UDO to Main Menu
                oUserObjectMD.MenuItem = BoYesNoEnum.tYES;
                oUserObjectMD.FatherMenuID = fatherMenuID;
                oUserObjectMD.Position = positionInMenu;
                oUserObjectMD.MenuUID = menuID;
                oUserObjectMD.MenuCaption = menuName;
            }
            else
            {
                oUserObjectMD.MenuItem = BoYesNoEnum.tNO;
            }
        }



        // Handle missing option prior 9.0
        public static ReportLayoutParams AddReportLayoutToMenu_FTExt(this ReportLayoutsService reportLayoutService, ReportLayout reportLayout, string menuID)
        {
            return reportLayoutService.AddReportLayoutToMenu(reportLayout, menuID);
        }


        // To display queries in status bar for debug
        public static bool IsDebug()
        {
            if (AddOn.AddOnSysParms.Parm3 == "DEBUG") return true;
            return false;
        }

        // Handle missing option prior 9.0
        public static bool IsHANA()
        {
            if (AddOn.CompanyInstance.Version >= 902001) return IsHANA_FTExt();
            return false;
        }

        public static bool IsHANA_FTExt()
        {
            return AddOn.CompanyInstance.DbServerType == BoDataServerTypes.dst_HANADB;
        }

        public static string ParseSQLSelect(string sql)
        {
            string parsedSQL = sql;

            if (FTS.SAP.Extension.FTExtension.IsHANA())
            {
                parsedSQL = sql.Replace("[", "\"").Replace("]", "\"");
                if (parsedSQL.Substring(parsedSQL.Length - 1, 1) != "") parsedSQL = parsedSQL + "";
                parsedSQL = parsedSQL.Replace("ISNULL", "IFNULL");
            }

            //AddOn.ApplicationInstance.StatusBar.SetSystemMessage(parsedSQL);
            return parsedSQL;
        }

        //public static string ParseSQLProc(string procName, string parameters)
        //{
        //    return ParseSQLProc(procName, parameters, false);
        //}

        public static string ParseSQLProc(string procName, string parameters = "", bool addArithAbort = false)
        {
            string parsedSQL = "";

            if (FTS.SAP.Extension.FTExtension.IsHANA())
            {
                parsedSQL = "CALL \"" + procName + "\" (" + parameters + ")";
            }
            else
            {
                parsedSQL = "EXEC \"dbo\".\"" + procName + "\" " + parameters + "";
                if (addArithAbort) parsedSQL = "SET ARITHABORT ON " + parsedSQL;
            }

            return parsedSQL;
        }

        // Handle Recordset.DoQuery to HANA
        public static void DoQuery_FTExt(this Recordset c, string query, string queryType = "SELECT", string parameters = "", bool addArithAbort = false)
        {
            string parsedQuery = "";
            switch (queryType)
            {
                case "PROC":
                    parsedQuery = FTS.SAP.Extension.FTExtension.ParseSQLProc(query, parameters, addArithAbort);
                    break;

                default:
                    parsedQuery = FTS.SAP.Extension.FTExtension.ParseSQLSelect(query);
                    break;
            }

            if (IsDebug()) AddOn.ApplicationInstance.StatusBar.SetText("DoQuery: " + parsedQuery, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            try
            {
                c.DoQuery(parsedQuery);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + " - " + parsedQuery);
            }

        }

        // Handle DataTable.ExecuteQuery to HANA
        public static void ExecuteQuery_FTExt(this DataTable c, string query, string queryType = "SELECT", string parameters = "", bool addArithAbort = false)
        {
            string parsedQuery = "";
            switch (queryType)
            {
                case "PROC":
                    parsedQuery = FTS.SAP.Extension.FTExtension.ParseSQLProc(query, parameters, addArithAbort);
                    break;

                default:
                    parsedQuery = FTS.SAP.Extension.FTExtension.ParseSQLSelect(query);
                    break;
            }

            if(IsDebug()) AddOn.ApplicationInstance.StatusBar.SetText("DoQuery: " + parsedQuery, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            try
            {
                c.ExecuteQuery(parsedQuery);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message + " - " + parsedQuery);
            }

        }


    }
}