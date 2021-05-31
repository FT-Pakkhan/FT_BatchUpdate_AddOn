using System;
using SAPbobsCOM;

namespace FTS.SAP
{
    public class UserDefinedObjectHandler
    {
        static private SAPbobsCOM.GeneralService UdoService = null; 
        
        static public Boolean Start(string udoName)
        {
            try
            {
                UdoService = AddOn.CompanyServiceInstance.GetGeneralService(udoName);
            }
            catch
            {
                return false;
            }
            return true;
        }

        static public void Stop()
        {
            UdoService = null;
        }

        static public void AddRecord(string code, string dataInString)
        {
            string[] column;

            if (dataInString != "")
            {
                try
                {
                    SAPbobsCOM.GeneralDataParams udoParams = (SAPbobsCOM.GeneralDataParams)UdoService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    SAPbobsCOM.GeneralData udoData = null;

                    try
                    {
                        udoParams.SetProperty("Code", code);
                        UdoService.GetByParams(udoParams);
                    }
                    catch
                    {
                        if (udoData == null)
                        {
                            udoData = (SAPbobsCOM.GeneralData)UdoService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                            // Loop through data string
                            foreach (string colName in dataInString.Split('|'))
                            {
                                column = colName.Split(':');
                                udoData.SetProperty(column[0], column[1]);
                            }
                            UdoService.Add(udoData);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }
        }

        static public void AddRecord(SAPbobsCOM.GeneralData udoData)
        {
            UdoService.Add(udoData);

            /*
                SAPbobsCOM.GeneralData udoData = (SAPbobsCOM.GeneralData)udoService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                udoData.SetProperty("Code", "AP01In");
                udoData.SetProperty("Name", "AP Bad Debt Recovered");
                udoData.SetProperty("U_Value", "");
                udoService.Add(udoData);
              
             */
        }

        static public Boolean DeleteRecord(string code)
        {
            try
            {
                SAPbobsCOM.GeneralDataParams udoParams = (SAPbobsCOM.GeneralDataParams)UdoService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                udoParams.SetProperty("Code", code);
                UdoService.Delete(udoParams);
            }
            catch { return false; }

            return true;
        }

        static public Boolean DeleteRecord(SAPbobsCOM.GeneralDataParams udoParams)
        {
            try
            {
                UdoService.Delete(udoParams);

                /*
                 
                udoParams.SetProperty("Code", "AP01Out");
                udoService.Delete(udoParams);
                 
                */
            }
            catch { return false; }

            return true;
        }

        static public Boolean UpdateRecord(string code, string property, string value)
        {
            try
            {
                SAPbobsCOM.GeneralDataParams udoParams = (SAPbobsCOM.GeneralDataParams)UdoService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                SAPbobsCOM.GeneralData udoData = null;

                udoParams.SetProperty("Code", code);
                udoData = UdoService.GetByParams(udoParams);
                udoData.SetProperty(property, value);
                UdoService.Update(udoData);
            }
            catch { return false; }

            return true;
        }

    }
}
