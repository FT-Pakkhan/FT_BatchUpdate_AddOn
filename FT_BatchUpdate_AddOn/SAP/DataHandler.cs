using System;
using System.Globalization;
using System.Text;

namespace FTS.SAP
{
    public class DataHandler
    {
        static public string Left(string data, int len)
        {
            return data.Substring(0, (data.Length > len ? len : data.Length));
        }

        static public string Right(string data, int len)
        {
            return data.Substring((data.Length > len ? 0 : data.Length - len), data.Length);
        }

        static public string GetApplicationDateFormat()
        {
            string dateSeparator = "";
            string format = "";

            dateSeparator = AddOn.AdminInfoInstance.DateSeparator;

            switch (AddOn.AdminInfoInstance.DateTemplate) // SAP Date Format
            {
                case SAPbobsCOM.BoDateTemplate.dt_CCYYMMDD:
                    format = "yyyy" + dateSeparator + "MM" + dateSeparator + "dd";
                    break;

                case SAPbobsCOM.BoDateTemplate.dt_DDMMCCYY:
                    format = "dd" + dateSeparator + "MM" + dateSeparator + "yyyy";
                    break;

                case SAPbobsCOM.BoDateTemplate.dt_DDMMYY:
                    format = "dd" + dateSeparator + "MM" + dateSeparator + "yy";
                    break;

                case SAPbobsCOM.BoDateTemplate.dt_DDMonthYYYY:
                    format = "dd" + dateSeparator + "MMMM" + dateSeparator + "yyyy";
                    break;

                case SAPbobsCOM.BoDateTemplate.dt_MMDDCCYY:
                    format = "MM" + dateSeparator + "dd" + dateSeparator + "yyyy";
                    break;

                case SAPbobsCOM.BoDateTemplate.dt_MMDDYY:
                    format = "MM" + dateSeparator + "dd" + dateSeparator + "yy";
                    break;

            }

            return format;
        }

        static public DateTime GetDateTimeFromString(string dateTimeInString, string dateTimeFormat)
        {
            DateTime result = DateTime.MinValue;
            try
            {
                //// This script cause error when parsing dd/MM/yyyy string
                //DateTime.TryParseExact(dateTimeInString, dateTimeFormat, DateTimeFormatInfo.CurrentInfo, DateTimeStyles.AdjustToUniversal, out result);
                DateTime.TryParseExact(dateTimeInString, dateTimeFormat, DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None, out result);
            }
            catch {}
            return result;
        }

        static public DateTime GetDateTimeFromStringInSystemCulture(string dateTimeInString)
        {
            /////
            // Convert datetime string get by GetValue() from oForm which formatted into Windows Date Format eg: M/d/yyyy
            // into Universal Date Format yyyyMMdd
            // 
            //  2011/10/31  oDataTable.GetValue now return MM/dd/yyyy format
            /////

            DateTime result = DateTime.MinValue;
            try
            {
                DateTime.TryParse(dateTimeInString, System.Threading.Thread.CurrentThread.CurrentCulture.DateTimeFormat, DateTimeStyles.AdjustToUniversal, out result);
            }
            catch { }
            return result;
        }

        static public bool IsNumeric(string val, System.Globalization.NumberStyles NumberStyle)
        {
            Double result;
            return Double.TryParse(val, NumberStyle,
                System.Globalization.CultureInfo.CurrentCulture, out result);
        }

        static public decimal CastStringToDecimal(string data)
        {
            try
            {
                if (string.IsNullOrEmpty(data))
                {
                    return 0;
                }

                return decimal.Parse(data);
            }
            catch { return 0; }


        }

        static public string GetSQLSafeString(string data)
        {
            if (string.IsNullOrEmpty(data) == true)
            {
                return "";
            }
            return data.Replace("'", "''");
        }
    }

}
