using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FT_BatchUpdate_AddOn
{
    static class FormExtensions
    {
        public static string GetUserSourceValue(this SAPbouiCOM.Form oForm, string source, bool sqlformat = false)
        {
            var datasource = oForm.DataSources.UserDataSources.Item(source);
            string value = datasource.ValueEx;

            if (!sqlformat) return value;

            switch (datasource.DataType)
            {
                case SAPbouiCOM.BoDataType.dt_DATE:
                case SAPbouiCOM.BoDataType.dt_SHORT_TEXT:
                case SAPbouiCOM.BoDataType.dt_LONG_TEXT:
                    return $"'{ value }'";
                default:
                    return value;
            }
        }

        public static void SetUserSourceValue(this SAPbouiCOM.Form oForm, string source, string value)
        {
            var datasource = oForm.DataSources.UserDataSources.Item(source);
            datasource.ValueEx = value;
        }

        public static bool HasUserSource(this SAPbouiCOM.Form oForm, string source)
        {
            if (oForm.DataSources.UserDataSources.Count == 0) return false;

            return oForm.DataSources.UserDataSources
                .OfType<SAPbouiCOM.UserDataSource>()
                .Where(uds => uds.UID == source)
                .Any();
        }

        public static bool HasDataSource(this SAPbouiCOM.Form oForm, string source)
        {
            if (oForm.DataSources.DBDataSources.Count == 0) return false;

            return oForm.DataSources.DBDataSources
                .OfType<SAPbouiCOM.DBDataSource>()
                .Where(ds => ds.TableName == source)
                .Any();
        }

        public static bool HasDataTable(this SAPbouiCOM.Form oForm, string source)
        {
            if (oForm.DataSources.DataTables.Count == 0) return false;

            for (int i = 0; i < oForm.DataSources.DataTables.Count; ++i)
            {
                if (oForm.DataSources.DataTables.Item(i).UniqueID == source) return true;
            }

            return false;
        }

        public static bool HasItem(this SAPbouiCOM.Form oForm, string uniqueid)
        {
            if (oForm.Items.Count == 0) return false;

            return oForm.Items
                .OfType<SAPbouiCOM.Item>()
                .Where(itm => itm.UniqueID == uniqueid)
                .Any();
        }
    }
}
