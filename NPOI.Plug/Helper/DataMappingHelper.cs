using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Specialized;
using System.Reflection;
using System.Data;

namespace NPOI.Plug.Helper
{
    public class DataMappingHelper
    {
        public static Dictionary<string, string> NameValues2Dictionary(NameValueCollection NVS)
        {
            if (!NVS.HasKeys()) return default(Dictionary<string, string>);
            Dictionary<string, string> entityDic = new Dictionary<string, string>();
            foreach (string key in NVS.AllKeys)
            {
                entityDic.Add(key, NVS[key]);
            }
            return entityDic;
        }

        public static T NameValues2Entity<T>(NameValueCollection NVS)
        {
            if (!NVS.HasKeys()) return default(T);
            T entity = Activator.CreateInstance<T>();
            PropertyInfo[] attrs = entity.GetType().GetProperties();
            foreach (PropertyInfo p in attrs)
            {
                foreach (string key in NVS.AllKeys)
                {
                    if (string.Compare(p.Name, key, true) == 0)
                    {
                        p.SetValue(entity, Convert.ChangeType(NVS[key], p.PropertyType), null);
                    }
                }
            }
            return entity;
        }

        public static List<T> DataTable2Entities<T>(DataTable table)
        {
            if (null == table || table.Rows.Count <= 0) return default(List<T>);
            List<T> list = new List<T>();
            List<string> keys = new List<string>();
            foreach (DataColumn c in table.Columns)
            {
                keys.Add(c.ColumnName.ToLower());
            }
            for (int i = 0; i < table.Rows.Count; i++)
            {
                T entity = Activator.CreateInstance<T>();
                PropertyInfo[] attrs = entity.GetType().GetProperties();
                foreach (PropertyInfo p in attrs)
                {
                    if (keys.Contains(p.Name.ToLower()))
                    {
                        if (!DBNull.Value.Equals(table.Rows[i][p.Name]))
                        {
                            p.SetValue(entity, Convert.ChangeType(table.Rows[i][p.Name], p.PropertyType), null);
                        }
                    }
                }
                list.Add(entity);
            }
            return list;
        }
    }
}
