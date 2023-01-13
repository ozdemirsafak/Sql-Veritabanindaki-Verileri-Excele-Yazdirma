using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;
using System.Text;

namespace Excel
{
    public static class Helpers
    { /// Gelen Listeyi t olarak alıp datatable nesnesine çeviririyor.
        public static DataTable ToDataTable<T>(List<T> items) 
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            //Tüm özellikleri al
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                ////Sütun adlarını Özellik adları olarak ayarlama
                dataTable.Columns.Add(prop.Name); 
            }
            foreach (T item in items) 
            {
                
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                 
                    values[i] = Props[i].GetValue(item, null); 
                }
                dataTable.Rows.Add(values);
            }
            //buraya bir kesme noktası koyun ve datatable'ı kontrol edin
            return dataTable;
        }
    }
}
