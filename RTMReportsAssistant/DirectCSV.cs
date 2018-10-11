using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace RTMReportsAssistant
{
   public class DirectCSV
    {
       public string ExportToCSV(DataTable table)
       {
           var result = new StringBuilder();
           for (int i = 0; i < table.Columns.Count; i++)
           {
               result.Append(table.Columns[i].ColumnName);
               result.Append(i == table.Columns.Count - 1 ? "\n" : ",");
           }

           foreach (DataRow row in table.Rows)
           {
               for (int i = 0; i < table.Columns.Count; i++)
               {
                   result.Append(row[i].ToString());
                   result.Append(i == table.Columns.Count - 1 ? "\n" : ",");
               }
           }

           return result.ToString();
       }
    }
}
