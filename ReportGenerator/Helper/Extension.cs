using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Helper
{
    public static class Extension
    {
        public static string GetExcelHeader(this DataTable table, int indexInExcel)
        {
            return table.Columns[indexInExcel].ColumnName;
        }
    }
}
