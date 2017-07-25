using Microsoft.Office.Interop.Excel;
using ReportGenerator.Helper;
using ReportGenerator.Provider;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Manager
{
    public class ReportManager
    {
        System.Data.DataTable rawData;
        Collection<Dictionary<string, string>> sortedData = new Collection<Dictionary<string, string>>();
        IConfigurationProvider configurationProvider;
        Workbook targetworkbook;
        Application xlapp ;

        public ReportManager(IConfigurationProvider provider)
        {
            configurationProvider = provider;
            rawData = new ExcelReader().ReadFile(configurationProvider.GetConfigurationValue("RawDataPath"));
            sortedData = new Collection<Dictionary<string, string>>();

            int ColumnIndexOfC0, ColumnIndexOfC2;
            int.TryParse(configurationProvider.GetConfigurationValue("C0InRawData"),out ColumnIndexOfC0);
            int.TryParse(configurationProvider.GetConfigurationValue("C2InRawData"), out ColumnIndexOfC2);
            var C0ExpectValue = configurationProvider.GetConfigurationValue("C0FilterValue");
            var C2ExpectValue = configurationProvider.GetConfigurationValue("C2FilterValue");

            var targetrows = rawData.Select($"{rawData.GetExcelHeader(ColumnIndexOfC0)} ='{C0ExpectValue}' and {rawData.GetExcelHeader(ColumnIndexOfC2)} = '{C2ExpectValue}'");
            foreach (var row in targetrows)
            {
                var datarow = new Dictionary<string, string>();
                foreach (DataColumn col in rawData.Columns)
                {
                    datarow.Add(col.ColumnName, row[col].ToString());
                }
                sortedData.Add(datarow);
            }
        }

        
        public void LoadTargetSheet()
        {
            xlapp = new Application();
            Workbook targetworkbook = xlapp.Workbooks.Open(configurationProvider.GetConfigurationValue("TargetDataPath"));
            var targetsheet = targetworkbook.Sheets[1];

            xlapp.Visible = true;

            Range copyRange = targetsheet.Range["G:I"];
            Range insertRange = targetsheet.Range["G:G"];

            insertRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, copyRange.Copy());

            targetsheet.Cells[1, 10] = "Last Week";
            targetsheet.Cells[1, 13] = "2 Weeks Ago";
            targetsheet.Cells[1, 16] = "3 Weeks Ago";
            targetsheet.Cells[1, 19] = string.Empty;

            targetsheet.Cells[2, 7] = sortedData.First().Where(p => p.Key == rawData.GetExcelHeader(5)).Select(p => p.Value).FirstOrDefault();


            string currentTopLevel = "";
            int totalCount = 0;
            for (int targetrow = 5; targetrow <= 62; targetrow++)
            {
                string category = targetsheet.Cells[targetrow, 3].Text?.ToString();
                string subCategory = targetsheet.Cells[targetrow, 4].Text.ToString();

                if (category != "" && subCategory == "" && category != "Project Desktop")
                {
                    currentTopLevel = category;
                    var row = sortedData.FirstOrDefault(p => p["C3"] == currentTopLevel);
                    totalCount += row == null ? 0 : int.Parse(row["TotalTests4"]);
                    targetsheet.Cells[targetrow, 5] = row == null ? "0" : row["TotalTests4"];
                    targetsheet.Cells[targetrow, 8] = row == null ? "0" : row["PassRate3"];

                    Console.WriteLine($"{currentTopLevel}---Num:{row["TotalTests4"]}---Rate:{row["PassRate3"]}");
                }
                else if (category == "" && subCategory != "")
                {
                    var subRow = sortedData.FirstOrDefault(p => p["C3"] == currentTopLevel && p["C4"] == subCategory);
                    targetsheet.Cells[targetrow, 5] = subRow == null ? "0" : subRow["TotalTests5"];
                    targetsheet.Cells[targetrow, 8] = subRow == null ? "0" : subRow["PassRate4"];

                    Console.WriteLine($"{currentTopLevel}---{subCategory}---Num:{targetsheet.Cells[targetrow, 5].Text.ToString()}-- -Rate:{targetsheet.Cells[targetrow, 8].Text.ToString()}");

                }
                targetworkbook.Save();
            }
            targetsheet.Cells[5, 5] = totalCount;
            targetsheet.Cells[5, 8] = targetsheet.Cells[2, 7];

            xlapp.Quit();
        }
    }
}
