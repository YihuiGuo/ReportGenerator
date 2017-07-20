using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.Data;
using ReportGenerator.Helper;

namespace ReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Collection<Dictionary<string, string>> datacontainer = new Collection<Dictionary<string, string>>();
            List<String> columnNames = new List<string>();
            System.Data.DataTable data;

            ExcelReader reader = new ExcelReader();
            data = reader.ReadFile(AppDomain.CurrentDomain.BaseDirectory + @"Pass Rate by Area.csv");

            var targetrows = data.Select($"{data.GetExcelHeader(0)} ='OfficeVSO' and {data.GetExcelHeader(6)} = 'WinProj'");
            foreach (var row in targetrows)
            {
                var datarow = new Dictionary<string, string>();
                foreach (DataColumn col in data.Columns)
                {
                    datarow.Add(col.ColumnName, row[col].ToString());
                }
                datacontainer.Add(datarow);
            }



            Application xlapp = new Application();
            Workbook targetworkbook = xlapp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + @"Project Desktop Automation Status.xlsx");

            var targetsheet = targetworkbook.Sheets[1];

            xlapp.Visible = true;

            Range copyRange = targetsheet.Range["G:I"];
            Range insertRange = targetsheet.Range["G:G"];

            insertRange.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, copyRange.Copy());

            targetsheet.Cells[1, 10] = "Last Week";
            targetsheet.Cells[1, 13] = "2 Weeks Ago";
            targetsheet.Cells[1, 16] = "3 Weeks Ago";
            targetsheet.Cells[1, 19] = string.Empty;

            targetsheet.Cells[2, 7] = datacontainer.First().Where(p => p.Key == data.GetExcelHeader(5)).Select(p => p.Value).FirstOrDefault();


            string currentTopLevel = "";
            int totalCount = 0;
            for (int targetrow = 5; targetrow <= 62; targetrow++)
            {
                string category = targetsheet.Cells[targetrow, 3].Text?.ToString();
                string subCategory = targetsheet.Cells[targetrow, 4].Text.ToString();

                if (category != "" && subCategory == "" && category != "Project Desktop")
                {
                    currentTopLevel = category;
                    var row = datacontainer.FirstOrDefault(p => p["C3"] == currentTopLevel);
                    totalCount += row == null ? 0 : int.Parse(row["TotalTests4"]);
                    targetsheet.Cells[targetrow, 5] = row == null ? "0" : row["TotalTests4"];
                    targetsheet.Cells[targetrow, 8] = row == null ? "0" : row["PassRate3"];
                    Console.WriteLine($"{currentTopLevel}---Num:{row["TotalTests4"]}---Rate:{row["PassRate3"]}");
                }
                else if (category == "" && subCategory != "")
                {
                    var subRow = datacontainer.FirstOrDefault(p => p["C3"] == currentTopLevel && p["C4"] == subCategory);
                    targetsheet.Cells[targetrow, 5] = subRow == null ? "0" : subRow["TotalTests5"];
                    targetsheet.Cells[targetrow, 8] = subRow == null ? "0" : subRow["PassRate4"];
                    Console.WriteLine($"{currentTopLevel}---{subCategory}---Num:{targetsheet.Cells[targetrow, 5].Text.ToString()}-- -Rate:{targetsheet.Cells[targetrow, 8].Text.ToString()}");

                }
                targetworkbook.Save();
            }
            targetsheet.Cells[5, 5] = totalCount;
            targetsheet.Cells[5, 8] = targetsheet.Cells[2, 7];

            xlapp.Quit();
            Console.Read();
        }

        public static string GetCellContentByIndex(Worksheet worksheet, int rowId, int colId)
        {
            return worksheet.Cells[rowId, colId].Text.ToString();
        }
    }
}
