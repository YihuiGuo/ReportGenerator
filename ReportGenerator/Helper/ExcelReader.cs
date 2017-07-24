using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Helper
{
    public class ExcelReader
    {
        private string filePath;
        private string fileName;
        private OleDbConnection conn;
        private DataTable readDataTable;
        private string connString;
        private FileType fileType = FileType.noset;

        private void SetFileInfo(string path)
        {
            filePath = path;

            fileName = this.filePath.Remove(0, this.filePath.LastIndexOf("\\") + 1);
            switch (fileName.Split('.')[1])
            {
                case "xls":
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'"; fileType = FileType.xls;
                    break;
                case "xlsx":
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'"; fileType = FileType.xlsx;
                    break;
                case "csv":
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath.Remove(filePath.LastIndexOf("\\") + 1) + ";Extended Properties='Text;FMT=Delimited;HDR=YES;'"; fileType = FileType.csv;
                    break;
            }
        }


        public DataTable ReadFile(string path)
        {
            if (System.IO.File.Exists(path))
            {
                SetFileInfo(path);
                OleDbDataAdapter myCommand = null;
                DataSet ds = null;

                using (conn = new OleDbConnection(connString))
                {
                    conn.Open();

                    DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                    string tableName = fileType == FileType.csv ? fileName : schemaTable.Rows[0][2].ToString().Trim();

                    string strExcel = string.Empty;

                    strExcel = "Select   *   From   [" + tableName + "]";
                    myCommand = new OleDbDataAdapter(strExcel, conn);

                    ds = new DataSet();

                    myCommand.Fill(ds, tableName);

                    readDataTable = ds.Tables[0];

                }
            }
            return readDataTable;
        }

        private enum FileType
        {
            noset,
            xls,
            xlsx,
            csv
        }

    }
}
