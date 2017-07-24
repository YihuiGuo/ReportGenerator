using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;
using System.Data;
using ReportGenerator.Helper;
using System.Configuration;
using ReportGenerator.Manager;
using ReportGenerator.Provider;

namespace ReportGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var reportMgr = new ReportManager(new LocalConfigurationProvider());
            reportMgr.LoadTargetSheet();

            Console.Read();
        }

    }
}
