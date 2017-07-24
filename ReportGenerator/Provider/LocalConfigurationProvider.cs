using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Provider
{
    public class LocalConfigurationProvider:IConfigurationProvider
    {
        public string GetConfigurationValue(string key)
        {
            return ConfigurationManager.AppSettings[key] ?? "";
        } 
    }
}
