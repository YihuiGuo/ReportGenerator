using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportGenerator.Provider
{
    public interface IConfigurationProvider
    {
        string GetConfigurationValue(string key);
    }
}
