using System;
using System.Configuration;

namespace OfficeApp.ConfigSections
{
    public class ExcelFormattersSection : ConfigurationSection
    {
        [ConfigurationProperty("",IsDefaultCollection = true)]
        public ExcelFormattersCollection Formatters => (ExcelFormattersCollection) base[""];
    }
}
