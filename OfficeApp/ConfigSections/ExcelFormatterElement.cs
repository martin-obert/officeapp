using System.Configuration;

namespace OfficeApp.ConfigSections
{
    public class ExcelFormatterElement : ConfigurationElement
    {

        [ConfigurationProperty("keyword", IsRequired = true, IsKey = true)]
        public string Keyword
        {
            get => this["keyword"] as string;
            set => this["keyword"] = value;
        }

        [ConfigurationProperty("replacement", IsRequired = true)]
        public string Replacement
        {
            get => this["replacement"] as string;
            set => this["replacement"] = value;
        }
    }
}