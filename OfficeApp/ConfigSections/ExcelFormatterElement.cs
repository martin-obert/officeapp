using System.Configuration;

namespace OfficeApp.ConfigSections
{
    public class ExcelFormatterElement : ConfigurationElement
    {
        [ConfigurationProperty("id", IsRequired = true, IsKey = true)]
        public string Id
        {
            get => this["id"] as string;
            set => this["id"] = value;
        }

        [ConfigurationProperty("keyword")]
        public string Keyword
        {
            get => this["keyword"] as string;
            set => this["keyword"] = value;
        }

        [ConfigurationProperty("replacement")]
        public string Replacement
        {
            get => this["replacement"] as string;
            set => this["replacement"] = value;
        }

        [ConfigurationProperty("numberFormat")]
        public string NumberFormat
        {
            get => this["numberFormat"] as string;
            set => this["numberFormat"] = value;
        }

        [ConfigurationProperty("range")]
        public string Range
        {
            get => this["range"] as string;
            set => this["range"] = value;
        }
    }
}