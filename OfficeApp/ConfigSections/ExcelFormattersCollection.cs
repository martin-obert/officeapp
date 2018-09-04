using System.Configuration;

namespace OfficeApp.ConfigSections
{
    public class ExcelFormattersCollection : ConfigurationElementCollection
    {
        protected override ConfigurationElement CreateNewElement()
        {
            return new ExcelFormatterElement();
        }

        public new ExcelFormatterElement this[string name] => (ExcelFormatterElement) BaseGet(name);


        public ExcelFormatterElement this[int index] => (ExcelFormatterElement) BaseGet(index);

        public int IndexOf(string name)
        {
            name = name.ToLower();

            for (int idx = 0; idx < base.Count; idx++)
            {
                if (this[idx].Keyword.ToLower() == name)
                    return idx;
            }
            return -1;
        }

        public override ConfigurationElementCollectionType CollectionType =>
            ConfigurationElementCollectionType.BasicMap;

        protected override object GetElementKey(ConfigurationElement element)
        {
            return ((ExcelFormatterElement) element).Keyword;
        }

        protected override string ElementName => "formatter";
    }
}