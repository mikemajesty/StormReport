using System;

namespace StormReport
{
    public class ExcelFileNameAttribute : Attribute
    {
        public string Name { get; set; }

        public ExcelFileNameAttribute(string name, ExcelExtensionEnum excelEnum = ExcelExtensionEnum.XLS, bool useDateTimeToAdditionalName = false)
        {
            Name = string.Format("{0}.{1}", name, GetEnumText(excelEnum));
            if (useDateTimeToAdditionalName)
            {
                Name += "-" + DateTime.Now;
            }
        }

        private static string GetEnumText(ExcelExtensionEnum excelEnum)
        {
            return excelEnum.ToString().ToLower();
        }
    }
}
