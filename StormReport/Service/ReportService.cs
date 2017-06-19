using StormReport.BuildTable;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace StormReport.Service
{
    public static class ReportService
    {
        public static void AddColumnGroup(IEnumerable<PropertyInfo> properties, TableFactory table)
        {
            var columnGroup = properties.Select(c => c.GetCustomAttributes(typeof(ExportableColumnGroupAttribute), false).FirstOrDefault()).ToList();

            if (columnGroup.FirstOrDefault(c => ((ExportableColumnGroupAttribute)c) != null) == null)
                return;

            table.AddRow();
            foreach (var prop in columnGroup.GroupBy(c => ((ExportableColumnGroupAttribute)c) == null ? null : ((ExportableColumnGroupAttribute)c).Description))
            {
                var style = properties.Select(c => c.GetCustomAttributes(typeof(ExportableColumnGroupAttribute), false)).SelectMany(c => c.Where(p => ((ExportableColumnGroupAttribute)p) != null && ((ExportableColumnGroupAttribute)p).Styles.Count() > 0)).Where(p => ((ExportableColumnGroupAttribute)p).Description == prop.Key);
                table.AddColumnGroup(prop.Key ?? "", prop.Count(), style.FirstOrDefault() != null ? ((ExportableColumnGroupAttribute)style.FirstOrDefault()).Styles : new string[] { });
            }
            table.EndRow();
        }

        public static void AddTableColumnCell<T>(IEnumerable<PropertyInfo> properties, TableFactory table, IList<T> listItems)
        {
            foreach (var row in listItems.Select(o => new { Properties = properties.Select(g => g), Value = o }).ToList())
            {
                table.AddRow();
                foreach (PropertyInfo cell in row.Properties)
                {
                    var cellValue = row.Properties.Select(g => cell.GetValue(row.Value)).FirstOrDefault();
                    var styleProperty = ((ExportableColumnContentStyleAttribute)cell.GetCustomAttributes(typeof(ExportableColumnContentStyleAttribute), false).FirstOrDefault());
                    var addtionalText = ((ExportableAddtionalTextAttribute)cell.GetCustomAttributes(typeof(ExportableAddtionalTextAttribute), false).FirstOrDefault());

                    table.AddColumnContentText(cellValue, styleProperty != null ? styleProperty.Styles : new string[] { }, addtionalText);
                }
                table.EndRow();
            }
        }

        public static void AddTableColumnHeader(IEnumerable<PropertyInfo> properties, TableFactory table)
        {
            table.AddRow();
            foreach (var headerCell in properties)
            {
                var headerText = ((ExportableColumnHeaderNameAttribute)headerCell.GetCustomAttributes(typeof(ExportableColumnHeaderNameAttribute), false).FirstOrDefault()).Description;
                var styleProperty = ((ExportableColumnHeaderStyleAttribute)headerCell.GetCustomAttributes(typeof(ExportableColumnHeaderStyleAttribute), false).FirstOrDefault());

                table.AddColumnTextHeader(headerText, styleProperty != null ? styleProperty.Styles : new string[] { });
            }
            table.EndRow();
        }

        public static IEnumerable<PropertyInfo> GetObjectPropertyInfo<T>()
        {
            return typeof(T).GetProperties().Where(f => ((ExportableColumnHeaderNameAttribute)f.GetCustomAttributes(typeof(ExportableColumnHeaderNameAttribute), true).FirstOrDefault()) != null);
        }

        public static string[] GetTitleStyles<T>()
        {
            var propExcelName = typeof(T).GetCustomAttributes(typeof(ReportTitleStyleAttribute), true).FirstOrDefault();
            return ((ReportTitleStyleAttribute)propExcelName) != null ? ((ReportTitleStyleAttribute)propExcelName).Styles : new string[] { };
        }
    }
}
