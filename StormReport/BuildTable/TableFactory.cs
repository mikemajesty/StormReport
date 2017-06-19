using System;
using System.Text;

namespace StormReport.BuildTable
{
    public class TableFactory
    {
        private StringBuilder table;

        public TableFactory()
        {
            table = new StringBuilder();
        }

        public void InitTable()
        {
            table.Append("<div>\n");
            table.Append("  <table cellspacing='0' rules='all' border='1'>\n");
        }

        public void EndTable()
        {
            table.Append("  </table>\n");
            table.Append("</div>\n");
        }

        public void AddRow()
        {
            table.Append("    <tr>\n");
        }

        public void EndRow()
        {
            table.Append("    </tr>\n");
        }

        public void AddReportTitle(string title, int columnCount, string[] styles)
        {
            StringBuilder style = new StringBuilder();

            Array.ForEach(styles, s => style.Append(s.Contains(";") ? s : s + ";"));

            table.Append(string.Format("      <th colspan='{0}' style='{2}'>{1}</th>\n", columnCount, title, style));
        }

        public void AddColumnGroup(string description, int colspan, string[] styles)
        {
            StringBuilder style = new StringBuilder();

            Array.ForEach(styles, s => style.Append(s.Contains(";") ? s : s + ";"));

            table.Append(string.Format("      <th colspan='{0}' style='{2}'>{1}</th>\n", colspan, description, style));
        }

        public void AddColumnTextHeader(object text, string[] style)
        {
            StringBuilder styles = new StringBuilder();

            Array.ForEach(style, s => styles.Append(s.Contains(";") ? s : s + ";"));

            table.Append(string.Format("      <th scope='col' style='{0}'>\n", styles));
            table.Append(text);
            table.Append("    </th>\n");
        }

        public void AddColumnContentText(object text, string[] style, ExportableAddtionalTextAttribute additionalText)
        {
            StringBuilder styles = new StringBuilder();

            Array.ForEach(style, s => styles.Append(s.Contains(";") ? s : s + ";"));

            table.Append(string.Format("      <td scope='row' style='{0}'>\n", styles));
            table.Append(FormatText(text, additionalText));
            table.Append("      </td>\n");
        }

        private string FormatText(object text, ExportableAddtionalTextAttribute additionalText)
        {
            if (additionalText == null || string.IsNullOrEmpty(additionalText.Description))
                return text.ToString();

            return additionalText.Direction == Model.AdditionalTextEnum.LEFT ? additionalText.Description + text.ToString() : text.ToString() + additionalText.Description;
        }

        public string ToHtml()
        {
            return table.ToString();
        }
    }
}
