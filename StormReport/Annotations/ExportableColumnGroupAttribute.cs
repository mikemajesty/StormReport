using System;

namespace StormReport.Test
{
    public class ExportableColumnGroupAttribute : Attribute
    {
        public string Description { get; set; }

        public string[] Styles { get; set; }

        public ExportableColumnGroupAttribute(string description, params string[] styles)
        {
            this.Description = description;
            this.Styles = styles;
        }
    }
}