using System;

namespace StormReport
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportableColumnHeaderStyleAttribute : Attribute
    {
        public string[] Styles { get; set; }

        public ExportableColumnHeaderStyleAttribute(params string[] styles)
        {
            this.Styles = styles;
        }
    }
}
