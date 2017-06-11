using System;

namespace StormReport
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportableColumnContentStyleAttribute : Attribute
    {
        public string[] Styles { get; set; }

        public ExportableColumnContentStyleAttribute(params string[] styles)
        {
            this.Styles = styles;
        }
    }
}
