using System;

namespace StormReport
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExcelTitleStyleAttribute : Attribute
    {
        public string[] Styles { get; set; }

        public ExcelTitleStyleAttribute(params string[] styles)
        {
            this.Styles = styles;
        }
    }
}