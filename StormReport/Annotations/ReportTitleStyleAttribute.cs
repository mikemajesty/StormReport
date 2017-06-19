using System;

namespace StormReport
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ReportTitleStyleAttribute : Attribute
    {
        public string[] Styles { get; set; }

        public ReportTitleStyleAttribute(params string[] styles)
        {
            this.Styles = styles;
        }
    }
}