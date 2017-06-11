using System;

namespace StormReport
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportableColumnHeaderNameAttribute : Attribute
    {
        public string Description { get; set; }
      
        public ExportableColumnHeaderNameAttribute(string description)
        {
            this.Description = string.IsNullOrEmpty(description) ? "Unknown name" : description;
        }
    }
}
