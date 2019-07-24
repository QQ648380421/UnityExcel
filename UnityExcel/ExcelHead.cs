using System;
namespace UnityExcel
{
    [AttributeUsage( AttributeTargets.Property)]
    public class ExcelHead:Attribute {
        private string name;

        public ExcelHead(string name)
        {
            Name = name; 
        }

        public string Name
        {
            get
            {
                return name;
            }

            set
            {
                name = value;
            }
        }
    }
}
