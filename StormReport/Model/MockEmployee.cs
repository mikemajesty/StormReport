using System.Collections.Generic;

namespace StormReport.Test
{
    [ExcelFileName("ExcelReport", useDateTimeToAdditionalName: true)]
    public class ClassFinalTest
    {
        private List<ClassFinalTest> List;

        public ClassFinalTest()
        {
            List = new List<ClassFinalTest>();
        }

        [ExportableColumnHeaderName("Nome do Cliente")]
        [ExportableColumnHeaderStyle("text-align: center;", "color: red;", "font-size: 17px;")]
        [ExportableColumnContentStyle("text-align: center;", "color: gray;", "font-size: 17px;")]
        public string Nome { get; set; }

        [ExportableColumnHeaderName("Idade do Cliente")]
        [ExportableColumnHeaderStyle("text-align: center;", "color: red;", "font-size: 17px;")]
        [ExportableColumnContentStyle("text-align: center;", "color: blue;", @"font-size: 17px; mso-number-format: '0\.000'")]
        public int Idade { get; set; }

        [ExportableColumnHeaderName("Cidade do Cliente")]
        [ExportableColumnHeaderStyle("text-align: center;", "color: red;", "font-size: 17px;")]
        [ExportableColumnContentStyle("text-align: center;", "color: gray;", "font-size: 17px;")]
        public string Cidade { get; set; }

        [ExportableColumnHeaderName("Estado do Cliente")]
        [ExportableColumnHeaderStyle("text-align: center;", "color: red;", "font-size: 17px;")]
        [ExportableColumnContentStyle("text-align: center;", "color: blue;", "font-size: 17px;")]
        public string Estado { get; set; }

        public void Add(ClassFinalTest e)
        {
            List.Add(e);
        }

        public List<ClassFinalTest> GetList()
        {
            return List;
        }
    }
}
