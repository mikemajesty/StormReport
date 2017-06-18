using StormReport.Model;
using System.Collections.Generic;

namespace StormReport.Test
{
    [ExcelTitleStyle("background-color:powderblue;")]
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
        [ExportableAddtionalText("R$ ", AdditionalTextEnum.LEFT)]
        [ExportableColumnGroup("Cliente")]
        public string Nome { get; set; }

        [ExportableColumnHeaderName("Idade do Cliente")]
        [ExportableColumnHeaderStyle("text-align: center;", "color: red;", "font-size: 17px;")]
        [ExportableColumnContentStyle("text-align: center;", "color: blue;", @"font-size: 17px; mso-number-format: '0\.000'")]
        [ExportableColumnGroup("Cliente")]
        public int Idade { get; set; }

        [ExportableColumnHeaderName("Cidade do Cliente")]
        [ExportableColumnHeaderStyle("text-align: center;", "color: red;", "font-size: 17px;")]
        [ExportableColumnContentStyle("text-align: center;", "color: gray;", "font-size: 17px;")]
        [ExportableColumnGroup("Endereço", "background-color: gray")]
        public string Cidade { get; set; }

        [ExportableColumnHeaderName("Estado do Cliente")]
        [ExportableColumnHeaderStyle("text-align: center;", "color: red;", "font-size: 17px;")]
        [ExportableColumnContentStyle("text-align: center;", "color: blue;", "font-size: 17px;")]
        [ExportableColumnGroup("Endereço")]
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
