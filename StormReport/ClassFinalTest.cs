using System.Collections;
using System.Collections.Generic;
using System.Web.UI.WebControls;

namespace StormReport.Test
{
    public class ClassFinalTest
    {
        private static Style a1;
        private static Style a2;
        private List<ClassFinalTest> List;
        private List<Style> style;

        public ClassFinalTest()
        {
            List = new List<ClassFinalTest>();
        }
  
        [ExportableColumnName("Nome do Cliente")]
        [ExportableColumnStyle("text-align: center;", "  color: red;", "font-size: 17px;")]
        public string Nome { get; set; }

        [ExportableColumnName("Idade do Cliente")]
        public int Idade { get; set; }

        [ExportableColumnName("Cidade do Cliente")]
        public string Cidade { get; set; }

        [ExportableColumnName("Estado do Cliente")]
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
