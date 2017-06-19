using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Web;
using Moq;
using System.IO;

namespace StormReport.Test
{
    [TestClass]
    public class ReportTest
    {
        private StormExcel report;
        private HttpResponseBase response;

        public ReportTest()
        {
            report = new StormExcel("Relatório de Cliente: " + DateTime.Now);
            response = Mock.Of<HttpResponseBase>();

            var textWriter = Mock.Of<TextWriter>();
            response.Output = textWriter;
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ResponseIsNull()
        {
            List<MockEmployeeTest> list = GetList();
            report.CreateExcel(list, null);
        }

        [TestMethod]
        public void GenerateReport()
        {
            StormPdf pdf = new StormPdf();
            List<MockEmployeeTest> list = GetList();
            pdf.CreatePdf(list, response);
        }

        [TestMethod]
        public void GenerateEmptyExcelRow()
        {
            report.CreateExcel(new List<MockEmployeeTest>(), response);
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void GenerateNullListObject()
        {
            List<MockEmployeeTest> list = null;
            report.CreateExcel(list, response);
        }

        private static List<MockEmployeeTest> GetList()
        {
            var class1 = new MockEmployeeTest { Cidade = "Ibiúna", Estado = "SP", Idade = 10, Nome = "Mike" };
            var class2 = new MockEmployeeTest { Cidade = "Sorocaba", Estado = "SP", Idade = 28, Nome = "Mike Lima" };
            var list = new List<MockEmployeeTest>();
            list.Add(class1);
            list.Add(class2);
            return list;
        }
    }
}
