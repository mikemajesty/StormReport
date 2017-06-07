using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Collections;
using static StormReport.Report;
using System.Web.Mvc;

namespace StormReport.Test
{
    [TestClass]
    public class ReportByTest
    {
        private Report report;
        public ReportByTest()
        {
            report = new Report();
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentNullException))]
        public void ResponseIsNull()
        {
            List<ClassTest> list = GetList();
            report.CreateExcelBase(list);
        }

        private static List<ClassTest> GetList()
        {
            var class1 = new ClassTest { Cidade = "Ibiúna", Estado = "SP", Idade = 10, Nome = "Mike" };
            var class2 = new ClassTest { Cidade = "Sorocaba", Estado = "SP", Idade = 28, Nome = "Mike Lima" };
            var list = new List<ClassTest>();
            list.Add(class1);
            list.Add(class2);
            return list;
        }
    }
}
