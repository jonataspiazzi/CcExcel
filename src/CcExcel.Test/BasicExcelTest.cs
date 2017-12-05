using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CcExcel.Test
{
    [TestClass]
    public class BasicExcelTest
    {
        [TestMethod]
        public void ShouldReadCellsAsString()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var excel = new Excel(rs, ExcelMode.OpenReadOnly))
            {
                Assert.AreEqual("general", (string)excel["Sheet1"].Values["B", 2]);
                Assert.AreEqual("12.4568", (string)excel["Sheet1"].Values["B", 3]);
                Assert.AreEqual("45.25", (string)excel["Sheet1"].Values["B", 4]);
                Assert.AreEqual("18.56", (string)excel["Sheet1"].Values["B", 5]);
                Assert.AreEqual("32408", (string)excel["Sheet1"].Values["B", 6]);
                Assert.AreEqual("42952", (string)excel["Sheet1"].Values["B", 7]);
                Assert.AreEqual("0.49", (string)excel["Sheet1"].Values["B", 8]);
                Assert.AreEqual("0.1845", (string)excel["Sheet1"].Values["B", 9]);
                Assert.AreEqual("0.2", (string)excel["Sheet1"].Values["B", 10]);
                Assert.AreEqual("10500000", (string)excel["Sheet1"].Values["B", 11]);
                Assert.AreEqual("text1", (string)excel["Sheet1"].Values["B", 12]);
                Assert.AreEqual("text2", (string)excel["Sheet1"].Values["B", 13]);
                Assert.AreEqual("text1", (string)excel["Sheet1"].Values["B", 14]);
                Assert.AreEqual("text2", (string)excel["Sheet1"].Values["B", 15]);
                Assert.AreEqual("a", (string)excel["Sheet1"].Values["B", 16]);
                Assert.AreEqual("1", (string)excel["Sheet1"].Values["B", 17]);
                Assert.IsNull(excel["Sheet1"].Values["B", 18]);
            }
        }

        [TestMethod]
        public void ShouldReadCellsAsCustomTypes()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var excel = new Excel(rs, ExcelMode.OpenReadOnly))
            {
                Assert.AreEqual("general", excel["Sheet1"].Values["B", 2]);
                Assert.AreEqual(12.4568, excel["Sheet1"].Values["B", 3]);
                Assert.AreEqual(45.25M, excel["Sheet1"].Values["B", 4]);
                Assert.AreEqual(18.56, excel["Sheet1"].Values["B", 5]);
                Assert.AreEqual(new DateTime(1988, 22, 09), excel["Sheet1"].Values["B", 6]);
                Assert.AreEqual(new DateTime(2017, 8, 5), excel["Sheet1"].Values["B", 7]);
                Assert.AreEqual(new TimeSpan(11, 45, 36), excel["Sheet1"].Values["B", 8]);
                Assert.AreEqual((float?)0.1845, excel["Sheet1"].Values["B", 9]);
                Assert.AreEqual((double?)0.2, excel["Sheet1"].Values["B", 10]);
                Assert.AreEqual(10500000, excel["Sheet1"].Values["B", 11]);
                Assert.AreEqual("text1", excel["Sheet1"].Values["B", 12]);
                Assert.AreEqual("text2", excel["Sheet1"].Values["B", 13]);
                Assert.AreEqual("text1", excel["Sheet1"].Values["B", 14]);
                Assert.AreEqual("text2", excel["Sheet1"].Values["B", 15]);
                Assert.AreEqual('a', excel["Sheet1"].Values["B", 16]);
                Assert.AreEqual(true, excel["Sheet1"].Values["B", 17]);
                Assert.IsNull(excel["Sheet1"].Values["B", 18]);
            }
        }

        [TestMethod]
        public void ShoudWriteAndReadACell()
        {
            using(var ms = new MemoryStream())
            {
                using (var excel = new Excel(ms, ExcelMode.Create))
                {
                    excel["test"].Values["b", 2] = "info";

                    excel.Save();
                }

                ms.Position = 0;

                using (var excel = new Excel(ms, ExcelMode.OpenReadOnly))
                {
                    string value = excel["test"].Values["b", 2];

                    Assert.AreEqual("info", value);
                }

                DumpGeneratedExcelFiles.Dump(ms);
            }
        }
    }
}
