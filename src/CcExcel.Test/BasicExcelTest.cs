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
