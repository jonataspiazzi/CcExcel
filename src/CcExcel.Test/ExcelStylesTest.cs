using CcExcel.Helpers;
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
    public class ExcelStylesTest
    {
        [TestMethod]
        public void ShouldGetCellStyles()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.Styles.xlsx"))
            using (var excel = new Excel(rs, ExcelMode.OpenReadOnly))
            {
                Assert.AreEqual("3", excel["Sheet1"].Styles["B", 2].ToString());
                Assert.AreEqual("1", excel["Sheet1"].Styles["B", 4].ToString());
                Assert.AreEqual("2", excel["Sheet1"].Styles["B", 6].ToString());
            }
        }

        [TestMethod]
        public void ShouldSetCellStyles()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.Styles.xlsx"))
            using (var ms = new MemoryStream())
            {
                rs.CopyTo(ms);
                ms.Position = 0;

                using (var excel = new Excel(ms, ExcelMode.Open))
                {
                    var style = excel["Sheet1"].Styles["B", 2];

                    for (var line = 3; line <= 6; line++)
                    {
                        excel["Sheet1"].Styles["B", line] = style;
                    }

                    excel.Save();
                }

                using (var excel = new Excel(ms, ExcelMode.OpenReadOnly))
                {
                    var sheet = SpreadsheetHelper.GetSheetData(excel.OpenXmlDocument, "Sheet1");
                    var b = BaseAZ.Parse("B");

                    Assert.AreEqual("3", SpreadsheetHelper.GetCell(sheet, b, 2).StyleIndex.InnerText);
                    Assert.AreEqual("3", SpreadsheetHelper.GetCell(sheet, b, 3).StyleIndex.InnerText);
                    Assert.AreEqual("3", SpreadsheetHelper.GetCell(sheet, b, 4).StyleIndex.InnerText);
                    Assert.AreEqual("3", SpreadsheetHelper.GetCell(sheet, b, 5).StyleIndex.InnerText);
                    Assert.AreEqual("3", SpreadsheetHelper.GetCell(sheet, b, 6).StyleIndex.InnerText);
                }

                DumpGeneratedExcelFiles.Dump(ms);
            }
        }
    }
}
