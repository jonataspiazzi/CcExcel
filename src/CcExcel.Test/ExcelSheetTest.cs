using CcExcel.Helpers;
using DocumentFormat.OpenXml.Packaging;
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
    public class ExcelSheetTest
    {
        [TestMethod]
        public void ShouldRemoveASheet()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.MultiTabs.xlsx"))
            using (var ms = new MemoryStream())
            {
                rs.CopyTo(ms);

                using (var excel = new Excel(ms, ExcelMode.Open))
                {
                    excel["Sheet2"].Remove();
                }

                using (var doc = SpreadsheetDocument.Open(ms, true))
                {
                    Assert.IsNull(SpreadsheetHelper.GetSheet(doc, "Sheet2"));
                    Assert.IsNull(SpreadsheetHelper.GetSheetData(doc, "Sheet2"));
                }

                DumpGeneratedExcelFiles.Dump(ms);
            }
        }
    }
}
