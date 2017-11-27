using CcExcel.Helpers;
using DocumentFormat.OpenXml;
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
    public class SpreadsheetHelperTest
    {
        [TestMethod]
        public void ShouldGetSheet()
        {
            var doc = SpreadsheetDocument.Open(GetType().Assembly
                .GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"), false);

            var sheet = SpreadsheetHelper.GetSheet(doc, "Sheet1");

            Assert.IsNotNull(sheet);
        }

        [TestMethod]
        public void ShouldGetSheetPart()
        {
            var doc = SpreadsheetDocument.Open(GetType().Assembly
                .GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"), false);

            var wsp = SpreadsheetHelper.GetSheetPart(doc, "Sheet1");

            Assert.IsNotNull(wsp);
        }

        [TestMethod]
        public void ShouldCreateSheet()
        {
            using (var ms = new MemoryStream())
            {
                var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true);

                SpreadsheetHelper.CreateWorkbook(doc);
                SpreadsheetHelper.CreateSheet(doc, 3, "plan3");

                doc.Save();
                doc.Dispose();
                ms.Position = 0;
                
                doc = SpreadsheetDocument.Open(ms, true);

                var sheet = SpreadsheetHelper.GetSheet(doc, "plan3");
                var wsp = SpreadsheetHelper.GetSheetPart(doc, "plan3");

                Assert.IsNotNull(wsp);
                Assert.AreEqual("plan3", (string)sheet.Name);

                doc.Dispose();

                DumpGeneratedExcelFiles.Dump(ms);
            }
        }
    }
}
