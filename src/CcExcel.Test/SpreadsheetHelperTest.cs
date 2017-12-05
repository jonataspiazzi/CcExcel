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
        public void ShouldGetSheetData()
        {
            var doc = SpreadsheetDocument.Open(GetType().Assembly
                .GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"), false);

            var sheetData = SpreadsheetHelper.GetSheetData(doc, "Sheet1");

            Assert.IsNotNull(sheetData);
        }

        [TestMethod]
        public void ShouldGetSheetDataInSteps()
        {
            var doc = SpreadsheetDocument.Open(GetType().Assembly
                .GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"), false);

            var sheet = SpreadsheetHelper.GetSheet(doc, "Sheet1");

            var sheetData = SpreadsheetHelper.GetSheetData(doc, sheet: sheet);

            Assert.IsNotNull(sheetData);
        }

        [TestMethod]
        public void ShouldGetCell()
        {
            var doc = SpreadsheetDocument.Open(GetType().Assembly
                .GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"), false);

            var sheetData = SpreadsheetHelper.GetSheetData(doc, "Sheet1");
            var cell = SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("B"), 2);

            Assert.IsNotNull(cell);
        }

        [TestMethod]
        public void ShouldGetMaxId()
        {
            var doc = SpreadsheetDocument.Open(GetType().Assembly
                .GetManifestResourceStream("CcExcel.Test.Resources.MultiTabs.xlsx"), false);

            var value = SpreadsheetHelper.GetMaxId(doc);

            Assert.AreEqual(3, value);
        }

        [TestMethod]
        public void ShouldCreateSheetData()
        {
            using (var ms = new MemoryStream())
            {
                using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true))
                {
                    SpreadsheetHelper.GetSheetData(doc, "plan3", createIfDoesntExists: true);

                    doc.Save();
                    doc.Dispose();
                    ms.Position = 0;
                }

                using (var doc = SpreadsheetDocument.Open(ms, true))
                {
                    var sheetData = SpreadsheetHelper.GetSheetData(doc, "plan3");

                    Assert.IsNotNull(sheetData);

                    doc.Dispose();
                }

                DumpGeneratedExcelFiles.Dump(ms);
            }
        }

        [TestMethod]
        public void ShouldCreateSheetCells()
        {
            using (var ms = new MemoryStream())
            {
                using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true))
                {
                    var sheetData = SpreadsheetHelper.GetSheetData(doc, "plan3", createIfDoesntExists: true);

                    Action<string, uint, int> setCell = (column, line, value) =>
                    {
                        var azColumn = BaseAZ.Parse(column);
                        var valueStr = value.ToString();

                        SpreadsheetHelper.GetCell(sheetData, azColumn, line, createIfDoesntExists: true).CellValue
                            = new DocumentFormat.OpenXml.Spreadsheet.CellValue(valueStr);
                    };

                    setCell("C", 3, 5);
                    setCell("C", 1, 8);
                    setCell("C", 5, 2);
                    setCell("A", 3, 4);
                    setCell("A", 1, 7);
                    setCell("A", 5, 1);
                    setCell("E", 3, 6);
                    setCell("E", 1, 9);
                    setCell("E", 5, 3);

                    doc.Save();
                    doc.Dispose();
                    ms.Position = 0;
                }

                using (var doc = SpreadsheetDocument.Open(ms, true))
                {
                    var sheetData = SpreadsheetHelper.GetSheetData(doc, "plan3");

                    Assert.AreEqual("1", SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("A"), 5).CellValue.InnerText);
                    Assert.AreEqual("2", SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("C"), 5).CellValue.InnerText);
                    Assert.AreEqual("3", SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("E"), 5).CellValue.InnerText);
                    Assert.AreEqual("4", SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("A"), 3).CellValue.InnerText);
                    Assert.AreEqual("5", SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("C"), 3).CellValue.InnerText);
                    Assert.AreEqual("6", SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("E"), 3).CellValue.InnerText);
                    Assert.AreEqual("7", SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("A"), 1).CellValue.InnerText);
                    Assert.AreEqual("8", SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("C"), 1).CellValue.InnerText);
                    Assert.AreEqual("9", SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("E"), 1).CellValue.InnerText);

                    doc.Dispose();
                }

                DumpGeneratedExcelFiles.Dump(ms);
            }
        }
    }
}
