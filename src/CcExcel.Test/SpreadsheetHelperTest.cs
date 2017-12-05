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
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace CcExcel.Test
{
    [TestClass]
    public class SpreadsheetHelperTest
    {
        [TestMethod]
        public void ShouldGetSheetData()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var doc = SpreadsheetDocument.Open(rs, false))
            {
                var sheetData = SpreadsheetHelper.GetSheetData(doc, "Sheet1");

                Assert.IsNotNull(sheetData);
            }
        }

        [TestMethod]
        public void ShouldGetSheetDataInSteps()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var doc = SpreadsheetDocument.Open(rs, false))
            {
                var sheet = SpreadsheetHelper.GetSheet(doc, "Sheet1");

                var sheetData = SpreadsheetHelper.GetSheetData(doc, sheet: sheet);

                Assert.IsNotNull(sheetData);
            }
        }

        [TestMethod]
        public void ShouldGetSharedStringTable()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var doc = SpreadsheetDocument.Open(rs, false))
            {
                var sharedStringTable = SpreadsheetHelper.GetSharedString(doc);

                Assert.IsNotNull(sharedStringTable);
            }
        }

        [TestMethod]
        public void ShouldGetCell()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var doc = SpreadsheetDocument.Open(rs, false))
            {
                var sheetData = SpreadsheetHelper.GetSheetData(doc, "Sheet1");
                var cell = SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("B"), 2);

                Assert.IsNotNull(cell);
            }
        }

        [TestMethod]
        public void ShouldGetMaxId()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.MultiTabs.xlsx"))
            using (var doc = SpreadsheetDocument.Open(rs, false))
            {
                var value = SpreadsheetHelper.GetMaxId(doc);

                Assert.AreEqual(3, value);
            }
        }

        [TestMethod]
        public void ShouldGetAllTypesOfValues()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var doc = SpreadsheetDocument.Open(rs, false))
            {
                var sd = SpreadsheetHelper.GetSheetData(doc, "Sheet1");
                var b = BaseAZ.Parse("B");

                Assert.AreEqual("general", SpreadsheetHelper.GetValue(doc, sd, b, 2));
                Assert.AreEqual("12.4568", SpreadsheetHelper.GetValue(doc, sd, b, 3));
                Assert.AreEqual("45.25", SpreadsheetHelper.GetValue(doc, sd, b, 4));
                Assert.AreEqual("18.56", SpreadsheetHelper.GetValue(doc, sd, b, 5));
                Assert.AreEqual("32408", SpreadsheetHelper.GetValue(doc, sd, b, 6));
                Assert.AreEqual("42952", SpreadsheetHelper.GetValue(doc, sd, b, 7));
                Assert.AreEqual("0.49", SpreadsheetHelper.GetValue(doc, sd, b, 8));
                Assert.AreEqual("0.1845", SpreadsheetHelper.GetValue(doc, sd, b, 9));
                Assert.AreEqual("0.2", SpreadsheetHelper.GetValue(doc, sd, b, 10));
                Assert.AreEqual("10500000", SpreadsheetHelper.GetValue(doc, sd, b, 11));
                Assert.AreEqual("text1", SpreadsheetHelper.GetValue(doc, sd, b, 12));
                Assert.AreEqual("text2", SpreadsheetHelper.GetValue(doc, sd, b, 13));
                Assert.AreEqual("text1", SpreadsheetHelper.GetValue(doc, sd, b, 14));
                Assert.AreEqual("text2", SpreadsheetHelper.GetValue(doc, sd, b, 15));
                Assert.AreEqual("a", SpreadsheetHelper.GetValue(doc, sd, b, 16));
                Assert.AreEqual("1", SpreadsheetHelper.GetValue(doc, sd, b, 17));
                Assert.IsNull(SpreadsheetHelper.GetValue(doc, sd, b, 18));
            }
        }

        [TestMethod]
        public void ShouldCreateSheetData()
        {
            using (var ms = new MemoryStream())
            {
                using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true))
                {
                    SpreadsheetHelper.GetSheetData(doc, "Sheet3", createIfDoesntExists: true);

                    doc.Save();
                    doc.Dispose();
                    ms.Position = 0;
                }

                using (var doc = SpreadsheetDocument.Open(ms, true))
                {
                    var sheetData = SpreadsheetHelper.GetSheetData(doc, "Sheet3");

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
                    var sheetData = SpreadsheetHelper.GetSheetData(doc, "Sheet3", createIfDoesntExists: true);

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
                    var sheetData = SpreadsheetHelper.GetSheetData(doc, "Sheet3");

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

        [TestMethod]
        public void ShouldInsertInSharedStringTable()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var ms = new MemoryStream())
            {
                rs.CopyTo(ms);
                ms.Position = 0;

                var doc = SpreadsheetDocument.Open(ms, true);

                var newId = SpreadsheetHelper.InsertInSharedString(doc, "new value");

                Assert.AreEqual(4, newId);

                var sheetData = SpreadsheetHelper.GetSheetData(doc, "Sheet1");
                var cell = SpreadsheetHelper.GetCell(sheetData, BaseAZ.Parse("B"), 12);

                cell.CellValue = new Spreadsheet.CellValue("4");

                doc.Save();

                Assert.AreEqual("4", cell.InnerText);

                DumpGeneratedExcelFiles.Dump(ms);
            }
        }
    }
}
