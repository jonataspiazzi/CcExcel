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
    public class ExcelValuesTest
    {
        [TestMethod]
        public void ShouldReadCellsAsString()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var excel = new Excel(rs, ExcelMode.OpenReadOnly))
            {
                Assert.AreEqual("general", excel["Sheet1"].Values["B", 2].ToString());
                Assert.AreEqual("12.4568", excel["Sheet1"].Values["B", 3].ToString());
                Assert.AreEqual("45.25", excel["Sheet1"].Values["B", 4].ToString());
                Assert.AreEqual("18.56", excel["Sheet1"].Values["B", 5].ToString());
                Assert.AreEqual("32408", excel["Sheet1"].Values["B", 6].ToString());
                Assert.AreEqual("42952", excel["Sheet1"].Values["B", 7].ToString());
                Assert.AreEqual("0.49", excel["Sheet1"].Values["B", 8].ToString());
                Assert.AreEqual("0.1845", excel["Sheet1"].Values["B", 9].ToString());
                Assert.AreEqual("0.2", excel["Sheet1"].Values["B", 10].ToString());
                Assert.AreEqual("10500000", excel["Sheet1"].Values["B", 11].ToString());
                Assert.AreEqual("text1", excel["Sheet1"].Values["B", 12].ToString());
                Assert.AreEqual("text2", excel["Sheet1"].Values["B", 13].ToString());
                Assert.AreEqual("text1", excel["Sheet1"].Values["B", 14].ToString());
                Assert.AreEqual("text2", excel["Sheet1"].Values["B", 15].ToString());
                Assert.AreEqual("a", excel["Sheet1"].Values["B", 16].ToString());
                Assert.AreEqual("1", excel["Sheet1"].Values["B", 17].ToString());
                Assert.IsNull(excel["Sheet1"].Values["B", 18].ToString());
            }
        }

        [TestMethod]
        public void ShouldReadCellsAsCustomTypes()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var excel = new Excel(rs, ExcelMode.OpenReadOnly))
            {
                Assert.AreEqual("general", excel["Sheet1"].Values["B", 2].ToString());
                Assert.AreEqual(12.4568, excel["Sheet1"].Values["B", 3].ToDouble());
                Assert.AreEqual(45.25M, excel["Sheet1"].Values["B", 4].ToDecimal());
                Assert.AreEqual(18.56f, excel["Sheet1"].Values["B", 5].ToSingle());
                Assert.AreEqual(new DateTime(1988, 9, 22), excel["Sheet1"].Values["B", 6].ToDateTime());
                Assert.AreEqual(new DateTime(2017, 8, 5), excel["Sheet1"].Values["B", 7].ToDateTime());
                Assert.AreEqual(new TimeSpan(11, 45, 36), excel["Sheet1"].Values["B", 8].ToTimeSpan());
                Assert.AreEqual((float?)0.1845, excel["Sheet1"].Values["B", 9].ToNullableSingle());
                Assert.AreEqual((double?)0.2, excel["Sheet1"].Values["B", 10].ToNullableDouble());
                Assert.AreEqual(10500000, excel["Sheet1"].Values["B", 11].ToInt32());
                Assert.AreEqual(10500000L, excel["Sheet1"].Values["B", 11].ToInt64());
                Assert.AreEqual("text1", excel["Sheet1"].Values["B", 12].ToString());
                Assert.AreEqual("text2", excel["Sheet1"].Values["B", 13].ToString());
                Assert.AreEqual("text1", excel["Sheet1"].Values["B", 14].ToString());
                Assert.AreEqual("text2", excel["Sheet1"].Values["B", 15].ToString());
                Assert.AreEqual('a', excel["Sheet1"].Values["B", 16].ToChar());
                Assert.AreEqual(true, excel["Sheet1"].Values["B", 17].ToBoolean());
            }
        }

        [TestMethod]
        public void ShouldHandleEmptyCells()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var excel = new Excel(rs, ExcelMode.OpenReadOnly))
            {
                var sheet = excel["Sheet1"];

                Assert.IsNull(sheet.Values["B", 18].ToNullableBoolean());
                Assert.IsNull(sheet.Values["B", 18].ToNullableByte());
                Assert.IsNull(sheet.Values["B", 18].ToNullableChar());
                Assert.IsNull(sheet.Values["B", 18].ToNullableDouble());
                Assert.IsNull(sheet.Values["B", 18].ToNullableInt16());
                Assert.IsNull(sheet.Values["B", 18].ToNullableInt32());
                Assert.IsNull(sheet.Values["B", 18].ToNullableInt64());
                Assert.IsNull(sheet.Values["B", 18].ToNullableSByte());
                Assert.IsNull(sheet.Values["B", 18].ToNullableSingle());
                Assert.IsNull(sheet.Values["B", 18].ToNullableUInt16());
                Assert.IsNull(sheet.Values["B", 18].ToNullableUInt32());
                Assert.IsNull(sheet.Values["B", 18].ToNullableUInt64());
                Assert.IsNull(sheet.Values["B", 18].ToNullableDecimal());
                Assert.IsNull(sheet.Values["B", 18].ToNullableDateTime());
                Assert.IsNull(sheet.Values["B", 18].ToNullableTimeSpan());
                Assert.IsNull(sheet.Values["B", 18].ToString());

                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToBoolean());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToByte());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToChar());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToDouble());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToInt16());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToInt32());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToInt64());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToSByte());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToSingle());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToUInt16());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToUInt32());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToUInt64());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToDecimal());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToDateTime());
                Assert.ThrowsException<EmptyValueException>(() => sheet.Values["B", 18].ToTimeSpan());
            }
        }

        [TestMethod]
        public void ShouldCleanCellsIfSetEmpty()
        {
            using (var rs = GetType().Assembly.GetManifestResourceStream("CcExcel.Test.Resources.AllTypes.xlsx"))
            using (var ms = new MemoryStream())
            {
                rs.CopyTo(ms);

                using (var excel = new Excel(ms, ExcelMode.Open))
                {
                    excel["Sheet1"].Values["B", 2] = null;
                    excel["Sheet1"].Values["B", 3] = null;
                    excel["Sheet1"].Values["B", 4] = null;
                    excel["Sheet1"].Values["B", 5] = null;
                    excel["Sheet1"].Values["B", 6] = null;
                    excel["Sheet1"].Values["B", 7] = null;
                    excel["Sheet1"].Values["B", 8] = null;
                    excel["Sheet1"].Values["B", 9] = null;
                    excel["Sheet1"].Values["B", 10] = null;
                    excel["Sheet1"].Values["B", 11] = null;
                    excel["Sheet1"].Values["B", 12] = null;
                    excel["Sheet1"].Values["B", 13] = null;
                    excel["Sheet1"].Values["B", 14] = null;
                    excel["Sheet1"].Values["B", 15] = null;
                    excel["Sheet1"].Values["B", 16] = null;
                    excel["Sheet1"].Values["B", 17] = null;
                    excel["Sheet1"].Values["B", 18] = null;

                    excel.Save();
                }

                using (var excel = new Excel(ms, ExcelMode.OpenReadOnly))
                {
                    Assert.IsNull(excel["Sheet1"].Values["B", 2].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 3].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 4].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 5].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 6].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 7].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 8].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 9].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 10].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 11].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 12].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 13].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 14].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 15].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 16].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 17].ToString());
                    Assert.IsNull(excel["Sheet1"].Values["B", 18].ToString());
                }

                DumpGeneratedExcelFiles.Dump(ms);
            }
        }

        [TestMethod]
        public void ShouldGetNullIfSheetDoesntExists()
        {
            using (var ms = new MemoryStream())
            using (var excel = new Excel(ms, ExcelMode.Create))
            {
                Assert.IsTrue(excel["Sheet1"].Values["B", 2].IsEmpty);
                Assert.IsTrue(excel["Sheet1"].Values["D", 2].IsEmpty);
                Assert.IsTrue(excel["Sheet1"].Values["B", 4].IsEmpty);
                Assert.IsTrue(excel["Sheet1"].Values["D", 4].IsEmpty);
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

                using (var excel = new Excel(ms, ExcelMode.OpenReadOnly))
                {
                    var value = excel["test"].Values["b", 2].ToString();

                    Assert.AreEqual("info", value);
                }

                DumpGeneratedExcelFiles.Dump(ms);
            }
        }
    }
}
