using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CcExcel.Helpers
{
    public static class SpreadsheetHelper
    {
        public static void CreateWorkbook(SpreadsheetDocument document)
        {
            if (document.WorkbookPart == null)
            {
                document.AddWorkbookPart();
            }

            if (document.WorkbookPart.Workbook == null)
            {
                document.WorkbookPart.Workbook = new Workbook();
            }
        }

        public static DocumentFormat.OpenXml.Spreadsheet.Sheet CreateSheet(SpreadsheetDocument document, int id, string name)
        {
            var wsp = document.WorkbookPart.AddNewPart<WorksheetPart>();
            wsp.Worksheet = new Worksheet(new SheetData());

            var sheets = document.WorkbookPart.Workbook.Sheets;

            if (sheets == null)
            {
                sheets = new Sheets();
                document.WorkbookPart.Workbook.AppendChild(sheets);
            }

            var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet
            {
                Id = document.WorkbookPart.GetIdOfPart(wsp),
                SheetId = UInt32Value.FromUInt32((uint)id),
                Name = StringValue.FromString(name)
            };

            sheets.Append(sheet);

            return sheet;
        }

        public static DocumentFormat.OpenXml.Spreadsheet.Sheet GetSheet(SpreadsheetDocument document, string name)
        {
            return document
                .WorkbookPart
                ?.Workbook
                ?.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                ?.FirstOrDefault(f => f.Name == name);
        }

        public static WorksheetPart GetSheetPart(SpreadsheetDocument document, string name)
        {
            var sheet = GetSheet(document, name);

            if (sheet == null) return null;

            return document.WorkbookPart?.GetPartById(sheet.Id) as WorksheetPart;
        }
    }
}
