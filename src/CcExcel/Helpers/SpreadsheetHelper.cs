using CcExcel.Messages;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace CcExcel.Helpers
{
#if TESTABLE
    public
#else
    internal
#endif
        static class SpreadsheetHelper
    {
        public static Workbook GetWorkbook(SpreadsheetDocument document, bool createIfDoesntExists = false)
        {
            if (document.WorkbookPart == null && createIfDoesntExists)
            {
                document.AddWorkbookPart();
            }

            if (document.WorkbookPart == null) return null;

            if (document.WorkbookPart.Workbook == null && createIfDoesntExists)
            {
                document.WorkbookPart.Workbook = new Workbook();
            }

            return document.WorkbookPart.Workbook;
        }

        public static SharedStringTable GetSharedString(SpreadsheetDocument document, bool createIfDoesntExists = false)
        {
            var wbp = GetWorkbook(document, createIfDoesntExists)?.WorkbookPart;

            if (wbp == null) return null;

            // TODO: consider use GetPartsOfType<> https://msdn.microsoft.com/pt-br/library/office/cc861607.aspx
            var sstp = wbp.SharedStringTablePart;

            if (sstp == null)
            {
                if (!createIfDoesntExists) return null;

                sstp = wbp.AddNewPart<SharedStringTablePart>();
            }

            if (sstp.SharedStringTable != null) return sstp.SharedStringTable;

            if (!createIfDoesntExists) return null;

            return sstp.SharedStringTable = new SharedStringTable();
        }

        public static Spreadsheet.Sheet GetSheet(SpreadsheetDocument document, string sheetName = null, int? sheetId = null, bool createIfDoesntExists = false)
        {
            var workbook = GetWorkbook(document, createIfDoesntExists);

            if (workbook == null) return null;

            // Get or create ~/workbook.xml/workbook/sheets

            var sheets = workbook?.GetFirstChild<Sheets>();

            if (sheets == null)
            {
                if (!createIfDoesntExists) return null;

                sheets = new Sheets();
                workbook.AppendChild(sheets);
            }

            // Get or create ~/workbook.xml/workbook/sheets/sheet

            var sheetCollection = sheets
                .Elements<Spreadsheet.Sheet>()
                .Where(w => w.SheetId == sheetId || w.Name == sheetName)
                .ToList();

            if (sheetCollection.Count > 1)
            {
                throw new ExcelBadFormatException(Texts.TheExcelFileIsProbablyCorrupted + " " + Texts.MultipleSheetsWithSameNameOrSameIdWereFound);
            }

            var sheet = sheetCollection.FirstOrDefault();

            if (sheet == null)
            {
                if (!createIfDoesntExists) return null;

                if (sheetId == null)
                {
                    sheetId = GetMaxId(document) + 1;
                }

                if (string.IsNullOrEmpty(sheetName))
                {
                    sheetName = Texts.DefaultSheetName + sheetId;
                }

                sheet = new Spreadsheet.Sheet
                {
                    SheetId = (uint)sheetId,
                    Name = sheetName
                };

                sheets.Append(sheet);

                return sheet;
            }
            else return sheet;
        }

        public static SheetData GetSheetData(SpreadsheetDocument document, string sheetName = null, int? sheetId = null, bool createIfDoesntExists = false, Spreadsheet.Sheet sheet = null)
        {
            sheet = sheet ?? GetSheet(document, sheetName, sheetId, createIfDoesntExists);
            SheetData sheetData = null;
            WorksheetPart wsp;

            if (sheet == null) return null;

            if (sheet.Id == null)
            {
                wsp = document.WorkbookPart.AddNewPart<WorksheetPart>();
                wsp.Worksheet = new Worksheet(sheetData = new SheetData());

                sheet.Id = document.WorkbookPart.GetIdOfPart(wsp);
            }

            // Get or create ~/worksheets/sheet0.xml/worksheet/sheetData

            if (sheetData != null) return sheetData;

            wsp = document.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;

            if (wsp == null)
            {
                throw new ExcelBadFormatException(Texts.TheExcelFileIsProbablyCorrupted + " " + Texts.TheWorksheetPartWasNotFound);
            }

            sheetData = wsp.Worksheet.GetFirstChild<SheetData>();

            if (sheetData == null)
            {
                throw new ExcelBadFormatException(Texts.TheExcelFileIsProbablyCorrupted + " " + Texts.TheWorksheetPartWasNotFound);
            }

            return sheetData;
        }

        public static Row GetRow(SheetData sheetData, uint line, bool createIfDoesntExists = false)
        {
            var row = sheetData?.Elements<Row>().FirstOrDefault(r => r.RowIndex == line);

            // Se existir a linha retorna.
            if (row != null)
            {
                row.Spans = null;
                return row;
            }
            else if (!createIfDoesntExists)
            {
                return null;
            }

            // Senao cria uma nova linha.
            row = new Row { RowIndex = line };

            var rows = sheetData.Elements<Row>().ToList();

            // Caso nao exista linhas pode simplesmente inserir
            if (!rows.Any())
            {
                sheetData.AppendChild(row);
                return row;
            }

            // Insere a linha na order correta.
            var before = rows.Where(w => w.RowIndex < line).OrderBy(o => o.RowIndex.Value).LastOrDefault();

            if (before != null) // Existem linhas anteriores a que sera inserida.
            {
                sheetData.InsertAfter(row, before);
            }
            else // Nao existem nenhuma linha anterior a que sera inserida.
            {
                var after = rows.Where(w => w.RowIndex > line).OrderBy(o => o.RowIndex.Value).FirstOrDefault();

                // Insere antes do primeiro.
                sheetData.InsertBefore(row, after);
            }

            return row;
        }

        public static Cell GetCell(SheetData sheetData, BaseAZ column, uint line, bool createIfDoesntExists = false)
        {
            var row = GetRow(sheetData, line, createIfDoesntExists);

            if (row == null) return null;

            var cellReference = column.ToString() + line;

            var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == cellReference);

            if (cell != null)
            {
                return cell;
            }
            else if (!createIfDoesntExists)
            {
                return null;
            }

            cell = new Cell { CellReference = cellReference };

            // Se nao existir outras celulas pode inserir em qualquer lugar.
            if (!row.Elements<Cell>().Any())
            {
                row.AppendChild(cell);
                return cell;
            }

            // Caso existam outras celular precisa inserir na posicao correta.
            var cells = row.Elements<Cell>()
                .Select(s => new
                {
                    Ref = CellReference.Parse(s.CellReference),
                    Cell = s
                })
                .OrderBy(o => o.Ref.Column)
                .ToList();

            var before = cells.LastOrDefault(w => w.Ref.Column < column);

            if (before != null) // Existem linhas anteriores a que sera inserida.
            {
                row.InsertAfter(cell, before.Cell);
            }
            else // Nao existem nenhuma linha anterior a que sera inserida.
            {
                var after = cells.First(f => f.Ref.Column > column);

                // Insere antes do primeiro.
                row.InsertBefore(cell, after.Cell);
            }

            return cell;
        }

        public static string GetValue(SpreadsheetDocument document, SheetData sheetData, Cell cell, BaseAZ? column = null, uint? line = null)
        {
            if (cell == null && (sheetData == null || column == null || line == null)) return null;

            cell = cell ?? GetCell(sheetData, column.Value, line.Value);

            if (cell?.DataType?.Value == CellValues.SharedString)
            {
                var sst = GetSharedString(document);

                return sst.ElementAt(int.Parse(cell.InnerText)).InnerText;
            }

            return cell?.InnerText;
        }

        public static int GetMaxId(SpreadsheetDocument document)
        {
            return document
                .WorkbookPart
                ?.Workbook
                ?.GetFirstChild<Sheets>()
                ?.Elements<Spreadsheet.Sheet>()
                ?.Select(s => (int)(uint)s.SheetId)
                ?.OrderByDescending(o => o)
                ?.FirstOrDefault() ?? 0;
        }

        public static int InsertInSharedString(SpreadsheetDocument document, string value)
        {
            var sst = GetSharedString(document, createIfDoesntExists: true);

            var i = 0;

            foreach (var item in sst.Elements<SharedStringItem>())
            {
                if (item.InnerText == value) return i;

                i++;
            }

            sst.AppendChild(new SharedStringItem(new Text(value)));
            sst.Save();

            return i;
        }

        public static void SetValue(SpreadsheetDocument document, Cell cell, string value, CellValues? type, SheetData sheetData = null, BaseAZ? column = null, uint? line = null)
        {
            if (cell == null && (sheetData == null || column == null || line == null))
            {
                throw new ArgumentNullException("cell");
            }

            cell = cell ?? GetCell(sheetData, column.Value, line.Value, createIfDoesntExists: true);

            if (type == CellValues.SharedString)
            {
                var index = InsertInSharedString(document, value);

                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.CellValue = new CellValue(index.ToString());

                return;
            }

            cell.DataType = type != null ? new EnumValue<CellValues>(type.Value) : null;
            cell.CellValue = new CellValue(value);
        }

        public static void RemoveSheet(SpreadsheetDocument document, string sheetToDelete)
        {
            // from: https://blogs.msdn.microsoft.com/vsod/2010/02/05/how-to-delete-a-worksheet-from-excel-using-open-xml-sdk-2-0/

            // Diferences
            // [Content_Types].xml/Types/
            // docProps\app.xml/Properties/HeadingPairs/vector/variant/i4
            // docProps\app.xml/Properties/TitlesOfParts/vector/:size
            // docProps\app.xml/Properties/TitlesOfParts/vector/lpstr
            // xl\_rels\workbook.xml.rels/Relationships/
            // xl\worksheets\
            // xl\workbook.xml/workbook/sheets/

            string Sheetid = "";
            //Open the workbook
            WorkbookPart wbPart = document.WorkbookPart;

            // Get the pivot Table Parts
            var pvtTableCacheParts = wbPart.PivotTableCacheDefinitionParts;
            var pvtTableCacheDefinationPart = new Dictionary<PivotTableCacheDefinitionPart, string>();
            foreach (var item in pvtTableCacheParts)
            {
                var pvtCacheDef = item.PivotCacheDefinition;
                //Check if this CacheSource is linked to SheetToDelete
                var pvtCahce = pvtCacheDef.Descendants<CacheSource>().Where(s => s.WorksheetSource.Sheet == sheetToDelete);
                if (pvtCahce.Count() > 0)
                {
                    pvtTableCacheDefinationPart.Add(item, item.ToString());
                }
            }
            foreach (var Item in pvtTableCacheDefinationPart)
            {
                wbPart.DeletePart(Item.Key);
            }
            //Get the SheetToDelete from workbook.xml
            var theSheet = wbPart.Workbook.Descendants<Spreadsheet.Sheet>().Where(s => s.Name == sheetToDelete).FirstOrDefault();
            if (theSheet == null)
            {
                // The specified sheet doesn't exist.
            }
            //Store the SheetID for the reference
            Sheetid = theSheet.SheetId;

            // Remove the sheet reference from the workbook.
            var worksheetPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
            theSheet.Remove();

            // Delete the worksheet part.
            wbPart.DeletePart(worksheetPart);

            //Get the DefinedNames
            var definedNames = wbPart.Workbook.Descendants<DefinedNames>().FirstOrDefault();
            if (definedNames != null)
            {
                var defNamesToDelete = new List<DefinedName>();

                foreach (DefinedName item in definedNames)
                {
                    // This condition checks to delete only those names which are part of Sheet in question
                    if (item.Text.Contains(sheetToDelete + "!"))
                        defNamesToDelete.Add(item);
                }

                foreach (DefinedName item in defNamesToDelete)
                {
                    item.Remove();
                }
            }
            // Get the CalculationChainPart 
            //Note: An instance of this part type contains an ordered set of references to all cells in all worksheets in the 
            //workbook whose value is calculated from any formula

            var calChainPart = wbPart.CalculationChainPart;
            if (calChainPart != null)
            {
                var calChainEntries = calChainPart.CalculationChain.Descendants<CalculationCell>().Where(c => c.SheetId == Sheetid);
                var calcsToDelete = new List<CalculationCell>();
                foreach (var item in calChainEntries)
                {
                    calcsToDelete.Add(item);
                }

                foreach (var item in calcsToDelete)
                {
                    item.Remove();
                }

                if (calChainPart.CalculationChain.Count() == 0)
                {
                    wbPart.DeletePart(calChainPart);
                }
            }
        }
    }
}
