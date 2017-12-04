using CcExcel.Messages;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

        public static SheetData GetSheetData(SpreadsheetDocument document, string sheetName = null, int? sheetId = null, bool createIfDoesntExists = false)
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
                .Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                .Where(w => w.SheetId == sheetId || w.Name == sheetName)
                .ToList();

            if (sheetCollection.Count > 1)
            {
                throw new ExcelBadFormatException(Texts.TheExcelFileIsProbablyCorrupted + " " + Texts.MultipleSheetsWithSameNameOrSameIdWereFound);
            }

            var sheet = sheetCollection.FirstOrDefault();
            SheetData sheetData = null;
            WorksheetPart wsp;

            if (sheet == null)
            {
                if (!createIfDoesntExists) return null;

                wsp = document.WorkbookPart.AddNewPart<WorksheetPart>();
                wsp.Worksheet = new Worksheet(sheetData = new SheetData());

                if (sheetId == null)
                {
                    sheetId = sheets
                        .Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                        .Select(s => (int)(uint)s.SheetId)
                        .OrderByDescending(o => o)
                        .FirstOrDefault();

                    sheetId++;
                }

                if (string.IsNullOrEmpty(sheetName))
                {
                    sheetName = Texts.DefaultSheetName + sheetId;
                }

                sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet
                {
                    Id = document.WorkbookPart.GetIdOfPart(wsp),
                    SheetId = (uint)sheetId,
                    Name = sheetName
                };

                sheets.Append(sheet);
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
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == line);

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
    }
}
