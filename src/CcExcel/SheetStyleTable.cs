using CcExcel.Helpers;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CcExcel
{
    public class SheetStyleTable
    {
        internal Sheet Owner { get; }

        internal SheetStyleTable(Sheet owner)
        {
            Owner = owner;
        }

        public uint? this[BaseAZ column, int line]
        {
            get
            {
                var cell = SpreadsheetHelper.GetCell(Owner.OpenXmlSheetData, column, (uint)line, createIfDoesntExists: false);

                return cell.StyleIndex;
            }
            set
            {
                Owner.Consolidate();

                var cell = SpreadsheetHelper.GetCell(Owner.OpenXmlSheetData, column, (uint)line, createIfDoesntExists: true);

                if (cell.CellValue == null) cell.CellValue = new CellValue(null);
                cell.StyleIndex = value;
            }
        }

        public uint? this[string column, int line]
        {
            get { return this[BaseAZ.Parse(column), line]; }
            set { this[BaseAZ.Parse(column), line] = value; }
        }

        public uint? this[int column, int line]
        {
            get { return this[(BaseAZ)column, line]; }
            set { this[(BaseAZ)column, line] = value; }
        }
    }
}
