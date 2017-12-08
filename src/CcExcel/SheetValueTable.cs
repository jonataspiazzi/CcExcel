using CcExcel.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CcExcel
{
    public class SheetValueTable
    {
        internal Sheet Owner { get; }

        internal SheetValueTable(Sheet owner)
        {
            Owner = owner;
        }

        public ExcelValue this[BaseAZ column, int line]
        {
            get
            {
                var cell = SpreadsheetHelper.GetCell(Owner.OpenXmlSheetData, column, (uint)line, createIfDoesntExists: false);

                var value = SpreadsheetHelper.GetValue(Owner.Owner.OpenXmlDocument, Owner.OpenXmlSheetData, cell);

                return new ExcelValue(value, cell?.DataType?.Value);
            }
            set
            {
                if (value == null || value.IsEmpty)
                {
                    var cell = SpreadsheetHelper.GetCell(Owner.OpenXmlSheetData, column, (uint)line, createIfDoesntExists: false);

                    cell?.Remove();
                }
                else
                {
                    Owner.Consolidate();

                    SpreadsheetHelper.SetValue(Owner.Owner.OpenXmlDocument, null, value.ToString(), value.ValueType, Owner.OpenXmlSheetData, column, (uint)line);
                }
            }
        }

        public ExcelValue this[string column, int line]
        {
            get { return this[BaseAZ.Parse(column), line]; }
            set { this[BaseAZ.Parse(column), line] = value; }
        }

        public ExcelValue this[int column, int line]
        {
            get { return this[(BaseAZ)column, line]; }
            set { this[(BaseAZ)column, line] = value; }
        }
    }
}
