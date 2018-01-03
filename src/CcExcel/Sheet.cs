using CcExcel.Helpers;
using CcExcel.Messages;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace CcExcel
{
    public class Sheet
    {
        #region Constructors

        internal Sheet(Excel owner, int id)
        {
            _inMemoryName = Texts.DefaultSheetName + id;
            Owner = owner;
            Id = id;
        }

        internal Sheet(Excel owner, Spreadsheet.Sheet sheet, SheetData sheetData)
        {
            Owner = owner;
            OpenXmlSheet = sheet;
            OpenXmlSheetData = sheetData;
            Id = (int)(uint)OpenXmlSheet.SheetId;
        }

        #endregion

        #region Non Public

        private SheetValueTable _sheetValueTable;
        private SheetStyleTable _sheetStyleTable;
        private string _inMemoryName;
        internal Excel Owner { get; }
        public Spreadsheet.Sheet OpenXmlSheet { get; private set; }
        public SheetData OpenXmlSheetData { get; private set; }

        internal void Consolidate()
        {
            if (OpenXmlSheetData != null) return;

            OpenXmlSheet = SpreadsheetHelper.GetSheet(Owner.OpenXmlDocument, Name, Id, createIfDoesntExists: true);
            OpenXmlSheetData = SpreadsheetHelper.GetSheetData(Owner.OpenXmlDocument, sheet: OpenXmlSheet);
        }

        #endregion

        #region Public

        public int Id { get; }

        public string Name
        {
            get { return OpenXmlSheet?.Name ?? _inMemoryName; }
            set
            {
                if (OpenXmlSheet != null) OpenXmlSheet.Name = value;
                _inMemoryName = value;
            }
        }

        public SheetValueTable Values
        {
            get { return _sheetValueTable ?? (_sheetValueTable = new SheetValueTable(this)); }
        }

        public SheetStyleTable Styles
        {
            get { return _sheetStyleTable ?? (_sheetStyleTable = new SheetStyleTable(this)); }
        }

        #endregion
    }
}
