using CcExcel.Messages;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CcExcel
{
    public class Sheet
    {
        #region Fields

        internal Excel Owner { get; }
        internal DocumentFormat.OpenXml.Spreadsheet.Sheet OpenXmlSheet { get; private set; }
        private string _onMemoryName;

        #endregion

        #region Constructors

        internal Sheet(Excel owner, int id)
        {
            Owner = owner;
            Id = id;
            _onMemoryName = Texts.DefaultSheetName + Id;
        }

        internal Sheet(Excel owner, DocumentFormat.OpenXml.Spreadsheet.Sheet openXmlSheet)
        {
            Owner = owner;
            OpenXmlSheet = openXmlSheet;
            Id = (int)(uint)OpenXmlSheet.SheetId;
            Name = OpenXmlSheet.Name;
        }

        #endregion

        #region Properties

        public int Id { get; }

        public string Name
        {
            get { return OpenXmlSheet?.Name ?? _onMemoryName; }
            set
            {
                if (OpenXmlSheet != null) OpenXmlSheet.Name = value;
                _onMemoryName = value;
            }
        }

        public SheetValueTable Values
        {
            get { throw new NotImplementedException(); }
        }

        public SheetValueTable Styles
        {
            get { throw new NotImplementedException(); }
        }

        #endregion

        #region Methods

        internal void CreateSheet()
        {
            if (OpenXmlSheet != null) return;

            var wsp = Owner.OpenXmlDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            wsp.Worksheet = new Worksheet(new SheetData());

            var sheets = Owner.OpenXmlDocument.WorkbookPart.Workbook.Sheets;

            if (sheets == null)
            {
                sheets = new Sheets();
                Owner.OpenXmlDocument.WorkbookPart.Workbook.AppendChild(sheets);
            }

            var sheet = new DocumentFormat.OpenXml.Spreadsheet.Sheet
            {
                Id = Owner.OpenXmlDocument.WorkbookPart.GetIdOfPart(wsp),
                SheetId = (uint)Id,
                Name = Name
            };

            sheets.Append(sheet);
        }

        internal Cell GetCell(BaseAZ column, int line, bool createIfNull)
        {
            throw new NotImplementedException();
        } 

        #endregion
    }
}
