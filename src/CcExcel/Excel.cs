using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OpenXmlSheet = DocumentFormat.OpenXml.Spreadsheet.Sheet;

namespace CcExcel
{
    public class Excel : IDisposable
    {
        #region Fields

        private readonly Stream _stream;
        private List<Sheet> _sheets;

        #endregion

        #region Constructors

        public Excel(string fileName, ExcelMode mode)
        {
            var fileMode = mode == ExcelMode.OpenReadOnly ? FileMode.Open : FileMode.OpenOrCreate;
            var fileAccess = mode == ExcelMode.OpenReadOnly ? FileAccess.Read : FileAccess.ReadWrite;

            _stream = new FileStream(fileName, fileMode, fileAccess);

            CanWrite = mode != ExcelMode.OpenReadOnly;

            OpenXmlDocument = LoadDocument(mode);
            _sheets = LoadSheets();
        }

        public Excel(Stream stream, ExcelMode mode)
        {
            CanWrite = mode != ExcelMode.OpenReadOnly;

            OpenXmlDocument = LoadDocument(mode);
            _sheets = LoadSheets();
        }

        private SpreadsheetDocument LoadDocument(ExcelMode mode)
        {
            var doc = mode == ExcelMode.Create
                ? SpreadsheetDocument.Create(_stream, SpreadsheetDocumentType.Workbook, true)
                : SpreadsheetDocument.Open(_stream, CanWrite);

            if (mode != ExcelMode.Create) return doc;

            doc.AddWorkbookPart();
            doc.WorkbookPart.Workbook = new Workbook();

            return doc;
        }

        private List<Sheet> LoadSheets()
        {
            var sheets = OpenXmlDocument
                .WorkbookPart
                .Workbook
                .GetFirstChild<Sheets>();

            if (sheets == null) return new List<Sheet>();

            return sheets
                .Elements<OpenXmlSheet>()
                .Select(s => new Sheet(this, s))
                .ToList();
        }

        #endregion

        #region Properties

        internal SpreadsheetDocument OpenXmlDocument { get; }

        public bool CanWrite { get; }

        public Sheet this[string sheetName]
        {
            get
            {
                var sheet = _sheets.FirstOrDefault(w => w.Name == sheetName);

                if (sheet != null) return sheet;

                var id = _sheets.Count > 1 ? _sheets.Max(m => m.Id) + 1 : 1;

                sheet = new Sheet(this, id);

                _sheets.Add(sheet);

                return sheet;
            }
        }

        public Sheet this[int sheetId]
        {
            get
            {
                var sheet = _sheets.FirstOrDefault(w => w.Id == sheetId);

                sheet = new Sheet(this, sheetId);

                _sheets.Add(sheet);

                return sheet;
            }
        }

        #endregion

        #region Methods

        public void Save()
        {
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }

        internal SharedStringTablePart GetSharedStringTable(bool createIfNull)
        {
            throw new NotImplementedException();
        }

        #endregion
    }
}
