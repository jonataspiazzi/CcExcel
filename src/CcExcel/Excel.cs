using CcExcel.Helpers;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;

namespace CcExcel
{
    public class Excel : IDisposable
    {
        #region Non Public

        private readonly Stream _stream;
        private readonly bool _streamOwner;
        private List<Sheet> _sheets = new List<Sheet>();

        public SpreadsheetDocument OpenXmlDocument { get; }

        #endregion

        #region Constructors

        public Excel(string fileName, ExcelMode mode)
        {
            var fileMode = mode == ExcelMode.OpenReadOnly ? FileMode.Open : FileMode.OpenOrCreate;
            var fileAccess = mode == ExcelMode.OpenReadOnly ? FileAccess.Read : FileAccess.ReadWrite;

            _stream = new FileStream(fileName, fileMode, fileAccess);
            _streamOwner = true;

            IsEditable = mode != ExcelMode.OpenReadOnly;

            OpenXmlDocument = LoadDocument(mode);
        }

        public Excel(Stream stream, ExcelMode mode)
        {
            IsEditable = mode != ExcelMode.OpenReadOnly;

            _stream = stream;
            _streamOwner = false;

            OpenXmlDocument = LoadDocument(mode);
        }

        private SpreadsheetDocument LoadDocument(ExcelMode mode)
        {
            var doc = mode == ExcelMode.Create
                ? SpreadsheetDocument.Create(_stream, SpreadsheetDocumentType.Workbook, true)
                : SpreadsheetDocument.Open(_stream, IsEditable);

            if (mode == ExcelMode.Create)
            {
                SpreadsheetHelper.GetWorkbook(doc, createIfDoesntExists: true);
            }

            return doc;
        }

        #endregion

        #region Public

        public bool IsEditable { get; }

        public Sheet this[string sheetName]
        {
            get
            {
                var sheet = _sheets.FirstOrDefault(f => f.Name == sheetName);

                if (sheet != null) return sheet;

                var openXmlSheet = SpreadsheetHelper.GetSheet(OpenXmlDocument, sheetName, null, createIfDoesntExists: false);
                var openXmlSheetData = SpreadsheetHelper.GetSheetData(OpenXmlDocument, sheet: openXmlSheet);

                if (openXmlSheetData != null)
                {
                    sheet = new Sheet(this, openXmlSheet, openXmlSheetData);
                }
                else
                {
                    var maxInFile = SpreadsheetHelper.GetMaxId(OpenXmlDocument);
                    var maxInMemory = _sheets.Any() ? _sheets.Max(m => m.Id) : 0;

                    var max = maxInFile > maxInMemory ? maxInFile : maxInMemory;

                    sheet = new Sheet(this, max + 1)
                    {
                        Name = sheetName
                    };
                }

                _sheets.Add(sheet);
                return sheet;
            }
        }

        public Sheet this[int sheetId]
        {
            get
            {
                var sheet = _sheets.FirstOrDefault(f => f.Id == sheetId);

                var openXmlSheet = SpreadsheetHelper.GetSheet(OpenXmlDocument, null, sheetId, createIfDoesntExists: false);
                var openXmlSheetData = SpreadsheetHelper.GetSheetData(OpenXmlDocument, sheet: openXmlSheet);

                sheet = openXmlSheetData != null
                    ? new Sheet(this, openXmlSheet, openXmlSheetData)
                    : new Sheet(this, sheetId);

                _sheets.Add(sheet);
                return sheet;
            }
        }

        #endregion

        #region Methods

        public void Save()
        {
            OpenXmlDocument.Save();
        }

        public void Dispose()
        {
            OpenXmlDocument.Dispose();
            if (_streamOwner) _stream.Dispose();
        }

        #endregion
    }
}
