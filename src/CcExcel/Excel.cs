using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CcExcel
{
    public class Excel : IDisposable
    {
        public Excel(string fileName, bool canWrite = true)
        {
        }

        public Excel(Stream stream, bool canWrite = true)
        {
        }

        public Sheet this[string sheetName]
        {
            get { throw new NotImplementedException(); }
        }

        public Sheet this[int sheetIndex]
        {
            get { throw new NotImplementedException(); }
        }

        public Sheet CreateSheet(string sheetName)
        {
            throw new NotImplementedException();
        }

        public void Save()
        {
        }

        public void Dispose()
        {
            throw new NotImplementedException();
        }
    }
}
