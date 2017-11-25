using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CcExcel
{
    public class CellValue
    {
        public static implicit operator CellValue(string value)
        {
            throw new NotImplementedException();
        }

        public static implicit operator string(CellValue value)
        {
            throw new NotImplementedException();
        }
    }
}
