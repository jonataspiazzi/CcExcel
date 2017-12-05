using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CcExcel
{
    public class SheetStyleTable
    {
        internal SheetStyleTable(Sheet owner)
        {
        }

        public StyleId this[BaseAZ column, int line]
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public StyleId this[string column, int line]
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public StyleId this[int column, int line]
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }
    }
}
