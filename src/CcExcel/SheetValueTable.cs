using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CcExcel
{
    public class SheetValueTable
    {
        private Sheet Owner;

        public CellValue this[BaseAZ column, int line]
        {
            get
            {
                //var cell = Owner.GetCell(column, line, false);
                //
                //if (cell == null) return null;
                //
                //if (cell.DataType?.Value == CellValues.SharedString)
                //{
                //    return Owner.Owner
                //        .GetSharedStringTable(false)
                //        ?.SharedStringTable
                //        ?.ElementAt(int.Parse(cell.InnerText))
                //        ?.InnerText;
                //}
                //
                //return cell?.InnerText;
                throw new NotImplementedException();
            }
            set
            {
                //Owner.CreateSheet();
                //
                //var cell = Owner.GetCell(column, line, true);

                //cell.

                throw new NotImplementedException();
            }
        }

        public CellValue this[string column, int line]
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }

        public CellValue this[int column, int line]
        {
            get { throw new NotImplementedException(); }
            set { throw new NotImplementedException(); }
        }
    }
}
