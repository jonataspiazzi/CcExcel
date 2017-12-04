using CcExcel.Messages;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CcExcel.Helpers
{
    #if TESTABLE
    public
    #else
    internal
    #endif
        class CellReference
    {
        #region Fields

        private readonly BaseAZ _column;
        private readonly uint _line;

        #endregion

        #region Constructors

        public CellReference(BaseAZ column, uint line)
        {
            _column = column;
            _line = line;
        }

        public static CellReference Parse(string cellReference)
        {
            var match = Regex.Match(cellReference, @"^(?<column>[A-Za-z]+)(?<line>[1-9]\d*)$");

            if (!match.Success)
            {
                throw new ArgumentException(Texts.TheParameterCellReferenceWasNotInACorrectFormat, nameof(cellReference));
            }

            var column = BaseAZ.Parse(match.Groups["column"].Value);
            var line = uint.Parse(match.Groups["line"].Value);

            return new CellReference(column, line);
        }

        #endregion

        #region Properties

        public BaseAZ Column => _column;

        public uint Line => _line;

        #endregion

        #region Methods

        public override string ToString()
        {
            return _column.ToString() + _line;
        }

        #endregion
    }
}
