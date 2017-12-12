using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CcExcel
{
    [Serializable]
    public class ExcelBadFormatException : Exception
    {
        public ExcelBadFormatException() { }
        public ExcelBadFormatException(string message) : base(message) { }
        public ExcelBadFormatException(string message, Exception inner) : base(message, inner) { }
        protected ExcelBadFormatException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }


    [Serializable]
    public class EmptyValueException : Exception
    {
        public EmptyValueException() { }
        public EmptyValueException(string message) : base(message) { }
        public EmptyValueException(string message, Exception inner) : base(message, inner) { }
        protected EmptyValueException(
          System.Runtime.Serialization.SerializationInfo info,
          System.Runtime.Serialization.StreamingContext context) : base(info, context) { }
    }
}
